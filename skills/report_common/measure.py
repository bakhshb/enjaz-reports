"""Measure the height each table row's text really needs, through LibreOffice.

LibreOffice draws a PowerPoint table row at exactly its declared height and lets
longer text overflow; it never grows the row. Reading row borders back from a
render therefore only echoes the declared heights. Instead, every measured row
is copied into its own one-row table on a probe slide, made taller than any
text it could hold and anchored to the top, and the distance from the row's top
edge to the bottom of its lowest glyph (plus the bottom cell inset) gives the
height the row needs. The source presentation is never modified.

Measuring from the row's top edge counts the top inset and the first line's
ascent. An earlier probe collapsed the row to 1 pt and measured only the glyph
extent, which left out that space and came out 0.05-0.07 in short on every row;
on a table of six rows that let the next table overlap the last one.
"""
from copy import deepcopy
from pathlib import Path
import sys
import tempfile

from pptx import Presentation

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from report_common.render import export_pdf
from report_common.pptx_helpers import add_slide
from report_common.table_typography import DEFAULT_INSET

EMU = 914400
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
PROBE_TOP = int(0.25 * EMU)  # the probe row starts here, anchored to the top
MIN_ROW = int(0.22 * EMU)  # one 11 pt line with insets


def _q(tag):
    return '{%s}%s' % (A, tag)


def _row_text(tr):
    return '\n'.join(''.join(t.text or '' for t in p.iter(_q('t')))
                     for tc in tr.findall(_q('tc')) for p in tc.iter(_q('p'))).strip()


def _bottom_inset(tr):
    """Largest bottom margin among the row's cells."""
    best = 0
    for tc in tr.findall(_q('tc')):
        pr = tc.find(_q('tcPr'))
        best = max(best, int(pr.get('marB', DEFAULT_INSET)) if pr is not None else DEFAULT_INSET)
    return best


def _anchor_top(tr):
    for tc in tr.findall(_q('tc')):
        pr = tc.find(_q('tcPr'))
        if pr is None:
            pr = tc.makeelement(_q('tcPr'), {})
            tc.append(pr)
        pr.set('anchor', 't')


def measure_rows(report, targets, padding=0.06):
    """Return {(slide_index, shape_id, row_index): needed height in EMU}.

    ``targets`` is an iterable of (slide_index, shape_id, row_indices). Rows
    without text are skipped (the caller keeps their declared height).
    ``padding`` (inches) is added to every measured row as a safety margin for
    renderer differences (glyph descent below the font box, PowerPoint vs
    LibreOffice line breaks)."""
    source = Presentation(report)
    probe = Presentation(report)
    layout = probe.slides[0].slide_layout
    jobs = []
    for slide_index, shape_id, rows in targets:
        frame = next(sh for sh in source.slides[slide_index].shapes if sh.shape_id == shape_id)
        trs = frame._element.find('.//' + _q('tbl')).findall(_q('tr'))
        for row_index in rows:
            tr = trs[row_index]
            if not _row_text(tr):
                continue
            slide = add_slide(probe, layout)
            for sh in list(slide.shapes):
                sh._element.getparent().remove(sh._element)
            element = deepcopy(frame._element)
            tbl = element.find('.//' + _q('tbl'))
            for k, other in enumerate(tbl.findall(_q('tr'))):
                if k != row_index:
                    tbl.remove(other)
            kept = tbl.find(_q('tr'))
            tall = probe.slide_height - 2 * PROBE_TOP
            kept.set('h', str(tall))
            _anchor_top(kept)
            # A merged-from-above cell would hide its text once alone.
            for tc in kept.findall(_q('tc')):
                for attr in ('vMerge', 'rowSpan'):
                    tc.attrib.pop(attr, None)
            slide.shapes._spTree.append(element)
            shape = slide.shapes[-1]
            shape.top = PROBE_TOP
            shape.height = tall
            jobs.append(((slide_index, shape_id, row_index), len(probe.slides) - 1,
                         shape.left, shape.left + shape.width, _bottom_inset(kept)))
    if not jobs:
        return {}
    # Original slides stay in the probe deck: layout nav links still reference
    # them, so dropping them would leave parts whose names collide on save.
    import pdfplumber
    result = {}
    with tempfile.TemporaryDirectory(prefix='measure-') as temporary:
        root = Path(temporary)
        probe.save(root / 'probe.pptx')
        pdf_path = export_pdf(root / 'probe.pptx', root / 'probe.pdf')
        with pdfplumber.open(pdf_path) as pdf:
            for key, page_index, left, right, insets in jobs:
                page = pdf.pages[page_index]
                x0, x1 = left / EMU * 72 - 2, right / EMU * 72 + 2
                chars = [c for c in page.chars if c['x1'] >= x0 and c['x0'] <= x1
                         and c['text'].strip()]
                if not chars:
                    continue
                text_bottom = max(c['bottom'] for c in chars) - PROBE_TOP / EMU * 72
                needed = int(text_bottom / 72 * EMU) + insets + int(padding * EMU)
                result[key] = max(MIN_ROW, needed)
    return result
