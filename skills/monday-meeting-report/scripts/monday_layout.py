"""Content-sized rows and continuous table flow for the Monday master."""
import copy
import json
import math
import os
from pathlib import Path
import subprocess
import tempfile

from pptx import Presentation
from pptx.oxml.ns import qn
from pptx.util import Pt
from report_common.pptx_helpers import fill_titled_table, remove_slide, set_cell_text
from report_common.update_format import format_table_updates
from report_common.table_typography import apply_table


def rtl_table(table):
    """Mirror LTR storage while preserving the visible column order and merges."""
    if table._tbl.tblPr.get('rtl') not in ('1', 'true'):
        grid = table._tbl.tblGrid
        for col in reversed(list(grid)):
            grid.append(col)
        for row in table.rows:
            groups = []
            cells = list(row._tr.tc_lst)
            i = 0
            while i < len(cells):
                span = int(cells[i].get('gridSpan', '1'))
                groups.append(cells[i:i + span])
                i += span
            for group in reversed(groups):
                for cell in group:
                    row._tr.insert_element_before(cell, 'a:extLst')
                    props = cell.find(qn('a:tcPr'))
                    if props is not None:
                        left, right = props.find(qn('a:lnL')), props.find(qn('a:lnR'))
                        for node in (left, right):
                            if node is not None: props.remove(node)
                        if left is not None:
                            left.tag = qn('a:lnR'); props.insert(0,left)
                        if right is not None:
                            right.tag = qn('a:lnL'); props.insert(0,right)
    table._tbl.tblPr.set('rtl', '1')
    for row in table.rows:
        for cell in row.cells:
            for p in cell.text_frame.paragraphs:
                p._p.get_or_add_pPr().set('rtl', '1')


def bullet_example(prs):
    for slide in prs.slides:
        for shape in slide.shapes:
            if not shape.has_table or len(shape.table.columns) != 6:
                continue
            if shape.table.cell(1,1).text.strip() != 'اسم المشروع':
                continue
            for row in list(shape.table.rows)[2:]:
                for p in row.cells[5].text_frame.paragraphs:
                    props = p._p.get_or_add_pPr()
                    if props.find(qn('a:buChar')) is not None:
                        return copy.deepcopy(props)
    raise ValueError('Suhail master has no approved bullet paragraph')


def populate(shape, rows, title=None, kind=None, styles=None, bullets=None):
    rtl_table(shape.table)
    fill_titled_table(shape, rows, title=title)
    for i, values in enumerate(rows, 2):
        for j, value in enumerate(values):
            cell = shape.table.cell(i, j)
            # Remove surplus template paragraphs, which otherwise create empty lines.
            wanted = max(1, len(str(value if value is not None else '').splitlines()))
            while len(cell.text_frame.paragraphs) > wanted:
                p = cell.text_frame.paragraphs[-1]._p
                p.getparent().remove(p)
        if kind:
            import monday_status_styles as status_styles
            cell = shape.table.cell(i, 4)
            status = status_styles.canonical(cell.text)
            allowed = status_styles.TASK_STATUSES if kind == 'task' else status_styles.PROJECT_STATUSES
            if status not in allowed or status not in styles:
                raise ValueError('Missing or unsupported approved status: ' + status)
            text = cell.text
            status_styles.apply(cell, styles[status])
            set_cell_text(cell, text)
    format_table_updates(shape.table, bullets, paragraph_items=True)
    for row in shape.table.rows:
        for cell in row.cells:
            for p in cell.text_frame.paragraphs:
                p._p.get_or_add_pPr().set('rtl','1')
                for run in p.runs:
                    if run.font.name not in ('Abar Mid', 'Abar Mid SemiBold'):
                        run.font.name = 'Abar Mid'
                    run.font.size = Pt(11)

    apply_table(shape.table)


def measure(specs, template, styles, bullets, engine='auto'):
    """Measure fresh disposable tables; no reusable/stale measurement cache."""
    native = engine == 'native' or (engine == 'auto' and os.name == 'nt')
    probe = Presentation(template)
    for slide in list(probe.slides):
        remove_slide(probe, slide)
    estimates = {}
    for key, shape, rows, title, kind in specs:
        if not rows:
            continue
        slide = probe.slides.add_slide(probe.slide_layouts[0])
        for sh in list(slide.shapes):
            sh.element.getparent().remove(sh.element)
        slide.shapes._spTree.insert_element_before(copy.deepcopy(shape.element), 'p:extLst')
        sh = slide.shapes[-1]
        sh.name = key
        populate(sh, rows, title, kind, styles, bullets)
        heights = []
        for row in list(sh.table.rows)[2:]:
            height = 20
            for j, cell in enumerate(row.cells):
                width = max(20, sh.table.columns[j].width / 12700 - (cell.margin_left + cell.margin_right) / 12700)
                lines = sum(max(1, math.ceil(len(line) * 4.8 / width)) for line in cell.text.splitlines())
                height = max(height, lines * 14 + (cell.margin_top + cell.margin_bottom) / 12700 + 2)
            heights.append(int(Pt(height)))
            row.height = Pt(1) if native else Pt(height)
        sh.height = sum(r.height for r in sh.table.rows)
        estimates[key] = ([r.height for r in list(sh.table.rows)[:2]], heights)
    if not native or not estimates:
        return estimates
    with tempfile.TemporaryDirectory(prefix='monday-layout-') as temporary:
        path = Path(temporary) / 'probe.pptx'
        probe.save(path)
        result = subprocess.run(['powershell.exe','-NoProfile','-File',str(Path(__file__).with_name('measure_rows.ps1')),str(path)], capture_output=True, text=True, timeout=60)
        if result.returncode:
            raise RuntimeError('Native PowerPoint row measurement failed: ' + result.stdout + result.stderr)
        actual = json.loads(path.with_suffix('.json').read_text(encoding='utf-8-sig'))
    return {key: ([max(int(Pt(v)), estimates[key][0][i]) for i,v in enumerate(values[:2])],
                  [max(int(Pt(v + 2)), int(Pt(20))) for v in values[2:]]) for key, values in actual.items()}


def pack(sections, start, bottom, gap=int(Pt(9)), continuation=None):
    """Pack whole records, counting a repeated title/header only once per table."""
    continuation = start if continuation is None else continuation
    pages = [[]]
    used = start
    for key, rows, header, heights in sections:
        if len(rows) != len(heights):
            raise ValueError('Row measurements do not match current records')
        chunk, sizes = [], []
        fixed = sum(header)
        for row, height in zip(rows, heights):
            if fixed + height > bottom - continuation:
                raise ValueError('A single record is too tall for an approved page: ' + key)
            needed = height if chunk else fixed + height + (gap if pages[-1] else 0)
            if used + needed > bottom:
                if chunk:
                    pages[-1].append((key, chunk, header, sizes, top))
                    chunk, sizes = [], []
                pages.append([])
                used = continuation
            if not chunk:
                top = used + (gap if pages[-1] else 0)
                used = top + fixed
            chunk.append(row)
            sizes.append(height)
            used += height
        if chunk:
            pages[-1].append((key, chunk, header, sizes, top))
    return pages


def size_shape(shape, header, heights, top):
    shape.top = top
    for row, height in zip(shape.table.rows, [*header, *heights]):
        row.height = height
    shape.height = sum(r.height for r in shape.table.rows)
