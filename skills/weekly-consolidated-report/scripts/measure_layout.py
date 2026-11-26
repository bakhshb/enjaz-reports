# /// script
# requires-python = ">=3.12"
# dependencies = ["openpyxl>=3.1", "python-pptx>=1.0.2", "lxml>=5.3", "pdfplumber>=0.11"]
# ///
"""Measure a weekly candidate's record rows through LibreOffice.

Cross-platform counterpart of measure_layout.ps1 (PowerPoint/Windows). Writes
<report>.layout.json in the same format, bound to the exact input hashes, for
`build_report.py --measurements`. The candidate is not modified. For the digital
model the digital-maturity rows are measured too and stored under ``digital``:
the summary «أبرز التحديثات» rows, and per sector its record rows in reviewed
order (continuation headings left out) and its update and challenge rows.
Install the template fonts (Abar Mid) first; LibreOffice substitutes wider fonts
otherwise.
"""
import argparse
import hashlib
import json
from pathlib import Path
import sys

from pptx import Presentation

sys.path.insert(0, str(Path(__file__).resolve().parents[2]))
sys.path.insert(0, str(Path(__file__).resolve().parent))
from report_common.measure import measure_rows
from weekly_artwork import is_digital_detail
from weekly_data import MASTER, SUMMARY_SIZES, PT
from report_common.pptx_helpers import TASK_SUMMARY_TITLES, SUHAIL_SUMMARY_TITLES
from report_common.table_typography import apply_table

# Summary tables stay on their one slide; when they do not fit at 11 pt the
# builder picks a smaller size, so their rows are measured at every size.
SUMMARY_TITLES = set(TASK_SUMMARY_TITLES) | set(SUHAIL_SUMMARY_TITLES) | {'المعاملات المتأخرة'}

DIGITAL_MASTER = Path(__file__).resolve().parents[1] / 'assets/weekly-report-digital-master.pptx'
TYPOGRAPHY = '11pt-suhail-paragraph-updates-v3'


def sha256(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest().upper()


def measure(report, tasks, suhail, transactions, template=None, digital_pdf=None, digital_data=None):
    if bool(digital_pdf) != bool(digital_data):
        raise ValueError('Provide both digital PDF and reviewed data when measuring the digital model')
    template = template or (DIGITAL_MASTER if digital_pdf else MASTER)
    hashes = {'report': sha256(report), 'tasks': sha256(tasks), 'suhail': sha256(suhail),
              'transactions': sha256(transactions), 'template': sha256(template)}
    if digital_pdf:
        hashes.update(digital_pdf=sha256(digital_pdf), digital_data=sha256(digital_data))
    prs = Presentation(report)
    order, targets, digital = [], [], []
    for si, slide in enumerate(prs.slides):
        # Digital summary and detail slides (incl. their own challenge tables,
        # whose titles repeat the Suhail ones) are kept apart under `digital`.
        is_digital = sum(sh.has_chart for sh in slide.shapes) >= 5 or is_digital_detail(slide)
        for sh in slide.shapes:
            if not sh.has_table or len(sh.table.rows) < 3:
                continue
            table = sh.table
            if any(run.font.size is not None and run.font.size.pt != 11
                   for row in table.rows for cell in row.cells
                   for p in cell.text_frame.paragraphs for run in p.runs if run.text.strip()):
                raise ValueError('Generate an 11 pt candidate before measuring')
            if is_digital:
                if digital_pdf:
                    digital.append((si, slide, sh))
                    targets.append((si, sh.shape_id, range(2, len(table.rows))))
                continue
            title = table.cell(0, 0).text.strip()
            rows = range(2, len(table.rows))
            order.append((title, si, sh.shape_id, rows, [table.rows[0].height, table.rows[1].height],
                          [r.height for r in list(table.rows)[2:]]))
            targets.append((si, sh.shape_id, rows))
    measured = measure_rows(report, targets)
    summary_targets = [(si, sid, rows) for title, si, sid, rows, *_ in order if title in SUMMARY_TITLES]
    summary_targets += [(si, sh.shape_id, range(2, len(sh.table.rows)))
                        for si, _, sh in digital if len(sh.table.columns) == 3]
    smaller = measure_sizes(report, summary_targets)
    tables = {}
    for title, si, shape_id, rows, header, declared in order:
        entry = tables.setdefault(title, {'header': [h / PT for h in header], 'body': []})
        for row, fallback in zip(rows, declared):
            entry['body'].append(measured.get((si, shape_id, row), fallback) / PT)
            if title in SUMMARY_TITLES:
                for size, found in smaller.items():
                    entry.setdefault('sizes', {}).setdefault(size, []).append(found.get((si, shape_id, row), fallback) / PT)
    if sha256(report) != hashes['report']:
        raise ValueError('Measuring changed the source presentation')
    result = {'typography': TYPOGRAPHY, 'hashes': hashes, 'tables': tables,
              'visualAcceptance': 'pending'}
    if digital_pdf:
        result['digital'] = digital_tables(digital, measured)
        summary = [(si, sh) for si, _, sh in digital if len(sh.table.columns) == 3]
        result['digital']['summary_sizes'] = {
            size: [found.get((si, sh.shape_id, r), sh.table.rows[r].height) / PT
                   for si, sh in summary for r in range(2, len(sh.table.rows))]
            for size, found in smaller.items()}
    return result


def measure_sizes(report, targets):
    """Measure `targets` again with their tables set to each smaller summary size."""
    import tempfile
    result = {}
    if not targets:
        return result
    with tempfile.TemporaryDirectory(prefix='measure-sizes-') as temporary:
        for size in SUMMARY_SIZES[1:]:
            prs = Presentation(report)
            for si, shape_id, _ in targets:
                shape = next(sh for sh in prs.slides[si].shapes if sh.shape_id == shape_id)
                apply_table(shape.table, size)
            copy = Path(temporary) / f'size-{size}.pptx'
            prs.save(copy)
            result[str(size)] = measure_rows(copy, targets)
    return result


def digital_tables(found, measured):
    """Group the measured digital rows by what they hold, in reviewed order."""
    from digital_maturity import split_title, is_continuation, _detail_columns, NOTE_TABLES
    height = lambda si, shape, row: measured.get((si, shape.shape_id, row), shape.table.rows[row].height) / PT
    keys = {title: key for key, title, _ in NOTE_TABLES}
    result = {'summary': [], 'sectors': {}}
    sector = None
    for si, slide, shape in found:
        table = shape.table
        rows = range(2, len(table.rows))
        if len(table.columns) == 3:
            result['summary'] = [height(si, shape, r) for r in rows]
        elif len(table.columns) == 4:
            title = split_title(table.cell(0, 0).text.strip())[0]
            sector = result['sectors'].setdefault(title, {'rows': [], **{k: [] for k in keys.values()}})
            for r in rows:
                row = table.rows[r]
                values = [row.cells[c].text for c in _detail_columns(table, row)]
                if r == 2 and is_continuation(values):
                    continue
                sector['rows'].append(height(si, shape, r))
        elif len(table.columns) == 2 and sector is not None:
            key = keys.get(table.cell(0, 0).text.strip())
            if key is None:
                raise ValueError('Unknown digital-maturity note table: ' + table.cell(0, 0).text)
            sector[key].extend(height(si, shape, r) for r in rows)
    return result


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--report', required=True)
    parser.add_argument('--tasks', required=True)
    parser.add_argument('--suhail', required=True)
    parser.add_argument('--transactions', required=True)
    parser.add_argument('--template')
    parser.add_argument('--digital-pdf')
    parser.add_argument('--digital-data')
    args = parser.parse_args(argv)
    result = measure(args.report, args.tasks, args.suhail, args.transactions, args.template,
                     args.digital_pdf, args.digital_data)
    output = Path(args.report).with_suffix('.layout.json')
    output.write_text(json.dumps(result, ensure_ascii=False, indent=1), encoding='utf-8')
    print(output)


if __name__ == '__main__':
    main()
