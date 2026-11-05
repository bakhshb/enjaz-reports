# /// script
# requires-python = ">=3.12"
# dependencies = ["python-pptx>=1.0.2", "lxml>=5.3", "pdfplumber>=0.11"]
# ///
"""Measure a weekly candidate's record rows through LibreOffice.

Cross-platform counterpart of measure_layout.ps1 (PowerPoint/Windows). Writes
<report>.layout.json in the same format, bound to the exact input hashes, for
`build_report.py --measurements`. The candidate is not modified. Digital-maturity
slides are skipped: they paginate with their own model. Install the template
fonts (Abar Mid) first; LibreOffice substitutes wider fonts otherwise.
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
from weekly_data import MASTER

DIGITAL_MASTER = Path(__file__).resolve().parents[1] / 'assets/weekly-report-digital-master.pptx'
TYPOGRAPHY = '11pt-suhail-paragraph-updates-v3'
PT = 12700


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
    order, targets = [], []
    for si, slide in enumerate(prs.slides):
        # Digital summary and detail slides (incl. their own challenge tables,
        # whose titles repeat the Suhail ones) follow the digital model.
        if sum(sh.has_chart for sh in slide.shapes) >= 5 or is_digital_detail(slide):
            continue
        for sh in slide.shapes:
            if not sh.has_table or len(sh.table.rows) < 3:
                continue
            table = sh.table
            if any(run.font.size is not None and run.font.size.pt != 11
                   for row in table.rows for cell in row.cells
                   for p in cell.text_frame.paragraphs for run in p.runs if run.text.strip()):
                raise ValueError('Generate an 11 pt candidate before measuring')
            title = table.cell(0, 0).text.strip()
            rows = range(2, len(table.rows))
            order.append((title, si, sh.shape_id, rows, [table.rows[0].height, table.rows[1].height],
                          [r.height for r in list(table.rows)[2:]]))
            targets.append((si, sh.shape_id, rows))
    measured = measure_rows(report, targets)
    tables = {}
    for title, si, shape_id, rows, header, declared in order:
        entry = tables.setdefault(title, {'header': [h / PT for h in header], 'body': []})
        for row, fallback in zip(rows, declared):
            entry['body'].append(measured.get((si, shape_id, row), fallback) / PT)
    if sha256(report) != hashes['report']:
        raise ValueError('Measuring changed the source presentation')
    return {'typography': TYPOGRAPHY, 'hashes': hashes, 'tables': tables,
            'visualAcceptance': 'pending'}


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
