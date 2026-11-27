# /// script
# requires-python = ">=3.12"
# dependencies = ["python-pptx>=1.0.2", "lxml>=5.3", "pdfplumber>=0.11", "openpyxl>=3.1"]
# ///
"""Put a report's tables back in order after the user edited them by hand.

PowerPoint does not move the table under one that grew, and a row the user
inserts copies the height of the row beside it. After a manual edit (a new
row, longer text, a deleted record) run this on the edited file: it writes a
new file in which every table keeps its own text size and the approved
typography, the «#» column counts again from the first slide of each table,
every row has the height its text needs (measured through LibreOffice), and
the tables of each slide are stacked again with the approved gap. The edited
file is not changed.

A slide whose tables no longer fit above the footer is not repaginated: it is
named in the notices, because moving records to another slide is a content
decision (rebuild from the Excel files, or move the rows in PowerPoint and run
this again).
"""
import argparse
import re
import subprocess
import sys
from collections import Counter
from pathlib import Path

from pptx import Presentation

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from report_common import notices
from report_common.header_heights import TABLE_GAP
from report_common.measure import measure_rows
from report_common.pptx_helpers import content_bottom, set_cell_text, table_title, tables
from report_common.render import abar_installed
from report_common.table_typography import TABLE_PT, apply_table, rtl_table, unique_row_ids

NUMBER = re.compile(r'\d+')
HEADER_ROWS = 2  # title strip and column headers (the default; see header_rows)
NUMBER_HEADERS = ('#', 'م')


def header_rows(table):
    """How many rows at the top of `table` are headers: 2 when row 1 is a title strip
    spanning every column (merged, or one filled cell with the rest empty), else 1."""
    rows = list(table.rows)
    if len(rows) < 2:
        return 1
    cells = list(rows[0].cells)
    first = cells[0]
    if len(cells) > 1 and first.is_merge_origin and first.span_width >= len(cells):
        return 2
    if len(cells) > 1 and sum(1 for c in cells if c.text.strip()) == 1:
        return 2
    return 1


def table_size(table):
    """The text size `table` uses most, so a summary that was shrunk keeps its size."""
    sizes = Counter(run.font.size.pt for row in table.rows for cell in row.cells
                    for p in cell.text_frame.paragraphs for run in p.runs
                    if run.text.strip() and run.font.size is not None)
    return sizes.most_common(1)[0][0] if sizes else TABLE_PT


def number_column(table):
    """Index of the «#» or «م» column when its body holds only running numbers, else None."""
    top = header_rows(table)
    if len(table.rows) <= top:
        return None
    headers = [cell.text.strip() for cell in table.rows[top - 1].cells]
    column = next((headers.index(h) for h in NUMBER_HEADERS if h in headers), None)
    if column is None:
        return None
    body = [row.cells[column].text.strip() for row in list(table.rows)[top:]]
    return column if all(not text or NUMBER.fullmatch(text) for text in body) else None


def renumber(prs):
    """Count the «#» column again; a table that continues on later slides keeps counting."""
    last = {}
    for slide in prs.slides:
        for shape in sorted(tables(slide), key=lambda sh: sh.top):
            column = number_column(shape.table)
            if column is None:
                continue
            title = table_title(shape)
            rows = list(shape.table.rows)[header_rows(shape.table):]
            first = rows[0].cells[column].text.strip() if rows else ''
            # A table whose first number is 1 starts a new count even under a repeated title.
            count = 0 if first == '1' else last.get(title, 0)
            for row in rows:
                count += 1
                if row.cells[column].text.strip() != str(count):
                    set_cell_text(row.cells[column], str(count))
            last[title] = count


def restack(slide, top, bottom, gap, tighten=True):
    """Give each table of `slide` the height of its rows and stack the tables again.

    A table follows the tables above it in the same column (tables that overlap
    horizontally); side-by-side tables stay apart and do not push each other.
    No table starts above `top`. With ``tighten`` (default) a table sits exactly
    `gap` below the table above it; with ``tighten=False`` it moves only when it
    would overlap that table. Row heights and text are untouched. Returns True
    when every table ends at or above `bottom`.
    """
    stack = sorted(tables(slide), key=lambda sh: sh.top)
    placed = []
    for shape in stack:
        shape.height = sum(row.height for row in shape.table.rows)
        above = [p for p in placed if p.left < shape.left + shape.width and shape.left < p.left + p.width]
        if above:
            floor = max(p.top + p.height for p in above)
            shape.top = floor + gap if tighten or shape.top < floor else shape.top
        elif shape.top < top:
            shape.top = top
        placed.append(shape)
    return all(sh.top + sh.height <= bottom for sh in stack)


def reflow(report, output):
    report, output = Path(report), Path(output)
    if output.resolve() == report.resolve():
        raise ValueError('Write the repaired report to a new file')
    notices.clear_stale(output)
    prs = Presentation(report)
    renumber(prs)
    targets = []
    for index, slide in enumerate(prs.slides):
        for shape in tables(slide):
            apply_table(shape.table, table_size(shape.table))
            rtl_table(shape.table)
            unique_row_ids(shape.table)
            top = header_rows(shape.table)
            if len(shape.table.rows) > top:
                targets.append((index, shape.shape_id, range(top, len(shape.table.rows))))
    output.parent.mkdir(parents=True, exist_ok=True)
    prs.save(output)
    try:
        if not abar_installed():
            notices.warn('خط «Abar Mid» غير مثبت على هذا الجهاز؛ القياس تم بخط بديل وقد تختلف ارتفاعات الصفوف في باوربوينت.', 'layout')
        measured = measure_rows(output, targets)
    except (RuntimeError, OSError, subprocess.SubprocessError) as exc:
        notices.warn(f'تعذر قياس الصفوف عبر LibreOffice ({exc})؛ أبقيت ارتفاعات الصفوف كما هي وأعدت ترتيب الجداول فقط.', 'layout')
        measured = {}
    for index, slide in enumerate(prs.slides):
        for shape in tables(slide):
            for row_index, row in enumerate(shape.table.rows):
                height = measured.get((index, shape.shape_id, row_index))
                if height:
                    row.height = height
        stack = tables(slide)
        if stack:
            try:
                limit = content_bottom(prs, slide)
            except ValueError:
                limit = None
            fits = restack(slide, min(sh.top for sh in stack), limit if limit is not None else 2 ** 62, TABLE_GAP)
            if limit is not None and not fits:
                notices.warn(f'الشريحة {index + 1}: الجداول تتجاوز التذييل بعد التعديل. انقل الصفوف الزائدة إلى شريحة تالية ثم أعد تشغيل الأداة، أو أعد بناء التقرير من ملفات Excel.', 'space')
    prs.save(output)
    notices.save(output)
    return output


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument('--report', required=True, help='The report as edited by hand in PowerPoint')
    parser.add_argument('--output', required=True, help='The repaired report, a new file')
    args = parser.parse_args()
    notices.reset()
    print(reflow(args.report, args.output))
