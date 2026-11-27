# /// script
# requires-python = ">=3.12"
# dependencies = ["openpyxl>=3.1,<4", "python-pptx>=1.0.2", "lxml>=5.3", "pdfplumber>=0.11"]
# ///
"""Tables stay sound when the user edits the delivered deck by hand in PowerPoint.

Two parts: what every builder writes into the file so an inserted row or an
emptied cell keeps the table's typography, and the reflow tool that restores
row heights, numbering and table stacking after such an edit.
"""
import contextlib
import io
import json
import shutil
import sys
import tempfile
import unittest
from collections import Counter
from copy import deepcopy
from pathlib import Path

from pptx import Presentation
from pptx.oxml.ns import qn

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / 'skills'))
from report_common import notices, reflow
from report_common.header_heights import TABLE_GAP
from report_common.pptx_helpers import content_bottom, tables
from report_common.table_typography import (ROW_ID, TABLE_FONT, all_tables, set_table_directions,
                                            set_table_sizes, unique_row_ids)

MASTER = ROOT / 'skills/weekly-consolidated-report/assets/weekly-report-master.pptx'


def body_copy(table, count):
    """Grow `table` the way the builders do: by copying one body row."""
    pattern = table._tbl.tr_lst[2]
    for _ in range(count):
        table._tbl.append(deepcopy(pattern))


class EditableTableTests(unittest.TestCase):
    def test_every_paragraph_carries_end_properties_with_the_table_font(self):
        prs = Presentation(MASTER)
        set_table_sizes(prs)
        for table in all_tables(prs):
            for paragraph in table._tbl.iter(qn('a:p')):
                end = paragraph.find(qn('a:endParaRPr'))
                self.assertIsNotNone(end)
                self.assertIs(paragraph[-1], end, 'end properties close the paragraph')
                self.assertEqual(end.get('sz'), '1100')
                self.assertEqual([end.find(qn('a:' + slot)).get('typeface') for slot in ('latin', 'ea', 'cs')],
                                 [TABLE_FONT] * 3)

    def test_copied_rows_get_their_own_row_ids(self):
        prs = Presentation(MASTER)
        table = next(all_tables(prs))
        body_copy(table, 5)
        ids = [node.get('val') for node in table._tbl.iter(ROW_ID)]
        self.assertGreater(max(Counter(ids).values()), 1, 'copies share the pattern row ID')
        kept = ids[:3]
        set_table_directions(prs)
        ids = [node.get('val') for node in table._tbl.iter(ROW_ID)]
        self.assertEqual(len(ids), len(set(ids)))
        self.assertEqual(ids[:3], kept, 'rows that already had their own ID keep it')
        before = list(ids)
        unique_row_ids(table)
        self.assertEqual([node.get('val') for node in table._tbl.iter(ROW_ID)], before)


@unittest.skipUnless(shutil.which('soffice'), 'LibreOffice is required to measure rows')
class HeaderShapeTests(unittest.TestCase):
    @staticmethod
    def table(rows, merge_title):
        from pptx.util import Emu
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        shape = slide.shapes.add_table(len(rows), len(rows[0]), Emu(0), Emu(0), Emu(3000000), Emu(1000000))
        table = shape.table
        for r, row in enumerate(rows):
            for c, text in enumerate(row):
                table.cell(r, c).text = text
        if merge_title:
            table.cell(0, 0).merge(table.cell(0, len(rows[0]) - 1))
        return table

    def test_two_header_rows_with_a_hash_column(self):
        table = self.table([['العنوان', '', ''], ['#', 'البند', 'الحالة'], ['1', 'أ', 'x'], ['5', 'ب', 'y']], True)
        self.assertEqual(reflow.header_rows(table), 2)
        self.assertEqual(reflow.number_column(table), 0)

    def test_one_header_row_with_the_arabic_number_header(self):
        table = self.table([['م', 'البند', 'الحالة'], ['01', 'افتتاحية', 'x'], ['7', 'ب', 'y']], False)
        self.assertEqual(reflow.header_rows(table), 1)
        self.assertEqual(reflow.number_column(table), 0)

    def test_renumber_counts_every_data_row_of_a_one_header_table(self):
        table = self.table([['م', 'البند'], ['01', 'افتتاحية'], ['7', 'ب'], ['', 'ج']], False)
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        shape = slide.shapes.add_table(4, 2, 0, 0, 3000000, 1000000)
        for r in range(4):
            for c in range(2):
                shape.table.cell(r, c).text = table.cell(r, c).text
        reflow.renumber(prs)
        self.assertEqual([shape.table.cell(r, 0).text for r in range(4)], ['م', '1', '2', '3'])


class ReflowTests(unittest.TestCase):
    def test_reflow_restores_heights_numbering_and_stacking(self):
        with tempfile.TemporaryDirectory() as directory:
            edited, fixed = Path(directory) / 'edited.pptx', Path(directory) / 'fixed.pptx'
            prs = Presentation(MASTER)
            set_table_sizes(prs); set_table_directions(prs)
            slide = prs.slides[1]
            stack = sorted(tables(slide), key=lambda sh: sh.top)
            first = stack[0].table
            # The user typed a long note and inserted a row; PowerPoint left the heights and the tables below alone.
            long_cell = first.cell(2, 1)
            long_cell.text_frame.paragraphs[0].runs[0].text = 'نص طويل أضافه المستخدم يدويا في باوربوينت ' * 12
            body_copy(first, 1)
            for row in first.rows:
                row.height = 100000
            prs.save(edited)
            before = edited.read_bytes()
            notices.reset()
            with contextlib.redirect_stderr(io.StringIO()):
                reflow.reflow(edited, fixed)
            self.assertEqual(edited.read_bytes(), before, 'the edited file is not changed')
            out = Presentation(fixed)
            slide = out.slides[1]
            stack = sorted(tables(slide), key=lambda sh: sh.top)
            table = stack[0].table
            numbers = [row.cells[[c.text.strip() for c in table.rows[1].cells].index('#')].text
                       for row in list(table.rows)[2:]]
            self.assertEqual(numbers, [str(i) for i in range(1, len(numbers) + 1)])
            self.assertGreater(table.rows[2].height, 400000, 'the long row takes the height its text needs')
            for shape in stack:
                self.assertEqual(shape.height, sum(row.height for row in shape.table.rows))
            for above, below in zip(stack, stack[1:]):
                self.assertEqual(below.top, above.top + above.height + TABLE_GAP)
            ids = [node.get('val') for node in table._tbl.iter(ROW_ID)]
            self.assertEqual(len(ids), len(set(ids)))
            # The slide now overflows or not; either way the tool says so only when it does.
            overflow = stack[-1].top + stack[-1].height > content_bottom(out, slide)
            found = json.loads(notices.notices_path(fixed).read_text(encoding='utf-8')) if notices.notices_path(fixed).exists() else []
            self.assertEqual(overflow, any(n['kind'] == 'space' and 'الشريحة 2' in n['message'] for n in found))



class SharedHelpersTests(unittest.TestCase):
    def test_style_key_and_neutral_fill(self):
        from report_common import status_styles
        styles = {'معلق': 1, 'متأخر': 2}
        self.assertEqual(status_styles.style_key('متأخر', {'متأخر', 'معلق'}, styles, 'معلق'), 'متأخر')
        self.assertEqual(status_styles.style_key('غريب', {'متأخر', 'معلق'}, styles, 'معلق'), 'معلق')
        self.assertEqual(status_styles.style_key('مكتملة', {'مكتملة'}, styles, 'معلق'), 'معلق')
        self.assertEqual(status_styles.NEUTRAL_FILL, status_styles.FILLS['لم تبدأ'])

    def test_restack_keeps_side_by_side_tables_apart(self):
        prs = Presentation(MASTER)
        slide = next(sl for sl in prs.slides if len(tables(sl)) >= 1)
        shapes = tables(slide)
        first = shapes[0]
        # Two copies: one below the first (same column), one beside it.
        below = deepcopy(first._element)
        beside = deepcopy(first._element)
        slide.shapes._spTree.append(below)
        slide.shapes._spTree.append(beside)
        stack = tables(slide)
        below_shape, beside_shape = stack[-2], stack[-1]
        for index, shape in enumerate((below_shape, beside_shape), start=9001):
            shape._element.nvGraphicFramePr.cNvPr.set('id', str(index))
        first.top = 1000000
        below_shape.top = 1000100  # overlaps the first
        beside_shape.left = first.left + first.width + 10
        beside_shape.top = 1000000
        gap = 120000
        fits = reflow.restack(slide, 0, 10 ** 9, gap, tighten=False)
        self.assertTrue(fits)
        self.assertEqual(below_shape.top, first.top + first.height + gap)
        self.assertEqual(beside_shape.top, 1000000)
        self.assertFalse(reflow.restack(slide, 0, 1000, gap))




if __name__ == '__main__':
    unittest.main()
