"""Every report table must carry native RTL direction, including empty tables."""
import sys
import tempfile
import unittest
from pathlib import Path

from pptx import Presentation
from pptx.util import Inches

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / 'skills'))
from report_common import table_typography as typography


class TableRTLTests(unittest.TestCase):
    def test_all_bundled_models_normalize_without_losing_cells(self):
        paths = list((ROOT / 'skills').glob('*/assets/*.pptx'))
        self.assertEqual(len(paths), 3)
        for path in paths:
            with self.subTest(template=path.name):
                prs = Presentation(path)
                tables = list(typography.all_tables(prs))
                before = [[(cell._tc, cell.text) for row in table.rows for cell in row.cells]
                          for table in tables]
                typography.set_table_directions(prs)
                for table, cells in zip(tables, before):
                    self.assertEqual(table._tbl.tblPr.get('rtl'), '1')
                    headers = [cell.text for cell in table.rows[1].cells]
                    number = next((h for h in headers if h in {'#', 'م', 'رقم المعاملة'}), None)
                    if number is not None:
                        self.assertEqual(headers[0], number)
                    self.assertEqual(sorted((id(tc), text) for tc, text in cells), sorted((id(cell._tc), cell.text)
                                     for row in table.rows for cell in row.cells))
                once = [table._tbl.xml for table in tables]
                typography.set_table_directions(prs)
                self.assertEqual(once, [table._tbl.xml for table in tables])
                with tempfile.TemporaryDirectory() as temporary:
                    output = Path(temporary) / 'rtl.pptx'
                    prs.save(output)
                    self.assertTrue(all(t._tbl.tblPr.get('rtl') == '1'
                                        for t in typography.all_tables(Presentation(output))))

    def test_validator_rejects_ltr_table_even_with_right_aligned_arabic(self):
        prs = Presentation()
        table = prs.slides.add_slide(prs.slide_layouts[6]).shapes.add_table(
            3, 2, 0, 0, Inches(4), Inches(2)).table
        table.cell(1, 0).text = 'القطاع'
        table.cell(1, 1).text = 'المهمة'
        typography.set_table_sizes(prs)
        errors = []
        typography.check_table_typography(prs, errors)
        self.assertIn('Table direction must be RTL', errors)


if __name__ == '__main__':
    unittest.main()
