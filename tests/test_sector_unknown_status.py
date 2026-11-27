"""A status the master has no table for is kept in the sector deck with its own text and a neutral fill."""
from pathlib import Path
import contextlib, io, json, shutil, sys, tempfile, unittest
from openpyxl import load_workbook
from pptx import Presentation
import run_weekly_workflow as workflow
from test_upstream_integrity import run

weekly = workflow.weekly
splitter = workflow.splitter
SECTOR = workflow.fixtures.SECTOR
quiet = lambda: (contextlib.redirect_stdout(io.StringIO()), contextlib.redirect_stderr(io.StringIO()))


def first_record_row(ws, header_text):
    header = next(r for r in ws.iter_rows() if any(c.value == header_text for c in r))
    column = next(c.column for c in header if c.value == 'الحالة')
    row = next(r for r in ws.iter_rows(min_row=header[0].row + 1) if isinstance(r[0].value, int) or r[0].value)
    return row, column


class UnknownStatusSectorTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.temp = tempfile.TemporaryDirectory()
        base = Path(cls.temp.name) / 'baseline'
        with contextlib.redirect_stdout(io.StringIO()):
            workflow.scenario(base, 'normal')
        cls.root = Path(cls.temp.name) / 'case'
        shutil.copytree(base, cls.root)
        cls.task_text = cls.mark(cls.root / 'tasks.xlsx', 'تفاصيل المهام', 'المهمة', 'قيد المراجعة', 1)
        cls.proj_text = cls.mark(cls.root / 'projects.xlsx', 'تفاصيل مشاريع سهيل', 'اسم المشروع', 'حالة غير معتمدة', 1)
        with contextlib.redirect_stdout(io.StringIO()), contextlib.redirect_stderr(io.StringIO()):
            weekly.main(['--tasks', str(cls.root / 'tasks.xlsx'), '--suhail', str(cls.root / 'projects.xlsx'),
                         '--transactions', str(cls.root / 'transactions.xlsx'), '--output', str(cls.root / 'weekly.pptx')])
        cls.out = cls.root / 'sectors2'
        cls.split = run(SECTOR / 'split_report.py', cls.root / 'weekly.pptx', '--transactions', cls.root / 'transactions.xlsx',
                        '--outdir', cls.out, '--no-refine')
        cls.qa = run(SECTOR / 'qa_report.py', cls.out, '--master', cls.root / 'weekly.pptx', '--transactions', cls.root / 'transactions.xlsx')

    @classmethod
    def mark(cls, path, sheet, name_header, status, which):
        wb = load_workbook(path)
        ws = wb[sheet]
        for r in ws.iter_rows():
            for c in r:
                if c.value == 'الحالة':
                    header = c.row
                    column = c.column
                    break
            else:
                continue
            break
        row = header + 1
        while ws.cell(row, 1).value in (None, ''):
            row += 1
        ws.cell(row, column).value = status
        wb.save(path)
        return row

    @classmethod
    def tearDownClass(cls):
        cls.temp.cleanup()

    def titles(self):
        found = []
        for deck in self.out.glob('*.pptx'):
            for sl in Presentation(deck).slides:
                for sh in sl.shapes:
                    if sh.has_table:
                        found.append(splitter.table_title(splitter.table_rows(sh)))
        return found

    def test_split_succeeds_and_has_a_table_per_unknown_status(self):
        self.assertEqual(self.split.returncode, 0, self.split.stderr)
        titles = self.titles()
        self.assertIn('المهام (قيد المراجعة)', titles)
        self.assertIn('المشاريع (حالة غير معتمدة)', titles)

    def test_every_source_record_is_in_a_sector_deck(self):
        tasks = projects = 0
        for deck in self.out.glob('*.pptx'):
            for sl in Presentation(deck).slides:
                for sh in sl.shapes:
                    if not sh.has_table:
                        continue
                    rows = splitter.table_rows(sh)
                    title = splitter.table_title(rows)
                    if title.startswith('المهام ('):
                        tasks += len(rows) - 2
                    if title.startswith('المشاريع ('):
                        projects += len(rows) - 2
        self.assertEqual((tasks, projects), (1, 1))

    def test_notice_names_the_statuses(self):
        found = json.loads((self.out / 'manifest.json.notices.json').read_text(encoding='utf-8'))
        text = ' '.join(n['message'] for n in found if n['kind'] == 'status')
        self.assertIn('قيد المراجعة', text)
        self.assertIn('حالة غير معتمدة', text)

    def test_project_status_cell_is_neutral_gray(self):
        fills = set()
        for deck in self.out.glob('*.pptx'):
            for sl in Presentation(deck).slides:
                for sh in sl.shapes:
                    if sh.has_table and splitter.table_title(splitter.table_rows(sh)) == 'المشاريع (حالة غير معتمدة)':
                        for row in list(sh.table.rows)[2:]:
                            fills.add(str(row.cells[4].fill.fore_color.rgb))
        self.assertEqual(fills, {splitter.NEUTRAL_FILL})

    def test_qa_does_not_crash_and_checks_the_new_tables(self):
        self.assertNotIn('Traceback', self.qa.stderr + self.qa.stdout)
        self.assertEqual(self.qa.returncode, 0, self.qa.stdout + self.qa.stderr)
        self.assertIn('ALL OK', self.qa.stdout)


if __name__ == '__main__':
    unittest.main()
