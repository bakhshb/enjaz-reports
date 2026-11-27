"""A weekly deck with appendix slides splits cleanly, and a bad transactions workbook only drops that section."""
from pathlib import Path
import contextlib, io, shutil, tempfile, unittest
from openpyxl import Workbook, load_workbook
from pptx import Presentation
import run_weekly_workflow as workflow
from test_upstream_integrity import run

weekly = workflow.weekly
splitter = workflow.splitter
SECTOR = workflow.fixtures.SECTOR
TRANS_HEAD = ['رقم المعاملة', 'موضوع المعاملة', 'الجهة الوارد منها المعاملة', 'تاريخ انشاء المعاملة', 'تاريخ الإنجاز المخطط له']


def deck_tables(path):
    return [(i, sh.top, splitter.table_title(splitter.table_rows(sh)))
            for i, sl in enumerate(Presentation(path).slides) for sh in sl.shapes if sh.has_table]


class AppendixMasterTests(unittest.TestCase):
    TRANSACTIONS = 40   # delayed transactions in the appendix deck

    @classmethod
    def setUpClass(cls):
        cls.temp = tempfile.TemporaryDirectory()
        cls.root = Path(cls.temp.name) / 'overflow'
        with contextlib.redirect_stdout(io.StringIO()), contextlib.redirect_stderr(io.StringIO()):
            workflow.scenario(cls.root, 'overflow')      # also splits the shrunk (reduced font floor) deck and runs qa on it
            cls.shrunk = cls.root / 'weekly.pptx'
            cls.appendix = cls.root / 'appendix.pptx'
            # At a line height of 1.2 times the font size the scenario's 24 delayed transactions fit the
            # summary slide, so the appendix deck gets its own workbook with enough of them to need an appendix.
            source = cls.root / 'raw-transactions' / 'source.xlsx'
            wb = load_workbook(source); ws = wb.active
            for r in range(29, 29 + cls.TRANSACTIONS - 24):
                for c in [20, 24, 27, 28, 29]: ws.cell(r, c).value = ws.cell(5, c).value
                ws.cell(r, 31).value = f'{r:06d}'
            wb.save(source)
            cls.transactions = cls.root / 'transactions-appendix.xlsx'
            workflow.require(run(workflow.TRANSACTIONS, '--input-dir', source.parent, '--output', cls.transactions))
            weekly.main(['--tasks', str(cls.root / 'tasks.xlsx'), '--suhail', str(cls.root / 'projects.xlsx'),
                         '--transactions', str(cls.transactions), '--output', str(cls.appendix),
                         '--measure', '--summary-overflow', 'appendix'])

    @classmethod
    def tearDownClass(cls):
        cls.temp.cleanup()

    def split_and_qa(self, master, name, transactions=None):
        out = self.root / name
        transactions = transactions or self.root / 'transactions.xlsx'
        split = run(SECTOR / 'split_report.py', master, '--transactions', transactions,
                    '--outdir', out, '--no-refine')
        self.assertEqual(split.returncode, 0, split.stderr)
        qa = run(SECTOR / 'qa_report.py', out, '--master', master, '--transactions', transactions)
        return out, qa

    def test_the_master_really_has_appendix_tables_above_the_kpi_zone(self):
        tops = [top for _, top, title in deck_tables(self.appendix)]
        self.assertTrue(any(t < splitter.KPI_ZONE for t in tops))

    def test_locate_takes_the_slide_with_the_kpi_not_the_appendix(self):
        prs = splitter.open_master(self.appendix)
        with contextlib.redirect_stdout(io.StringIO()):
            role = splitter.locate(prs)
        slides = list(prs.slides)
        self.assertIn(role['trans'], range(len(slides)))
        self.assertIn('إجمالي', splitter.kpi_shapes(slides[role['trans']]))
        self.assertTrue(role['trans_slides'])   # the transactions table sits on an appendix, not on the KPI slide
        self.assertTrue(all(i > role['trans'] for i in role['trans_slides']))

    def test_split_of_an_appendix_master_passes_qa(self):
        out, qa = self.split_and_qa(self.appendix, 'sectors-appendix', self.transactions)
        self.assertEqual(qa.returncode, 0, qa.stdout + qa.stderr)
        self.assertNotIn('ISSUES FOUND', qa.stdout)
        for deck in out.glob('*.pptx'):
            prs = Presentation(deck)
            rows = [len(sh.table.rows) - 2 for sl in prs.slides for sh in sl.shapes
                    if sh.has_table and splitter.table_title(splitter.table_rows(sh)) == 'المعاملات المتأخرة']
            self.assertEqual(sum(rows), self.TRANSACTIONS, rows)   # each transaction once, not the master's table again

    def test_split_of_a_reduced_font_master_passes_qa(self):
        out, qa = self.split_and_qa(self.shrunk, 'sectors-shrunk')
        self.assertEqual(qa.returncode, 0, qa.stdout + qa.stderr)
        self.assertNotIn('ISSUES FOUND', qa.stdout)


class TransactionsWorkbookTests(unittest.TestCase):
    def items(self):
        return [{'head': list(TRANS_HEAD), 'row': ['1', 'موضوع', 'جهة', '01-01-2026', '02-01-2026']}]

    def workbook(self, directory, rows):
        path = Path(directory) / 'tx.xlsx'
        wb = Workbook()
        for r in rows:
            wb.active.append(r)
        wb.save(path)
        return str(path)

    def attach(self, rows, items=None):
        from report_common import notices
        notices.reset()
        with tempfile.TemporaryDirectory() as td:
            result = splitter.attach_transaction_sectors(items or self.items(), self.workbook(td, rows))
        return result, [n for n in notices.collected() if n['kind'] == 'source']

    def test_a_workbook_with_no_recognisable_header_warns_and_returns_nothing(self):
        result, notes = self.attach([['a', 'b'], [1, 2]])
        self.assertEqual(result, [])
        self.assertEqual(len(notes), 1)
        self.assertIn('خارج', notes[0]['message'])

    def test_a_workbook_missing_one_of_the_five_columns_warns_and_returns_nothing(self):
        result, notes = self.attach([['رقم المعاملة', 'موضوع المعاملة', 'القطاع'], ['1', 'موضوع', 'قطاع']])
        self.assertEqual(result, [])
        self.assertEqual(len(notes), 1)

    def test_an_unexpected_header_in_the_deck_table_warns_and_returns_nothing(self):
        items = [{'head': ['رقم المعاملة', 'موضوع المعاملة', 'عمود آخر'], 'row': ['1', 'موضوع', 'x']}]
        result, notes = self.attach([TRANS_HEAD + ['القطاع'], ['1', 'موضوع', 'جهة', '01-01-2026', '02-01-2026', 'قطاع']], items)
        self.assertEqual(result, [])
        self.assertEqual(len(notes), 1)


if __name__ == '__main__':
    unittest.main()
