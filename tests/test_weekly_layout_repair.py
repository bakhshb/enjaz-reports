"""The build repairs the builder's own table overlap once, reports what remains, and clears stale notices."""
from pathlib import Path
import contextlib, io, json, shutil, tempfile, unittest
from unittest.mock import patch
from pptx import Presentation
import run_weekly_workflow as workflow

weekly = workflow.weekly


def overlaps(path):
    found = []
    for number, sl in enumerate(Presentation(path).slides, 1):
        tables = sorted(weekly.builder.tables(sl), key=lambda sh: sh.top)
        if any(a.top + a.height > b.top for a, b in zip(tables, tables[1:])):
            found.append(number)
    return found


class LayoutRepairTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.temp = tempfile.TemporaryDirectory()
        cls.root = Path(cls.temp.name) / 'baseline'
        with contextlib.redirect_stdout(io.StringIO()):
            workflow.scenario(cls.root, 'normal')

    @classmethod
    def tearDownClass(cls):
        cls.temp.cleanup()

    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.case = Path(self.temp.name)
        for path in self.root.glob('*.xlsx'):
            shutil.copy2(path, self.case / path.name)
        self.out = self.case / 'weekly.pptx'

    def tearDown(self):
        self.temp.cleanup()

    def args(self):
        return ['--tasks', str(self.case / 'tasks.xlsx'), '--suhail', str(self.case / 'projects.xlsx'),
                '--transactions', str(self.case / 'transactions.xlsx'), '--output', str(self.out)]

    def notices(self):
        path = weekly.notices.notices_path(self.out)
        return json.loads(path.read_text(encoding='utf-8')) if path.exists() else []

    def build_with_overlap(self):
        real = weekly.build_candidate

        def broken(args):
            deck = real(args)
            prs = Presentation(deck)
            for sl in prs.slides:
                tables = sorted(weekly.builder.tables(sl), key=lambda sh: sh.top)
                if len(tables) >= 2:
                    tables[1].top = tables[0].top + tables[0].height // 2
                    break
            prs.save(deck)
            return deck
        return patch.object(weekly, 'build_candidate', broken)

    def test_overlap_made_by_the_build_is_repaired_before_publishing(self):
        with self.build_with_overlap(), contextlib.redirect_stdout(io.StringIO()), contextlib.redirect_stderr(io.StringIO()):
            weekly.main(self.args())
        self.assertEqual(overlaps(self.out), [])
        self.assertFalse([n for n in self.notices() if 'متداخلان' in n['message']])

    def test_overlap_that_cannot_be_repaired_is_a_notice_naming_the_slide(self):
        with self.build_with_overlap(), patch.object(weekly, 'repair_layout', lambda deck, slides: None), \
                contextlib.redirect_stdout(io.StringIO()), contextlib.redirect_stderr(io.StringIO()):
            weekly.main(self.args())
        found = [n for n in self.notices() if n['kind'] == 'layout' and 'متداخلان' in n['message']]
        self.assertTrue(found)
        self.assertRegex(found[0]['message'], r'الشريحة \d+')

    def test_layout_slides_reads_both_findings(self):
        found = weekly.layout_slides(['Tables overlap on slide 4', 'Table frame reaches footer artwork on slide 7; rendered correction required', 'other'])
        self.assertEqual(found, {weekly.LAYOUT_FINDINGS[0]: [4], weekly.LAYOUT_FINDINGS[1]: [7]})

    def test_stale_notices_are_removed_when_a_build_fails(self):
        stale = weekly.notices.notices_path(self.out)
        stale.write_text('[{"kind":"layout","message":"old"}]', encoding='utf-8')
        with patch.object(weekly, 'build_candidate', side_effect=RuntimeError('injected')):
            with self.assertRaises(RuntimeError):
                weekly.main(self.args())
        self.assertFalse(stale.exists())


if __name__ == '__main__':
    unittest.main()
