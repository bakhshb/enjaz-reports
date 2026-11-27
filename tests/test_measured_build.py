# /// script
# requires-python = ">=3.12"
# dependencies = ["pandas>=2.2,<3", "openpyxl>=3.1,<4", "python-pptx>=1.0.2", "lxml>=5.3", "pillow", "xlrd>=2", "pdfplumber>=0.11"]
# ///
"""`build_report.py --measure`: one command builds, measures through LibreOffice, rebuilds and validates.

A deck built with estimated row heights overlapped and ran past the footer in
PowerPoint (26-11-2026), so the delivery path never depends on a separate
measuring step being remembered.
"""
import contextlib
import io
import json
import shutil
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from pptx import Presentation

import run_weekly_workflow as workflow

weekly = workflow.weekly


def run(arguments):
    out, err = io.StringIO(), io.StringIO()
    with contextlib.redirect_stdout(out), contextlib.redirect_stderr(err):
        weekly.main(arguments)
    return out.getvalue()


class MeasuredBuildTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.temp = tempfile.TemporaryDirectory()
        cls.root = Path(cls.temp.name)
        with contextlib.redirect_stdout(io.StringIO()), contextlib.redirect_stderr(io.StringIO()):
            workflow.scenario(cls.root, 'overflow')

    @classmethod
    def tearDownClass(cls):
        cls.temp.cleanup()

    def arguments(self, output):
        return ['--tasks', str(self.root / 'tasks.xlsx'), '--suhail', str(self.root / 'projects.xlsx'),
                '--transactions', str(self.root / 'transactions.xlsx'), '--output', str(output)]

    @unittest.skipUnless(shutil.which('soffice'), 'LibreOffice is required to measure rows')
    def test_one_command_delivers_a_measured_validated_report(self):
        output = self.root / 'measured.pptx'
        printed = run(self.arguments(output) + ['--measure'])
        self.assertIn('measured build', printed)
        layout = json.loads(output.with_suffix('.layout.json').read_text(encoding='utf-8'))
        self.assertTrue(layout['tables'])
        found = json.loads(weekly.notices.notices_path(output).read_text(encoding='utf-8')) if weekly.notices.notices_path(output).exists() else []
        self.assertFalse(any('تقديرية' in n['message'] for n in found), found)
        result = workflow.weekly_validation.validate(output, self.root / 'tasks.xlsx', self.root / 'projects.xlsx',
                                                     self.root / 'transactions.xlsx', strict=False)
        self.assertEqual(result['programmatic'], 'passed')
        prs = Presentation(output)
        for slide in prs.slides:
            tables = sorted(weekly.builder.tables(slide), key=lambda sh: sh.top)
            self.assertTrue(all(a.top + a.height < b.top for a, b in zip(tables, tables[1:])))
            for shape in tables:
                self.assertEqual(shape.height, sum(row.height for row in shape.table.rows))
                self.assertLessEqual(shape.top + shape.height, weekly.content_bottom(prs, slide))
        # Each summary is one slide unless a notice says its overflow went to an appendix.
        appendix = any(n['kind'] == 'space' and 'ملحق' in n['message'] for n in found)
        charts = [sl for sl in prs.slides if any(sh.has_chart for sh in sl.shapes)]
        summary_tables = [sl for sl in prs.slides if any(weekly.builder.table_title(sh) in weekly.builder.TASK_SUMMARY_TITLES
                                                         for sh in weekly.builder.tables(sl, 4))]
        self.assertTrue(appendix or all(sl in charts for sl in summary_tables))

    def test_without_libreoffice_the_report_is_still_built_and_flagged(self):
        output = self.root / 'unmeasured.pptx'
        import measure_layout
        with patch.object(measure_layout, 'measure', side_effect=RuntimeError('LibreOffice is required for PDF rendering')):
            printed = run(self.arguments(output) + ['--measure'])
        self.assertIn('PROVISIONAL', printed)
        found = json.loads(weekly.notices.notices_path(output).read_text(encoding='utf-8'))
        self.assertTrue(any(n['kind'] == 'layout' and 'LibreOffice' in n['message'] for n in found))
        self.assertTrue(output.exists())

    def test_unusable_measurements_become_a_notice_and_the_build_continues(self):
        cases = {'malformed': {}, 'stale hash': {'hashes': {'tasks': '0', 'suhail': '0', 'transactions': '0', 'template': '0'},
                                                   'typography': 'x', 'tables': {}},
                 'not json at all': None}
        for name, content in cases.items():
            with self.subTest(case=name):
                layout = self.root / f'{name.replace(" ", "-")}.layout.json'
                layout.write_text('{not json' if content is None else json.dumps(content), encoding='utf-8')
                output = self.root / f'{layout.stem}.pptx'
                printed = run(self.arguments(output) + ['--measurements', str(layout)])
                self.assertIn('PROVISIONAL', printed)
                found = json.loads(weekly.notices.notices_path(output).read_text(encoding='utf-8'))
                self.assertTrue(any(n['kind'] == 'layout' and 'ملف القياس' in n['message'] for n in found), found)
                self.assertTrue(output.exists())

    def test_measure_failing_with_value_or_key_error_falls_back_to_the_estimate(self):
        import measure_layout
        for error in (ValueError('x'), KeyError('tables')):
            with self.subTest(error=type(error).__name__):
                output = self.root / f'fallback-{type(error).__name__}.pptx'
                with patch.object(measure_layout, 'measure', side_effect=error):
                    printed = run(self.arguments(output) + ['--measure'])
                self.assertIn('PROVISIONAL', printed)
                found = json.loads(weekly.notices.notices_path(output).read_text(encoding='utf-8'))
                self.assertTrue(any(n['kind'] == 'layout' and 'تعذر قياس' in n['message'] for n in found))
                self.assertTrue(any('الملف ناقص' in n['message'] for n in found), found)

    def test_measure_and_measurements_are_exclusive(self):
        with self.assertRaisesRegex(ValueError, 'not both'):
            weekly.main(self.arguments(self.root / 'x.pptx') + ['--measure', '--measurements', 'layout.json'])


if __name__ == '__main__':
    unittest.main()
