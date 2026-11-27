# /// script
# requires-python = ">=3.12"
# dependencies = ["pandas>=2.2,<3", "openpyxl>=3.1,<4", "python-pptx>=1.0.2", "lxml>=5.3", "pillow", "xlrd>=2", "pdfplumber>=0.11"]
# ///
"""The four summary slides share one look (owner's rule, 27-11-2026).

Margins and size are decided once for the tasks, Suhail, transactions and
digital summaries together: every summary fits at 11 pt with the template
margins and nothing changes; otherwise every summary table takes the compact
margins of the template's detail tables and the three standard summaries share
one size (floor 10); a table that still does not fit moves whole to an appendix.
The digital summary alone may go smaller, down to 8. The tests build real decks
from the bundled templates with made-up measurements (no LibreOffice).
"""
import contextlib
import hashlib
import io
import json
import re
import sys
import tempfile
import unittest
from pathlib import Path

from pptx import Presentation

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent
WEEKLY = ROOT / 'skills/weekly-consolidated-report/scripts'
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(ROOT / 'skills'))
sys.path.insert(0, str(WEEKLY))
import run_weekly_workflow as workflow
import digital_maturity as digital
import validate_report
from weekly_data import (MASTER, SUMMARY_SIZES, SUMMARY_LABELS, TABLE_GAP, compact_margin, size_key, plan_summaries,
                         summary_fits, Summary)
from report_common import pptx_helpers as helpers
from report_common.validation import sha
from test_weekly_digital_maturity import master_sectors

weekly = workflow.weekly
PT = 12700
KEYS = ('tasks', 'suhail', 'transactions')
SUMMARY_TEXT = {'tasks': 'إجمالي المهام', 'suhail': 'إجمالي مشاريع سهيل', 'transactions': 'إجمالي المعاملات'}


def ramp(first, step=0.1):
    """Fractions of the available space the compact tables use at each summary size, falling by `step` per size."""
    return {size_key(size): first - index * step for index, size in enumerate(SUMMARY_SIZES)}


class Inputs:
    """Synthetic inputs of one model, a provisional deck, and builds with made-up measured heights."""

    def __init__(self, with_digital=False):
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name)
        with contextlib.redirect_stdout(io.StringIO()), contextlib.redirect_stderr(io.StringIO()):
            workflow.scenario(self.root, 'normal')
        self.digital = with_digital
        self.args = ['--tasks', str(self.root / 'tasks.xlsx'), '--suhail', str(self.root / 'projects.xlsx'),
                     '--transactions', str(self.root / 'transactions.xlsx')]
        self.template = digital.MASTER if with_digital else MASTER
        self.sectors = master_sectors()
        if with_digital:
            pdf = self.root / 'source.pdf'
            pdf.write_bytes(b'synthetic source identity')
            review = self.root / 'review.json'
            review.write_text(json.dumps({
                'source_pdf_sha256': hashlib.sha256(pdf.read_bytes()).hexdigest(), 'reviewed': True,
                'metrics': {'experience': 15, 'satisfaction': 10, 'complaints': 20, 'tools': 65},
                'sectors': [{'name': 'القطاع الأول', 'note': None, 'metrics': {'experience': 15, 'satisfaction': 10, 'complaints': 20, 'tools': 65}}],
                'updates': [{'perspective': 'عام', 'text': f'تحديث رقم {i} عن سير العمل في المنصات'} for i in range(1, 7)],
                'detail_sectors': self.sectors}, ensure_ascii=False), encoding='utf-8')
            self.args += ['--model', 'digital', '--digital-pdf', str(pdf), '--digital-data', str(review)]
            self.pdf, self.review = pdf, review
        provisional = self.run(self.root / 'provisional.pptx')
        prs = Presentation(provisional)
        self.slides = {}
        self.room = {}
        for key in KEYS:
            slide = helpers.find_slide_with_text(prs, SUMMARY_TEXT[key])
            shapes = sorted(helpers.tables(slide), key=lambda sh: sh.top)
            self.slides[key] = [(helpers.table_title(sh), len(sh.table.rows) - 2, [r.height for r in list(sh.table.rows)[:2]]) for sh in shapes]
            self.room[key] = weekly.content_bottom(prs, slide) - shapes[0].top
        if with_digital:
            slide = digital.digital_slide(prs)
            shape = digital.summary_table_shape(slide)
            self.slides['digital'] = [('أبرز التحديثات', len(shape.table.rows) - 2, [r.height for r in list(shape.table.rows)[:2]])]
            self.room['digital'] = weekly.content_bottom(prs, slide) - shape.top

    def close(self):
        self.temp.cleanup()

    def run(self, output, *extra):
        with contextlib.redirect_stdout(io.StringIO()), contextlib.redirect_stderr(io.StringIO()):
            weekly.main(self.args + ['--output', str(output), *extra])
        return output

    def pt(self, key, fraction):
        """Height of each body row (points) so the tables of `key` use `fraction` of its space at 11 pt."""
        tables = self.slides[key]
        fixed = sum(sum(h) for _, _, h in tables) + TABLE_GAP * (len(tables) - 1)
        count = sum(n for _, n, _ in tables)
        return max(fraction * self.room[key] - fixed, 10 * PT * count) / count / PT

    def pt_of(self, key, fraction):
        """Height (points) of one body row that makes it `fraction` of the room of `key`."""
        return fraction * self.room[key] / PT

    def layout(self, profile, compact=True):
        """A layout file. `profile` maps a summary to ``(template fraction, {size: compact fraction})``."""
        tables = {}
        digital_part = None
        for key, found in self.slides.items():
            template, by_size = profile.get(key, (0.6, ramp(0.5, 0.0)))
            for title, count, header in found:
                body = [self.pt(key, template)] * count
                entry = {'header': [h / PT for h in header], 'body': body}
                if compact:
                    entry['compact'] = {size: [self.pt(key, fraction)] * count for size, fraction in by_size.items()}
                if key == 'digital':
                    digital_part = {'summary': body, 'sectors': {
                        s['title']: {'rows': [20] * len(s['rows']), 'updates': [20] * len(s['updates']),
                                     'challenges': [20] * len(s['challenges'])} for s in self.sectors}}
                    if compact:
                        digital_part['summary_compact'] = entry['compact']
                else:
                    tables[title] = entry
        hashes = {'tasks': sha(self.root / 'tasks.xlsx'), 'suhail': sha(self.root / 'projects.xlsx'),
                  'transactions': sha(self.root / 'transactions.xlsx'), 'template': sha(self.template)}
        data = {'typography': '11pt-line-1.2-updates-v4', 'hashes': hashes, 'tables': tables}
        if self.digital:
            hashes.update(digital_pdf=sha(self.pdf), digital_data=sha(self.review))
            data['digital'] = digital_part
        return data

    def build(self, profile, name='deck', compact=True, extra=()):
        layout = self.root / f'{name}.layout.json'
        layout.write_text(json.dumps(self.layout(profile, compact)), encoding='utf-8')
        output = self.root / f'{name}.pptx'
        self.run(output, '--measurements', str(layout), *extra)
        path = weekly.notices.notices_path(output)
        found = json.loads(path.read_text(encoding='utf-8')) if path.exists() else []
        return Result(output, found, self)


class Result:
    def __init__(self, path, found, inputs):
        self.path, self.notices, self.inputs = path, found, inputs
        self.prs = Presentation(path)
        self.master = Presentation(inputs.template)
        self.compact = compact_margin(self.master)

    def space(self):
        return [n['message'] for n in self.notices if n['kind'] == 'space']

    def summary_slides(self):
        slides = {key: helpers.find_slide_with_text(self.prs, SUMMARY_TEXT[key]) for key in KEYS}
        if self.inputs.digital:
            slides['digital'] = digital.digital_slide(self.prs)
        return slides

    def summary_tables(self, key):
        return sorted(helpers.tables(self.summary_slides()[key]), key=lambda sh: sh.top)

    def appendix_tables(self, key):
        """Tables of the slides right after summary `key` that carry no chart."""
        slides = list(self.prs.slides)
        start = slides.index(self.summary_slides()[key]) + 1
        found = []
        for slide in slides[start:]:
            if any(sh.has_chart for sh in slide.shapes) or helpers.detail_tables(slide, 'task') or helpers.detail_tables(slide, 'project') or not helpers.tables(slide):
                break
            found.extend(helpers.tables(slide))
        return found

    def summary_tables_in(self, prs, key):
        """Summary tables of `key` in another copy `prs` of this deck."""
        slide = digital.digital_slide(prs) if key == 'digital' else helpers.find_slide_with_text(prs, SUMMARY_TEXT[key])
        return sorted(helpers.tables(slide), key=lambda sh: sh.top)

    @staticmethod
    def margins(shape):
        return {(c.margin_left, c.margin_right, c.margin_top, c.margin_bottom)
                for row in list(shape.table.rows)[2:] for c in row.cells}

    @staticmethod
    def header_margins(shape):
        return {(c.margin_left, c.margin_right, c.margin_top, c.margin_bottom)
                for row in list(shape.table.rows)[:2] for c in row.cells}

    @staticmethod
    def sizes(shape):
        return {run.font.size.pt for row in shape.table.rows for c in row.cells
                for p in c.text_frame.paragraphs for run in p.runs if run.text.strip()}

    def template_margins(self, shape):
        title = helpers.table_title(shape)
        pattern = next(sh for sl in self.master.slides for sh in helpers.tables(sl)
                       if len(sh.table.columns) == len(shape.table.columns) and (len(shape.table.columns) == 3 or helpers.table_title(sh) == title))
        return {(c.margin_left, c.margin_right, c.margin_top, c.margin_bottom) for c in list(pattern.table.rows)[2].cells}

    def all_template(self):
        return all(self.margins(sh) == self.template_margins(sh) for key in self.summary_slides() for sh in self.summary_tables(key))

    def all_compact(self):
        return all(self.margins(sh) == {(self.compact,) * 4} for key in self.summary_slides() for sh in self.summary_tables(key))

    def findings(self):
        """Validator findings about margins or sizes (they become notices of kind check)."""
        return [n['message'] for n in self.notices if n['kind'] == 'check']


class SharedLookBuildTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.standard = Inputs()

    @classmethod
    def tearDownClass(cls):
        cls.standard.close()

    def test_everything_fits_changes_nothing(self):
        result = self.standard.build({})
        self.assertEqual(result.space(), [])
        self.assertTrue(result.all_template())
        for key in KEYS:
            for shape in result.summary_tables(key):
                self.assertEqual(result.sizes(shape), {11})
        self.assertEqual(result.findings(), [])

    def test_one_summary_needing_compact_margins_switches_every_summary(self):
        # The Suhail summary only fits with the compact margins; the others would fit as they are.
        result = self.standard.build({'suhail': (1.3, ramp(0.9, 0.0))})
        self.assertTrue(result.all_compact())
        for key in KEYS:
            for shape in result.summary_tables(key):
                self.assertEqual(result.sizes(shape), {11})
        self.assertEqual(result.space(), [
            'ملخص مشاريع سهيل: لم تتسع الجداول في الشريحة بخط 11. تم تقليص هوامش خلايا الجداول في جميع شرائح الملخص '
            'ليبقى شكلها موحدا، ويبقى الخط 11. المقترح: اختصار النص في الملف.'])
        self.assertEqual(result.findings(), [])

    def test_title_and_header_rows_keep_the_template_margins(self):
        result = self.standard.build({'tasks': (1.3, ramp(0.9, 0.0))})
        for key in KEYS:
            for shape in result.summary_tables(key):
                template = Presentation(MASTER)
                pattern = next(sh for sl in template.slides for sh in helpers.tables(sl) if helpers.table_title(sh) == helpers.table_title(shape))
                expected = {(c.margin_left, c.margin_right, c.margin_top, c.margin_bottom) for row in list(pattern.table.rows)[:2] for c in row.cells}
                self.assertEqual(result.header_margins(shape), expected)

    def test_compact_and_smaller_font_are_shared_by_the_three_standard_summaries(self):
        # Tasks need 10 pt; Suhail and transactions would fit at 11 but follow.
        result = self.standard.build({'tasks': (1.4, ramp(1.15, 0.1))})
        self.assertTrue(result.all_compact())
        for key in KEYS:
            for shape in result.summary_tables(key):
                self.assertEqual(result.sizes(shape), {10}, key)
        self.assertEqual(result.space(), [
            'ملخص المهام: لم تتسع الجداول في الشريحة بخط 11. تم تقليص هوامش خلايا الجداول وتصغير الخط إلى 10 '
            'في جميع شرائح الملخص ليبقى شكلها موحدا. المقترح: اختصار النص في الملف.'])
        self.assertEqual(result.findings(), [])

    def test_past_the_floor_whole_tables_go_to_an_appendix_and_nothing_is_below_10(self):
        result = self.standard.build({'tasks': (1.6, ramp(1.6, 0.02))})
        self.assertTrue(result.all_compact())
        for key in KEYS:
            for shape in result.summary_tables(key):
                self.assertEqual(result.sizes(shape), {10}, key)
        kept = [helpers.table_title(sh) for sh in result.summary_tables('tasks')]
        moved = [helpers.table_title(sh) for sh in result.appendix_tables('tasks')]
        self.assertTrue(moved)
        self.assertEqual(sorted(kept + moved), sorted(title for title, _, _ in self.standard.slides['tasks']))
        for shape in result.appendix_tables('tasks'):
            self.assertEqual(result.sizes(shape), {11})
            self.assertEqual(result.margins(shape), result.template_margins(shape))
            self.assertEqual(len(shape.table.rows) - 2, dict((t, n) for t, n, _ in self.standard.slides['tasks'])[helpers.table_title(shape)])
        names = '، '.join(f'«{t}»' for t in moved)
        shown = f'جدول {names}' if len(moved) == 1 else f'الجداول ({names})'
        whole = 'كاملا' if len(moved) == 1 else 'كاملة'
        self.assertIn(f'ملخص المهام: لم تتسع الجداول في الشريحة بعد تقليص الهوامش وتصغير الخط إلى 10. تم نقل {shown} {whole} '
                      'إلى شريحة ملحق بعدها بخط 11. المقترح: اختصار النص في الملف.', result.space())
        self.assertEqual(result.findings(), [])

    def build_suhail(self, name, shapes):
        """A build where the two Suhail tables get their own made-up heights.

        `shapes` maps a table title to ``(template fraction, compact fraction)`` of the room,
        for its one row, whatever the size.
        """
        inputs = self.standard
        layout = inputs.layout({})
        for title, (template, compact) in shapes.items():
            entry = layout['tables'][title]
            entry['compact'] = {size: [inputs.pt_of('suhail', compact)] for size in entry['compact']}
            entry['body'] = [inputs.pt_of('suhail', template)]
        path = inputs.root / f'{name}.layout.json'
        path.write_text(json.dumps(layout), encoding='utf-8')
        output = inputs.root / f'{name}.pptx'
        inputs.run(output, '--measurements', str(path))
        found = weekly.notices.notices_path(output)
        return Result(output, json.loads(found.read_text(encoding='utf-8')) if found.exists() else [], inputs)

    def check_keep_what_fits(self, result, kept, moved):
        titles = lambda shapes: [helpers.table_title(sh) for sh in shapes]
        self.assertEqual(titles(result.summary_tables('suhail')), [kept])
        self.assertEqual(titles(result.appendix_tables('suhail')), [moved])
        self.assertEqual(result.sizes(result.summary_tables('suhail')[0]), {10})
        self.assertEqual(result.margins(result.summary_tables('suhail')[0]), {(result.compact,) * 4})
        big = result.appendix_tables('suhail')[0]
        self.assertEqual(result.sizes(big), {11})
        self.assertEqual(result.margins(big), result.template_margins(big))
        self.assertEqual(len(big.table.rows) - 2, 1)
        slides = list(result.prs.slides)
        after = slides[slides.index(result.summary_slides()['suhail']) + 1:]
        self.assertEqual(sum(1 for sl in after[:1] if helpers.tables(sl) and not any(sh.has_chart for sh in sl.shapes)), 1)
        self.assertFalse(result.appendix_tables('suhail')[1:])
        notice = [n for n in result.space() if n.startswith('ملخص مشاريع سهيل: لم تتسع الجداول في الشريحة بعد')]
        self.assertEqual(len(notice), 1, result.space())
        self.assertIn(f'تم نقل جدول «{moved}» كاملا إلى شريحة ملحق بعدها بخط 11.', notice[0])
        self.assertNotIn(kept, notice[0])
        self.assertEqual(result.findings(), [])

    def test_a_small_table_after_one_that_cannot_fit_stays_on_the_summary_slide(self):
        # Reported case: the long first table moves, the one-row table still fits under the chart at 10 pt.
        result = self.build_suhail('keep-small-second', {'المشاريع المتأخرة': (0.95, 1.1), 'التحديات': (0.2, 0.1)})
        self.check_keep_what_fits(result, 'التحديات', 'المشاريع المتأخرة')

    def test_a_small_first_table_stays_and_the_big_second_one_moves(self):
        result = self.build_suhail('keep-small-first', {'المشاريع المتأخرة': (0.2, 0.1), 'التحديات': (0.95, 1.1)})
        self.check_keep_what_fits(result, 'المشاريع المتأخرة', 'التحديات')

    def test_explicit_floor_is_the_floor_for_all(self):
        profile = {'tasks': (1.6, ramp(1.6, 0.02))}
        result = self.standard.build(profile, 'floor-9', extra=['--summary-min-pt', '9'])
        for key in KEYS:
            for shape in result.summary_tables(key):
                self.assertEqual(result.sizes(shape), {9}, key)
        result = self.standard.build({'tasks': (1.6, ramp(1.25, 0.05))}, 'floor-9-fits', extra=['--summary-min-pt', '9'])
        for key in KEYS:
            for shape in result.summary_tables(key):
                self.assertTrue(result.sizes(shape) <= {9, 9.5, 10}, (key, result.sizes(shape)))
        self.assertFalse(result.appendix_tables('tasks'))

    def test_appendix_mode_keeps_11_and_shares_the_compact_margins(self):
        result = self.standard.build({'tasks': (1.6, ramp(1.3, 0.1))}, 'appendix-mode', extra=['--summary-overflow', 'appendix'])
        self.assertTrue(result.all_compact())
        for key in KEYS:
            for shape in result.summary_tables(key):
                self.assertEqual(result.sizes(shape), {11}, key)
        moved = [helpers.table_title(sh) for sh in result.appendix_tables('tasks')]
        self.assertTrue(moved)
        names = '، '.join(f'«{t}»' for t in moved)
        shown = f'جدول {names}' if len(moved) == 1 else f'الجداول ({names})'
        whole = 'كاملا' if len(moved) == 1 else 'كاملة'
        self.assertIn(f'ملخص المهام: لم تتسع الجداول في الشريحة بخط 11 بعد تقليص الهوامش. تم نقل {shown} {whole} إلى شريحة ملحق بعدها حسب الطلب. '
                      'المقترح: اختصار النص في الملف.', result.space())
        self.assertEqual(result.findings(), [])

    def test_floor_of_11_moves_tables_without_claiming_a_shrink(self):
        result = self.standard.build({'tasks': (1.6, ramp(1.3, 0.1))}, 'floor-11', extra=['--summary-min-pt', '12'])
        moved = [helpers.table_title(sh) for sh in result.appendix_tables('tasks')]
        self.assertTrue(moved)
        notice = [n for n in result.space() if n.startswith('ملخص المهام: لم تتسع الجداول في الشريحة بخط 11 بعد تقليص')]
        self.assertEqual(len(notice), 1, result.space())
        self.assertNotIn('تصغير الخط', notice[0])
        self.assertNotIn('حسب الطلب', notice[0])
        names = '، '.join(f'«{t}»' for t in moved)
        shown = f'جدول {names} كاملا' if len(moved) == 1 else f'الجداول ({names}) كاملة'
        self.assertIn(f'تم نقل {shown} إلى شريحة ملحق بعدها. المقترح: اختصار النص في الملف.', notice[0])

    def test_moved_names_word_follows_the_count(self):
        import build_report
        self.assertEqual(build_report.moved_whole(['أ']), 'جدول «أ» كاملا')
        self.assertEqual(build_report.moved_whole(['أ', 'ب']), 'الجداول («أ»، «ب») كاملة')

    def test_appendix_mode_with_everything_fitting_changes_nothing(self):
        result = self.standard.build({}, 'appendix-fits', extra=['--summary-overflow', 'appendix'])
        self.assertTrue(result.all_template())
        self.assertEqual(result.space(), [])

    def test_a_layout_file_without_compact_data_still_builds_with_the_bound(self):
        # Template-margin heights only (an older file or the PowerShell measurer): the bound is used.
        result = self.standard.build({'tasks': (1.15, {}), 'suhail': (0.5, {}), 'transactions': (0.5, {})}, 'older', compact=False)
        self.assertTrue(result.all_compact())
        sizes = {size for key in KEYS for shape in result.summary_tables(key) for size in result.sizes(shape)}
        self.assertEqual(len(sizes), 1)
        self.assertLessEqual(sizes.pop(), 11)
        self.assertEqual(result.findings(), [])
        for key in KEYS:
            tables = result.summary_tables(key)
            self.assertLessEqual(tables[-1].top + tables[-1].height, weekly.content_bottom(result.prs, result.summary_slides()[key]))

    def test_compact_margin_comes_from_the_template(self):
        template = Presentation(MASTER)
        self.assertEqual(compact_margin(template), 9144)
        shape = helpers.detail_tables(helpers.detail_pattern_slides(template, 'task')[0], 'task')[0]
        for row in list(shape.table.rows)[2:]:
            row.cells[0].margin_left = 20000
        self.assertEqual(compact_margin(template), 20000)
        # The build uses the template it is given: a copy with another detail margin gives the summaries that margin.
        changed = self.standard.root / 'template-20000.pptx'
        for row in list(shape.table.rows)[2:]:
            for cell in row.cells:
                cell.margin_left = 20000
        template.save(changed)
        output = self.standard.root / 'own-template.pptx'
        with contextlib.redirect_stdout(io.StringIO()), contextlib.redirect_stderr(io.StringIO()):
            layout = self.standard.layout({'tasks': (1.3, ramp(0.9, 0.0))})
            layout['hashes']['template'] = sha(changed)
            path = self.standard.root / 'own-template.layout.json'
            path.write_text(json.dumps(layout), encoding='utf-8')
            weekly.main(self.standard.args + ['--template', str(changed), '--measurements', str(path), '--output', str(output)])
        prs = Presentation(output)
        shapes = [sh for key in KEYS for sh in helpers.tables(helpers.find_slide_with_text(prs, SUMMARY_TEXT[key]))]
        for shape in shapes:
            self.assertEqual(Result.margins(shape), {(20000,) * 4})

    def test_notices_carry_no_latin_letters_no_double_dash_no_em_dash(self):
        cases = [self.standard.build({'suhail': (1.3, ramp(0.9, 0.0))}, 'n1'),
                 self.standard.build({'tasks': (1.4, ramp(1.15, 0.1))}, 'n2'),
                 self.standard.build({'tasks': (1.6, ramp(1.6, 0.02))}, 'n3'),
                 self.standard.build({'tasks': (1.6, ramp(1.3, 0.1))}, 'n4', extra=['--summary-overflow', 'appendix'])]
        for result in cases:
            self.assertTrue(result.space())
            for message in result.space():
                self.assertIsNone(re.search('[A-Za-z]', message), message)
                self.assertNotIn('--', message)
                self.assertNotIn('—', message)


class DigitalSharedLookTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.inputs = Inputs(with_digital=True)

    @classmethod
    def tearDownClass(cls):
        cls.inputs.close()

    def test_everything_fits_changes_nothing(self):
        result = self.inputs.build({})
        self.assertEqual(result.space(), [])
        self.assertTrue(result.all_template())
        self.assertEqual(result.sizes(result.summary_tables('digital')[0]), {11})

    def test_digital_alone_needing_compact_margins_switches_every_slide(self):
        result = self.inputs.build({'digital': (1.3, ramp(0.9, 0.0))})
        self.assertTrue(result.all_compact())
        for key in result.summary_slides():
            for shape in result.summary_tables(key):
                self.assertEqual(result.sizes(shape), {11}, key)
        number = list(result.prs.slides).index(result.summary_slides()['digital']) + 1
        self.assertEqual(result.space(), [
            f'ملخص التجربة الرقمية (الشريحة {number}): لم تتسع الجداول في الشريحة بخط 11. تم تقليص هوامش خلايا الجداول في جميع شرائح الملخص '
            'ليبقى شكلها موحدا، ويبقى الخط 11. المقترح: اختصار النص في الملف.'])
        self.assertEqual(result.findings(), [])

    def test_digital_takes_the_shared_smaller_size_when_it_fits_there(self):
        # Tasks need 10 pt; the digital table would fit at 10 and so follows the shared size.
        result = self.inputs.build({'tasks': (1.4, ramp(1.15, 0.1))}, 'digital-follows')
        self.assertTrue(result.all_compact())
        for key in result.summary_slides():
            for shape in result.summary_tables(key):
                self.assertEqual(result.sizes(shape), {10}, key)
        self.assertEqual(result.findings(), [])

    def test_digital_notices_carry_no_latin_letters_no_double_dash_no_em_dash(self):
        cases = [self.inputs.build({'digital': (1.3, ramp(0.9, 0.0))}, 'dn1'),
                 self.inputs.build({'digital': (1.6, ramp(1.3, 0.1))}, 'dn2'),
                 self.inputs.build({'digital': (3.0, ramp(3.0, 0.0))}, 'dn3'),
                 self.inputs.build({'tasks': (1.4, ramp(1.15, 0.1))}, 'dn4')]
        for result in cases:
            self.assertTrue(result.space())
            for message in result.space():
                self.assertIsNone(re.search('[A-Za-z]', message), message)
                self.assertNotIn('--', message)
                self.assertNotIn('—', message)

    def test_digital_alone_needing_a_smaller_font_goes_smaller_alone(self):
        # The digital table fits at 9 only; the standard summaries stay at the shared 11.
        result = self.inputs.build({'digital': (1.6, ramp(1.3, 0.1))}, 'digital-smaller')
        self.assertTrue(result.all_compact())
        for key in KEYS:
            for shape in result.summary_tables(key):
                self.assertEqual(result.sizes(shape), {11}, key)
        digital_size = result.sizes(result.summary_tables('digital')[0])
        self.assertEqual(len(digital_size), 1)
        size = digital_size.pop()
        self.assertTrue(8 <= size < 11)
        number = list(result.prs.slides).index(result.summary_slides()['digital']) + 1
        label = f'ملخص التجربة الرقمية (الشريحة {number})'
        self.assertIn(f'{label}: لم يتسع جدول «أبرز التحديثات» بخط 11 بعد تقليص الهوامش. تم تصغير خط هذا الجدول وحده إلى {size:g}. '
                      'المقترح: اختصار التحديثات في ملف المراجعة.', result.space())
        self.assertEqual(result.findings(), [])

    def test_digital_below_10_is_allowed_while_standard_stay_at_the_shared_size(self):
        result = self.inputs.build({'tasks': (1.4, ramp(1.15, 0.1)), 'digital': (1.8, ramp(1.8, 0.1))}, 'digital-low')
        for key in KEYS:
            for shape in result.summary_tables(key):
                self.assertEqual(result.sizes(shape), {10}, key)
        size = result.sizes(result.summary_tables('digital')[0]).pop()
        self.assertTrue(8 <= size < 10, size)
        self.assertEqual(result.findings(), [])

    def test_digital_never_goes_below_8_and_is_squeezed_with_a_notice(self):
        result = self.inputs.build({'digital': (3.0, ramp(3.0, 0.0))}, 'digital-squeeze')
        self.assertEqual(result.sizes(result.summary_tables('digital')[0]), {8})
        number = list(result.prs.slides).index(result.summary_slides()['digital']) + 1
        self.assertIn(f'ملخص التجربة الرقمية (الشريحة {number}): لم يتسع جدول «أبرز التحديثات» حتى بخط 8. تم ضغط صفوفه في المساحة المتاحة، '
                      'وقد لا يظهر نص الصفوف كاملا. المقترح: مراجعة هذه الشريحة واختصار التحديثات في ملف المراجعة.', result.space())
        for message in result.space():
            self.assertIsNone(re.search('[A-Za-z]', message), message)

    def test_explicit_floor_applies_to_the_digital_summary_too(self):
        result = self.inputs.build({'digital': (3.0, ramp(3.0, 0.0))}, 'digital-floor', extra=['--summary-min-pt', '9.5'])
        self.assertEqual(result.sizes(result.summary_tables('digital')[0]), {9.5})


def set_margins(shape, margin, first=2):
    for row in list(shape.table.rows)[first:]:
        for cell in row.cells:
            cell.margin_left = cell.margin_right = cell.margin_top = cell.margin_bottom = margin


def set_size(shape, pt):
    for row in shape.table.rows:
        for cell in row.cells:
            for paragraph in cell.text_frame.paragraphs:
                for run in paragraph.runs:
                    run.font.size = int(pt * PT)


def findings_after(result, change, name):
    """Validator findings for a copy of the built deck after `change(prs)` edited it."""
    prs = Presentation(result.path)
    change(prs, result)
    edited = result.path.with_name(name + '.pptx')
    prs.save(edited)
    inputs = result.inputs
    root = inputs.root
    extra = {}
    if inputs.digital:
        extra = dict(model='digital', digital_pdf=inputs.pdf, digital_data=inputs.review)
    return validate_report.validate(edited, root / 'tasks.xlsx', root / 'projects.xlsx', root / 'transactions.xlsx',
                                    template=inputs.template, strict=False, **extra)['warnings']


class ValidatorSharedLookTests(unittest.TestCase):
    """validate_report rejects a deck whose summaries do not share one look."""

    @classmethod
    def setUpClass(cls):
        cls.inputs = Inputs(with_digital=True)
        cls.fitting = cls.inputs.build({}, 'v-fit')
        cls.compact = cls.inputs.build({'suhail': (1.3, ramp(0.9, 0.0))}, 'v-compact')
        cls.smaller = cls.inputs.build({'tasks': (1.4, ramp(1.15, 0.1))}, 'v-smaller')

    @classmethod
    def tearDownClass(cls):
        cls.inputs.close()

    def test_the_untouched_decks_pass(self):
        for result in (self.fitting, self.compact, self.smaller):
            self.assertEqual(findings_after(result, lambda prs, r: None, 'same'), [])

    def test_compact_margins_on_a_detail_table_are_rejected(self):
        def change(prs, result):
            shape = helpers.detail_tables(helpers.detail_pattern_slides(prs, 'task')[0], 'task')[0]
            set_margins(shape, result.compact)
        found = findings_after(self.fitting, change, 'detail-compact')
        self.assertIn('Table cell margins changed', found)

    def test_a_summary_margin_that_is_neither_template_nor_compact_is_rejected(self):
        def change(prs, result):
            set_margins(result.summary_tables_in(prs, 'suhail')[0], 20000)
        found = findings_after(self.fitting, change, 'odd-margin')
        self.assertIn('Table cell margins changed', found)
        self.assertIn('Summary tables differ in cell margins', found)

    def test_summary_tables_that_disagree_on_margins_are_rejected(self):
        def change(prs, result):
            set_margins(result.summary_tables_in(prs, 'tasks')[0], result.compact)
        found = findings_after(self.fitting, change, 'disagree')
        self.assertIn('Summary tables differ in cell margins', found)
        # The digital summary is part of the comparison too.
        def digital_only(prs, result):
            set_margins(digital.summary_table_shape(digital.digital_slide(prs)), result.compact)
        self.assertIn('Summary tables differ in cell margins', findings_after(self.fitting, digital_only, 'disagree-digital'))

    def test_standard_summaries_at_different_sizes_are_rejected(self):
        def change(prs, result):
            for shape in result.summary_tables_in(prs, 'tasks'):
                set_size(shape, 10.5)
        found = findings_after(self.compact, change, 'sizes')
        self.assertIn('Standard summary tables use different font sizes', found)

    def test_a_smaller_font_with_the_template_margins_is_rejected(self):
        def change(prs, result):
            for key in KEYS:
                for shape in result.summary_tables_in(prs, key):
                    set_size(shape, 10)
        found = findings_after(self.fitting, change, 'small-template')
        self.assertTrue(any(f.startswith('Summary table below 11 pt keeps the template cell margins') for f in found), found)

    def test_a_digital_summary_larger_than_the_shared_size_is_rejected(self):
        self.assertEqual(Result.sizes(self.smaller.summary_tables('tasks')[0]), {10})
        def change(prs, result):
            set_size(digital.summary_table_shape(digital.digital_slide(prs)), 11)
        found = findings_after(self.smaller, change, 'digital-larger')
        self.assertIn('Digital summary table is larger than the shared summary size', found)


if __name__ == '__main__':
    unittest.main()
