# /// script
# requires-python = ">=3.12"
# dependencies = ["pandas>=2.2,<3", "openpyxl>=3.1,<4", "python-pptx>=1.0.2", "lxml>=5.3", "pillow", "pdfplumber>=0.11", "xlrd>=2"]
# ///
"""The two Minister messages come from the delivered report (user request, 26-11-2026)."""
import contextlib
import io
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / 'skills'))
sys.path.insert(0, str(ROOT / 'skills/weekly-consolidated-report/scripts'))
import minister_messages as messages


class AgreementTests(unittest.TestCase):
    def test_number_noun_agreement(self):
        cases = [((1, 'task', 'مكتملة'), 'مهمة واحدة مكتملة'), ((2, 'task', 'مكتملة'), 'مهمتان مكتملتان'),
                 ((3, 'task', 'متأخرة'), '3 مهام متأخرة'), ((15, 'task', 'على المخطط'), '15 مهمة على المخطط'),
                 ((7, 'project', 'مكتملة'), '7 مشاريع مكتملة'), ((12, 'project', 'على المخطط'), '12 مشروعا على المخطط'),
                 ((2, 'project', 'لم تبدأ'), 'مشروعان لم يبدآ'), ((19, 'project', None), '19 مشروعا'),
                 ((1032, 'transaction', None), '1032 معاملة'), ((105, 'transaction', None), '105 معاملات')]
        for args, expected in cases:
            self.assertEqual(messages.counted(*args), expected)

    def test_counts_sentences_match_the_approved_wording(self):
        self.assertEqual(messages.counts_sentence('task', 17, [('مكتملة', 2), ('على المخطط', 15), ('متأخرة', 0), ('معلقة', 0)]),
                         'بلغ إجمالي المهام 17 مهمة، منها مهمتان مكتملتان و15 مهمة على المخطط، ولا توجد مهام متأخرة أو معلقة ضمن جدول متابعة المهام.')
        self.assertEqual(messages.counts_sentence('project', 19, [('مكتملة', 7), ('على المخطط', 12), ('متأخرة', 0), ('لم تبدأ', 0)]),
                         'بلغ إجمالي مشاريع سهيل 19 مشروعا، منها 7 مشاريع مكتملة و12 مشروعا على المخطط، ولا توجد مشاريع متأخرة أو لم تبدأ.')
        self.assertEqual(messages.counts_sentence('transaction', 1032, [('مكتملة', 676), ('على المخطط', 356), ('متأخرة', 0)]),
                         'بلغ إجمالي المعاملات 1032 معاملة، منها 676 معاملة مكتملة و356 معاملة على المخطط، ولا توجد معاملات متأخرة.')
        self.assertEqual(messages.counts_sentence('task', 4, [('مكتملة', 1), ('على المخطط', 1), ('متأخرة', 1), ('معلقة', 1)]),
                         'بلغ إجمالي المهام 4 مهام، منها مهمة واحدة مكتملة ومهمة واحدة على المخطط ومهمة واحدة متأخرة ومهمة واحدة معلقة.')
        self.assertEqual(messages.counts_sentence('task', 0, []), 'لم يتضمن التقرير مهام.')


class ReportMessageTests(unittest.TestCase):
    def build(self, name):
        import run_weekly_workflow as workflow
        root = Path(self.temp.name) / name
        with contextlib.redirect_stdout(io.StringIO()):
            workflow.scenario(root, name)
        return root / 'weekly.pptx'

    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)

    def test_messages_follow_the_report(self):
        deck = self.build('normal')
        data = messages.facts(deck)
        short = messages.short_message(data).splitlines()
        self.assertEqual((short[0], short[-1]), ('معالي الوزير حفظه الله', 'مع التحية والتقدير'))
        self.assertEqual(len(short), 5, 'standard model: no digital line')
        self.assertIn('عدا', short[1])  # the scenario has a delayed and a suspended task
        long = messages.long_draft(data)
        t = data['tasks']
        self.assertIn(f'بلغ إجمالي المهام {messages.counted(t["total"], "task")}', long)
        for heading in ('أولا: المهام', 'ثانيا: مشاريع سهيل', 'ثالثا: المعاملات', 'رابعا: طلبات الدعم المرفوعة لمعاليكم'):
            self.assertIn(heading, long)
        self.assertNotIn('مؤشر نضج التجربة الرقمية', long)
        self.assertIn('المهام المتأخرة:', long)
        self.assertIn('طلب دعم تجريبي', long)
        self.assertTrue(long.startswith('معالي الوزير حفظه الله\n\nالسلام عليكم'))
        self.assertTrue(long.rstrip().endswith('وتفضلوا بقبول خالص التحية والتقدير.'))
        self.assertIn(data['date'], long)

    def test_empty_sections_get_a_formal_sentence(self):
        long = messages.long_draft(messages.facts(self.build('empty')))
        for key in ('completed_tasks', 'suhail_updates', 'suhail_challenges'):
            self.assertIn(messages.EMPTY[key], long)
        self.assertIn(messages.EMPTY['support'], long)
        self.assertIn('لم يتضمن التقرير مهام.', long)


class DigitalMessageTests(unittest.TestCase):
    def test_digital_facts_and_sections(self):
        import json
        import hashlib
        from pptx import Presentation
        import digital_maturity as digital
        from test_weekly_digital_maturity import master_sectors
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            pdf = root / 'source.pdf'; pdf.write_bytes(b'synthetic source identity')
            sectors = master_sectors()
            sectors[0]['challenges'] = ['تأخر التقدم: لم يسجل أي تقدم في الملف حتى الآن.']
            review = root / 'review.json'
            review.write_text(json.dumps({
                'source_pdf_sha256': hashlib.sha256(pdf.read_bytes()).hexdigest(), 'reviewed': True,
                'metrics': {'experience': 66, 'satisfaction': 45, 'complaints': 67, 'tools': 77},
                'sectors': [{'name': 'القطاع الأول', 'note': None, 'metrics': {'experience': 83, 'satisfaction': 50, 'complaints': None, 'tools': 79}},
                            {'name': 'القطاع الثاني', 'note': 'قيد الدراسة', 'metrics': dict.fromkeys(digital.AXES)}],
                'updates': [{'perspective': 'تجربة المستخدم', 'text': 'اكتمل متطلب توافرية المنصات.'}],
                'detail_sectors': sectors}, ensure_ascii=False), encoding='utf-8')
            prs = Presentation(digital.MASTER)
            digital.populate(prs, digital.load_reviewed(review, pdf))
            deck = root / 'digital.pptx'; prs.save(deck)
            data = messages.facts(deck)
        d = data['digital']
        self.assertEqual(d['metrics'], {'experience': '66%', 'satisfaction': '45%', 'complaints': '67%', 'tools': '77%'})
        self.assertEqual([s['name'] for s in d['sectors']], ['القطاع الأول', 'القطاع الثاني'])
        first = d['sectors_detail'][sectors[0]['title']]
        body = [r for r in sectors[0]['rows'] if not digital.is_heading(r)]
        self.assertEqual(sum(len(v) for v in first['rows'].values()), len(body))
        short = messages.short_message(data)
        self.assertIn(f'مع تأخر لدى {sectors[0]["title"]}', short)
        # Only the fact: no invented judgement and no placeholder, and the generated message passes --check.
        self.assertNotIn('التصعيد', short)
        self.assertNotIn('…', short)
        self.assertEqual(messages.check_message(short, data), [])
        self.assertEqual(sum('مؤشر نضج التجربة الرقمية' in line for line in short.splitlines()), 1)
        long = messages.long_draft(data)
        self.assertIn(messages.DISCLAIMER, long)
        self.assertIn('خامسا: طلبات الدعم المرفوعة لمعاليكم', long)
        for number, label in enumerate(digital.LABELS, 1):
            self.assertIn(f'{number}. منظور {label} (نسبة الإنجاز', long)
        self.assertIn('بلغت نسبة الإنجاز لدى القطاع الأول 83%', long)
        self.assertIn('القطاع الثاني: قيد الدراسة', long)
        self.assertIn(messages.EMPTY['perspective_sectors'], long)  # no sector has a complaints value
        # Every fixed phrase the script writes is free of machine-sounding patterns, colloquial words and dashes.
        from report_common import wording
        for text in (short, long):
            self.assertEqual([], [f for line in text.splitlines() for f in wording.review(line, 'x')])
        # The draft is not a finished message: its brackets are reported, and so is an invented number.
        findings = messages.check_message(long, data)
        self.assertTrue(any('المسودة' in f for f in findings))
        ready = 'بلغت نسبة الإنجاز لدى القطاع الأول 83%.\n1. منظور تجربة المستخدم (نسبة الإنجاز 66%)'
        self.assertEqual(messages.check_message(ready, data), [])
        self.assertTrue(any('91' in f for f in messages.check_message('بلغت النسبة 91%.', data)))
        self.assertTrue(any('عامية' in f for f in messages.check_message('شغالين على الملف.', data)))
        self.assertTrue(any('آلية' in f for f in messages.check_message('وتعكس النسب تقدم العمل.', data)))


if __name__ == '__main__':
    unittest.main()
