# /// script
# requires-python = ">=3.12"
# dependencies = []
# ///
"""Warn-and-continue notices and the wording checks shared by every skill."""
import contextlib
import datetime as dt
import io
import json
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / 'skills'))
from report_common import notices, wording


def quiet(function, *args, **kwargs):
    with contextlib.redirect_stderr(io.StringIO()):
        return function(*args, **kwargs)


class NoticeTests(unittest.TestCase):
    def setUp(self):
        notices.reset()

    def test_notices_are_kept_once_and_saved_next_to_the_output(self):
        with tempfile.TemporaryDirectory() as directory:
            output = Path(directory) / 'report.xlsx'
            quiet(notices.warn, 'ملاحظة أولى')
            quiet(notices.warn, 'ملاحظة أولى')
            quiet(notices.warn, 'ملاحظة ثانية', 'space')
            path = quiet(notices.save, output)
            self.assertEqual(path.name, 'report.xlsx.notices.json')
            self.assertEqual(json.loads(path.read_text(encoding='utf-8')),
                             [{'kind': 'assumption', 'message': 'ملاحظة أولى'}, {'kind': 'space', 'message': 'ملاحظة ثانية'}])
            # A later clean run removes the stale file.
            notices.reset()
            quiet(notices.save, output)
            self.assertFalse(path.exists())

    def test_unreadable_date_is_kept_as_written(self):
        self.assertEqual(notices.tolerant_date(dt.date(2026, 11, 26), 'الصف 2'), '26-11-2026')
        self.assertEqual(notices.collected(), [])
        self.assertEqual(quiet(notices.tolerant_date, ' قريبا ', 'الصف 3'), 'قريبا')
        found = notices.collected()
        self.assertEqual(found[0]['kind'], 'date')
        self.assertIn('الصف 3', found[0]['message'])
        text, key = quiet(notices.tolerant_date_parts, 'قريبا', 'الصف 4')
        self.assertEqual((text, key), ('قريبا', (9999, 99, 99)))
        self.assertEqual(notices.tolerant_date_parts('15/3/1448', 'الصف 5', hijri=True), ('15-03-1448', (1448, 3, 15)))
        self.assertEqual(notices.tolerant_date_parts(None, 'الصف 6'), ('', (9999, 99, 99)))

    def test_unknown_status_is_placed_by_wording_or_default(self):
        known = {s: s for s in ('مكتملة', 'على المخطط', 'متأخر', 'معلق')}
        self.assertEqual(quiet(notices.guess_status, 'متأخر جدا', known, 'على المخطط', 'الصف 2'), 'متأخر')
        self.assertEqual(quiet(notices.guess_status, 'تم الانتهاء، منجز', known, 'على المخطط', 'الصف 3'), 'مكتملة')
        self.assertEqual(quiet(notices.guess_status, 'قيد المراجعة', known, 'على المخطط', 'الصف 4'), 'على المخطط')
        # A category this report does not have falls to the default, and labels follow the report.
        transactions = {'مكتملة': 'مكتملة', 'على المخطط': 'على المخطط', 'متأخر': 'متأخرة'}
        self.assertEqual(quiet(notices.guess_status, 'معلق', transactions, 'على المخطط', 'الصف 5'), 'على المخطط')
        self.assertEqual(quiet(notices.guess_status, 'متاخره', transactions, 'على المخطط', 'الصف 6'), 'متأخرة')
        self.assertEqual([n['kind'] for n in notices.collected()], ['status'] * 5)
        self.assertTrue(all('الصف' in n['message'] for n in notices.collected()))


class WordingTests(unittest.TestCase):
    def test_machine_sounding_patterns_are_found(self):
        for text in ('وتعكس النسب تقدم العمل.', 'يلعب المركز دورا محوريا.', 'علاوة على ذلك اكتمل الربط.', 'اكتمل الربط، مما يسهم في رفع النسبة.',
                     'تجدر الإشارة إلى اكتمال الربط.', 'يُعدّ المشروع من أبرز المشاريع.', 'رغم التقدم إلا أنه يواجه تحديات.'):
            with self.subTest(text=text):
                self.assertTrue(wording.issues(text), text)

    def test_formal_government_sentences_pass(self):
        for text in ('اكتمل نقل محتوى نسك عمرة، ويبقى موعد الإطلاق 15 ديسمبر.',
                     'تأخر الاعتماد: لم يصدر قرار توحيد مركز الاتصال حتى تاريخه.',
                     'أرفع لمعاليكم ملخص تقرير الإنجاز الأسبوعي المحدث بتاريخ 26 نوفمبر 2026م، ويشمل المهام ومشاريع سهيل والمعاملات.',
                     'تطوير رحلة المستفيد في منصة نسك.', 'يجري العمل على توحيد المحتوى.',
                     'بلغت نسبة الإنجاز لدى وكالة التحول الرقمي 83%، ولدى مركز منصة نسك الرقمية 74%.'):
            with self.subTest(text=text):
                self.assertEqual(wording.review(text, 'x'), [])

    def test_colloquial_markup_and_latin(self):
        self.assertEqual(wording.colloquial('شغالين على الملف الحين'), ['الحين', 'شغالين'])
        self.assertEqual(wording.colloquial('يُراجَع الملف'), [])
        self.assertTrue(wording.markup('**مهم** — عاجل'))
        self.assertEqual(wording.markup('اكتمل الربط (المرحلة الأولى).'), [])
        self.assertEqual(wording.latin_words('إطلاق nusuk.sa وعقد Contentsquare'), ['nusuk.sa', 'Contentsquare'])

    def test_new_banned_phrases_are_found(self):
        for text in ('بكل سرور.', 'بالتأكيد! ابدأنا.', 'سؤال رائع عن المهام.', 'أتمنى أن يكون هذا مفيداً.',
                     'أتمنى أن يكون هذا مفيدا.', 'لا تتردد بالسؤال.', 'إليك الملخص.', 'تم الربط. هل تريد تفاصيل أكثر؟',
                     'نعتذر عن أي إزعاج.', 'يرى الخبراء أن النسبة جيدة.', 'يرى المختصون ذلك.', 'يشير المراقبون إلى ذلك.',
                     'التقرير ليس مجرد جدول بل أداة.', 'يُعدّ المشروع مهما.', 'يعد المشروع من أبرز المشاريع.',
                     'تعد الوكالة نموذجا يحتذى.', 'بالإضافة إلى ذلك اكتمل الربط.', 'بالإضافة إلى المهام، اكتمل الربط.'):
            with self.subTest(text=text):
                self.assertTrue(wording.issues(text), text)

    def test_allowed_phrases_are_not_flagged(self):
        for text in ('يعد الفريق التقرير كل أسبوع.', 'اكتملت المهام بالإضافة إلى المشاريع.',
                     'وصلت إليكم الملفات.', 'تتواصل أعمال رفع المؤشر.', 'تبرز الحاجة إلى التوجيه.',
                     'هل تريد الوكالة التقرير في المساء؟ نعم.'):
            with self.subTest(text=text):
                self.assertEqual(wording.issues(text), [], text)


class WordingFalsePositiveTests(unittest.TestCase):
    def test_ordinary_formal_arabic_is_not_flagged(self):
        for text in ('في ظل غياب القرار تأخر الاعتماد.', 'رغم اكتمال الربط إلا أن الاعتماد لم يصدر.',
                     'رغم تأخر التسليم اكتملت المرحلة الأولى.'):
            with self.subTest(text=text):
                self.assertEqual(wording.issues(text), [], text)

    def test_puffery_and_new_phrases_are_flagged(self):
        for text in ('في ظل المشهد المتغير اكتمل الربط.', 'في ظل التطورات المتسارعة تم الربط.',
                     'رغم التقدم إلا أنه يواجه تحديات.', 'رغم التقدم إلا أنها تواجه تحديات كبيرة.',
                     'رغم الجهود لم يكتمل الربط.', 'ليس فقط الربط بل الاعتماد.',
                     'يشار إلى أن الربط اكتمل.', 'يُشار إلى أن الربط اكتمل.', 'ختاما اكتمل الربط.', 'وفي الختام اكتمل الربط.'):
            with self.subTest(text=text):
                self.assertTrue(wording.issues(text), text)


class EndedStatusTests(unittest.TestCase):
    KNOWN = {'مكتملة': 'مكتملة', 'على المخطط': 'على المخطط', 'متأخر': 'متأخر', 'معلق': 'معلق'}

    def guess(self, text):
        notices.reset()
        return quiet(notices.guess_status, text, self.KNOWN, 'على المخطط', 'x')

    def test_whole_word_forms_are_completed(self):
        for text in ('منتهي', 'منتهية', 'منته', 'منتهية بنجاح'):
            with self.subTest(text=text):
                self.assertEqual(self.guess(text), 'مكتملة')

    def test_expired_is_not_completed(self):
        for text in ('منتهي المدة', 'منتهية الصلاحية', 'منتهي الصلاحية'):
            with self.subTest(text=text):
                self.assertEqual(self.guess(text), 'على المخطط')
                self.assertIn('افتراضا', notices.collected()[0]['message'])


class StatusNegationTests(unittest.TestCase):
    KNOWN = {'مكتملة': 'مكتملة', 'على المخطط': 'على المخطط', 'متأخر': 'متأخرة', 'معلق': 'معلق', 'لم تبدأ': 'لم تبدأ'}

    def setUp(self):
        notices.reset()

    def guess(self, text):
        with contextlib.redirect_stderr(io.StringIO()):
            return notices.guess_status(text, self.KNOWN, 'على المخطط', 'x')

    def test_negated_statuses_use_the_default_with_a_notice(self):
        for text in ('غير مكتمل', 'غير منتهي', 'غير متأخر', 'لم ينجز', 'ليس متأخرا', 'بدون إنجاز مكتمل'):
            with self.subTest(text=text):
                notices.reset()
                self.assertEqual(self.guess(text), 'على المخطط')
                self.assertEqual(len(notices.collected()), 1)
                self.assertIn('افتراضا', notices.collected()[0]['message'])

    def test_plain_stems_still_match_and_stems_match_at_word_start_only(self):
        self.assertEqual(self.guess('مكتمل جزئيا'), 'مكتملة')
        self.assertEqual(self.guess('متأخر عن الخطة'), 'متأخرة')
        self.assertEqual(self.guess('لم يبدأ بعد'), 'لم تبدأ')
        notices.reset()
        self.assertEqual(self.guess('أعمال المجاري'), 'على المخطط')
        self.assertIn('افتراضا', notices.collected()[0]['message'])
        self.assertEqual(self.guess('جاري التنفيذ'), 'على المخطط')
        self.assertIn('لتشابه', notices.collected()[-1]['message'])


class ClearStaleTests(unittest.TestCase):
    def test_clear_stale_removes_an_old_notices_file(self):
        with tempfile.TemporaryDirectory() as tmp:
            out = Path(tmp) / 'r.xlsx'
            old = notices.notices_path(out)
            old.write_text('[]', encoding='utf-8')
            notices.clear_stale(out)
            self.assertFalse(old.exists())
            notices.clear_stale(out)  # nothing left: no error


if __name__ == '__main__':
    unittest.main()
