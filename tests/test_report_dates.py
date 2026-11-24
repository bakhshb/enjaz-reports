# /// script
# requires-python = ">=3.12"
# dependencies = ["pandas>=2.2,<3", "openpyxl>=3.1,<4", "python-pptx>=1.0.2", "lxml>=5.3"]
# ///
"""Every report date comes out as dd-mm-yyyy whatever form the export used."""

import datetime as dt
import re
import sys
import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook, load_workbook

sys.path.insert(0, str(Path(__file__).resolve().parent))
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'skills'))
from report_common.dates import DateError, report_date
from test_upstream_integrity import SUHAIL, TASKS, TRANSACTIONS, run, suhail_source, task_source, transaction_source

DD_MM_YYYY = re.compile(r'\d{2}-\d{2}-\d{4}')

# Raw value as it may appear in an export -> the report date.
MIXED = [
    (dt.datetime(2026, 12, 11), '11-12-2026'),
    (dt.date(2026, 3, 5), '05-03-2026'),
    ('11-12-2026', '11-12-2026'),
    ('5/3/2026', '05-03-2026'),
    ('05.03.2026', '05-03-2026'),
    ('2026-12-11', '11-12-2026'),
    ('2026-12-11 00:00:00', '11-12-2026'),
    ('11/12/2026 00:00:00', '11-12-2026'),
    (46367, '11-12-2026'),
    (46367.0, '11-12-2026'),
    ('11 ديسمبر 2026', '11-12-2026'),
    ('٥ مارس ٢٠٢٦م', '05-03-2026'),
    ('١١/١٢/٢٠٢٦', '11-12-2026'),
    ('‏11-12-2026‏', '11-12-2026'),
]
UNREADABLE = ['15/03/1448', '1448-03-15', '15 رمضان 1447', 'قريبا', '31/02/2026', 12]


class SharedDateTests(unittest.TestCase):
    def test_mixed_formats_become_dd_mm_yyyy(self):
        for raw, expected in MIXED:
            with self.subTest(raw=raw):
                self.assertEqual(report_date(raw), expected)

    def test_empty_stays_empty(self):
        for raw in (None, '', '  ', float('nan')):
            self.assertEqual(report_date(raw), '')

    def test_unreadable_values_stop(self):
        for raw in UNREADABLE:
            with self.subTest(raw=raw), self.assertRaises(DateError):
                report_date(raw)

    def test_hijri_allowed_only_when_asked(self):
        self.assertEqual(report_date('15/3/1448', hijri=True), '15-03-1448')
        self.assertEqual(report_date('1448/05/17هـ', hijri=True), '17-05-1448')
        with self.assertRaises(DateError):
            report_date('35/03/1448', hijri=True)


class BuilderDateTests(unittest.TestCase):
    raws = [raw for raw, _ in MIXED]

    def assert_all_report_dates(self, values, count):
        self.assertEqual(len(values), count)
        for value in values:
            self.assertRegex(value, DD_MM_YYYY)
        self.assertEqual(sorted(values), sorted(expected for _, expected in MIXED))

    def test_suhail_project_dates(self):
        with tempfile.TemporaryDirectory() as td:
            source, output = Path(td) / 'source.xlsx', Path(td) / 'out.xlsx'
            suhail_source(source, 'على المخطط')
            wb = load_workbook(source); ws = wb.active
            template = [c.value for c in ws[2]]; ws.delete_rows(2)
            for i, raw in enumerate(self.raws):
                ws.append(template[:1] + [f'مشروع {i}', raw, raw] + template[4:])
            wb.save(source)
            result = run(SUHAIL, source, output)
            self.assertEqual(result.returncode, 0, result.stderr + result.stdout)
            detail = load_workbook(output)['تفاصيل مشاريع سهيل']
            rows = [r for r in detail.iter_rows(values_only=True) if r[1] and str(r[1]).startswith('مشروع ')]
            self.assert_all_report_dates([r[2] for r in rows], len(MIXED))
            self.assert_all_report_dates([r[3] for r in rows], len(MIXED))

    def test_suhail_stops_on_unreadable_date(self):
        with tempfile.TemporaryDirectory() as td:
            source, output = Path(td) / 'source.xlsx', Path(td) / 'out.xlsx'
            suhail_source(source, 'على المخطط')
            wb = load_workbook(source); wb.active.cell(2, 4).value = '15/03/1448'; wb.save(source)
            result = run(SUHAIL, source, output)
            self.assertNotEqual(result.returncode, 0)
            self.assertIn('row 2', result.stderr)
            self.assertIn('15/03/1448', result.stderr)
            self.assertFalse(output.exists())

    def test_task_planned_dates(self):
        with tempfile.TemporaryDirectory() as td:
            source, output = Path(td) / 'source.xlsx', Path(td) / 'out.xlsx'
            task_source(source, support=False)
            wb = load_workbook(source); ws = wb['المهام']
            template = [c.value for c in ws[2]]; ws.delete_rows(2)
            for i, raw in enumerate(self.raws):
                ws.append([f'مهمة {i}'] + template[1:3] + [raw] + template[4:])
            wb.save(source)
            result = run(TASKS, '--input', source, '--output', output)
            self.assertEqual(result.returncode, 0, result.stderr + result.stdout)
            detail = load_workbook(output)['تفاصيل المهام']
            dates = [r[3] for r in detail.iter_rows(values_only=True) if r[1] and str(r[1]).startswith('مهمة ')]
            self.assert_all_report_dates(dates, len(MIXED))

    def test_task_stops_on_unreadable_date(self):
        with tempfile.TemporaryDirectory() as td:
            source, output = Path(td) / 'source.xlsx', Path(td) / 'out.xlsx'
            task_source(source, support=False)
            wb = load_workbook(source); ws = wb['المهام']
            ws.append(['مهمة ثانية', 'وكالة شؤون الحج', 'اجتماع القيادات', 'قريبا', 'مكتمل', ''])
            wb.save(source)
            result = run(TASKS, '--input', source, '--output', output)
            self.assertNotEqual(result.returncode, 0)
            self.assertIn('3', result.stderr)
            self.assertIn('قريبا', result.stderr)
            self.assertFalse(output.exists())

    def test_transaction_dates(self):
        hijri = [('15/3/1448', '15-03-1448'), ('1448-03-16', '16-03-1448')]
        with tempfile.TemporaryDirectory() as td:
            input_dir, output = Path(td) / 'input', Path(td) / 'out.xlsx'
            input_dir.mkdir()
            transaction_source(input_dir / 'one.xlsx')
            wb = load_workbook(input_dir / 'one.xlsx'); ws = wb.active
            template = [c.value for c in ws[5]]; ws.delete_rows(5)
            raws = self.raws + [raw for raw, _ in hijri]
            for i, raw in enumerate(raws):
                row = list(template); row[23] = raw; row[27] = raw; row[30] = f'{1000 + i}'
                ws.append(row)
            wb.save(input_dir / 'one.xlsx')
            result = run(TRANSACTIONS, '--input-dir', input_dir, '--output', output)
            self.assertEqual(result.returncode, 0, result.stderr + result.stdout)
            ws = load_workbook(output).active
            rows = [r for r in ws.iter_rows(values_only=True) if r[0] and str(r[0]).isdigit() and len(str(r[0])) == 4]
            self.assertEqual(len(rows), len(raws))
            expected = sorted([e for _, e in MIXED] + [e for _, e in hijri])
            for column in (4, 5):
                values = [r[column] for r in rows]
                for value in values:
                    self.assertRegex(value, DD_MM_YYYY)
                self.assertEqual(sorted(values), expected)

    def test_transaction_stops_on_unreadable_date(self):
        with tempfile.TemporaryDirectory() as td:
            input_dir, output = Path(td) / 'input', Path(td) / 'out.xlsx'
            input_dir.mkdir()
            transaction_source(input_dir / 'one.xlsx')
            wb = load_workbook(input_dir / 'one.xlsx'); wb.active.cell(5, 24).value = 'قريبا'; wb.save(input_dir / 'one.xlsx')
            result = run(TRANSACTIONS, '--input-dir', input_dir, '--output', output)
            self.assertNotEqual(result.returncode, 0)
            self.assertIn('one.xlsx row 5', result.stderr)
            self.assertIn('قريبا', result.stderr)
            self.assertFalse(output.exists())


class SectorTransactionMatchTests(unittest.TestCase):
    """The sector splitter joins deck transactions to FINAL.xlsx by date too."""

    def test_real_excel_date_matches_deck_text(self):
        sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'skills/weekly-report-by-sector/scripts'))
        import split_report as splitter
        head = ['#', 'رقم المعاملة', 'موضوع المعاملة', 'الجهة الوارد منها المعاملة',
                'تاريخ إنشاء المعاملة', 'تاريخ الإنجاز المخطط له']
        with tempfile.TemporaryDirectory() as td:
            final = Path(td) / 'FINAL.xlsx'
            wb = Workbook(); ws = wb.active; ws.title = 'المعاملات'
            ws.append(['رقم المعاملة', 'موضوع المعاملة', 'الجهة الوارد منها المعاملة',
                       'تاريخ انشاء المعاملة', 'تاريخ الإنجاز المخطط له', 'القطاع'])
            ws.append([' 4471 ', 'طلب تصريح', 'إمارة منطقة مكة', dt.datetime(2026, 3, 5),
                       dt.datetime(2026, 12, 11), 'وكالة شؤون الحج'])
            ws.append(['4472', 'طلب آخر', 'جهة أخرى', '15/3/1448', '16/3/1448', 'وكالة شؤون العمرة'])
            wb.save(final)
            self.assertIsInstance(load_workbook(final).active['D2'].value, dt.datetime)
            trans = [
                {'head': head, 'row': ['1', '4471', 'طلب تصريح', 'إمارة منطقة مكة', '05-03-2026', '11-12-2026']},
                {'head': head, 'row': ['2', '4472', 'طلب آخر', 'جهة أخرى', '15-03-1448', '16-03-1448']},
            ]
            splitter.attach_transaction_sectors(trans, str(final))
            self.assertEqual([t['sector'] for t in trans], ['وكالة شؤون الحج', 'وكالة شؤون العمرة'])

            trans = [{'head': head, 'row': ['1', '4471', 'طلب تصريح', 'إمارة منطقة مكة', '06-03-2026', '11-12-2026']}]
            with self.assertRaisesRegex(ValueError, 'does not match'):
                splitter.attach_transaction_sectors(trans, str(final))


if __name__ == '__main__':
    unittest.main()
