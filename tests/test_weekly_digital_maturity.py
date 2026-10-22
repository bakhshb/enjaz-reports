"""Digital weekly model keeps reviewed PDF values and rejects unsafe inputs."""
import hashlib
import json
import sys
import tempfile
import unittest
from pathlib import Path

from pptx import Presentation

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / 'skills/weekly-consolidated-report/scripts'))
import digital_maturity as digital
import build_report as weekly_build


class DigitalMaturityTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        self.pdf = self.root / 'source.pdf'
        self.pdf.write_bytes(b'synthetic source identity')
        self.data = {
            'source_pdf_sha256': hashlib.sha256(self.pdf.read_bytes()).hexdigest(),
            'source_date_hijri': '1448/1/1',
            'reviewed': True,
            'metrics': {'experience': 15, 'satisfaction': 10, 'complaints': 20, 'tools': 65},
            'sectors': [
                {'name': 'القطاع الأول', 'metrics': {'experience': 15, 'satisfaction': 10, 'complaints': 20, 'tools': 65}, 'note': None},
                {'name': 'القطاع الثاني', 'metrics': {'experience': None, 'satisfaction': None, 'complaints': None, 'tools': None}, 'note': 'قيد الدراسة'},
                {'name': 'القطاع الثالث', 'metrics': {'experience': 25, 'satisfaction': None, 'complaints': None, 'tools': 70}, 'note': None},
                {'name': 'القطاع الرابع', 'metrics': {'experience': 5, 'satisfaction': 15, 'complaints': 30, 'tools': None}, 'note': None},
                {'name': 'القطاع الخامس', 'metrics': {'experience': 10, 'satisfaction': 20, 'complaints': 25, 'tools': 75}, 'note': None},
            ],
            'challenges': [{'perspective': 'عام', 'text': 'تحد أول'}, {'perspective': 'تجربة المستخدم', 'text': 'تحد ثان'}],
        }
        master = Presentation(digital.MASTER)
        self.data['detail_pages'] = []
        for number, slide in enumerate(digital.detail_slides(master), 2):
            table = digital._detail_table(slide)
            challenge_table = digital._challenge_table(slide)
            self.data['detail_pages'].append({
                'page': number,
                'title': table.cell(0, 0).text,
                'cards': {axis: digital.helpers.shape_by_id(slide, shape_id).text_frame.paragraphs[1].text
                          for axis, shape_id in digital.DETAIL_CARD_IDS.items()},
                'rows': [[cell.text for cell in row.cells] for row in list(table.rows)[2:]],
                'challenges': [] if challenge_table is None else
                [row.cells[1].text for row in list(challenge_table.rows)[2:]],
            })
        self.review = self.root / 'review.json'

    def load(self):
        self.review.write_text(json.dumps(self.data, ensure_ascii=False), encoding='utf-8')
        return digital.load_reviewed(self.review, self.pdf)

    def test_review_and_pdf_identity_are_required(self):
        self.data['reviewed'] = False
        with self.assertRaisesRegex(ValueError, 'Review'):
            self.load()
        self.data['reviewed'] = True
        self.load()
        self.pdf.write_bytes(b'a different PDF')
        with self.assertRaisesRegex(ValueError, 'does not match'):
            digital.load_reviewed(self.review, self.pdf)

    def test_capacity_and_conflicting_note_rejected(self):
        self.data['challenges'] *= 3
        with self.assertRaisesRegex(ValueError, 'exceed'):
            self.load()
        self.data['challenges'] = self.data['challenges'][:2]
        self.data['sectors'][1]['metrics']['experience'] = 5
        with self.assertRaisesRegex(ValueError, 'cannot have percentages'):
            self.load()

    def test_detail_capacity_and_missing_page_rejected(self):
        last = self.data['detail_pages'][-1]
        self.data['detail_pages'] = self.data['detail_pages'][:-1]
        with self.assertRaisesRegex(ValueError, 'Six reviewed'):
            self.load()
        self.data['detail_pages'].append(last)
        self.data['detail_pages'][0]['rows'].append(['جاهزية', '15%', 'مهمة إضافية', 'X-99'])
        data = self.load()
        with self.assertRaisesRegex(ValueError, 'exceeds its detail table capacity'):
            digital.populate(Presentation(digital.MASTER), data)

    def test_charts_and_challenges_match_reviewed_data(self):
        data = self.load()
        prs = Presentation(digital.MASTER)
        digital.populate(prs, data)
        report = self.root / 'report.pptx'
        prs.save(report)
        errors = []
        digital.check(Presentation(report), report, data, errors)
        self.assertEqual(errors, [])
        prs = Presentation(report)
        card = digital.helpers.shape_by_id(digital.digital_slide(prs), digital.CARDS['experience'][1])
        card.text = '99%'
        prs.save(report)
        errors = []
        digital.check(Presentation(report), report, data, errors)
        self.assertTrue(any('card label differs' in error for error in errors))
        card.text = '15%'
        sector = digital.helpers.shape_by_id(digital.digital_slide(prs), 93)
        sector.chart.series[0].data_labels.number_format = 'General'
        prs.save(report)
        errors = []
        digital.check(Presentation(report), report, data, errors)
        self.assertTrue(any('must display percentages' in error for error in errors))

    def test_page_six_challenges_and_latin_codes(self):
        self.data['detail_pages'][4]['challenges'] = [
            'تحد مرتبط بالمهمة W-03', 'تحد ثان مرتبط بالمهمة S-02'
        ]
        self.data['detail_pages'][4]['rows'][1][3] = 'W-01'
        data = self.load()
        prs = Presentation(digital.MASTER)
        digital.populate(prs, data)
        slide = digital.detail_slides(prs)[4]
        challenges = digital._challenge_table(slide)
        self.assertEqual([row.cells[1].text for row in list(challenges.rows)[2:]], data['detail_pages'][4]['challenges'])
        code = digital._detail_table(slide).cell(3, 3)
        self.assertEqual(digital._plain(code.text), 'W-01')
        self.assertEqual(code.text_frame.paragraphs[0]._p.get_or_add_pPr().get('rtl'), '0')

    def test_summary_tables_include_auto_height_headers_in_spacing(self):
        prs = Presentation(digital.MASTER)
        slide = prs.slides[1]
        patterns = list(weekly_build.builder.tables(slide))
        sections = [(weekly_build.builder.table_title(pattern),
                     [['1', 'مهمة تجريبية', 'قطاع تجريبي', 'ملاحظة تجريبية']], pattern)
                    for pattern in patterns]
        weekly_build.flow(prs, slide, sections, summary=True)
        tables = sorted(weekly_build.builder.tables(slide), key=lambda shape: shape.top)
        self.assertEqual(len(tables), len(patterns))
        self.assertTrue(all(table.table.rows[0].height > 0 and table.table.rows[1].height > 0
                            for table in tables))
        self.assertTrue(all(first.top + first.height < second.top
                            for first, second in zip(tables, tables[1:])))


if __name__ == '__main__':
    unittest.main()
