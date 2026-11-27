# /// script
# requires-python = ">=3.12"
# dependencies = ["openpyxl>=3.1,<4", "python-pptx>=1.0.2", "lxml>=5.3", "pdfplumber>=0.11"]
# ///
"""Lack of space never stops the digital detail slides; what was done instead is a notice."""
import contextlib
import io
import sys
import unittest
from pathlib import Path

from pptx import Presentation

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / 'skills'))
sys.path.insert(0, str(ROOT / 'skills/weekly-consolidated-report/scripts'))
import digital_maturity as digital

PT = 12700


def sector(records, updates=1, challenges=1):
    rows = [['تجربة المستخدم: 50%', '', '', '']] + [['أساس', '50%', f'مهمة {i}', f'UX-{i:02d}'] for i in range(1, records + 1)]
    return {'title': 'قطاع تجريبي', 'cards': dict.fromkeys(digital.AXES, '50%'), 'rows': rows,
            'updates': [f'اكتمل البند {i}.' for i in range(updates)],
            'challenges': [f'بند {i}: لم يكتمل.' for i in range(challenges)]}


def measured(data, row_pt, note_pt):
    return {'sectors': {data['title']: {'rows': [row_pt] * len(data['rows']), 'updates': [note_pt] * len(data['updates']),
                                         'challenges': [note_pt] * len(data['challenges'])}}}


class DigitalSpaceTests(unittest.TestCase):
    def setUp(self):
        self.prs = Presentation(digital.MASTER)
        self.pattern = digital.DetailPattern(self.prs, digital.detail_slides(self.prs)[0])
        self.room = self.pattern.bottom - self.pattern.top
        digital.notices.reset()

    def paginate(self, data, row_pt, note_pt):
        with contextlib.redirect_stderr(io.StringIO()):
            return digital.paginate_sector(self.pattern, data, measured(data, row_pt, note_pt))

    def records(self, pages):
        return [row for rows, _, _ in pages for row in rows if not digital.is_continuation(row)]

    def test_normal_sector_has_no_notice(self):
        data = sector(4)
        pages = self.paginate(data, 30, 30)
        self.assertEqual(len(pages), 1)
        self.assertEqual(self.records(pages), data['rows'])
        self.assertEqual([note[4] for note in pages[-1][2]], [11, 11])
        self.assertEqual(digital.notices.collected(), [])

    def test_record_taller_than_a_slide_takes_its_own_slide(self):
        data = sector(3)
        huge = self.room / PT * 2
        heights = measured(data, 30, 30)
        heights['sectors'][data['title']]['rows'][2] = huge
        with contextlib.redirect_stderr(io.StringIO()):
            pages = digital.paginate_sector(self.pattern, data, heights)
        self.assertEqual(self.records(pages), data['rows'])
        for rows, row_heights, _ in pages:
            self.assertLessEqual(self.pattern.title_height + self.pattern.header_height + sum(row_heights), self.room)
        self.assertTrue(any(n['kind'] == 'space' and 'أطول من مساحة شريحة' in n['message'] for n in digital.notices.collected()))

    def test_notes_that_cannot_follow_the_last_record_shrink_instead_of_stopping(self):
        data = sector(3, updates=6, challenges=6)
        note_pt = self.room / PT / 11  # twelve note rows alone pass a slide at 11 pt
        pages = self.paginate(data, 30, note_pt)
        self.assertEqual(self.records(pages), data['rows'])
        notes = pages[-1][2]
        self.assertEqual([note[0] for note in notes], ['أبرز التحديثات', 'التحديات'])
        self.assertTrue(all(note[4] < 11 for note in notes))
        self.assertTrue(all(not page[2] for page in pages[:-1]))
        self.assertTrue(any(not digital.is_heading(row) for row in pages[-1][0]), 'the notes keep a record with them')
        self.assertTrue(any(n['kind'] == 'space' and 'صغرت خطهما' in n['message'] for n in digital.notices.collected()))

    def test_stale_sector_measurement_falls_back_to_the_estimate(self):
        data = sector(3)
        with contextlib.redirect_stderr(io.StringIO()):
            pages = digital.paginate_sector(self.pattern, data, {'sectors': {}})
        self.assertEqual(self.records(pages), data['rows'])
        self.assertTrue(any(n['kind'] == 'layout' for n in digital.notices.collected()))

    def test_fallback_builds_slides_that_reconcile(self):
        """The shrunk note tables are written to the slide and the reviewed rows and notes still reconcile."""
        data = sector(3, updates=6, challenges=6)
        note_pt = self.room / PT / 11
        sizes = []
        with contextlib.redirect_stderr(io.StringIO()):
            digital.populate_details(self.prs, {'detail_sectors': [data]}, measured(data, 30, note_pt), sizes)
        self.assertEqual(len(sizes), 2)
        errors, layout = [], []
        digital.check_details(self.prs, {'detail_sectors': [data]}, errors, self.prs and Presentation(digital.MASTER), layout)
        self.assertEqual(errors, [])


if __name__ == '__main__':
    unittest.main()
