# /// script
# requires-python = ">=3.12"
# dependencies = ["openpyxl>=3.1,<4", "python-pptx>=1.0.2", "lxml>=5.3", "pdfplumber>=0.11"]
# ///
"""Each summary stays on one slide (user request, 26-11-2026).

With measurements, summary tables that do not fit at 11 pt shrink on their own
slide (never a continuation slide), down to 8 pt; beyond that the build stops.
"""
import sys
import unittest
from pathlib import Path

from pptx import Presentation

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / 'skills'))
sys.path.insert(0, str(ROOT / 'skills/weekly-consolidated-report/scripts'))
import build_report as weekly
from weekly_data import scaled_height, SUMMARY_SIZES
from report_common.table_typography import apply_table, check_table_typography, set_table_sizes, set_table_directions

PT = 12700


def sections(slide, count):
    patterns = list(weekly.builder.tables(slide))
    return [(weekly.builder.table_title(p), [[str(i), 'مهمة', 'قطاع', 'ملاحظة'] for i in range(1, count + 1)], p)
            for p in patterns]


def measurements(prs, slide, rows, body_pt, smaller=None):
    bottom = weekly.content_bottom(prs, slide)
    out = {}
    for title, records, pattern in sections(slide, rows):
        entry = {'header': [h / PT for h in weekly.header_heights(pattern)], 'body': [body_pt] * len(records)}
        if smaller:
            entry['sizes'] = {k: [v] * len(records) for k, v in smaller.items()}
        out[title] = entry
    return out


class SummaryOneSlideTests(unittest.TestCase):
    def setUp(self):
        self.prs = Presentation(weekly.MASTER)
        self.slide = self.prs.slides[1]

    def test_overflow_shrinks_on_the_same_slide(self):
        before = len(self.prs.slides)
        data = sections(self.slide, 3)
        sizes = []
        targets = weekly.flow(self.prs, self.slide, data, True,
                              measurements(self.prs, self.slide, 3, 90, {'9': 25, '8': 20}), sizes=sizes)
        self.assertEqual(targets, [self.slide])
        self.assertEqual(len(self.prs.slides), before)
        tables = sorted(weekly.builder.tables(self.slide), key=lambda sh: sh.top)
        self.assertEqual(len(tables), len(data))
        self.assertLessEqual(tables[-1].top + tables[-1].height, weekly.content_bottom(self.prs, self.slide))
        self.assertTrue(all(a.top + a.height < b.top for a, b in zip(tables, tables[1:])))
        self.assertTrue(sizes)
        self.assertEqual({size for _, size in sizes}, {9})
        # Body rows use the heights measured at that size.
        self.assertEqual(tables[0].table.rows[2].height, 25 * PT)

    def test_fitting_summary_keeps_11_pt(self):
        sizes = []
        weekly.flow(self.prs, self.slide, sections(self.slide, 1), True,
                    measurements(self.prs, self.slide, 1, 30), sizes=sizes)
        self.assertEqual(sizes, [])

    def test_too_long_even_at_8_pt_stops(self):
        with self.assertRaisesRegex(ValueError, 'do not fit on one slide even at 8 pt'):
            weekly.flow(self.prs, self.slide, sections(self.slide, 4), True,
                        measurements(self.prs, self.slide, 4, 400, {str(s): 300 for s in SUMMARY_SIZES[1:]}), sizes=[])

    def test_provisional_pass_still_continues_for_measurement(self):
        targets = weekly.flow(self.prs, self.slide, sections(self.slide, 12), True)
        self.assertGreater(len(targets), 1)

    def test_scaled_height_bound(self):
        self.assertEqual(scaled_height(1000000, 11), 1000000)
        self.assertLess(scaled_height(1000000, 8), 1000000)
        self.assertGreater(scaled_height(1000000, 8), 1000000 * 8 / 11)

    def test_reduced_size_allowed_only_where_granted(self):
        set_table_sizes(self.prs);set_table_directions(self.prs)
        apply_table(weekly.builder.tables(self.slide)[0].table, 9.5)
        errors = []
        check_table_typography(self.prs, errors)
        self.assertTrue(any('11 pt' in e for e in errors))
        table = weekly.builder.tables(self.slide)[0].table
        errors = []
        check_table_typography(self.prs, errors, reduced={table._tbl: SUMMARY_SIZES})
        self.assertFalse(any('font size' in e for e in errors), errors)


if __name__ == '__main__':
    unittest.main()
