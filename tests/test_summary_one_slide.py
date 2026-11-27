# /// script
# requires-python = ">=3.12"
# dependencies = ["openpyxl>=3.1,<4", "python-pptx>=1.0.2", "lxml>=5.3", "pdfplumber>=0.11"]
# ///
"""Each summary stays on one slide (user request, 26-11-2026).

With measurements, summary tables that do not fit at 11 pt with the template
margins take the compact margins and shrink on their own slide, down to 10 pt
(8 pt for the digital summary only; see test_summary_shared_look.py for the
rule shared by all four summaries, 27-11-2026). Beyond the floor the build does
not stop: the tables that do not fit move whole to an appendix slide at 11 pt
and a notice says so (user request, 26-11-2026).
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
from report_common import notices
from report_common.table_typography import apply_table, check_table_typography, set_table_sizes, set_table_directions, TABLE_FONT
from pptx.oxml.ns import qn

PT = 12700


def sections(slide, count):
    patterns = list(weekly.builder.tables(slide))
    return [(weekly.builder.table_title(p), [[str(i), 'مهمة', 'قطاع', 'ملاحظة'] for i in range(1, count + 1)], p)
            for p in patterns]


def measurements(prs, slide, rows, body_pt, compact=None):
    """Layout-file entries: `body_pt` is the 11 pt row height with template margins, `compact` maps a size key to the row height with compact margins."""
    bottom = weekly.content_bottom(prs, slide)
    out = {}
    for title, records, pattern in sections(slide, rows):
        entry = {'header': [h / PT for h in weekly.header_heights(pattern)], 'body': [body_pt] * len(records)}
        if compact:
            entry['compact'] = {k: [v] * len(records) for k, v in compact.items()}
        out[title] = entry
    return out


class SummaryOneSlideTests(unittest.TestCase):
    def setUp(self):
        self.prs = Presentation(weekly.MASTER)
        self.slide = self.prs.slides[1]

    def test_overflow_shrinks_on_the_same_slide(self):
        # Fits only at 10 pt with the compact margins (rows of 25 pt there).
        before = len(self.prs.slides)
        data = sections(self.slide, 3)
        sizes = []
        targets = weekly.flow(self.prs, self.slide, data, True,
                              measurements(self.prs, self.slide, 3, 90, {'11': 80, '10.5': 60, '10': 25}), sizes=sizes)
        self.assertEqual(targets, [self.slide])
        self.assertEqual(len(self.prs.slides), before)
        tables = sorted(weekly.builder.tables(self.slide), key=lambda sh: sh.top)
        self.assertEqual(len(tables), len(data))
        self.assertLessEqual(tables[-1].top + tables[-1].height, weekly.content_bottom(self.prs, self.slide))
        self.assertTrue(all(a.top + a.height < b.top for a, b in zip(tables, tables[1:])))
        self.assertTrue(sizes)
        margin = weekly.compact_margin(self.prs)
        self.assertEqual({(size, used) for _, size, used in sizes}, {(10, margin)})
        # Body rows use the heights measured at that size with the compact margins (the caller applies the margins to the cells, `sizes` carries them).
        self.assertEqual(tables[0].table.rows[2].height, 25 * PT)

    def test_fitting_summary_keeps_11_pt(self):
        sizes = []
        weekly.flow(self.prs, self.slide, sections(self.slide, 1), True,
                    measurements(self.prs, self.slide, 1, 30), sizes=sizes)
        self.assertEqual(sizes, [])

    def test_too_long_even_at_the_floor_moves_whole_tables_to_an_appendix(self):
        notices.reset()
        data = sections(self.slide, 4)
        sizes = []
        measured = measurements(self.prs, self.slide, 4, 400, {str(s): 300 for s in SUMMARY_SIZES})
        # The first table is short and stays on the summary slide; every other one is too tall and moves.
        first = measured[data[0][0]]
        first['body'] = [20] * 4
        first['compact'] = {str(s): [20] * 4 for s in SUMMARY_SIZES}
        targets = weekly.flow(self.prs, self.slide, data, True, measured, sizes=sizes)
        self.assertGreater(len(targets), 1)
        self.assertTrue(any(n['kind'] == 'space' and 'ملحق' in n['message'] for n in notices.collected()))
        placed = [row.cells[0].text for sl in targets for sh in sorted(weekly.builder.tables(sl), key=lambda s: s.top)
                  for row in list(sh.table.rows)[2:]]
        self.assertEqual(placed, [r[0] for _, records, _ in data for r in records])
        for sl in targets:
            bottom = weekly.content_bottom(self.prs, sl)
            tables = sorted(weekly.builder.tables(sl), key=lambda sh: sh.top)
            self.assertTrue(all(sh.top + sh.height <= bottom for sh in tables))
            self.assertTrue(all(a.top + a.height < b.top for a, b in zip(tables, tables[1:])))
        # Only the summary slide itself holds changed tables (compact, at the floor of 10); the appendix is at 11 pt with template margins.
        shrunk = {sh.shape_id for sh, _, _ in sizes}
        self.assertTrue(shrunk)
        self.assertEqual(shrunk, {sh.shape_id for sh in weekly.builder.tables(self.slide)})
        self.assertEqual({size for _, size, _ in sizes}, {10})
        for sl in targets[1:]:
            for sh in weekly.builder.tables(sl):
                self.assertNotIn(sh.shape_id, shrunk)

    def test_appendix_moves_whole_tables_never_a_cut_one(self):
        # Each table alone fits a slide, but not all of them on the summary slide.
        notices.reset()
        data = sections(self.slide, 5)
        sizes = []
        targets = weekly.flow(self.prs, self.slide, data, True,
                              measurements(self.prs, self.slide, 5, 50, {str(s): 45 for s in SUMMARY_SIZES}), sizes=sizes)
        self.assertGreater(len(targets), 1)
        seen = {}
        for sl in targets:
            for sh in weekly.builder.tables(sl):
                seen.setdefault(weekly.builder.table_title(sh), []).append(len(sh.table.rows) - 2)
        for title, records, _ in data:
            self.assertEqual(seen[title], [len(records)], f'«{title}» was cut across slides: {seen[title]}')
        for sl in targets:
            bottom = weekly.content_bottom(self.prs, sl)
            tables = sorted(weekly.builder.tables(sl), key=lambda sh: sh.top)
            self.assertTrue(all(sh.top + sh.height <= bottom for sh in tables))
            self.assertTrue(all(a.top + a.height < b.top for a, b in zip(tables, tables[1:])))

    def test_a_table_taller_than_a_slide_is_split_with_its_header_repeated(self):
        notices.reset()
        data = sections(self.slide, 4)
        targets = weekly.flow(self.prs, self.slide, data, True,
                              measurements(self.prs, self.slide, 4, 400, {str(s): 300 for s in SUMMARY_SIZES}), sizes=[])
        for sl in targets[1:]:
            for sh in weekly.builder.tables(sl):
                self.assertIn('#', [c.text.strip() for c in sh.table.rows[1].cells])
                self.assertTrue(weekly.builder.table_title(sh))

    def test_appendix_choice_keeps_11_pt(self):
        notices.reset()
        sizes = []
        targets = weekly.flow(self.prs, self.slide, sections(self.slide, 3), True,
                              measurements(self.prs, self.slide, 3, 90, {'11': 80, '10.5': 60, '10': 25}), sizes=sizes, overflow='appendix')
        self.assertGreater(len(targets), 1)
        # The tables kept on the summary slide take the compact margins but stay at 11 pt.
        self.assertEqual({size for _, size, _ in sizes}, {11})
        self.assertTrue(any('ملحق' in n['message'] for n in notices.collected()))

    def test_minimum_size_is_respected(self):
        # Only 9 pt fits, but the floor is 10: nothing goes below it, the tables move to an appendix instead.
        notices.reset()
        sizes = []
        targets = weekly.flow(self.prs, self.slide, sections(self.slide, 3), True,
                              measurements(self.prs, self.slide, 3, 90, {'11': 80, '10': 60, '9': 25, '8': 20}), sizes=sizes, min_pt=10)
        self.assertTrue(all(size >= 10 for _, size, _ in sizes))
        self.assertGreater(len(targets), 1)

    def test_an_explicit_floor_allows_smaller(self):
        notices.reset()
        sizes = []
        targets = weekly.flow(self.prs, self.slide, sections(self.slide, 3), True,
                              measurements(self.prs, self.slide, 3, 90, {'11': 80, '10': 60, '9': 25, '8': 20}), sizes=sizes, min_pt=9)
        self.assertEqual(targets, [self.slide])
        self.assertEqual({size for _, size, _ in sizes}, {9})

    def test_shrinking_is_reported(self):
        notices.reset()
        weekly.flow(self.prs, self.slide, sections(self.slide, 3), True,
                    measurements(self.prs, self.slide, 3, 90, {'11': 80, '10.5': 60, '10': 25}), sizes=[])
        self.assertTrue(any(n['kind'] == 'space' and 'تصغير الخط إلى 10' in n['message'] for n in notices.collected()))

    def test_stale_table_measurement_falls_back_to_the_estimate(self):
        notices.reset()
        found = measurements(self.prs, self.slide, 1, 30)
        first = next(iter(found))
        found[first]['body'] = [30, 30, 30]
        weekly.flow(self.prs, self.slide, sections(self.slide, 1), True, found, sizes=[])
        self.assertTrue(any(first in n['message'] for n in notices.collected()))

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


class TableFontTests(unittest.TestCase):
    def test_every_table_run_uses_abar(self):
        """26-11-2026: note cells whose master example names no font fell back to the theme font."""
        prs = Presentation(weekly.MASTER)
        cell = weekly.builder.tables(prs.slides[1])[0].table.cell(2, 3)
        for r in list(cell._tc.iter(qn('a:r'))):
            r.getparent().remove(r)
        weekly.set_cell_exact(cell, 'ملاحظة بلا خط محدد')
        set_table_sizes(prs);set_table_directions(prs)
        for table in (sh.table for sl in prs.slides for sh in sl.shapes if sh.has_table):
            for r in table._tbl.iter(qn('a:r')):
                props = r.find(qn('a:rPr'))
                self.assertEqual([props.find(qn('a:' + k)).get('typeface') for k in ('latin', 'ea', 'cs')], [TABLE_FONT] * 3)
        errors = []
        check_table_typography(prs, errors)
        self.assertFalse(any('must use' in e for e in errors), errors)
        props = cell.text_frame.paragraphs[0].runs[0]._r.find(qn('a:rPr'))
        props.find(qn('a:cs')).set('typeface', 'Arial')
        errors = []
        check_table_typography(prs, errors)
        self.assertTrue(any(f'must use {TABLE_FONT}' in e for e in errors))


if __name__ == '__main__':
    unittest.main()
