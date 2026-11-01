"""Digital weekly model re-flows the landscape PDF per sector and rejects unsafe inputs."""
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
from weekly_artwork import prepare_navigation, DIGITAL_DETAIL_LABEL
from pptx.oxml.ns import qn
from report_common.table_typography import all_tables


def master_sectors():
    """Build reviewed sectors from the master's own detail examples (design samples only).

    The master shows one sector split over two sample slides; the reviewed data
    keeps one merged record set per sector, as the extractor produces.
    """
    master = Presentation(digital.MASTER)
    sectors = []
    for number, slide in enumerate(digital.detail_slides(master), 2):
        table = digital._detail_table(slide)
        challenge_table = digital._challenge_table(slide)
        challenges = [] if challenge_table is None else [row.cells[1].text for row in list(challenge_table.rows)[2:]]
        title = digital.split_title(table.cell(0, 0).text)[0]
        rows = [[cell.text for cell in row.cells] for row in list(table.rows)[2:]]
        if sectors and sectors[-1]['title'] == title:
            sector = sectors[-1]
            if digital.is_heading(rows[0]):
                rows[0][0] = digital.heading_base(rows[0][0])
            sector['rows'].extend(rows)
            sector['source_pages'].append(number)
            sector['challenges'].extend(challenges)
            continue
        sectors.append({
            'title': title,
            'source_pages': [number],
            'cards': {axis: digital.helpers.shape_by_id(slide, shape_id).text_frame.paragraphs[1].text
                      for axis, shape_id in digital.DETAIL_CARD_IDS.items()},
            'rows': rows,
            'challenge_title': 'التحديات' if challenges else None,
            'challenge_header': 'التحدي' if challenges else None,
            'challenges': challenges,
        })
    for sector in sectors:
        sector['challenge_title'] = 'التحديات' if sector['challenges'] else None
        sector['challenge_header'] = 'التحدي' if sector['challenges'] else None
    return sectors


class VisualTextTests(unittest.TestCase):
    """The landscape PDF stores visual glyph order; extraction restores logical text."""

    def test_arabic_latin_digits_and_brackets(self):
        cases = {
            '(1/3) ﻲﻤﻗﺮﻟا لﻮﺤﺘﻟا ﺔﻟﺎﻛو': 'وكالة التحول الرقمي (1/3)',
            '%15 :مﺪﺨﺘﺴﻤﻟا ﺔﺑﺮﺠﺗ': 'تجربة المستخدم: 15%',
            '(ﺔﻤﺘﺗ) %15 :تاودﻷاو تﺎﻴﻨﻘﺘﻟا': 'التقنيات والأدوات: 15% (تتمة)',
            '(ﺔﻈﺣﻼﻣ 18) ﺞﺣ ﻚﺴﻧ ﺔﺼﻨﻣ': 'منصة نسك حج (18 ملاحظة)',
            'ﻰﻠﻋﺄﻓ %99.99 ﱃإ': 'إلى 99.99% فأعلى',
            '(CS-01..09) ةﺎﻨﻗ ﻞﻜﻟ': 'لكل قناة (CS-01..09)',
            'SIP Trunk (CS-33..34) ﺔﻴﻔﺗﺎﻬﻟا': 'الهاتفية SIP Trunk (CS-33..34)',
            '(SIP Trunk) ﺔﻴﻔﺗﺎﻬﻟا': 'الهاتفية (SIP Trunk)',
            'INF-01': 'INF-01',
            'U-01..09': 'U-01..09',
            'nusuk.sa مﺎﻬﻣ': 'مهام nusuk.sa',
        }
        for visual, logical in cases.items():
            self.assertEqual(digital._logical(visual), logical)
        self.assertEqual(digital._rtl('CS-06/13/\n21/77'), 'CS-06/13/21/77')
        self.assertEqual(digital.split_title('مركز منصة نسك الرقمية (2/6)'), ('مركز منصة نسك الرقمية', 2, 6))
        self.assertEqual(digital.split_title('وكالة شؤون العمرة والزيارة'), ('وكالة شؤون العمرة والزيارة', None, None))


class MergeTests(unittest.TestCase):
    """Source pages of one sector fold into a single record set without source pagination."""

    def page(self, number, title, rows, challenges=(), cards=None):
        return {'page': number, 'title': title, 'rows': rows,
                'cards': cards or {axis: '15%' for axis in digital.AXES},
                'challenge_title': 'التحديات' if challenges else None,
                'challenge_header': 'التحدي' if challenges else None,
                'challenges': list(challenges)}

    def test_pages_merge_and_continuation_headings_drop(self):
        pages = [
            self.page(2, 'جهة أ (1/2)', [['تجربة المستخدم: 15%', '', '', ''], ['أ', '15%', 'مهمة 1', 'X-01']], ['تحد 1']),
            self.page(3, 'جهة أ (2/2)', [['تجربة المستخدم: 15% (تتمة)', '', '', ''], ['أ', '15%', 'مهمة 2', 'X-02'],
                                        ['رضا المستفيد: 15%', '', '', ''], ['أ', '15%', 'مهمة 3', 'X-03']], ['تحد 2']),
            self.page(4, 'جهة ب', [['أ', 'معلق', 'مهمة 4', 'Y-01']]),
        ]
        sectors = digital.merge_detail_pages(pages)
        self.assertEqual([s['title'] for s in sectors], ['جهة أ', 'جهة ب'])
        self.assertEqual(sectors[0]['source_pages'], [2, 3])
        self.assertEqual([row[0] for row in sectors[0]['rows'] if digital.is_heading(row)],
                         ['تجربة المستخدم: 15%', 'رضا المستفيد: 15%'])
        self.assertEqual([row[3] for row in sectors[0]['rows'] if not digital.is_heading(row)], ['X-01', 'X-02', 'X-03'])
        self.assertEqual(sectors[0]['challenges'], ['تحد 1', 'تحد 2'])
        self.assertEqual(sectors[1]['challenges'], [])
        self.assertIsNone(sectors[1]['challenge_title'])

    def test_broken_sequences_and_card_conflicts_rejected(self):
        with self.assertRaisesRegex(ValueError, 'page sequence'):
            digital.merge_detail_pages([self.page(2, 'جهة أ (1/3)', [['أ', '1%', 'م', 'X-01']]),
                                        self.page(3, 'جهة أ (3/3)', [['أ', '1%', 'م', 'X-02']])])
        with self.assertRaisesRegex(ValueError, 'missing source pages'):
            digital.merge_detail_pages([self.page(2, 'جهة أ (1/2)', [['أ', '1%', 'م', 'X-01']])])
        with self.assertRaisesRegex(ValueError, 'cards differ'):
            digital.merge_detail_pages([self.page(2, 'جهة أ (1/2)', [['أ', '1%', 'م', 'X-01']]),
                                        self.page(3, 'جهة أ (2/2)', [['أ', '1%', 'م', 'X-02']],
                                                  cards={axis: '20%' for axis in digital.AXES})])
        with self.assertRaisesRegex(ValueError, 'appears twice'):
            digital.merge_detail_pages([self.page(2, 'جهة أ', [['أ', '1%', 'م', 'X-01']]),
                                        self.page(3, 'جهة ب', [['أ', '1%', 'م', 'X-02']]),
                                        self.page(4, 'جهة أ', [['أ', '1%', 'م', 'X-03']])])


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
            'detail_sectors': master_sectors(),
        }
        self.review = self.root / 'review.json'

    def load(self):
        self.review.write_text(json.dumps(self.data, ensure_ascii=False), encoding='utf-8')
        return digital.load_reviewed(self.review, self.pdf)

    def build(self):
        data = self.load()
        prs = Presentation(digital.MASTER)
        digital.populate(prs, data)
        prepare_navigation(prs)
        report = self.root / 'report.pptx'
        prs.save(report)
        return data, Presentation(report), report

    def test_review_and_pdf_identity_are_required(self):
        self.data['reviewed'] = False
        with self.assertRaisesRegex(ValueError, 'Review'):
            self.load()
        self.data['reviewed'] = True
        self.load()
        self.pdf.write_bytes(b'a different PDF')
        with self.assertRaisesRegex(ValueError, 'does not match'):
            digital.load_reviewed(self.review, self.pdf)

    def test_calculation_note_is_required(self):
        slide = digital.digital_slide(Presentation(digital.MASTER))
        self.assertIn('آلية الاحتساب', digital.calculation_note(slide).text)
        digital.helpers.shape_by_id(slide, digital.CALCULATION_NOTE_ID).text = 'ملاحظة ناقصة'
        with self.assertRaisesRegex(ValueError, 'calculation note'):
            digital.calculation_note(slide)

    def test_capacity_and_conflicting_note_rejected(self):
        self.data['challenges'] *= 3
        with self.assertRaisesRegex(ValueError, 'exceed'):
            self.load()
        self.data['challenges'] = self.data['challenges'][:2]
        self.data['sectors'][1]['metrics']['experience'] = 5
        with self.assertRaisesRegex(ValueError, 'cannot have percentages'):
            self.load()

    def test_source_pagination_artifacts_rejected(self):
        self.data['detail_sectors'][0]['title'] += ' (1/3)'
        with self.assertRaisesRegex(ValueError, 'without a page suffix'):
            self.load()
        self.data['detail_sectors'][0]['title'] = self.data['detail_sectors'][0]['title'][:-6]
        self.data['detail_sectors'][0]['rows'].insert(2, ['تجربة المستخدم: 15% (تتمة)', '', '', ''])
        with self.assertRaisesRegex(ValueError, 'continuation headings'):
            self.load()
        del self.data['detail_sectors'][0]['rows'][2]
        self.data['detail_sectors'][0]['challenge_title'] = 'عنوان آخر'
        with self.assertRaisesRegex(ValueError, 'unapproved'):
            self.load()
        self.data['detail_sectors'][0]['challenge_title'] = 'أبرز التحديثات'
        self.data['detail_sectors'][0]['challenge_header'] = 'التحديث'
        legacy = dict(self.data, detail_pages=[])
        self.review.write_text(json.dumps(legacy, ensure_ascii=False), encoding='utf-8')
        with self.assertRaisesRegex(ValueError, 'detail_sectors'):
            digital.load_reviewed(self.review, self.pdf)

    def test_charts_and_challenges_match_reviewed_data(self):
        data, prs, report = self.build()
        self.assertTrue(all(t._tbl.tblPr.get('rtl') == '1' for t in all_tables(prs)))
        errors = []
        digital.check(prs, report, data, errors)
        self.assertEqual(errors, [])
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

    def test_sector_overflow_reflows_into_continuation_slides(self):
        """A long sector paginates by the portrait slide, never by the source pages."""
        first = self.data['detail_sectors'][0]
        body = [row for row in first['rows'] if not digital.is_heading(row)]
        first['rows'] = ([['تجربة المستخدم: 15%', '', '', '']]
                         + [[row[0], row[1], f'مهمة طويلة رقم {i} ' + row[2], f'X-{i:02d}'] for i, row in enumerate(body * 3, 1)]
                         + [['رضا المستفيد: 15%', '', '', '']]
                         + [[row[0], row[1], row[2], f'Y-{i:02d}'] for i, row in enumerate(body, 1)])
        first['challenge_title'], first['challenge_header'] = 'أبرز التحديثات', 'التحديث'
        first['challenges'] = ['تحديث أول عن المهمة W-03 وسير العمل', 'تحديث ثان']
        data, prs, report = self.build()
        slides = digital.detail_slides(prs)
        group = [slide for slide in slides if digital.split_title(digital._detail_table(slide).cell(0, 0).text)[0] == first['title']]
        self.assertGreater(len(group), 1)
        titles = [digital._detail_table(slide).cell(0, 0).text for slide in group]
        self.assertEqual(titles, [f'{first["title"]} ({i}/{len(group)})' for i in range(1, len(group)+1)])
        for slide in group:
            table = digital._detail_table(slide)
            shape = digital._detail_shape(slide)
            self.assertLessEqual(shape.top + shape.height, digital.content_bottom(prs, slide))
            self.assertTrue(all(c.text == first['cards'][axis] for axis, c in
                                ((axis, digital.helpers.shape_by_id(slide, sid).text_frame.paragraphs[1]) for axis, sid in digital.DETAIL_CARD_IDS.items())))
            self.assertFalse(digital.is_heading(digital._slide_rows(table)[-1]))
        continuation = digital._slide_rows(digital._detail_table(group[1]))[0]
        self.assertTrue(digital.is_heading(continuation))
        self.assertTrue(continuation[0].endswith(digital.CONTINUATION))
        self.assertEqual([digital._challenge_table(slide) is None for slide in group], [True] * (len(group)-1) + [False])
        challenges = digital._challenge_table(group[-1])
        self.assertEqual(challenges.cell(0, 0).text, 'أبرز التحديثات')
        self.assertEqual(challenges.cell(1, 1).text, 'التحديث')
        self.assertEqual([digital._plain(row.cells[1].text) for row in list(challenges.rows)[2:]], first['challenges'])
        # Sector order and the closing slide are preserved; sample master pages are gone.
        self.assertEqual(list(prs.slides).index(slides[-1]), len(prs.slides) - 2)
        errors = []
        digital.check(prs, report, data, errors)
        self.assertEqual(errors, [])

    def test_premature_page_break_is_rejected(self):
        """Moving a record that still fits onto the next slide is the last-week failure."""
        first = self.data['detail_sectors'][0]
        body = [row for row in first['rows'] if not digital.is_heading(row)]
        first['rows'] = [['تجربة المستخدم: 15%', '', '', '']] + [[row[0], row[1], row[2], f'X-{i:02d}'] for i, row in enumerate(body * 3, 1)]
        data, prs, report = self.build()
        group = [slide for slide in digital.detail_slides(prs)
                 if digital.split_title(digital._detail_table(slide).cell(0, 0).text)[0] == first['title']]
        self.assertGreater(len(group), 1)
        source, target = digital._detail_shape(group[0]), digital._detail_shape(group[1])
        moved = source.table._tbl.tr_lst[-1]
        source.table._tbl.remove(moved)
        source.height = sum(row.height for row in source.table.rows)
        target.table._tbl.tr_lst[2].addnext(moved)
        target.height = sum(row.height for row in target.table.rows)
        prs.save(report)
        errors = []
        digital.check(Presentation(report), report, data, errors)
        self.assertTrue(any('breaks the page early' in error for error in errors), errors)
        self.assertFalse(any('rows differ' in error for error in errors), errors)

    def test_stale_source_suffix_and_misplaced_challenges_rejected(self):
        data, prs, report = self.build()
        slides = digital.detail_slides(prs)
        table = digital._detail_table(slides[0])
        digital.set_cell_exact(table.cell(0, 0), table.cell(0, 0).text + ' (1/3)')
        prs.save(report)
        errors = []
        digital.check(Presentation(report), report, data, errors)
        self.assertTrue(any('suffix' in error or 'differ' in error for error in errors), errors)

    def test_latin_codes_keep_direction(self):
        self.data['detail_sectors'][4]['challenges'] = ['تحد مرتبط بالمهمة W-03', 'تحد ثان مرتبط بالمهمة S-02']
        self.data['detail_sectors'][4]['challenge_title'] = 'التحديات'
        self.data['detail_sectors'][4]['challenge_header'] = 'التحدي'
        self.data['detail_sectors'][4]['rows'][1][3] = 'U-01..09'
        data, prs, report = self.build()
        slide = [s for s in digital.detail_slides(prs)
                 if digital.split_title(digital._detail_table(s).cell(0, 0).text)[0] == data['detail_sectors'][4]['title']][-1]
        challenges = digital._challenge_table(slide)
        self.assertEqual([row.cells[1].text for row in list(challenges.rows)[2:]], data['detail_sectors'][4]['challenges'])
        group = [s for s in digital.detail_slides(prs)
                 if digital.split_title(digital._detail_table(s).cell(0, 0).text)[0] == data['detail_sectors'][4]['title']]
        code = next(row.cells[list(digital._detail_columns(table, row))[3]] for s in group for table in [digital._detail_table(s)]
                    for row in list(table.rows)[2:] if digital._plain(row.cells[list(digital._detail_columns(table, row))[3]].text) == 'U-01..09')
        self.assertEqual(code.text_frame.paragraphs[0]._p.get_or_add_pPr().get('rtl'), '0')

    def test_all_indicator_details_appear_in_their_section(self):
        for number, sector in enumerate(self.data['detail_sectors'], 2):
            sector['title'] = f'تفاصيل الجهة {number}'
            sector['rows'][1][2] = f'مهمة الصفحة {number}'
        data, prs, report = self.build()
        slides = digital.detail_slides(prs)
        indexes = [list(prs.slides).index(slide) for slide in slides]
        self.assertGreaterEqual(len(slides), len(data['detail_sectors']))
        self.assertEqual(indexes, list(range(11, 11 + len(slides))))
        self.assertEqual(indexes[-1], len(prs.slides) - 2)
        firsts = [slide for slide in slides if digital.split_title(digital._detail_table(slide).cell(0, 0).text)[1] in (None, 1)]
        self.assertEqual([digital.split_title(digital._detail_table(s).cell(0, 0).text)[0] for s in firsts],
                         [sector['title'] for sector in data['detail_sectors']])
        for slide, sector in zip(firsts, data['detail_sectors']):
            self.assertEqual(digital._detail_table(slide).cell(3, 1).text, sector['rows'][1][2])
        divider = prs.slides[5]
        button = next(shape for shape in divider.shapes if getattr(shape, 'has_text_frame', False)
                      and shape.text == DIGITAL_DETAIL_LABEL)
        links = list(button.element.iter(qn('a:hlinkClick')))
        self.assertEqual(len(links), 1)
        self.assertEqual(divider.part.rels[links[0].get(qn('r:id'))].target_part, slides[0].part)

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

    def test_full_digital_model_builds_rtl_tables_and_rejects_disabled_direction(self):
        import contextlib
        import io
        import run_weekly_workflow as workflow
        inputs = self.root / 'inputs'
        first = self.data['detail_sectors'][0]
        body = [row for row in first['rows'] if not digital.is_heading(row)]
        first['rows'] = [['تجربة المستخدم: 15%', '', '', '']] + [[row[0], row[1], row[2], f'X-{i:02d}'] for i, row in enumerate(body * 3, 1)]
        self.load()
        report = self.root / 'digital.pptx'
        with contextlib.redirect_stdout(io.StringIO()):
            workflow.scenario(inputs, 'normal')
            weekly_build.main(['--tasks', str(inputs / 'tasks.xlsx'),
                               '--suhail', str(inputs / 'projects.xlsx'),
                               '--transactions', str(inputs / 'transactions.xlsx'),
                               '--model', 'digital', '--digital-pdf', str(self.pdf),
                               '--digital-data', str(self.review), '--output', str(report)])
        prs = Presentation(report)
        tables = list(all_tables(prs))
        self.assertTrue(all(t._tbl.tblPr.get('rtl') == '1' for t in tables))
        self.assertGreater(len(digital.detail_slides(prs)), len(self.data['detail_sectors']))
        table = digital._detail_table(digital.detail_slides(prs)[0])
        table._tbl.tblPr.set('rtl', '0')
        prs.save(report)
        before = report.read_bytes()
        with self.assertRaisesRegex(ValueError, 'Table direction must be RTL'):
            workflow.weekly_validation.validate(report, inputs / 'tasks.xlsx',
                inputs / 'projects.xlsx', inputs / 'transactions.xlsx',
                digital.MASTER, 'digital', self.pdf, self.review)
        self.assertEqual(report.read_bytes(), before)


if __name__ == '__main__':
    unittest.main()
