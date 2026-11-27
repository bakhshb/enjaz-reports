# /// script
# requires-python = ">=3.12"
# dependencies = ["pandas>=2.2,<3", "openpyxl>=3.1,<4", "python-pptx>=1.0.2", "lxml>=5.3", "pillow", "pdfplumber>=0.11"]
# ///
"""Monday build keeps going on survivable conditions: unknown status, KPI mismatch, tight agenda."""
import json
import shutil
import sys
import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook, load_workbook
from pptx import Presentation

from test_deck_integrity import MONDAY, gate, run
from test_upstream_integrity import SUHAIL, TASKS, suhail_source, task_source

sys.path.insert(0, str(MONDAY))
import monday_layout as layout
import monday_status_styles as status_styles

NEUTRAL_FILL = status_styles.FILLS if hasattr(status_styles, 'FILLS') else None


def set_unknown_status(path, sheet, summary, row_of_status_col=5):
    """Give the first detail record an unknown status and keep the summary counts consistent with it."""
    wb = load_workbook(path)
    ws = wb[sheet]
    row = next(r for r in range(1, ws.max_row + 1) if str(ws.cell(r, 1).value or '').isdigit())
    old = {'مكتمل': 'مكتملة', 'متأخرة': 'متأخر', 'معلقة': 'معلق'}.get(ws.cell(row, 5).value, ws.cell(row, 5).value)
    sector = ws.cell(row, 3).value if sheet == 'تفاصيل المهام' else next(
        ws.cell(r, 1).value for r in range(row, 0, -1) if ws.cell(r, 1).value and not str(ws.cell(r, 1).value).isdigit()
        and ws.cell(r, 2).value is None)
    ws.cell(row, 5, 'قيد المراجعة')
    sm = wb[summary]
    head = next(r for r in range(1, sm.max_row + 1) if sm.cell(r, 1).value == 'القطاع')
    col = next(c for c in range(1, 8) if sm.cell(head, c).value == old)
    line = next(r for r in range(head + 1, sm.max_row + 1) if sm.cell(r, 1).value == sector)
    sm.cell(line, col).value = (sm.cell(line, col).value or 1) - 1 or None
    wb.save(path)


def fill_of(cell):
    props = cell._tc.tcPr
    return props.xpath('./a:solidFill/a:srgbClr/@val')[0] if props is not None else None


class MondayRobustTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.temp = tempfile.TemporaryDirectory()
        cls.root = Path(cls.temp.name)
        raw_tasks, raw_projects = cls.root / 'raw-tasks.xlsx', cls.root / 'raw-projects.xlsx'
        task_source(raw_tasks)
        suhail_source(raw_projects, 'متأخر')
        cls.tasks, cls.projects = cls.root / 'tasks.xlsx', cls.root / 'projects.xlsx'
        for script, args in [(TASKS, ['--input', raw_tasks, '--output', cls.tasks]), (SUHAIL, [raw_projects, cls.projects])]:
            result = run(script, *args)
            if result.returncode:
                raise AssertionError(result.stderr + result.stdout)

    @classmethod
    def tearDownClass(cls):
        cls.temp.cleanup()

    def topics(self, directory, count, text='موضوع'):
        path = Path(directory) / 'topics.xlsx'
        wb = Workbook()
        ws = wb.active
        ws.title = 'Sheet1'
        ws.append(['أسبوع تجريبي'])
        ws.append(['الترتيب', 'الموضوع', 'المتحدث', 'الوقت المطلوب من المالك'])
        for i in range(count):
            ws.append([i + 1, f'{text} {i + 1}', 'مالك', 5])
        ws.append([count + 1, 'ملخص الاجتماع', 'مالك', 5])
        wb.save(path)
        return path

    def build(self, directory, topics=None, tasks=None, projects=None):
        output = Path(directory) / 'report.pptx'
        result = run(MONDAY / 'build_report.py', '--topics', topics or self.topics(directory, 1),
                     '--tasks', tasks or self.tasks, '--suhail', projects or self.projects,
                     '--output', output, '--layout-engine', 'estimate')
        notes = Path(str(output) + '.notices.json')
        return result, output, json.loads(notes.read_text(encoding='utf-8')) if notes.exists() else []

    def agenda_tables(self, output):
        prs = Presentation(output)
        return prs, [sh for sl in prs.slides for sh in sl.shapes if sh.has_table
                     and [c.text.strip() for c in sh.table.rows[0].cells] == gate.AGENDA_HEADERS]

    def test_unknown_status_keeps_text_neutral_fill_and_a_notice(self):
        with tempfile.TemporaryDirectory() as td:
            tasks, projects = Path(td) / 'tasks.xlsx', Path(td) / 'projects.xlsx'
            for source, sheet, summary, target in [(self.tasks, 'تفاصيل المهام', 'ملخص المهام', tasks),
                                                   (self.projects, 'تفاصيل مشاريع سهيل', 'ملخص مشاريع سهيل', projects)]:
                shutil.copy2(source, target)
                set_unknown_status(target, sheet, summary)
            result, output, notes = self.build(td, tasks=tasks, projects=projects)
            self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
            self.assertTrue(any(n['kind'] == 'status' and 'قيد المراجعة' in n['message'] for n in notes), notes)
            # One wording only: the builder's. The gate adds no second or third notice for the same status.
            status_notes = [n['message'] for n in notes if n['kind'] == 'status']
            self.assertFalse([m for m in status_notes if 'تفاصيل الشرائح' in m or 'عدّ الحالات' in m], status_notes)
            self.assertEqual(len(status_notes), len(set(status_notes)), status_notes)
            prs = Presentation(output)
            approved = status_styles.examples()
            found = 0
            for sl in prs.slides:
                for sh in sl.shapes:
                    if sh.has_table and len(sh.table.columns) == 6 and sh.table.cell(1, 4).text.strip() == 'الحالة':
                        for row in list(sh.table.rows)[2:]:
                            cell = row.cells[4]
                            if cell.text.strip() == 'قيد المراجعة':
                                found += 1
                                kind = 'task' if sh.table.cell(1, 1).text.strip() == 'المهمة' else 'suhail'
                                self.assertEqual(fill_of(cell), fill_of(approved[status_styles.NEUTRAL_STATUS[kind]]))
            self.assertEqual(found, 2)

    def test_agenda_floor_and_step_are_defined_once(self):
        self.assertEqual(gate.AGENDA_FLOOR_PT, layout.AGENDA_FLOOR_PT)
        self.assertFalse(hasattr(gate, 'AGENDA_STEP_PT'))

    def test_style_status_uses_the_shared_style_key(self):
        approved = {'معلق': 1, 'مكتملة': 1}
        self.assertEqual(status_styles.style_status('مكتملة', 'task', approved), 'مكتملة')
        self.assertEqual(status_styles.style_status('قيد المراجعة', 'task', approved), 'معلق')
        self.assertEqual(status_styles.style_status('مكتملة', 'suhail', {'لم تبدأ': 1}), 'لم تبدأ')

    def test_gate_still_rejects_a_wrong_fill_on_an_unknown_status(self):
        with tempfile.TemporaryDirectory() as td:
            tasks = Path(td) / 'tasks.xlsx'
            shutil.copy2(self.tasks, tasks)
            set_unknown_status(tasks, 'تفاصيل المهام', 'ملخص المهام')
            result, output, _ = self.build(td, tasks=tasks)
            self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
            prs = Presentation(output)
            approved = status_styles.examples()
            cell = next(r.cells[4] for sl in prs.slides for sh in sl.shapes if sh.has_table
                        and len(sh.table.columns) == 6 and sh.table.cell(1, 1).text.strip() == 'المهمة'
                        for r in list(sh.table.rows)[2:] if r.cells[4].text.strip() == 'قيد المراجعة')
            status_styles.apply(cell, approved['مكتملة'])
            from report_common.pptx_helpers import set_cell_text
            set_cell_text(cell, 'قيد المراجعة')
            prs.save(output)
            gated = run(MONDAY / 'report_gate_v2.py', '--report', output, '--topics', self.topics(td, 1),
                        '--tasks', tasks, '--suhail', self.projects)
            self.assertNotEqual(gated.returncode, 0, gated.stdout)
            self.assertIn('differs from approved template', gated.stdout)

    def test_kpi_disagreement_reaches_notices_and_does_not_stop_the_build(self):
        with tempfile.TemporaryDirectory() as td:
            tasks = Path(td) / 'tasks.xlsx'
            wb = load_workbook(self.tasks)
            wb['ملخص المهام']['B2'] = 99
            wb.save(tasks)
            result, output, notes = self.build(td, tasks=tasks)
            self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
            self.assertTrue(any(n['kind'] == 'source' and 'المؤشرات' in n['message'] for n in notes), notes)

    def test_stale_notices_file_is_removed_at_start(self):
        with tempfile.TemporaryDirectory() as td:
            stale = Path(td) / 'report.pptx.notices.json'
            stale.write_text('[{"kind":"x","message":"old"}]', encoding='utf-8')
            result, output, notes = self.build(td, topics=self.topics(td, 1))
            self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
            self.assertFalse(any(n['message'] == 'old' for n in notes))

    def test_agenda_shrinks_in_half_point_steps_before_a_second_slide(self):
        with tempfile.TemporaryDirectory() as td:
            # Long enough to overflow at 14 pt, short enough to fit above the floor.
            result, output, notes = self.build(td, topics=self.topics(td, 17, 'متابعة موضوع'))
            self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
            prs, agendas = self.agenda_tables(output)
            self.assertEqual(len(agendas), 1)
            sizes = {r.font.size.pt for row in agendas[0].table.rows for c in row.cells
                     for p in c.text_frame.paragraphs for r in p.runs if r.text.strip()}
            self.assertEqual(len(sizes), 1)
            size = sizes.pop()
            self.assertTrue(layout.AGENDA_FLOOR_PT < size < layout.AGENDA_PT and (size * 2).is_integer(), size)
            self.assertLessEqual(agendas[0].top + sum(r.height for r in agendas[0].table.rows),
                                 prs.slide_height - int(55 * 12700))
            self.assertEqual(len(agendas[0].table.rows), 19)
            self.assertTrue(any(n['kind'] == 'space' and 'جدول الأعمال' in n['message'] and f'{size:g}' in n['message']
                                for n in notes), notes)

    def test_agenda_continues_on_a_second_slide_at_the_floor(self):
        with tempfile.TemporaryDirectory() as td:
            result, output, notes = self.build(td, topics=self.topics(td, 60))
            self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
            prs, agendas = self.agenda_tables(output)
            self.assertGreaterEqual(len(agendas), 2)
            for table in agendas:
                self.assertEqual([c.text.strip() for c in table.table.rows[0].cells], gate.AGENDA_HEADERS)
                self.assertLessEqual(table.top + sum(r.height for r in table.table.rows), prs.slide_height - int(55 * 12700))
                self.assertEqual({r.font.size.pt for row in table.table.rows for c in row.cells
                                  for p in c.text_frame.paragraphs for r in p.runs if r.text.strip()},
                                 {float(layout.AGENDA_FLOOR_PT)})
            rows = [r for table in agendas for r in list(table.table.rows)[1:]]
            self.assertEqual([r.cells[0].text for r in rows], [f'{i:02d}' for i in range(1, 62)])
            self.assertTrue(any(n['kind'] == 'space' and 'شريحة جدول أعمال ثانية' in n['message'] for n in notes), notes)

    def test_gate_rejects_a_split_agenda_above_the_floor(self):
        with tempfile.TemporaryDirectory() as td:
            topics = self.topics(td, 1)
            result, output, _ = self.build(td, topics=topics)
            self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
            from report_common.pptx_helpers import clone_slide
            prs = Presentation(output)
            slide = next(sl for sl in prs.slides if any(sh.has_table and sh.table.cell(0, 0).text == 'م' for sh in sl.shapes))
            clone_slide(prs, slide, after=slide)
            prs.save(output)
            gated = run(MONDAY / 'report_gate_v2.py', '--report', output, '--topics', topics,
                        '--tasks', self.tasks, '--suhail', self.projects)
            self.assertNotEqual(gated.returncode, 0)
            self.assertIn('Topics must occupy exactly one page', gated.stdout)


if __name__ == '__main__':
    unittest.main()
