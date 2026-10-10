"""Observable acceptance of manual tables, planned dates and Suhail rich text."""
import contextlib, io, shutil, tempfile, unittest
from pathlib import Path
from openpyxl import load_workbook
from pptx import Presentation
from pptx.oxml.ns import qn
import run_weekly_workflow as workflow

class RequestedReportChangesTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.temp=tempfile.TemporaryDirectory();cls.root=Path(cls.temp.name)
        with contextlib.redirect_stdout(io.StringIO()):workflow.scenario(cls.root,'normal')

    @classmethod
    def tearDownClass(cls):cls.temp.cleanup()

    def test_sector_planned_dates_only_for_on_plan_and_in_detail_order(self):
        master=Presentation(self.root/'weekly.pptx')
        expected=[r for sl in master.slides for sh in workflow.weekly.builder.detail_tables(sl,'task') for r in list(sh.table.rows)[2:] if r.cells[4].text=='على المخطط']
        self.assertTrue(expected)
        for deck in (self.root/'sectors').glob('*.pptx'):
            prs=Presentation(deck)
            seen=[]
            for sl in prs.slides:
                for sh in workflow.weekly.builder.tables(sl):
                    title=workflow.weekly.builder.table_title(sh)
                    if title=='المهام على المخطط':
                        self.assertEqual([c.text for c in sh.table.rows[1].cells],['#','المهمة','القطاع','تاريخ الإنجاز المخطط','الحالة','ملاحظات'])
                        seen.extend([[c.text for c in r.cells][1:] for r in list(sh.table.rows)[2:]])
                    elif title in workflow.splitter.TASK_TITLE.values():
                        self.assertNotIn('تاريخ الإنجاز المخطط',[c.text for c in sh.table.rows[1].cells])
            self.assertEqual(seen,[[c.text for c in r.cells][1:] for r in expected])

    def test_sector_date_corruption_rejected(self):
        with tempfile.TemporaryDirectory() as td:
            out=Path(td)/'sectors';shutil.copytree(self.root/'sectors',out)
            deck=next(out.glob('*.pptx'));prs=Presentation(deck)
            sh=next(sh for sl in prs.slides for sh in workflow.weekly.builder.tables(sl) if workflow.weekly.builder.table_title(sh)=='المهام على المخطط')
            sh.table.cell(2,3).text='تاريخ خاطئ';prs.save(deck)
            import json
            manifest=json.loads((out/'manifest.json').read_text(encoding='utf-8'))
            for entry in manifest:entry['out']=str(deck)
            (out/'manifest.json').write_text(json.dumps(manifest),encoding='utf-8')
            with contextlib.redirect_stdout(io.StringIO()):
                self.assertFalse(workflow.fixtures.sector_qa.check(out,self.root/'weekly.pptx',self.root/'transactions.xlsx'))

    def rich_case(self,folder):
        folder.mkdir(exist_ok=True)
        for file in self.root.glob('*.xlsx'):shutil.copy2(file,folder/file.name)
        wb=load_workbook(folder/'projects.xlsx');ws=wb['تفاصيل مشاريع سهيل']
        values=[
            'تاريخ التحديث 2 أكتوبر 2026 نص بعد التاريخ\n• تجهيز المتطلبات\n• التنسيق مع الجهات',
            'تاريخ التحديث: 02/10/2026\nفقرة عادية تمتد إلى سطر آخر\nمتابعة التنفيذ والتاريخ 03/10/2026',
            'تاريخ التحديث ٢ أكتوبر ٢٠٢٦\n- نقطة واحدة فقط',
            'تحديث بلا تاريخ\n1. متابعة التنفيذ\n2. معالجة الملاحظات',
        ]
        rows=[r for r in range(3,ws.max_row+1) if ws.cell(r,1).value]
        for row,value in zip(rows,values):ws.cell(row,6,value)
        wb.save(folder/'projects.xlsx')
        with contextlib.redirect_stdout(io.StringIO()):workflow.populate(folder)
        return values

    def test_update_date_and_actual_bullets_preserve_prose(self):
        with tempfile.TemporaryDirectory() as td:
            folder=Path(td);values=self.rich_case(folder)
            prs=Presentation(folder/'weekly.pptx')
            cells=[r.cells[5] for sl in prs.slides for sh in workflow.weekly.builder.detail_tables(sl,'project') for r in list(sh.table.rows)[2:]]
            self.assertEqual(len(cells),4)
            self.assertEqual(cells[0].text,'تاريخ التحديث 2 أكتوبر 2026 نص بعد التاريخ\nتجهيز المتطلبات\nالتنسيق مع الجهات')
            expected_bold=['تاريخ التحديث 2 أكتوبر 2026','تاريخ التحديث: 02/10/2026','تاريخ التحديث ٢ أكتوبر ٢٠٢٦','']
            for cell,bold,bullet_count in zip(cells,expected_bold,[2,2,0,2]):
                bold_text=''.join(r.text for p in cell.text_frame.paragraphs for r in p.runs if r.font.bold)
                self.assertEqual(bold_text,bold)
                bullets=[p for p in cell.text_frame.paragraphs if p._p.find('./'+qn('a:pPr')+'/'+qn('a:buChar')) is not None]
                self.assertEqual(len(bullets),bullet_count)
                self.assertTrue(all(p._p.get_or_add_pPr().get('rtl')=='1' for p in cell.text_frame.paragraphs))
            self.assertEqual(cells[1].text,values[1]);self.assertEqual(cells[2].text,values[2].replace('- نقطة','نقطة'))

    def test_rich_text_corruption_is_rejected(self):
        with tempfile.TemporaryDirectory() as td:
            folder=Path(td);self.rich_case(folder)
            deck=folder/'weekly.pptx';original=deck.read_bytes()
            for change in ('whole_paragraph_bold','missing_bullet','body_bullet'):
                with self.subTest(change=change):
                    deck.write_bytes(original);prs=Presentation(deck)
                    cell=next(r.cells[5] for sl in prs.slides for sh in workflow.weekly.builder.detail_tables(sl,'project') for r in list(sh.table.rows)[2:])
                    if change=='whole_paragraph_bold':
                        for r in cell.text_frame.paragraphs[0].runs:r.font.bold=True
                    elif change=='missing_bullet':
                        p=cell.text_frame.paragraphs[1];node=p._p.find('./'+qn('a:pPr')+'/'+qn('a:buChar'));node.getparent().remove(node)
                    else:
                        from pptx.oxml.xmlchemy import OxmlElement
                        cell.text_frame.paragraphs[0]._p.get_or_add_pPr().append(OxmlElement('a:buChar'))
                    prs.save(deck)
                    with self.assertRaises(ValueError):workflow.weekly_validation.validate(deck,folder/'tasks.xlsx',folder/'projects.xlsx',folder/'transactions.xlsx')

if __name__=='__main__':unittest.main()
