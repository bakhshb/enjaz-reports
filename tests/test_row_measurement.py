# /// script
# requires-python = ">=3.12"
# dependencies = ["pandas>=2.2,<3", "openpyxl>=3.1,<4", "python-pptx>=1.0.2", "lxml>=5.3", "pillow", "pdfplumber>=0.11", "xlrd>=2"]
# ///
"""Row heights come from measurement, not from echoing declared heights.

LibreOffice draws a table row at exactly its declared height and lets longer
text overflow, so reading row borders from a render never corrects an estimate.
These tests pin the probe measurement (report_common/measure.py), the
consolidated LibreOffice measurement (measure_layout.py) and the sector
splitter's measure-and-rebuild pass, plus the upstream date-text fix.
"""
import contextlib
import datetime
import io
import json
from pathlib import Path
import shutil
import sys
import tempfile
import unittest

from openpyxl import load_workbook
from pptx import Presentation

ROOT=Path(__file__).resolve().parents[1]
sys.path.insert(0,str(ROOT/'skills'))
from report_common.measure import measure_rows
from report_common.pptx_helpers import val
from test_upstream_integrity import run, suhail_source, SUHAIL

SECTOR=ROOT/'skills/weekly-report-by-sector/scripts'
WEEKLY=ROOT/'skills/weekly-consolidated-report/scripts'
DIGITAL=ROOT/'skills/weekly-consolidated-report/assets/weekly-report-digital-master.pptx'
LIBREOFFICE=shutil.which('soffice') is not None
EMU=914400


def spacing_faults(deck):
    """Table paragraphs of a saved deck whose first paragraph property is not a fixed line spacing of 1.2 times every size in them."""
    from pptx.oxml.ns import qn
    faults=[];count=0
    for si,slide in enumerate(Presentation(deck).slides):
        for tbl in slide._element.iter(qn('a:tbl')):
            for p in tbl.iter(qn('a:p')):
                count+=1
                props=p.find(qn('a:pPr'))
                exact=props[0].find(qn('a:spcPts')) if props is not None and len(props) and props[0].tag==qn('a:lnSpc') else None
                sizes={int(node.get('sz')) for node in p.iter(qn('a:rPr'),qn('a:endParaRPr')) if node.get('sz')}
                if exact is None or len(sizes)!=1 or int(exact.get('val'))!=round(sizes.pop()*1.2):
                    faults.append((Path(deck).name,si+1,''.join(t.text or '' for t in p.iter(qn('a:t')))[:30]))
    assert count,deck
    return faults


def first_table(prs,min_rows=3):
    for i,slide in enumerate(prs.slides):
        for sh in slide.shapes:
            if sh.has_table and len(sh.table.rows)>=min_rows:return i,sh
    raise AssertionError('no table')


@unittest.skipUnless(LIBREOFFICE,'LibreOffice is required')
class ProbeMeasurementTests(unittest.TestCase):
    def test_measures_text_not_declared_height(self):
        with tempfile.TemporaryDirectory() as directory:
            deck=Path(directory)/'deck.pptx'
            prs=Presentation(DIGITAL);si,sh=first_table(prs,4)
            rows=list(sh.table.rows)
            long_text=' '.join(['نص طويل للاختبار']*60)
            rows[2].cells[1].text=long_text;rows[2].height=int(0.05*EMU)   # far too short
            rows[3].cells[1].text='قصير';rows[3].height=int(6*EMU)          # far too tall
            prs.save(deck);before=deck.read_bytes()
            measured=measure_rows(deck,[(si,sh.shape_id,[2,3])])
            self.assertGreater(measured[(si,sh.shape_id,2)],int(0.6*EMU))
            self.assertLess(measured[(si,sh.shape_id,3)],int(0.8*EMU))
            self.assertEqual(before,deck.read_bytes())


    def test_known_line_count_measures_lines_times_1_2_font_size_plus_margins(self):
        """27-11-2026: LibreOffice drew Abar lines at 1.8 times the size (19.8 pt at 11) and the user's PDF at 1.2 (13.2 pt),

        so a 19-line row measured 5.39 in against his 3.58 in and a summary that fits went to an appendix.
        The tables now carry a fixed spacing of 1.2 times the size, and the measure step reads that height.
        """
        sys.path.insert(0,str(WEEKLY))
        from weekly_data import set_cell_exact, set_body_margins
        from report_common.table_typography import set_table_sizes, apply_table
        lines=7;padding=0.06
        with tempfile.TemporaryDirectory() as directory:
            for size,margin in [(11,None),(10.5,None),(10,9144),(8,9144)]:
                deck=Path(directory)/f'deck-{size:g}.pptx'
                prs=Presentation(DIGITAL);set_table_sizes(prs);si,sh=first_table(prs,4)
                table=sh.table
                for tr in list(table._tbl.tr_lst)[3:]:table._tbl.remove(tr)
                # Short paragraphs, one per line: the line count does not depend on wrapping.
                for cell in table.rows[2].cells:set_cell_exact(cell,'')
                set_cell_exact(table.rows[2].cells[3],'\n'.join(f'سطر {n}' for n in range(1,lines+1)))
                apply_table(table,size)
                if margin is not None:set_body_margins(table,margin)
                cell=table.rows[2].cells[3]
                margins=(cell.margin_top+cell.margin_bottom)/EMU*72
                prs.save(deck)
                key=(si,sh.shape_id,2)
                bare=measure_rows(deck,[(si,sh.shape_id,[2])],padding=0)[key]/EMU*72
                expected=lines*1.2*size+margins
                # Within one point of the exact height; at the font's own 1.8 the row would be 0.6 x size x lines taller.
                self.assertAlmostEqual(bare,expected,delta=1.0,msg=f'{size} pt')
                padded=measure_rows(deck,[(si,sh.shape_id,[2])])[key]/EMU*72
                self.assertAlmostEqual(padded,expected+padding*72,delta=1.0,msg=f'{size} pt')

    def test_measured_height_holds_the_rendered_text(self):
        """26-11-2026: rows sized by the old collapsed probe were 0.05-0.07 in short, so tables overlapped."""
        sys.path.insert(0,str(ROOT/'skills'))
        from report_common.render import export_pdf
        import pdfplumber
        with tempfile.TemporaryDirectory() as directory:
            deck=Path(directory)/'deck.pptx'
            sys.path.insert(0,str(WEEKLY))
            from weekly_data import set_cell_exact
            from report_common.table_typography import set_table_sizes
            prs=Presentation(DIGITAL);set_table_sizes(prs);si,sh=first_table(prs,4)
            table=sh.table
            for tr in list(table._tbl.tr_lst)[3:]:table._tbl.remove(tr)
            # Filled the way the builder fills cells: the template run keeps Abar 11 pt.
            set_cell_exact(table.rows[2].cells[3],'\n'.join(' '.join(['سطر طويل لقياس ارتفاع الصف']*2) for _ in range(3)))
            prs.save(deck)
            needed=measure_rows(deck,[(si,sh.shape_id,[2])])[(si,sh.shape_id,2)]
            prs=Presentation(deck);frame=next(x for x in prs.slides[si].shapes if x.shape_id==sh.shape_id)
            # Roomy title/header rows, so only the measured row decides where the text ends.
            frame.table.rows[0].height=frame.table.rows[1].height=EMU
            frame.table.rows[2].height=needed;frame.height=sum(r.height for r in frame.table.rows)
            prs.save(deck)
            pdf=export_pdf(deck,Path(directory)/'deck.pdf')
            with pdfplumber.open(pdf) as document:
                page=document.pages[si]
                left,right=frame.left/EMU*72,(frame.left+frame.width)/EMU*72
                bottom=(frame.top+frame.height)/EMU*72
                # Text overflowing the row would land within an inch below it; the footer is far lower.
                chars=[c for c in page.chars if c['text'].strip() and left<=c['x0']<=right and frame.top/EMU*72<c['top']<bottom+72]
                self.assertTrue(chars)
                self.assertLessEqual(max(c['bottom'] for c in chars),bottom)


class DateTextTests(unittest.TestCase):
    def test_iso_timestamp_text_is_shown_as_report_date(self):
        self.assertEqual(val('2026-12-11 00:00:00'),'11-12-2026')
        self.assertEqual(val('2026-12-11'),'11-12-2026')
        self.assertEqual(val('30-08-2026'),'30-08-2026')
        self.assertEqual(val('ملاحظة 2026-12-11'),'ملاحظة 2026-12-11')

    def test_suhail_export_writes_excel_dates_as_text_dates(self):
        with tempfile.TemporaryDirectory() as directory:
            root=Path(directory);source=root/'raw.xlsx';output=root/'out.xlsx'
            suhail_source(source,status='على المخطط')
            wb=load_workbook(source);ws=wb.active
            ws.cell(2,4).value=datetime.datetime(2026,12,11)                 # تاريخ النهاية
            wb.save(source)
            result=run(SUHAIL,source,output)
            self.assertEqual(result.returncode,0,result.stdout+result.stderr)
            ws=load_workbook(output)['تفاصيل مشاريع سهيل']
            values=[c.value for row in ws.iter_rows() for c in row if c.value is not None]
            self.assertIn('11-12-2026',values)
            self.assertFalse(any('00:00:00' in str(v) for v in values))


@unittest.skipUnless(LIBREOFFICE,'LibreOffice is required')
class SectorMeasurementTests(unittest.TestCase):
    def test_rows_fit_measured_text_after_rebuild(self):
        sys.path.insert(0,str(SECTOR))
        import split_report as splitter
        with tempfile.TemporaryDirectory() as directory:
            out=Path(directory)/'out'
            result=run(SECTOR/'split_report.py',DIGITAL,'--outdir',out,'--only','وكالة شؤون الحج')
            self.assertEqual(result.returncode,0,result.stdout+result.stderr)
            self.assertIn('rebuilding',result.stdout)
            deck=next(out.glob('*.pptx'))
            prs=Presentation(deck);targets=[]
            for si,slide in enumerate(prs.slides):
                for sh in slide.shapes:
                    if sh.has_table and len(sh.table.rows)>2:
                        targets.append((si,sh.shape_id,range(2,len(sh.table.rows))))
                        self.assertEqual(sh.height,sum(r.height for r in sh.table.rows))
            measured=measure_rows(deck,targets)
            for (si,sid,ri),needed in measured.items():
                frame=next(s for s in prs.slides[si].shapes if s.shape_id==sid)
                declared=frame.table.rows[ri].height
                # Neither clipped nor padded beyond renderer noise.
                self.assertGreaterEqual(declared,needed-int(0.05*EMU),(si,ri))
                self.assertLessEqual(declared,needed+int(0.05*EMU),(si,ri))
            result=run(SECTOR/'qa_report.py',out,'--master',DIGITAL,'--only','وكالة شؤون الحج')
            self.assertEqual(result.returncode,0,result.stdout+result.stderr)


@unittest.skipUnless(LIBREOFFICE,'LibreOffice is required')
class ConsolidatedMeasurementTests(unittest.TestCase):
    def test_libreoffice_measurements_rebuild_and_validate(self):
        import run_weekly_workflow as workflow
        with tempfile.TemporaryDirectory() as directory:
            root=Path(directory)
            with contextlib.redirect_stdout(io.StringIO()):workflow.scenario(root,'normal')
            # The files the builders save (weekly, Monday, every sector deck) carry the fixed spacing in every table paragraph.
            decks=[root/'weekly.pptx',root/'monday.pptx',*sorted((root/'sectors').glob('*.pptx'))]
            self.assertGreater(len(decks),2)
            for deck in decks:self.assertEqual(spacing_faults(deck),[])
            inputs=['--tasks',root/'tasks.xlsx','--suhail',root/'projects.xlsx','--transactions',root/'transactions.xlsx']
            result=run(WEEKLY/'measure_layout.py','--report',root/'weekly.pptx',*inputs)
            self.assertEqual(result.returncode,0,result.stdout+result.stderr)
            layout=root/'weekly.layout.json'
            data=json.loads(layout.read_text(encoding='utf-8'))
            self.assertEqual(data['typography'],'11pt-line-1.2-updates-v4')
            self.assertTrue(data['tables'])
            # Summary tables carry rows measured with the compact margins at every summary size, 11 included.
            summaries=[entry for entry in data['tables'].values() if 'compact' in entry]
            self.assertTrue(summaries)
            for entry in summaries:
                self.assertEqual(set(entry['compact']),{'11','10.5','10','9.5','9','8.5','8'})
                for rows in entry['compact'].values():self.assertEqual(len(rows),len(entry['body']))
            result=run(WEEKLY/'build_report.py',*inputs,'--measurements',layout,'--output',root/'measured.pptx')
            self.assertEqual(result.returncode,0,result.stdout+result.stderr)
            self.assertLessEqual(len(Presentation(root/'measured.pptx').slides),len(Presentation(root/'weekly.pptx').slides))
            self.assertEqual(spacing_faults(root/'measured.pptx'),[])

@unittest.skipUnless(LIBREOFFICE,'LibreOffice is required')
class DigitalMeasurementTests(unittest.TestCase):
    def test_digital_rows_are_measured_and_rebuilt(self):
        import hashlib
        import run_weekly_workflow as workflow
        from test_weekly_digital_maturity import master_sectors
        with tempfile.TemporaryDirectory() as directory:
            root=Path(directory)
            with contextlib.redirect_stdout(io.StringIO()):workflow.scenario(root,'normal')
            pdf=root/'source.pdf';pdf.write_bytes(b'synthetic source identity')
            review=root/'review.json'
            review.write_text(json.dumps({
                'source_pdf_sha256':hashlib.sha256(pdf.read_bytes()).hexdigest(),'reviewed':True,
                'metrics':{'experience':15,'satisfaction':10,'complaints':20,'tools':65},
                'sectors':[{'name':'القطاع الأول','note':None,'metrics':{'experience':15,'satisfaction':10,'complaints':20,'tools':65}}],
                'updates':[{'perspective':'عام','text':f'تحديث رقم {i} عن سير العمل في المنصات'} for i in range(1,7)],
                'detail_sectors':master_sectors()},ensure_ascii=False),encoding='utf-8')
            inputs=['--tasks',root/'tasks.xlsx','--suhail',root/'projects.xlsx','--transactions',root/'transactions.xlsx']
            digital=['--digital-pdf',pdf,'--digital-data',review]
            result=run(WEEKLY/'build_report.py',*inputs,'--model','digital',*digital,'--output',root/'d0.pptx')
            self.assertEqual(result.returncode,0,result.stdout+result.stderr)
            result=run(WEEKLY/'measure_layout.py','--report',root/'d0.pptx',*inputs,*digital)
            self.assertEqual(result.returncode,0,result.stdout+result.stderr)
            data=json.loads((root/'d0.layout.json').read_text(encoding='utf-8'))
            self.assertEqual(len(data['digital']['summary']),6)
            self.assertEqual(set(data['digital']['summary_compact']),{'11','10.5','10','9.5','9','8.5','8'})
            sectors=master_sectors()
            self.assertEqual(sorted(data['digital']['sectors']),sorted(s['title'] for s in sectors))
            for sector in sectors:
                self.assertEqual(len(data['digital']['sectors'][sector['title']]['rows']),len(sector['rows']))
            result=run(WEEKLY/'build_report.py',*inputs,'--model','digital',*digital,
                       '--measurements',root/'d0.layout.json','--output',root/'d1.pptx')
            self.assertEqual(result.returncode,0,result.stdout+result.stderr)
            # Older layout files without digital rows are stale for the digital model.
            del data['digital'];stale=root/'stale.layout.json'
            stale.write_text(json.dumps(data,ensure_ascii=False),encoding='utf-8')
            result=run(WEEKLY/'build_report.py',*inputs,'--model','digital',*digital,
                       '--measurements',stale,'--output',root/'d2.pptx')
            self.assertEqual(result.returncode,0,result.stdout+result.stderr)
            self.assertTrue((root/'d2.pptx').exists())
            notices=json.loads((root/'d2.pptx.notices.json').read_text(encoding='utf-8'))
            self.assertTrue(any('ملف القياس' in json.dumps(n,ensure_ascii=False) for n in notices),notices)

if __name__=='__main__':
    unittest.main()
