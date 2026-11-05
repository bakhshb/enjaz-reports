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
            inputs=['--tasks',root/'tasks.xlsx','--suhail',root/'projects.xlsx','--transactions',root/'transactions.xlsx']
            result=run(WEEKLY/'measure_layout.py','--report',root/'weekly.pptx',*inputs)
            self.assertEqual(result.returncode,0,result.stdout+result.stderr)
            layout=root/'weekly.layout.json'
            data=json.loads(layout.read_text(encoding='utf-8'))
            self.assertEqual(data['typography'],'11pt-suhail-paragraph-updates-v3')
            self.assertTrue(data['tables'])
            result=run(WEEKLY/'build_report.py',*inputs,'--measurements',layout,'--output',root/'measured.pptx')
            self.assertEqual(result.returncode,0,result.stdout+result.stderr)
            self.assertLessEqual(len(Presentation(root/'measured.pptx').slides),len(Presentation(root/'weekly.pptx').slides))


if __name__=='__main__':
    unittest.main()
