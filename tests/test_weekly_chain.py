# /// script
# requires-python = ">=3.12"
# dependencies = ["pandas>=2.2,<3", "openpyxl>=3.1,<4", "python-pptx>=1.0.2", "lxml>=5.3", "pillow"]
# ///
"""Weekly-source handoff and independent sector identity checks."""
import contextlib
import io
import json
from pathlib import Path
import shutil
import tempfile
import unittest
import zipfile
from pptx import Presentation

import run_weekly_workflow as workflow

qa=workflow.fixtures.sector_qa
run=workflow.run
SECTOR=workflow.fixtures.SECTOR


class WeeklyChainTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.temp=tempfile.TemporaryDirectory();cls.root=Path(cls.temp.name)
        workflow.scenario(cls.root/'baseline','normal')

    @classmethod
    def tearDownClass(cls):cls.temp.cleanup()

    def setUp(self):
        self.case=tempfile.TemporaryDirectory();self.path=Path(self.case.name)
        self.master=self.root/'baseline/weekly.pptx';self.transactions=self.root/'baseline/transactions.xlsx'
        self.out=self.path/'out';shutil.copytree(self.root/'baseline/sectors',self.out)
        self.deck=next(self.out.glob('*.pptx'))
        manifest=json.loads((self.out/'manifest.json').read_text(encoding='utf-8'))
        for entry in manifest:entry['out']=str(self.deck)
        (self.out/'manifest.json').write_text(json.dumps(manifest,ensure_ascii=False),encoding='utf-8')

    def tearDown(self):self.case.cleanup()

    def check(self):
        with contextlib.redirect_stdout(io.StringIO()):
            return qa.validate(self.out) and qa.check(self.out,self.master,self.transactions)

    def test_normal_chain_and_transaction_identifier(self):
        self.assertTrue(self.check())
        data=qa.actual_records(Presentation(self.deck))
        self.assertEqual(data['المعاملات المتأخرة'][0][0],'000123')

    def test_cover_corruption_rejected(self):
        prs=Presentation(self.deck)
        group=prs.slides[0].shapes[0]
        next(sh for sh in group.shapes if sh.name=='SectorName').text='قطاع خاطئ'
        prs.save(self.deck);self.assertFalse(self.check())

    def test_task_kpi_corruption_rejected(self):
        prs=Presentation(self.deck)
        next(sh for sh in prs.slides[1].shapes if sh.shape_id==11).text='999'
        prs.save(self.deck);self.assertFalse(self.check())

    def test_transaction_kpi_corruption_rejected(self):
        prs=Presentation(self.deck)
        slide=next(sl for sl in prs.slides if any(getattr(sh,'has_text_frame',False) and sh.text=='إجمالي المعاملات المتأخرة' for sh in sl.shapes))
        next(sh for sh in slide.shapes if sh.shape_id==24).text='999'
        prs.save(self.deck);self.assertFalse(self.check())

    def test_missing_transaction_mapping_rejected(self):
        result=run(SECTOR/'split_report.py',self.master,'--outdir',self.path/'missing','--no-refine')
        self.assertNotEqual(result.returncode,0);self.assertIn('--transactions',result.stderr)

    def test_alias_and_only_selection(self):
        aliases=self.path/'aliases.json';aliases.write_text(json.dumps({'وكالة شؤون الحج':'قطاع موحد'},ensure_ascii=False),encoding='utf-8')
        out=self.path/'alias'
        args=['--transactions',self.transactions,'--aliases',aliases,'--only','قطاع موحد']
        workflow.require(run(SECTOR/'split_report.py',self.master,'--outdir',out,'--no-refine',*args))
        workflow.require(run(SECTOR/'qa_report.py',out,'--master',self.master,*args))
        self.assertEqual(len(list(out.glob('*.pptx'))),1)

    def test_selection_without_records_is_valid(self):
        out=self.path/'empty';args=['--transactions',self.transactions,'--only','لا يطابق أي قطاع']
        workflow.require(run(SECTOR/'split_report.py',self.master,'--outdir',out,'--no-refine',*args))
        workflow.require(run(SECTOR/'qa_report.py',out,'--master',self.master,*args))
        self.assertEqual(list(out.glob('*.pptx')),[])

    def test_unlisted_slide_rejected(self):
        with zipfile.ZipFile(self.deck,'a') as package:
            package.writestr('ppt/slides/slide999.xml',package.read('ppt/slides/slide1.xml'))
        self.assertFalse(self.check())

    def test_extra_deck_rejected(self):
        shutil.copy2(self.deck,self.out/'stale.pptx');self.assertFalse(self.check())


if __name__=='__main__':unittest.main()
