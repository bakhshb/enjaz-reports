# /// script
# requires-python = ">=3.12"
# dependencies = ["openpyxl>=3.1,<4", "python-pptx>=1.0.2", "lxml>=5.3", "pillow"]
# ///
"""The sector splitter must survive master-deck structure changes.

Regression cases from the 2026-11-05 master: a grouped tasks KPI card,
digital-experience التحديات tables without a القطاع column, and a five-tab
nav bar at new coordinates. The bundled digital master has all three.
"""
import contextlib
from copy import deepcopy
import io
from pathlib import Path
import sys
import tempfile
import unittest

from pptx import Presentation
from pptx.util import Emu

ROOT=Path(__file__).resolve().parents[1]
SECTOR=ROOT/'skills/weekly-report-by-sector/scripts'
ASSETS=ROOT/'skills/weekly-consolidated-report/assets'
DIGITAL=ASSETS/'weekly-report-digital-master.pptx'
CLASSIC=ASSETS/'weekly-report-master.pptx'
sys.path.insert(0,str(SECTOR))
import split_report as splitter
import qa_report as sector_qa
from test_upstream_integrity import run


def quiet(fn,*args):
    with contextlib.redirect_stdout(io.StringIO()):return fn(*args)


class GroupedKpiTests(unittest.TestCase):
    def test_tasks_kpi_inside_group_is_located_and_flattened(self):
        raw=Presentation(DIGITAL)
        self.assertTrue(any(sh.shape_type==6 for sh in raw.slides[1].shapes))
        with self.assertRaises(SystemExit):quiet(splitter.locate,raw)
        prs=splitter.open_master(DIGITAL)
        role=quiet(splitter.locate,prs)
        self.assertEqual(role['tasks'],1)
        kpi=splitter.kpi_shapes(prs.slides[role['tasks']])
        self.assertTrue({'إجمالي','مكتملة','على المخطط','متأخرة','معلقة'}<=set(kpi))

    def test_flatten_keeps_geometry_and_skips_transformed_groups(self):
        prs=Presentation(DIGITAL);slide=prs.slides[1]
        group=next(sh for sh in slide.shapes if sh.shape_type==6)
        before={sh.shape_id:(sh.left,sh.top,sh.width,sh.height) for sh in group.shapes}
        moved=deepcopy(group._element)          # same group with a real offset
        moved.find(splitter.qp('grpSpPr')).find(splitter.q('xfrm')).find(splitter.q('off')).set('x','0')
        slide.shapes._spTree.append(moved)
        splitter.flatten_kpi_groups(prs)
        top={sh.shape_id:(sh.left,sh.top,sh.width,sh.height) for sh in slide.shapes}
        for sid,geom in before.items():self.assertEqual(top.get(sid),geom)
        self.assertEqual(sum(sh.shape_type==6 for sh in slide.shapes),1)


class DigitalChallengeTests(unittest.TestCase):
    def test_challenge_tables_without_sector_are_not_project_summaries(self):
        prs=splitter.open_master(DIGITAL)
        role=quiet(splitter.locate,prs)
        self.assertEqual(role['project_summaries'],[2])
        quiet(splitter.read_data,prs,role)


class NavTests(unittest.TestCase):
    def check_nav(self,master):
        prs=Presentation(master);splitter.strip_and_centre_nav(prs)
        seen=0
        for layout in prs.slide_masters[0].slide_layouts:
            names,span=splitter.nav_tabs(layout)
            if not names:continue
            seen+=1
            self.assertEqual(sorted(names),sorted(splitter.NAV_KEEP_LABELS),layout.name)
            self.assertLessEqual(abs(span[0]+span[1]-prs.slide_width),2,layout.name)
        self.assertGreater(seen,0)
        # Links are wired per sector deck in build(); see test_navigation_links.py.
        self.assertEqual([e for e in sector_qa.nav_errors(prs) if 'jump' not in e and 'links' not in e],[])

    def test_five_tab_nav_is_reduced_and_centred(self):self.check_nav(DIGITAL)

    def test_four_tab_nav_is_reduced_and_centred(self):self.check_nav(CLASSIC)

    def test_qa_rejects_extra_or_uncentred_tabs(self):
        prs=Presentation(DIGITAL)
        self.assertTrue(sector_qa.nav_errors(prs))             # five tabs
        splitter.strip_and_centre_nav(prs)
        layout=prs.slides[1].slide_layout
        for sh in layout.shapes:
            if sh.top is not None and sh.top<splitter.NAV_ZONE and sh.width==1252892:
                sh.left=Emu(sh.left+400000)
        self.assertIn(f'{layout.name}: nav tabs are not centred',sector_qa.nav_errors(prs))


class EndToEndTests(unittest.TestCase):
    def test_digital_master_splits_without_preprocessing(self):
        with tempfile.TemporaryDirectory() as directory:
            out=Path(directory)/'out';before=DIGITAL.read_bytes()
            result=run(SECTOR/'split_report.py',DIGITAL,'--outdir',out,'--no-refine')
            self.assertEqual(result.returncode,0,result.stdout+result.stderr)
            result=run(SECTOR/'qa_report.py',out,'--master',DIGITAL)
            self.assertEqual(result.returncode,0,result.stdout+result.stderr)
            self.assertIn('ALL OK',result.stdout)
            self.assertEqual(before,DIGITAL.read_bytes())
            src=splitter.open_master(DIGITAL);role=quiet(splitter.locate,src)
            tasks=quiet(splitter.read_data,src,role)[0]
            total=0
            for deck in out.glob('*.pptx'):
                prs=Presentation(deck)
                for slide in prs.slides:
                    kpi=splitter.kpi_shapes(slide)
                    if 'معلقة' in kpi:total+=int(kpi['إجمالي']['num'].text_frame.text)
            self.assertEqual(total,len(tasks))


if __name__=='__main__':
    unittest.main()
