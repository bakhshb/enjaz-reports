# /// script
# requires-python = ">=3.12"
# dependencies = ["openpyxl>=3.1,<4", "python-pptx>=1.0.2", "lxml>=5.3", "pillow"]
# ///
"""Nav tabs must jump to the right page in consolidated and sector decks.

The approved masters carry two link faults: the المعاملات label has no link
(only its background does) and, on the digital model, the التجربة الرقمية label
jumps to the التفاصيل divider. Builders must rewire every nav layout.
Sector decks always show their three tabs; a tab whose section is absent
(usually المعاملات) stays visible with no link.
"""
import contextlib
import io
from pathlib import Path
import sys
import tempfile
import unittest

from pptx import Presentation
from pptx.oxml.ns import qn

ROOT=Path(__file__).resolve().parents[1]
WEEKLY=ROOT/'skills/weekly-consolidated-report/scripts'
SECTOR=ROOT/'skills/weekly-report-by-sector/scripts'
ASSETS=ROOT/'skills/weekly-consolidated-report/assets'
DIGITAL=ASSETS/'weekly-report-digital-master.pptx'
CLASSIC=ASSETS/'weekly-report-master.pptx'
for path in (WEEKLY,SECTOR,ROOT/'skills'):sys.path.insert(0,str(path))
import weekly_artwork as artwork
import split_report as splitter
import qa_report as sector_qa
from test_upstream_integrity import run


def jump_targets(layout,shape):
    out=[]
    for link in shape.element.iter(qn('a:hlinkClick')):
        out.append(layout.part.rels[link.get(qn('r:id'))].target_part)
    return out


class ConsolidatedNavigationTests(unittest.TestCase):
    def check(self,master,names):
        raw=Presentation(master);errors=[]
        artwork.check_navigation_links(raw,errors)
        self.assertTrue(any('المعاملات' in e for e in errors),'master fault no longer present?')
        prs=Presentation(master);artwork.prepare_navigation(prs)
        errors=[];artwork.check_navigation_links(prs,errors);self.assertEqual(errors,[])
        destinations=artwork.targets(prs)
        layouts=artwork.navigation_layouts(prs)
        # Summary, divider and detail layouts are all rewired, not only detail ones.
        self.assertEqual(len(layouts),len({sl.slide_layout.part for sl in prs.slides
                                           if artwork.has_navigation(sl.slide_layout)}))
        for layout in layouts.values():
            for name,shapes in artwork.buttons(layout,destinations):
                self.assertIn(name,names)
                for shape in shapes:
                    self.assertEqual(jump_targets(layout,shape),[destinations[name]],(layout.name,name))

    def test_standard_master_links(self):
        self.check(CLASSIC,{'ملخص المهام','ملخص مشاريع سهيل','المعاملات','التفاصيل'})

    def test_digital_master_links(self):
        self.check(DIGITAL,{'ملخص المهام','ملخص مشاريع سهيل','المعاملات','التفاصيل','التجربة الرقمية'})
        prs=Presentation(DIGITAL);artwork.prepare_navigation(prs)
        digital=artwork.targets(prs)[artwork.DIGITAL_NAME]
        self.assertTrue(sum(sh.has_chart for sh in next(sl for sl in prs.slides if sl.part is digital).shapes)==5)

    def test_validator_rejects_wrong_destination(self):
        prs=Presentation(DIGITAL);artwork.prepare_navigation(prs)
        destinations=artwork.targets(prs)
        layout=next(iter(artwork.navigation_layouts(prs).values()))
        name,(text,_)=next((n,s) for n,s in artwork.buttons(layout,destinations) if n=='التجربة الرقمية')
        link=next(text.element.iter(qn('a:hlinkClick')))
        link.set(qn('r:id'),layout.part.relate_to(destinations['التفاصيل'],'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide'))
        errors=[];artwork.check_navigation_links(prs,errors)
        self.assertIn(f'Layout {layout.name}: التجربة الرقمية tab does not jump to its section',errors)


class SectorNavigationTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.temp=tempfile.TemporaryDirectory();cls.out=Path(cls.temp.name)/'out'
        result=run(SECTOR/'split_report.py',DIGITAL,'--outdir',cls.out,'--no-refine')
        assert result.returncode==0,result.stdout+result.stderr

    @classmethod
    def tearDownClass(cls):cls.temp.cleanup()

    def deck(self,sector):
        return Presentation(next(self.out.glob(f'*{sector}.pptx')))

    def test_tabs_stay_and_only_present_sections_link(self):
        # This sector has tasks only: all three tabs show, only ملخص المهام links.
        prs=self.deck('مركز التخطيط الاستراتيجي')
        tasks=next(sl for sl in prs.slides if any(getattr(sh,'has_text_frame',False)
                   and sh.text_frame.text.strip()=='إجمالي المهام' for sh in sl.shapes))
        for layout in {sl.slide_layout.part:sl.slide_layout for sl in prs.slides}.values():
            cols=dict(splitter.nav_columns(layout))
            if not cols:continue
            self.assertEqual(sorted(cols),sorted(splitter.NAV_KEEP_LABELS))
            for shape in cols['ملخص المهام']:self.assertEqual(jump_targets(layout,shape),[tasks.part])
            for name in ('ملخص مشاريع سهيل','المعاملات'):
                for shape in cols[name]:self.assertEqual(jump_targets(layout,shape),[],name)
        self.assertEqual(sector_qa.nav_errors(prs),[])

    def test_every_deck_passes_navigation_qa(self):
        for deck in self.out.glob('*.pptx'):
            with self.subTest(deck=deck.name):
                self.assertEqual(sector_qa.nav_errors(Presentation(deck)),[])

    def test_qa_rejects_dead_or_wrong_links(self):
        prs=self.deck('وكالة شؤون الحج')            # Suhail projects only
        layout=prs.slides[1].slide_layout;cols=dict(splitter.nav_columns(layout))
        for c in cols['ملخص مشاريع سهيل'][0].element.iter(qn('p:cNvPr')):
            for link in c.findall(qn('a:hlinkClick')):c.remove(link)
        errors=sector_qa.nav_errors(prs)
        self.assertIn(f'{layout.name}: ملخص مشاريع سهيل tab does not jump to its section',errors)
        prs=self.deck('وكالة شؤون الحج');layout=prs.slides[1].slide_layout
        splitter._set_jump(layout,dict(splitter.nav_columns(layout))['المعاملات'][0],prs.slides[1].part)
        self.assertIn(f'{layout.name}: المعاملات tab links although its section is absent',
                      sector_qa.nav_errors(prs))


if __name__=='__main__':
    unittest.main()
