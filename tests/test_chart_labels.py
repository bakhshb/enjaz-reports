# /// script
# requires-python = ">=3.12"
# dependencies = ["openpyxl>=3.1,<4", "python-pptx>=1.0.2", "lxml>=5.3"]
# ///
"""Every bar segment in the summary charts must show its value.

Both weekly masters hide labels on individual points chosen for their sample
data, and switch them off for the delayed and suspended series. After the data
is replaced those switches land on arbitrary bars, so builders must clear them.
"""
from pathlib import Path
import sys
import tempfile
import unittest

from pptx import Presentation

ROOT=Path(__file__).resolve().parents[1]
WEEKLY=ROOT/'skills/weekly-consolidated-report/scripts'
for path in (WEEKLY,ROOT/'skills'):sys.path.insert(0,str(path))
import build_report
from report_common.validation import chart_labels_visible

ASSETS=ROOT/'skills/weekly-consolidated-report/assets'
STATUSES=['مكتملة','على المخطط','متأخر','معلق']


class ChartLabelTests(unittest.TestCase):
    def test_masters_hide_some_labels(self):
        for master in (ASSETS/'weekly-report-master.pptx',ASSETS/'weekly-report-digital-master.pptx'):
            errors=[]
            chart_labels_visible(master,'ppt/charts/chart1.xml',errors)
            self.assertTrue(errors,master.name)

    def test_every_segment_labelled_after_update(self):
        for master in (ASSETS/'weekly-report-master.pptx',ASSETS/'weekly-report-digital-master.pptx'):
            prs=Presentation(master)
            charts=[sh for sl in prs.slides for sh in sl.shapes if sh.has_chart][:2]
            for shape in charts:
                names=[str(c.label) for c in shape.chart.plots[0].categories]
                records=[[name,1,2,3,4] for name in names]
                build_report.update_chart(shape,records,STATUSES)
            with tempfile.TemporaryDirectory() as tmp:
                out=Path(tmp)/'deck.pptx';prs.save(out)
                for shape in charts:
                    errors=[]
                    chart_labels_visible(out,str(shape.chart.part.partname).lstrip('/'),errors)
                    self.assertEqual(errors,[],master.name)


if __name__=='__main__':
    unittest.main()
