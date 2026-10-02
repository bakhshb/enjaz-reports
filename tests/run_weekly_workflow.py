# /// script
# requires-python = ">=3.12"
# dependencies = ["pandas>=2.2,<3", "openpyxl>=3.1,<4", "python-pptx>=1.0.2", "lxml>=5.3", "pillow"]
# ///
"""Synthetic acceptance scenarios exercising the supported report entry points."""
import argparse
from collections import defaultdict
from copy import deepcopy
import datetime as dt
import hashlib
import json
from pathlib import Path
import tempfile
import zipfile

from lxml import etree
from openpyxl import load_workbook
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.oxml.ns import qn

import test_deck_integrity as fixtures
from test_upstream_integrity import TRANSACTIONS, TASKS, SUHAIL, run, transaction_source

ROOT=fixtures.ROOT
builder=fixtures.load_module('monday_builder',fixtures.MONDAY/'build_report.py')
gate=fixtures.gate
splitter=fixtures.sector_qa.splitter
MASTER=ROOT/'skills/weekly-consolidated-report/assets/weekly-report-master.pptx'
MASTER_HASH='3e9b018c223c76cbb0e435cd543732801ec8685fc4ec827d4a7f99b29e0423e5'


def require(result):
    if result.returncode:raise AssertionError(result.stdout+result.stderr)
    return result.stdout


# Tests exercise production code; production never imports this test module.
weekly=fixtures.load_module('weekly_builder',ROOT/'skills/weekly-consolidated-report/scripts/build_report.py')
import validate_report as weekly_validation
flow=weekly.flow
update_chart=weekly.update_chart
transaction_data=weekly.transaction_data

def populate(root):
    weekly.main(['--tasks',str(root/'tasks.xlsx'),'--suhail',str(root/'projects.xlsx'),
                 '--transactions',str(root/'transactions.xlsx'),'--output',str(root/'weekly.pptx')])
    return root/'weekly.pptx'

def verify(root,deck,chart_specs=None):
    weekly_validation.validate(deck,root/'tasks.xlsx',root/'projects.xlsx',root/'transactions.xlsx')
    return deck

def scenario(root,name):
    count=0 if name=='empty' else 32 if name=='overflow' else 4
    result,_=fixtures.build_scenario(root,name,count,count,name=='new')
    require(result)
    # Exercise all approved task/project statuses with independent summary sections.
    for raw,script,output in [('raw-tasks.xlsx',TASKS,'tasks.xlsx'),('raw-projects.xlsx',SUHAIL,'projects.xlsx')]:
        wb=load_workbook(root/raw);ws=wb.active
        wb.save(root/raw)
        require(run(script,*(['--input',root/raw,'--output',root/output] if script==TASKS else [root/raw,root/output])))
    # Simulate the user's manual entries after upstream generation.
    # The upstream builders themselves must leave these sections blank.
    if count:
        for file,sheet,title,text in [('tasks.xlsx','ملخص المهام','طلبات الدعم','طلب دعم تجريبي'),
                                      ('projects.xlsx','ملخص مشاريع سهيل','سادسا: التحديات','تحدي تجريبي')]:
            wb=load_workbook(root/file);ws=wb[sheet]
            row=next(c.row for cells in ws for c in cells if c.value==title)+2
            record_name='مهمة تجريبية 1' if file=='tasks.xlsx' else 'مشروع تجريبي 1'
            for col,value in enumerate([1,record_name,'قطاع جديد تجريبي' if name=='new' else 'وكالة شؤون الحج',text],1):ws.cell(row,col,value)
            wb.save(root/file)
    # Rebuild Monday from the exact final workbooks used by the weekly path.
    require(run(fixtures.MONDAY/'build_report.py','--topics',root/'topics.xlsx','--tasks',root/'tasks.xlsx','--suhail',root/'projects.xlsx','--output',root/'monday.pptx'))
    require(run(fixtures.MONDAY/'report_gate_v2.py','--report',root/'monday.pptx','--topics',root/'topics.xlsx','--tasks',root/'tasks.xlsx','--suhail',root/'projects.xlsx'))
    txdir=root/'raw-transactions';txdir.mkdir(exist_ok=True)
    sector='قطاع جديد تجريبي' if name=='new' else 'وكالة شؤون الحج'
    transaction_source(txdir/'source.xlsx',sector=sector,empty=count==0)
    if name=='overflow':
        wb=load_workbook(txdir/'source.xlsx');ws=wb.active
        for r in range(6,29):
            for c in [20,24,27,28,29]:ws.cell(r,c).value=ws.cell(5,c).value
            ws.cell(r,31).value=f'{r:06d}'
        wb.save(txdir/'source.xlsx')
    require(run(TRANSACTIONS,'--input-dir',txdir,'--output',root/'transactions.xlsx'))
    deck=populate(root)
    out=root/'sectors'
    log=require(run(fixtures.SECTOR/'split_report.py',deck,'--transactions',root/'transactions.xlsx','--outdir',out,'--no-refine'))
    log+=require(run(fixtures.SECTOR/'qa_report.py',out,'--master',deck,'--transactions',root/'transactions.xlsx'))
    evidence={'scenario':name,'programmatic':'passed','visual':'pending','split_log':log,
              'hashes':{str(p.relative_to(root)):hashlib.sha256(p.read_bytes()).hexdigest() for p in root.rglob('*') if p.suffix in ('.pptx','.xlsx')}}
    (root/'evidence.json').write_text(json.dumps(evidence,ensure_ascii=False,indent=2),encoding='utf-8')
    return evidence


if __name__=='__main__':
    parser=argparse.ArgumentParser();parser.add_argument('--outdir',type=Path,required=True);parser.add_argument('--scenario',choices=['normal','empty','new','overflow'])
    args=parser.parse_args()
    for name in ([args.scenario] if args.scenario else ['normal','empty','new','overflow']):
        scenario(args.outdir/name,name);print(name+': programmatic checks passed; visual acceptance pending',flush=True)
