# /// script
# requires-python = ">=3.12"
# dependencies = ["pandas>=2.2,<3", "openpyxl>=3.1,<4", "python-pptx>=1.0.2", "lxml>=5.3", "pillow"]
# ///
"""Reproducible synthetic demonstration of the manual weekly workflow.

This is a test harness, not a replacement for the skill's render-driven layout
decisions. It emits candidates and source reconciliation evidence; visual
acceptance is a separate final step. No ministry data is included.
"""
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


def transaction_data(path):
    wb=load_workbook(path,data_only=True);ws=wb['المعاملات']
    kpis=[ws.cell(r,2).value for r in range(2,6)]
    hr=next(r for r in range(1,ws.max_row+1) if ws.cell(r,1).value=='القطاع')
    chart=[]
    for r in range(hr+1,ws.max_row+1):
        if not ws.cell(r,1).value:break
        chart.append([ws.cell(r,c).value for c in range(1,5)])
    hr=next(r for r in range(1,ws.max_row+1) if ws.cell(r,1).value=='رقم المعاملة')
    records=[]
    for r in range(hr+1,ws.max_row+1):
        if ws.cell(r,1).value is not None:records.append([builder.val(ws.cell(r,c).value) for c in range(1,7)])
    wb.close()
    assert sum(kpis[1:])==kpis[0] and len(records)==kpis[3]
    assert [sum(int(row[c] or 0) for row in chart) for c in range(1,4)]==kpis[1:]
    return kpis,chart,records


def remove(shape):shape._element.getparent().remove(shape._element)


def flow(prs,seed,sections,summary=False):
    """Initial placement only; final visual acceptance must confirm pagination."""
    patterns={builder.table_title(sh):sh for sh in builder.tables(seed)}
    top=min(sh.top for sh in patterns.values())
    bottom=prs.slide_height-650000
    # Continuations preserve layout/navigation, but not the source KPI/chart.
    def keep(sh):
        outside_summary = sh.top + sh.height <= 700000 or sh.top >= bottom
        return (not sh.has_table and not sh.has_chart
                and not getattr(sh,'has_text_frame',False)
                and (not summary or outside_summary))
    current=seed;y=top;targets=[seed]
    for sh in list(builder.tables(seed)):remove(sh)
    for title,records,pattern in sections:
        if not records:continue
        fixed=sum(r.height for r in list(pattern.table.rows)[:2])
        base=builder._body_base_height(pattern.table,2)
        chunk=[];weights=[]
        def place():
            nonlocal y,chunk,weights
            element=deepcopy(pattern._element)
            current.shapes._spTree.insert_element_before(element,'p:extLst')
            sh=list(current.shapes)[-1];sh.top=y
            builder.fill_titled_table(sh,chunk,title=title,weights=weights)
            y+=sh.height+90000;chunk=[];weights=[]
        for row in records:
            weight=builder.row_units(row,pattern.table,11)
            if fixed+base*weight>bottom-top:raise ValueError('Fixture row exceeds available page')
            used=fixed+base*(sum(weights)+weight)
            if y+used>bottom:
                if chunk:place()
                current=builder.clone_slide(prs,seed,after=current,shape_predicate=keep)
                targets.append(current);y=900000 if summary else top
            chunk.append(row);weights.append(weight)
        if chunk:place()
    return targets


def update_chart(shape,records,statuses):
    data=CategoryChartData()
    # Retain the template category order for matching sectors, then append new ones.
    order=[str(cat.label) for cat in shape.chart.plots[0].categories]
    rank={name:i for i,name in enumerate(order)}
    records=sorted([row for row in records if any(float(value or 0)>0 for value in row[1:])],key=lambda row:rank.get(str(row[0]),len(rank)))
    rows=records or [[None]*(len(statuses)+1)]
    data.categories=[row[0] or '' for row in rows]
    for c,status in enumerate(statuses,1):data.add_series(status,[row[c] or None for row in rows])
    shape.chart.replace_data(data)
    chart=shape.chart._chartSpace
    blank=chart.find('.//'+qn('c:dispBlanksAs'))
    if blank is not None:blank.set('val','gap')
    return records


def populate(root):
    assert hashlib.sha256(MASTER.read_bytes()).hexdigest()==MASTER_HASH
    tasks=root/'tasks.xlsx';projects=root/'projects.xlsx';transactions=root/'transactions.xlsx'
    tk,tc,ts,td=builder.task_data(tasks);pk,pc,ps,pd=builder.suhail_data(projects);xk,xc,xd=transaction_data(transactions)
    prs=Presentation(MASTER);slides=list(prs.slides)
    date=dt.datetime.now().astimezone().date()
    builder.set_text(builder.shape_by_id(slides[0],19),f'{date.day} {builder.MONTHS[date.month-1]} {date.year}')
    chart_specs=[]
    for sl,kpis,ids,records,statuses,chart_name in [
        (slides[1],tk,[11,14,16,7,6],tc,['مكتملة','على المخطط','متأخر','معلق'],'chart1.xml'),
        (slides[2],pk,[10,12,14,8,7],pc,['مكتملة','على المخطط','متأخر','لم تبدأ'],'chart2.xml'),
        (slides[3],xk,[24,29,31,26],xc,['مكتملة','على المخطط','متأخر'],'chart3.xml')]:
        for sid,value in zip(ids,kpis):builder.set_text(builder.shape_by_id(sl,sid),value)
        ordered=update_chart(next(sh for sh in sl.shapes if sh.has_chart),records,statuses)
        chart_specs.append((chart_name,ordered,statuses))
    for sl,sections in [(slides[1],ts),(slides[2],ps)]:
        # Map source header order to the approved template header order.
        seq=[]
        for title,records in sections.items():
            pattern=next(sh for sh in builder.tables(sl) if builder.table_title(sh)==title)
            headers=[c.text.strip() for c in pattern.table.rows[1].cells]
            src=['#','المهمة' if sl==slides[1] else 'المشروع','القطاع',
                 'طلب الدعم' if title=='طلبات الدعم' else 'التحديث' if title=='أبرز التحديثات' else 'التحدي' if title=='التحديات' else 'ملاحظات']
            seq.append((title,[[row[src.index(h)] for h in headers] for row in records],pattern))
        flow(prs,sl,seq,True)
    tframe=next(sh for sh in builder.tables(slides[3]))
    headings=[splitter.field_key(c.text) for c in tframe.table.rows[1].cells]
    xheaders=[splitter.field_key(s) for s in ['رقم المعاملة','موضوع المعاملة','القطاع','الجهة الوارد منها المعاملة','تاريخ إنشاء المعاملة','تاريخ الإنجاز المخطط']]
    flow(prs,slides[3],[('المعاملات المتأخرة',[[row[xheaders.index(h)] for h in headings] for row in xd],tframe)],True)
    styles={}
    for sl in slides:
        for sh in builder.tables(sl,6):
            for row in list(sh.table.rows)[2:]:
                status=gate.canon(row.cells[4].text)
                if status in ('مكتملة','على المخطط'):styles[status]=deepcopy(row.cells[4]._tc.get_or_add_tcPr())
    generated={};unused=[]
    for kind,indices,sections in [('task',[5,6],td),('project',[7,8,9],pd)]:
        seed=slides[indices[0]];pattern=builder.tables(seed,6)[0]
        # Remove any additional sample tables before filling a single flowing stream.
        seq=[((f'تفاصيل المهام ({title})' if kind=='task' else title),[row[:6] for row in rows],pattern) for title,rows in sections]
        generated[kind]=flow(prs,seed,seq)
        if not any(rows for _,rows in sections):generated[kind]=[]
        for sl in generated[kind]:
            for sh in builder.tables(sl,6):
                for row in list(sh.table.rows)[2:]:
                    cell=row.cells[4];status=gate.canon(cell.text)
                    if status not in styles:raise ValueError('Fixture needs approved status style: '+status)
                    old=cell._tc.get_or_add_tcPr();cell._tc.remove(old);cell._tc.append(deepcopy(styles[status]))
                    builder.set_cell_text(cell,status)
                    if kind=='project':
                        for p in row.cells[5].text_frame.paragraphs:
                            for run in p.runs:run.font.bold=p.text.startswith('تاريخ التحديث ')
        unused.extend(slides[i] for i in indices[1:])
        if not generated[kind]:unused.append(seed)
    for slide in unused:builder.remove_slide(prs,slide)
    # Repoint all internal slide links from removed examples to their surviving section.
    retained={sl.part for sl in prs.slides}
    for part in list(prs.part.package.iter_parts()):
        for rel in list(part.rels.values()):
            if rel.is_external or not rel.reltype.endswith('/slide') or rel.target_part in retained:continue
            fallback=generated['task'] if rel.target_part in {slides[5].part,slides[6].part} else generated['project']
            if fallback:rel._target=fallback[0].part
            else:
                for el in list(part._element.iter()) if hasattr(part,'_element') else []:
                    if el.get(qn('r:id'))==rel.rId and el.tag in (qn('a:hlinkClick'),qn('a:hlinkMouseOver')):el.getparent().remove(el)
                part.drop_rel(rel.rId)
    for i,sl in enumerate(prs.slides,1):
        for sh in sl.shapes:
            if sh._element.find('.//'+qn('a:fld')) is not None and getattr(sh,'has_text_frame',False):
                for t in sh._element.findall('.//'+qn('a:t')):t.text=str(i)
    deck=root/'weekly.pptx';prs.save(deck)
    with tempfile.TemporaryDirectory() as td:
        with zipfile.ZipFile(deck) as z:z.extractall(td)
        included=splitter.reachable_parts(td)
        with zipfile.ZipFile(deck,'w',zipfile.ZIP_DEFLATED) as z:
            for name in sorted(included):z.write(Path(td)/name,name)
    verify(root,deck,chart_specs)
    return deck


def verify(root,deck,chart_specs=None):
    """Recheck the final bytes after manual/PowerPoint layout adjustments."""
    tasks=root/'tasks.xlsx';projects=root/'projects.xlsx'
    tk,tc,ts,_=builder.task_data(tasks);pk,pc,ps,_=builder.suhail_data(projects)
    xk,xc,xd=transaction_data(root/'transactions.xlsx')
    if chart_specs is None:
        chart_specs=[];template=Presentation(MASTER)
        for index,records,statuses in [(1,tc,['مكتملة','على المخطط','متأخر','معلق']),(2,pc,['مكتملة','على المخطط','متأخر','لم تبدأ']),(3,xc,['مكتملة','على المخطط','متأخر'])]:
            shape=next(sh for sh in template.slides[index].shapes if sh.has_chart)
            order={str(c.label):i for i,c in enumerate(shape.chart.plots[0].categories)}
            ordered=sorted([row for row in records if any(float(value or 0)>0 for value in row[1:])],key=lambda row:order.get(str(row[0]),len(order)))
            chart_specs.append((f'chart{index}.xml',ordered,statuses))
    errors=[];gate.package(deck,errors)
    for chart_name,records,statuses in chart_specs:
        gate.chart_metrics(deck,'ppt/charts/'+chart_name,records,statuses,errors)
    actual_tasks=[];actual_projects=[];actual_summaries=defaultdict(list);actual_transactions=[]
    final=Presentation(deck)
    for label,kpis,ids in [('إجمالي المهام',tk,[11,14,16,7,6]),('إجمالي مشاريع سهيل',pk,[10,12,14,8,7]),('إجمالي العاملات',xk,[24,29,31,26])]:
        candidates=[sl for sl in final.slides if any(getattr(sh,'has_text_frame',False) and sh.text.strip()==label for sh in sl.shapes)]
        if len(candidates)!=1:errors.append(label+': expected one KPI slide')
        else:
            for sid,value in zip(ids,kpis):
                if builder.shape_by_id(candidates[0],sid).text.strip()!=str(value):errors.append(label+': KPI differs')
    if sum(sh.has_chart for sl in final.slides for sh in sl.shapes)!=3:errors.append('Expected exactly three charts')
    for sl in final.slides:
        for sh in builder.tables(sl):
            rows=[[gate.norm(c.text) for c in row.cells] for row in sh.table.rows]
            if len(rows)<3:errors.append('Empty table')
            title=next((c for c in rows[0] if c),'')
            if len(rows[1])==6:
                if 'المهمة' in rows[1]:actual_tasks.extend(rows[2:])
                else:actual_projects.extend([(title,*row) for row in rows[2:]])
            elif 'رقم المعاملة' in rows[1]:
                head=[splitter.field_key(h) for h in rows[1]]
                actual_transactions.extend([[row[head.index(splitter.field_key(h))] for h in splitter.TRANS_FIELDS] for row in rows[2:]])
            else:
                headers=rows[1];field='المهمة' if title in ts else 'المشروع'
                last='طلب الدعم' if title=='طلبات الدعم' else 'التحديث' if title=='أبرز التحديثات' else 'التحدي' if title=='التحديات' else 'ملاحظات'
                actual_summaries[title].extend([[row[headers.index(h)] for h in ['#',field,'القطاع',last]] for row in rows[2:]])
    norm=lambda row:tuple(gate.canon(v) for v in row)
    def compare(label,a,b):
        if list(map(norm,a))!=list(map(norm,b)):errors.append(label+' records/order differ')
    compare('tasks',[r[1:] for r in gate.read_tasks(tasks)],actual_tasks)
    compare('projects',gate.read_suhail(projects),actual_projects)
    for title,records in {**ts,**ps}.items():compare(title,records,actual_summaries[title])
    compare('transactions',[[row[i] for i in [0,1,3,4,5]] for row in xd],actual_transactions)
    if errors:raise AssertionError('\n'.join(errors))
    return deck


def scenario(root,name):
    count=0 if name=='empty' else 32 if name=='overflow' else 4
    result,_=fixtures.build_scenario(root,name,count,count,name=='new')
    require(result)
    # Weekly status exemplars currently cover completed/on-plan. Other statuses
    # remain exercised in Monday tests; they require visual style approval weekly.
    for raw,script,output in [('raw-tasks.xlsx',TASKS,'tasks.xlsx'),('raw-projects.xlsx',SUHAIL,'projects.xlsx')]:
        wb=load_workbook(root/raw);ws=wb.active
        for row in range(2,ws.max_row+1):ws.cell(row,5 if script==TASKS else 7).value='مكتمل' if row%2 else 'على المخطط'
        if script==TASKS and count:wb['طلبات الدعم'].append(['مهمة تجريبية 1',ws.cell(2,2).value,'طلب دعم تجريبي'])
        if script==SUHAIL and count:ws.cell(2,10).value='تحدي تجريبي'
        wb.save(root/raw)
        require(run(script,*(['--input',root/raw,'--output',root/output] if script==TASKS else [root/raw,root/output])))
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
