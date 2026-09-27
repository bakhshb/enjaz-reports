#!/usr/bin/env python3
# /// script
# requires-python = ">=3.12"
# dependencies = [
#   "lxml>=5.3",
#   "openpyxl>=3.1",
#   "python-pptx>=1.0",
# ]
# ///
import argparse, copy, datetime as dt, os, re, shutil, subprocess, sys, tempfile, zipfile
from pathlib import Path
from lxml import etree
import openpyxl
from pptx import Presentation

NS_C = {'c':'http://schemas.openxmlformats.org/drawingml/2006/chart'}
MONTHS = ['يناير','فبراير','مارس','أبريل','مايو','يونيو','يوليو','أغسطس','سبتمبر','أكتوبر','نوفمبر','ديسمبر']

def val(v):
    if v is None: return ''
    if isinstance(v, (dt.datetime, dt.date)): return v.strftime('%d-%m-%Y')
    if isinstance(v, float) and v.is_integer(): return str(int(v))
    return str(v)

def set_text(shape, value):
    tf=shape.text_frame; text=val(value)
    if tf.paragraphs and tf.paragraphs[0].runs:
        runs=tf.paragraphs[0].runs; runs[0].text=text
        for r in runs[1:]: r.text=''
        for p in tf.paragraphs[1:]:
            for r in p.runs: r.text=''
    else: tf.text=text

def shape_by_id(slide, sid):
    def walk(items):
        for sh in items:
            if sh.shape_id==sid: return sh
            if getattr(sh,'shape_type',None)==6:
                x=walk(sh.shapes)
                if x: return x
    x=walk(slide.shapes)
    if not x: raise KeyError(f'shape {sid} not found')
    return x

def resize_table(table, rows_needed):
    tbl=table._tbl
    while len(table.rows)<rows_needed:
        tbl.append(copy.deepcopy(tbl.tr_lst[-1]))
    while len(table.rows)>rows_needed:
        tbl.remove(tbl.tr_lst[-1])

def fill_table(shape, rows, title=None):
    t=shape.table; resize_table(t, 2+len(rows))
    if title is not None:
        for j,c in enumerate(t.rows[0].cells): c.text = title if j==0 else ''
    for i,row in enumerate(rows,2):
        for j,x in enumerate(row): set_text(t.cell(i,j),x)

def header_map(ws, row):
    return {str(c.value).strip():c.column for c in ws[row] if c.value is not None}

def agenda_rows(path):
    ws=openpyxl.load_workbook(path,data_only=True)['Sheet1']; hm=header_map(ws,2); out=[]; found=False
    for r in range(3,ws.max_row+1):
        topic=val(ws.cell(r,hm['الموضوع']).value).strip()
        order=ws.cell(r,hm['الترتيب']).value
        if order not in (None,''):
            out.append((f'{int(order):02d}',topic,val(ws.cell(r,hm['المتحدث']).value),val(ws.cell(r,hm['الوقت المطلوب من المالك']).value)))
        if re.sub(r'\s+',' ',topic)=='ملخص الاجتماع': found=True; break
    if not found: raise ValueError('ملخص الاجتماع is missing from topics workbook')
    return out

def find_header(ws, labels, start=1):
    labels=[x.strip() for x in labels]
    for r in range(start,ws.max_row+1):
        got=[val(ws.cell(r,c).value).strip() for c in range(1,len(labels)+1)]
        if got==labels: return r
    raise ValueError(f'header not found: {labels}')

def read_block(ws, header_row, cols, stop_titles=()):
    out=[]
    for r in range(header_row+1,ws.max_row+1):
        first=val(ws.cell(r,1).value).strip()
        if first in stop_titles or (first and all(ws.cell(r,c).value is None for c in range(2,cols+1))): break
        row=[ws.cell(r,c).value for c in range(1,cols+1)]
        if any(x not in (None,'') for x in row): out.append(row)
    return out

def detail_sections(ws, cols):
    out=[]; r=1
    expected_task=['#','المهمة','القطاع','تاريخ الانجاز المخطط','الحالة','ملاحظات','الجهات ذات العلاقة']
    expected_proj=['#','اسم المشروع','تاريخ البداية','تاريخ النهاية','الحالة','ما تم حتى تاريخه']
    expected=expected_task if cols==7 else expected_proj
    while r<=ws.max_row:
        got=[val(ws.cell(r,c).value).strip() for c in range(1,cols+1)]
        if got==expected:
            title=val(ws.cell(r-1,1).value).strip()
            title=re.sub(r'^مهام مصدر\s*\((.*)\)$',r'\1',title)
            out.append((title,read_block(ws,r,cols)))
        r+=1
    return out

def task_data(path):
    wb=openpyxl.load_workbook(path,data_only=True); s=wb['ملخص المهام']
    k=[s.cell(r,2).value for r in range(2,7)]
    hr=find_header(s,['القطاع','مكتملة','على المخطط','متأخر','معلق'])
    chart=read_block(s,hr,5)
    cr=find_header(s,['#','المهمة','القطاع'],hr+1); completed=read_block(s,cr,3,('المهام المتأخرة','طلبات الدعم'))
    return k,chart,completed,detail_sections(wb['تفاصيل المهام'],7)

def suhail_data(path):
    wb=openpyxl.load_workbook(path,data_only=True); s=wb['ملخص مشاريع سهيل']
    k=[s.cell(3,c).value for c in range(1,6)]
    hr=find_header(s,['القطاع','مكتملة','على المخطط','متأخر','لم تبدأ'])
    chart=read_block(s,hr,5)
    ur=find_header(s,['#','المشروع','القطاع','التحديث'],hr+1); updates=read_block(s,ur,4)
    return k,chart,updates,detail_sections(wb['تفاصيل مشاريع سهيل'],6)

def find_tables_by_title(prs, slide_nums):
    d={}
    for n in slide_nums:
        for sh in prs.slides[n-1].shapes:
            if sh.has_table:
                title=sh.table.cell(0,0).text.strip()
                d[title]=(n,sh)
    return d

def save_chart_payload(rows, path):
    wb=openpyxl.load_workbook(path); ws=wb.active
    maxr=max(ws.max_row,len(rows)+1)
    for r in range(2,maxr+1):
        for c in range(1,6): ws.cell(r,c).value=None
    for r,row in enumerate(rows,2):
        for c,x in enumerate(row,1): ws.cell(r,c).value=x
    wb.save(path)

def update_chart_xml(xml_bytes, rows):
    root=etree.fromstring(xml_bytes)
    root_range=len(rows)+1
    for ser_i,ser in enumerate(root.xpath('.//c:ser',namespaces=NS_C)):
        if ser_i>=4: continue
        for f in ser.xpath('.//c:f',namespaces=NS_C):
            f.text=re.sub(r'\$\d+$',f'${root_range}',f.text)
        cats=[val(r[0]) for r in rows]
        nums=[r[ser_i+1] for r in rows]
        for cache,values,numeric in [(ser.find('.//c:cat//c:strCache',NS_C),cats,False),(ser.find('.//c:val//c:numCache',NS_C),nums,True)]:
            if cache is None: continue
            for p in list(cache.findall('c:pt',NS_C)): cache.remove(p)
            pc=cache.find('c:ptCount',NS_C)
            if pc is not None: pc.set('val',str(len(values)))
            for idx,x in enumerate(values):
                if x in (None,''): continue
                pt=etree.SubElement(cache,'{%s}pt'%NS_C['c'],idx=str(idx)); v=etree.SubElement(pt,'{%s}v'%NS_C['c']); v.text=val(x)
    disp=root.find('.//c:dispBlanksAs',NS_C)
    if disp is None: disp=etree.SubElement(root,'{%s}dispBlanksAs'%NS_C['c'])
    disp.set('val','gap')
    return etree.tostring(root,xml_declaration=True,encoding='UTF-8',standalone=True)

def patch_charts(pptx, task_rows, suhail_rows):
    tmp=Path(tempfile.mkdtemp()); src=tmp/'src'; src.mkdir()
    with zipfile.ZipFile(pptx) as z: z.extractall(src)
    for chart_name,book_name,rows in [('chart1.xml','Microsoft_Excel_Worksheet.xlsx',task_rows),('chart2.xml','Microsoft_Excel_Worksheet1.xlsx',suhail_rows)]:
        save_chart_payload(rows,src/'ppt'/'embeddings'/book_name)
        cp=src/'ppt'/'charts'/chart_name; cp.write_bytes(update_chart_xml(cp.read_bytes(),rows))
    out=tmp/'out.pptx'
    with zipfile.ZipFile(out,'w',zipfile.ZIP_DEFLATED) as z:
        for p in src.rglob('*'):
            if p.is_file(): z.write(p,p.relative_to(src))
    shutil.copy2(out,pptx); shutil.rmtree(tmp)

def main():
    ap=argparse.ArgumentParser(); ap.add_argument('--topics',required=True); ap.add_argument('--tasks',required=True); ap.add_argument('--suhail',required=True); ap.add_argument('--output',required=True); ap.add_argument('--template')
    a=ap.parse_args(); template=Path(a.template) if a.template else Path(__file__).resolve().parents[1]/'template'/'monday-meeting-master.pptx'
    if not template.exists() and not a.template:
        subprocess.run([sys.executable,str(Path(__file__).with_name('restore_template.py'))],check=True)
    if not template.exists(): raise FileNotFoundError(template)
    out=Path(a.output); shutil.copy2(template,out)
    prs=Presentation(out)
    today=dt.datetime.now().astimezone().date(); set_text(shape_by_id(prs.slides[0],4),f'{today.day} {MONTHS[today.month-1]} {today.year}')
    agenda=agenda_rows(a.topics); fill_table(shape_by_id(prs.slides[1],7),agenda,title='م')
    tk,tchart,completed,tdetails=task_data(a.tasks)
    for sid,x in zip([21,23,25,19,18],tk): set_text(shape_by_id(prs.slides[4],sid),x)
    fill_table(shape_by_id(prs.slides[4],48),[(r[2],r[1],r[0]) for r in completed],title='المهام المكتملة')
    task_targets={'اجتماع القيادات':[(11,5),(12,3)],'اجتماع التجربة الرقمية':[(12,18)],'لجنة المتابعة':[(13,4)]}
    for title,rows in tdetails:
        targets=task_targets.get(title,[]); cap=sum(len(shape_by_id(prs.slides[s-1],i).table.rows)-2 for s,i in targets)
        if len(rows)>cap: raise ValueError(f'task detail overflow requires a new continuation pattern: {title}')
        pos=0
        for s,i in targets:
            sh=shape_by_id(prs.slides[s-1],i); n=min(len(rows)-pos,len(sh.table.rows)-2)
            fill_table(sh,rows[pos:pos+n],title=f'تفاصيل المهام ({title})'); pos+=n
    sk,schart,updates,sdetails=suhail_data(a.suhail)
    for sid,x in zip([5,8,11,17,13],sk): set_text(shape_by_id(prs.slides[14],sid),x)
    set_text(shape_by_id(prs.slides[14],14),'لم تبدأ')
    fill_table(shape_by_id(prs.slides[14],44),[(r[3],r[2],r[1],r[0]) for r in updates],title='أبرز التحديثات')
    ptargets=find_tables_by_title(prs,[16,17,18,19])
    for title,rows in sdetails:
        if title not in ptargets: raise ValueError(f'new Suhail sector requires a template pattern: {title}')
        _,sh=ptargets[title]
        if len(rows)>len(sh.table.rows)-2: raise ValueError(f'Suhail detail overflow requires continuation: {title}')
        fill_table(sh,rows,title=title)
    prs.save(out); patch_charts(out,tchart,schart)
    print(out.resolve())
if __name__=='__main__': main()
