#!/usr/bin/env python3
# /// script
# requires-python = ">=3.12"
# dependencies = [
#   "lxml>=5.3",
#   "openpyxl>=3.1",
#   "python-pptx>=1.0",
# ]
# ///
import argparse, copy, datetime as dt, hashlib, math, re, shutil, statistics, tempfile, zipfile
from pathlib import Path
from lxml import etree
import openpyxl
from pptx import Presentation
from pptx.oxml.xmlchemy import OxmlElement
from pptx.oxml.ns import qn

NS_C = {'c':'http://schemas.openxmlformats.org/drawingml/2006/chart'}
MONTHS = ['يناير','فبراير','مارس','أبريل','مايو','يونيو','يوليو','أغسطس','سبتمبر','أكتوبر','نوفمبر','ديسمبر']
AGENDA_HEADERS=['م','جدول الأعمال','المسؤول','المدة الزمنية (بالدقيقة)']
MASTER_SHA256='86a1ed90844059c88d9da10a33c81cd14902a6b99c02cfa37661f527e774ab4a'
TASK_SUMMARY_TITLES=('المهام المكتملة','المهام المتأخرة','المهام المعلقة','طلبات الدعم')
SUHAIL_SUMMARY_TITLES=('أبرز التحديثات','المشاريع المتأخرة','التحديات')


def val(v):
    if v is None: return ''
    if isinstance(v, (dt.datetime, dt.date)): return v.strftime('%d-%m-%Y')
    if isinstance(v, float) and v.is_integer(): return str(int(v))
    return str(v)


def set_text(shape, value):
    tf=shape.text_frame; text=val(value)
    lines=text.splitlines() or ['']
    while len(tf.paragraphs)<len(lines):
        src=tf.paragraphs[1]._p if len(tf.paragraphs)>1 else tf.paragraphs[0]._p
        tf._txBody.append(copy.deepcopy(src))
    for i,p in enumerate(tf.paragraphs):
        line=lines[i] if i<len(lines) else ''
        if p.runs:
            p.runs[0].text=line
            for r in p.runs[1:]: r.text=''
        elif line:
            p.add_run().text=line


def clear_cell_bullets(cell):
    for p in cell.text_frame.paragraphs:
        pPr=p._p.get_or_add_pPr()
        for child in list(pPr):
            if child.tag in {qn('a:buChar'),qn('a:buAutoNum'),qn('a:buBlip'),qn('a:buNone')}:
                pPr.remove(child)
        pPr.insert(0,OxmlElement('a:buNone'))

def set_cell_text(cell,value):
    set_text(cell,value); clear_cell_bullets(cell)

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


def recursive_text(shape):
    out=[]
    if getattr(shape,'has_text_frame',False) and shape.text.strip(): out.append(shape.text.strip())
    if getattr(shape,'shape_type',None)==6:
        for child in shape.shapes: out.extend(recursive_text(child))
    return out


def slide_texts(slide):
    out=[]
    for sh in slide.shapes: out.extend(recursive_text(sh))
    return out


def resize_table(table, rows_needed, body_template_index=None):
    tbl=table._tbl
    if body_template_index is None:
        body_template_index=max(0,len(table.rows)-1)
    while len(table.rows)<rows_needed:
        src=table.rows[min(body_template_index,len(table.rows)-1)]._tr
        tbl.append(copy.deepcopy(src))
    while len(table.rows)>rows_needed:
        tbl.remove(tbl.tr_lst[-1])


def _body_base_height(table, fixed_rows):
    hs=[table.rows[i].height for i in range(fixed_rows,len(table.rows)) if table.rows[i].height]
    if hs: return int(statistics.median(hs))
    return 520000


def fill_titled_table(shape, rows, title=None, weights=None):
    t=shape.table
    base=_body_base_height(t,2)
    title_h=t.rows[0].height; header_h=t.rows[1].height
    resize_table(t,2+len(rows),2 if len(t.rows)>2 else len(t.rows)-1)
    if title is not None:
        for j,c in enumerate(t.rows[0].cells): set_cell_text(c,title if j==0 else '')
    for i,row in enumerate(rows,2):
        for j,x in enumerate(row): set_cell_text(t.cell(i,j),x)
    if weights is None: weights=[1]*len(rows)
    for i,w in enumerate(weights,2): t.rows[i].height=max(base,int(base*w))
    if title_h: t.rows[0].height=title_h
    if header_h: t.rows[1].height=header_h
    shape.height=sum((r.height or 0) for r in t.rows)


def fill_agenda_table(shape, rows, weights=None):
    t=shape.table
    base=_body_base_height(t,1); header_h=t.rows[0].height
    resize_table(t,1+len(rows),1 if len(t.rows)>1 else 0)
    for j,x in enumerate(AGENDA_HEADERS): set_cell_text(t.cell(0,j),x)
    for i,row in enumerate(rows,1):
        for j,x in enumerate(row): set_cell_text(t.cell(i,j),x)
    if weights is None: weights=[1]*len(rows)
    for i,w in enumerate(weights,1): t.rows[i].height=max(base,int(base*w))
    if header_h: t.rows[0].height=header_h
    shape.height=sum((r.height or 0) for r in t.rows)


def estimate_lines(text,width_inches,font_pt):
    text=val(text).strip()
    if not text: return 1
    # Arabic at the approved font sizes is typically around 10-13 visible
    # characters per inch; use a conservative estimate so long records move
    # to a continuation page before they become cramped.
    chars_per_line=max(6,int(width_inches*(12 if font_pt<=11 else 10)))
    return sum(max(1,math.ceil(len(part)/chars_per_line)) for part in text.splitlines() or [''])


def row_units(row, table, font_pt=11):
    widths=[c.width/914400 for c in table.columns]
    lines=max(estimate_lines(x,widths[i] if i<len(widths) else 1.5,font_pt) for i,x in enumerate(row))
    return max(1,math.ceil(lines/3))


def paginate_rows(rows, shape, fixed_rows=2, font_pt=11, capacity_units=None):
    t=shape.table
    if capacity_units is None: capacity_units=max(1,len(t.rows)-fixed_rows)
    pages=[]; cur=[]; weights=[]; used=0
    for row in rows:
        w=row_units(row,t,font_pt)
        if w>capacity_units:
            raise ValueError('A single row needs %d units but the template page holds %d; '
                             'shorten the source text or use a taller approved template row.'
                             % (w,capacity_units))
        if cur and used+w>capacity_units:
            pages.append((cur,weights)); cur=[]; weights=[]; used=0
        cur.append(row); weights.append(w); used+=w
    if cur or not pages: pages.append((cur,weights))
    return pages


def slide_index(prs,target):
    for i,sl in enumerate(prs.slides):
        if sl.part is target.part: return i
    raise ValueError('slide is not in presentation')


def clone_slide(prs,source,after=None,shape_predicate=None):
    dest=prs.slides.add_slide(source.slide_layout)
    for sh in list(dest.shapes):
        el=sh.element; el.getparent().remove(el)
    for sh in source.shapes:
        if shape_predicate is None or shape_predicate(sh):
            dest.shapes._spTree.insert_element_before(copy.deepcopy(sh.element),'p:extLst')
    for k,v in source._element.attrib.items(): dest._element.set(k,v)
    if after is not None:
        lst=prs.slides._sldIdLst; el=lst[-1]; lst.remove(el); lst.insert(slide_index(prs,after)+1,el)
    return dest


def remove_slide(prs,slide):
    idx=slide_index(prs,slide); sldId=prs.slides._sldIdLst[idx]; rid=sldId.rId
    prs.slides._sldIdLst.remove(sldId); prs.part.drop_rel(rid)


def tables(slide,cols=None):
    out=[sh for sh in slide.shapes if getattr(sh,'has_table',False)]
    return [sh for sh in out if cols is None or len(sh.table.columns)==cols]


def table_title(shape):
    return next((c.text.strip() for c in shape.table.rows[0].cells if c.text.strip()),'')


def detail_tables(slide,kind):
    marker='المهمة' if kind=='task' else 'اسم المشروع'
    return [sh for sh in tables(slide,6) if len(sh.table.rows)>1
            and sh.table.cell(1,0).text.strip()=='#'
            and sh.table.cell(1,1).text.strip()==marker]


def detail_pattern_slides(prs,kind):
    return [sl for sl in prs.slides if detail_tables(sl,kind)]


def first_table(slide,cols):
    xs=tables(slide,cols)
    if not xs: raise ValueError(f'{cols}-column table pattern missing')
    return xs[0]


def find_slide_with_text(prs,text,require_table=None):
    for sl in prs.slides:
        if text in slide_texts(sl):
            if require_table is None or bool(tables(sl))==require_table: return sl
    raise ValueError(f'slide pattern missing: {text}')


def fill_summary_tables(prs,source_slide,sections,titles,section_title):
    """Populate only nonempty template tables, flowing records onto extra pages."""
    patterns={table_title(sh):sh for sh in tables(source_slide) if table_title(sh) in titles}
    if set(patterns)!=set(titles):raise ValueError(f'{section_title}: incomplete table patterns')
    start_top=min(sh.top for sh in patterns.values())
    bottom=prs.slide_height-int(.9*914400)
    gap=int(.12*914400)
    pages=[[]]; used=start_top
    for title in titles:
        source_rows=sections[title]
        if not source_rows:continue
        # All new summary tables use four columns, displayed right-to-left.
        rows=[tuple(reversed(row)) for row in source_rows]
        shape=patterns[title]; table=shape.table
        base=_body_base_height(table,2)
        fixed=sum((table.rows[i].height or 0) for i in range(2))
        chunk=[]; weights=[]
        for row in rows:
            weight=row_units(row,table,11)
            row_h=base*weight
            if fixed+row_h>bottom-start_top:
                raise ValueError(f'{title}: one record is too tall for a summary continuation page')
            next_top=used+(gap if pages[-1] else 0)
            if next_top+fixed+row_h>bottom:
                if chunk:
                    pages[-1].append((title,chunk,weights,table_top))
                    chunk=[];weights=[]
                pages.append([]);used=start_top;next_top=used
            if not chunk:
                table_top=next_top
                used=table_top+fixed
            chunk.append(row);weights.append(weight);used+=row_h
        if chunk:pages[-1].append((title,chunk,weights,table_top))
    # Clone before editing the seed page; cloned pages retain its approved table styles.
    def keep(sh):
        if getattr(sh,'has_table',False):return table_title(sh) in titles
        if getattr(sh,'has_text_frame',False) and sh.text.strip()==section_title:return True
        return getattr(sh,'shape_type',None)==9 and sh.top<int(.65*914400)
    targets=[source_slide];last=source_slide
    for _ in pages[1:]:
        last=clone_slide(prs,source_slide,after=last,shape_predicate=keep);targets.append(last)
    for page_idx,(sl,items) in enumerate(zip(targets,pages)):
        by_title={table_title(sh):sh for sh in tables(sl) if table_title(sh) in titles}
        used_titles=set()
        for title,chunk,weights,top in items:
            sh=by_title[title];sh.top=top
            fill_titled_table(sh,chunk,weights=weights)
            used_titles.add(title)
        for title,sh in by_title.items():
            if title not in used_titles:sh.element.getparent().remove(sh.element)
        if page_idx:
            for sh in sl.shapes:
                if getattr(sh,'has_text_frame',False) and sh.text.strip()==section_title:
                    set_text(sh,section_title+' (تابع)')
    return targets


def validate_master_patterns(prs):
    agenda=[sh for sl in prs.slides for sh in tables(sl,4) if len(sh.table.rows)>=1]
    task_detail=detail_pattern_slides(prs,'task'); suhail_detail=detail_pattern_slides(prs,'suhail')
    if not agenda: raise ValueError('agenda table pattern missing from master')
    if not task_detail: raise ValueError('task detail table pattern missing from master')
    if not suhail_detail: raise ValueError('Suhail detail table pattern missing from master')
    if not any(getattr(sh,'has_chart',False) for sl in prs.slides for sh in sl.shapes):
        raise ValueError('summary chart patterns missing from master')
    for titles in (TASK_SUMMARY_TITLES,SUHAIL_SUMMARY_TITLES):
        missing=[title for title in titles if not any(table_title(sh)==title for sl in prs.slides for sh in tables(sl))]
        if missing:raise ValueError(f'summary table patterns missing: {missing}')


def header_map(ws,row):
    return {str(c.value).strip():c.column for c in ws[row] if c.value is not None}


def agenda_rows(path):
    ws=openpyxl.load_workbook(path,data_only=True)['Sheet1']; hm=header_map(ws,2); out=[]; found=False
    for r in range(3,ws.max_row+1):
        topic=val(ws.cell(r,hm['الموضوع']).value).strip(); order=ws.cell(r,hm['الترتيب']).value
        if order not in (None,''):
            out.append((f'{int(order):02d}',topic,val(ws.cell(r,hm['المتحدث']).value),val(ws.cell(r,hm['الوقت المطلوب من المالك']).value)))
        if re.sub(r'\s+',' ',topic)=='ملخص الاجتماع': found=True; break
    if not found: raise ValueError('ملخص الاجتماع is missing from topics workbook')
    return out


def find_header(ws,labels,start=1):
    labels=[x.strip() for x in labels]
    for r in range(start,ws.max_row+1):
        got=[val(ws.cell(r,c).value).strip() for c in range(1,len(labels)+1)]
        if got==labels:return r
    raise ValueError(f'header not found: {labels}')


def read_block(ws,header_row,cols,stop_titles=()):
    out=[]
    for r in range(header_row+1,ws.max_row+1):
        first=val(ws.cell(r,1).value).strip()
        if first in stop_titles or (first and all(ws.cell(r,c).value is None for c in range(2,cols+1))):break
        row=[ws.cell(r,c).value for c in range(1,cols+1)]
        if any(x not in (None,'') for x in row):out.append(row)
    return out


def detail_sections(ws,cols):
    out=[]; r=1
    expected_task=['#','المهمة','القطاع','تاريخ الانجاز المخطط','الحالة','ملاحظات','الجهات ذات العلاقة']
    expected_proj=['#','اسم المشروع','تاريخ البداية','تاريخ النهاية','الحالة','ما تم حتى تاريخه']
    expected=expected_task if cols==7 else expected_proj
    while r<=ws.max_row:
        got=[val(ws.cell(r,c).value).strip() for c in range(1,cols+1)]
        if got==expected:
            title=val(ws.cell(r-1,1).value).strip(); title=re.sub(r'^مهام مصدر\s*\((.*)\)$',r'\1',title)
            out.append((title,read_block(ws,r,cols)))
        r+=1
    return out


def summary_records(ws,title,headers):
    """Read one titled Excel section; example/empty markers are not records."""
    for r in range(1,ws.max_row):
        heading=val(ws.cell(r,1).value).strip()
        if heading==title or heading.endswith(': '+title):
            got=[val(ws.cell(r+1,c).value).strip() for c in range(1,len(headers)+1)]
            if got!=headers:continue
            out=[]
            for rr in range(r+2,ws.max_row+1):
                first=val(ws.cell(rr,1).value).strip()
                if not first or first=='لا يوجد':break
                if not re.fullmatch(r'\d+(?:\.0)?',first):break
                out.append([ws.cell(rr,c).value for c in range(1,len(headers)+1)])
            return out
    raise ValueError(f'missing summary section: {title}')


def task_data(path):
    wb=openpyxl.load_workbook(path,data_only=True); s=wb['ملخص المهام']
    k=[s.cell(r,2).value for r in range(2,7)]
    hr=find_header(s,['القطاع','مكتملة','على المخطط','متأخر','معلق']); chart=read_block(s,hr,5)
    specs=[('#','المهمة','القطاع','ملاحظات')]*3+[('#','المهمة','القطاع','طلب الدعم')]
    summaries={title:summary_records(s,title,list(headers)) for title,headers in zip(TASK_SUMMARY_TITLES,specs)}
    return k,chart,summaries,detail_sections(wb['تفاصيل المهام'],7)


def suhail_data(path):
    wb=openpyxl.load_workbook(path,data_only=True); s=wb['ملخص مشاريع سهيل']
    k=[s.cell(3,c).value for c in range(1,6)]
    hr=find_header(s,['القطاع','مكتملة','على المخطط','متأخر','لم تبدأ']); chart=read_block(s,hr,5)
    specs=[['#','المشروع','القطاع','التحديث'],['#','المشروع','القطاع','ملاحظات'],['#','المشروع','القطاع','التحدي']]
    summaries={title:summary_records(s,title,headers) for title,headers in zip(SUHAIL_SUMMARY_TITLES,specs)}
    return k,chart,summaries,detail_sections(wb['تفاصيل مشاريع سهيل'],6)


def save_chart_payload(rows,path):
    wb=openpyxl.load_workbook(path); ws=wb.active; maxr=max(ws.max_row,len(rows)+1)
    for r in range(2,maxr+1):
        for c in range(1,6):ws.cell(r,c).value=None
    for r,row in enumerate(rows,2):
        for c,x in enumerate(row,1):ws.cell(r,c).value=x
    wb.save(path)


def update_chart_xml(xml_bytes,rows):
    root=etree.fromstring(xml_bytes); root_range=len(rows)+1
    for ser_i,ser in enumerate(root.xpath('.//c:ser',namespaces=NS_C)):
        if ser_i>=4:continue
        for f in ser.xpath('.//c:f',namespaces=NS_C):f.text=re.sub(r'\$\d+$',f'${root_range}',f.text)
        cats=[val(r[0]) for r in rows]; nums=[r[ser_i+1] for r in rows]
        for cache,values in [(ser.find('.//c:cat//c:strCache',NS_C),cats),(ser.find('.//c:val//c:numCache',NS_C),nums)]:
            if cache is None:continue
            for p in list(cache.findall('c:pt',NS_C)):cache.remove(p)
            pc=cache.find('c:ptCount',NS_C)
            if pc is not None:pc.set('val',str(len(values)))
            for idx,x in enumerate(values):
                if x in (None,''):continue
                pt=etree.SubElement(cache,'{%s}pt'%NS_C['c'],idx=str(idx)); v=etree.SubElement(pt,'{%s}v'%NS_C['c']); v.text=val(x)
    disp=root.find('.//c:dispBlanksAs',NS_C)
    if disp is None:disp=etree.SubElement(root,'{%s}dispBlanksAs'%NS_C['c'])
    disp.set('val','gap')
    return etree.tostring(root,xml_declaration=True,encoding='UTF-8',standalone=True)


def patch_charts(pptx,task_rows,suhail_rows):
    tmp=Path(tempfile.mkdtemp()); src=tmp/'src'; src.mkdir()
    with zipfile.ZipFile(pptx) as z:z.extractall(src)
    for chart_name,book_name,rows in [('chart1.xml','Microsoft_Excel_Worksheet.xlsx',task_rows),('chart2.xml','Microsoft_Excel_Worksheet1.xlsx',suhail_rows)]:
        save_chart_payload(rows,src/'ppt'/'embeddings'/book_name)
        cp=src/'ppt'/'charts'/chart_name; cp.write_bytes(update_chart_xml(cp.read_bytes(),rows))
    out=tmp/'out.pptx'
    with zipfile.ZipFile(out,'w',zipfile.ZIP_DEFLATED) as z:
        for p in src.rglob('*'):
            if p.is_file():z.write(p,p.relative_to(src))
    shutil.copy2(out,pptx); shutil.rmtree(tmp)


def build_dynamic_detail_section(prs,pattern_slides,sections,kind,title_prefix=None):
    pattern=max(pattern_slides,key=lambda sl:max(len(sh.table.rows) for sh in detail_tables(sl,kind)))
    pattern_shape=max(detail_tables(pattern,kind),key=lambda sh:len(sh.table.rows))
    cols=6
    pages=[]
    for section,rows in sections:
        if not rows:
            continue
        if kind=='task':rows=[row[:6] for row in rows]  # User-approved omission of related agencies.
        for chunk,weights in paginate_rows(rows,pattern_shape,2,11):
            pages.append((section,chunk,weights))
    if not pages:
        return [],list(pattern_slides)
    # Allocate all generated slide parts before deleting any seed slides. This
    # keeps OOXML part names unique and avoids duplicate slide*.xml entries.
    targets=[pattern]; last=pattern
    for _ in range(1,len(pages)):
        last=clone_slide(prs,pattern,after=last); targets.append(last)
    for sl,(section,chunk,weights) in zip(targets,pages):
        sh=detail_tables(sl,kind)[0]
        title=f'{title_prefix} ({section})' if title_prefix else section
        fill_titled_table(sh,chunk,title=title,weights=weights)
    unused=[sl for sl in pattern_slides if sl.part is not pattern.part]
    return targets,unused

def main():
    ap=argparse.ArgumentParser(); ap.add_argument('--topics',required=True); ap.add_argument('--tasks',required=True); ap.add_argument('--suhail',required=True); ap.add_argument('--output',required=True); ap.add_argument('--template')
    a=ap.parse_args(); template=Path(a.template) if a.template else Path(__file__).resolve().parents[1]/'assets'/'monday-meeting-master.pptx'
    if not template.exists():raise FileNotFoundError(template)
    if not a.template and hashlib.sha256(template.read_bytes()).hexdigest()!=MASTER_SHA256:
        raise ValueError('Bundled Monday meeting template checksum mismatch')

    agenda=agenda_rows(a.topics); tk,tchart,task_summaries,tdetails=task_data(a.tasks); sk,schart,suhail_summaries,sdetails=suhail_data(a.suhail)
    out=Path(a.output); shutil.copy2(template,out); prs=Presentation(out); validate_master_patterns(prs)

    # Capture all master patterns before any slide is cloned or removed.
    agenda_slide=next(sl for sl in prs.slides if any(len(sh.table.columns)==4 and sh.table.cell(0,0).text.strip()=='م' for sh in tables(sl)))
    agenda_shape=next(sh for sh in tables(agenda_slide,4) if sh.table.cell(0,0).text.strip()=='م')
    task_summary=find_slide_with_text(prs,'ملخص المهام')
    suhail_summary=find_slide_with_text(prs,'ملخص مشاريع سهيل')
    task_patterns=detail_pattern_slides(prs,'task'); suhail_patterns=detail_pattern_slides(prs,'suhail')

    # Cover + agenda. Agenda automatically adds continuation pages when needed.
    set_text(shape_by_id(prs.slides[0],4),f'{dt.datetime.now().astimezone().date().day} {MONTHS[dt.datetime.now().astimezone().date().month-1]} {dt.datetime.now().astimezone().date().year}')
    agenda_pages=paginate_rows(agenda,agenda_shape,1,14)
    agenda_targets=[agenda_slide]; last=agenda_slide
    for _ in range(1,len(agenda_pages)):
        last=clone_slide(prs,agenda_slide,after=last); agenda_targets.append(last)
    for sl,(chunk,weights) in zip(agenda_targets,agenda_pages):fill_agenda_table(first_table(sl,4),chunk,weights)

    # Task summary can expand down the left side; when that area fills, create a
    # clean continuation page containing only the approved heading/table pattern.
    for sid,x in zip([21,23,25,19,18],tk):set_text(shape_by_id(task_summary,sid),x)
    fill_summary_tables(prs,task_summary,task_summaries,TASK_SUMMARY_TITLES,'ملخص المهام')

    # Task details are data-driven. The master contributes one approved detail
    # pattern; every source and overflow page is generated from it.
    task_generated,task_unused=build_dynamic_detail_section(prs,task_patterns,tdetails,'task','تفاصيل المهام')

    # Suhail summary and updates.
    for sid,x in zip([5,8,11,17,13],sk):set_text(shape_by_id(suhail_summary,sid),x)
    set_text(shape_by_id(suhail_summary,14),'لم تبدأ')
    fill_summary_tables(prs,suhail_summary,suhail_summaries,SUHAIL_SUMMARY_TITLES,'ملخص مشاريع سهيل')

    # Suhail details are also data-driven: new sectors and any amount of projects
    # simply create more pages from the approved six-column detail pattern.
    suhail_generated,suhail_unused=build_dynamic_detail_section(prs,suhail_patterns,sdetails,'suhail',None)

    # Only after every dynamic page has been allocated do we remove the unused
    # seed-pattern slides from the master.
    for sl in task_unused+suhail_unused:
        remove_slide(prs,sl)

    prs.save(out); patch_charts(out,tchart,schart); print(out.resolve())


if __name__=='__main__':main()
