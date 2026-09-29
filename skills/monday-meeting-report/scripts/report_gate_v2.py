# /// script
# requires-python = ">=3.11"
# dependencies = ["python-pptx>=1.0.2", "openpyxl>=3.1.5"]
# ///
"""Mandatory post-build delivery gate for monday-meeting-report.

Run after build_report.py and before delivery. This script normalizes approved
table typography and blocks delivery when the deck is structurally invalid,
uses stale/current-input-mismatched data, loses table titles/headers, or contains
external PowerPoint relationships.
"""
from __future__ import annotations
import argparse, collections, hashlib, io, os, posixpath, re, shutil, tempfile, zipfile
from pathlib import Path, PurePosixPath
from xml.etree import ElementTree as ET
from openpyxl import load_workbook
from openpyxl.utils.cell import range_boundaries
from pptx import Presentation
from pptx.util import Pt

AGENDA_HEADERS=["م","جدول الأعمال","المسؤول","المدة الزمنية (بالدقيقة)"]
TASK_HEADERS=["#","المهمة","القطاع","تاريخ الإنجاز المخطط","الحالة","ملاحظات"]
SUHAIL_HEADERS=["#","اسم المشروع","تاريخ البداية","تاريخ النهاية","الحالة","ما تم حتى تاريخه"]
SUMMARY_HEADERS={
    "المهام المكتملة":["ملاحظات","القطاع","المهمة","#"],
    "المهام المتأخرة":["ملاحظات","القطاع","المهمة","#"],
    "المهام المعلقة":["ملاحظات","القطاع","المهمة","#"],
    "طلبات الدعم":["طلب الدعم","القطاع","المهمة","#"],
    "أبرز التحديثات":["التحديث","القطاع","المشروع","#"],
    "المشاريع المتأخرة":["ملاحظات","القطاع","المشروع","#"],
    "التحديات":["التحدي","القطاع","المشروع","#"],
}
FONT="Abar Mid"; ALLOWED={"Abar Mid","Abar Mid SemiBold"}
AGENDA_PT=14.0; TABLE_PT=11.0

def norm(v):
    if v is None:return ""
    if isinstance(v,float) and v.is_integer():v=int(v)
    return re.sub(r"\s+"," ",str(v).replace("\u00a0"," ")).strip()

def sha(path):
    h=hashlib.sha256()
    with open(path,"rb") as f:
        for b in iter(lambda:f.read(1<<20),b""):h.update(b)
    return h.hexdigest()

def header_row(ws,required,max_scan=30):
    for r in range(1,min(ws.max_row,max_scan)+1):
        m={norm(ws.cell(r,c).value):c for c in range(1,ws.max_column+1)}
        if required.issubset(m):return r,m
    return None,None

def read_agenda(path):
    wb=load_workbook(io.BytesIO(Path(path).read_bytes()),read_only=True,data_only=False); ws=wb["Sheet1"]
    req={"الترتيب","الموضوع","المتحدث","الوقت المطلوب من المالك"}; hr,c=header_row(ws,req)
    if not hr:raise RuntimeError("topics workbook headers not found")
    out=[]; stop=False
    for r in range(hr+1,ws.max_row+1):
        order=ws.cell(r,c["الترتيب"]).value; topic=norm(ws.cell(r,c["الموضوع"]).value)
        if order in (None,"") or not topic:continue
        try:o=f"{int(float(order)):02d}"
        except:o=norm(order)
        dur=ws.cell(r,c["الوقت المطلوب من المالك"]).value
        if isinstance(dur,float) and dur.is_integer():dur=int(dur)
        out.append((o,topic,norm(ws.cell(r,c["المتحدث"]).value),norm(dur)))
        if topic=="ملخص الاجتماع":stop=True;break
    if not stop:raise RuntimeError("topics workbook missing final item ملخص الاجتماع")
    return out

def read_tasks(path):
    wb=load_workbook(io.BytesIO(Path(path).read_bytes()),read_only=True,data_only=False); ws=wb["تفاصيل المهام"]
    out=[]; r=1
    while r<=ws.max_row:
        m=re.fullmatch(r"مهام مصدر \((.+)\)",norm(ws.cell(r,1).value))
        if not m:r+=1;continue
        sec=norm(m.group(1)); r+=2
        while r<=ws.max_row:
            a=norm(ws.cell(r,1).value)
            if re.fullmatch(r"مهام مصدر \((.+)\)",a):break
            if re.fullmatch(r"\d+(?:\.0)?",a):
                vals=[norm(ws.cell(r,c).value) for c in range(1,7)]
                out.append((sec,*vals))
            r+=1
    return out

def read_suhail(path):
    wb=load_workbook(io.BytesIO(Path(path).read_bytes()),read_only=True,data_only=False); ws=wb["تفاصيل مشاريع سهيل"]
    out=[]; r=1
    while r<=ws.max_row:
        vals=[norm(ws.cell(r,c).value) for c in range(1,7)]
        if vals[0] and not any(vals[1:]) and r<ws.max_row:
            nxt=[norm(ws.cell(r+1,c).value) for c in range(1,7)]
            if nxt[:2]==["#","اسم المشروع"]:
                sec=vals[0]; r+=2
                while r<=ws.max_row:
                    row=[norm(ws.cell(r,c).value) for c in range(1,7)]
                    if not re.fullmatch(r"\d+(?:\.0)?",row[0]):break
                    out.append((sec,*row)); r+=1
                continue
        r+=1
    return out

def normalize_typography(prs):
    for sl in prs.slides:
        for sh in sl.shapes:
            if getattr(sh,"has_table",False):
                tt=sh.table; rr=rows(tt); is_ag=bool(rr and rr[0]==AGENDA_HEADERS); pt=AGENDA_PT if is_ag else TABLE_PT
                for rw in tt.rows:
                    for cell in rw.cells:
                        for p in cell.text_frame.paragraphs:
                            for run in p.runs:
                                if run.text:
                                    if run.font.name not in ALLOWED:run.font.name=FONT
                                    run.font.size=Pt(pt)
            elif getattr(sh,"has_text_frame",False):
                for p in sh.text_frame.paragraphs:
                    for run in p.runs:
                        if run.text and run.font.name in {"Arial","Calibri","+mn-lt","+mj-lt"}:run.font.name=FONT


def canon(v):
    s=norm(v)
    s=re.sub(r"تاريخ التحديث\s*[:：]\s*","تاريخ التحديث ",s)
    return norm(s)

def canon_record(rec):
    return tuple(canon(x) for x in rec)

def rows(t):return [[norm(c.text) for c in r.cells] for r in t.rows]
def isnum(s):return bool(re.fullmatch(r"\d+(?:\.0)?",norm(s)))


def read_summary(path,sheet,titles):
    ws=load_workbook(io.BytesIO(Path(path).read_bytes()),read_only=True,data_only=False)[sheet]
    result={title:[] for title in titles}
    found=set()
    for r in range(1,ws.max_row):
        heading=norm(ws.cell(r,1).value)
        title=next((t for t in titles if heading==t or heading.endswith(': '+t)),None)
        if title is None:continue
        if ws.cell(r,2).value is not None:continue  # KPI label, not a section band.
        expected=list(reversed(SUMMARY_HEADERS[title]))
        got=[norm(ws.cell(r+1,c).value) for c in range(1,5)]
        if got!=expected:raise ValueError(f'{sheet}: invalid headers for {title}')
        if title in found:raise ValueError(f'{sheet}: duplicate summary section {title}')
        found.add(title)
        for rr in range(r+2,ws.max_row+1):
            first=norm(ws.cell(rr,1).value)
            if not isnum(first):break
            result[title].append(tuple(norm(ws.cell(rr,c).value) for c in range(4,0,-1)))
    if set(titles)!=found:raise ValueError(f'{sheet}: missing summary sections {sorted(set(titles)-found)}')
    return result

def package(path,errors):
    try:
        with zipfile.ZipFile(path) as z:
            bad=z.testzip(); names=set(z.namelist())
            if bad:errors.append(f"ZIP CRC failure: {bad}")
            if len(names)!=len(z.namelist()):errors.append('Duplicate package members')
            relns="{http://schemas.openxmlformats.org/package/2006/relationships}"
            for n in names:
                if n.endswith((".xml",".rels")):
                    try:ET.fromstring(z.read(n))
                    except Exception as e:errors.append(f"malformed XML {n}: {e}")
                if n.endswith('.xml'):
                    relpath=posixpath.join(posixpath.dirname(n),'_rels',posixpath.basename(n)+'.rels')
                    ids={r.get('Id') for r in ET.fromstring(z.read(relpath))} if relpath in names else set()
                    for el in ET.fromstring(z.read(n)).iter():
                        if el.tag=='{http://schemas.openxmlformats.org/drawingml/2006/main}pPr':
                            tags=[child.tag.rsplit('}',1)[-1] for child in el]
                            if 'buNone' in tags and any(tag in tags and tags.index(tag)>tags.index('buNone') for tag in ('lnSpc','spcBef','spcAft','buFont','buFontTx')):
                                errors.append(f'invalid bullet property order: {n}')
                        for key,value in el.attrib.items():
                            if key.startswith('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}') and value and value not in ids:
                                errors.append(f'missing relationship ID: {n} -> {value}')
                if not n.endswith(".rels"):continue
                root=ET.fromstring(z.read(n)); p=PurePosixPath(n)
                source=PurePosixPath("") if p.name==".rels" and str(p.parent)=="_rels" else p.parent.parent/p.name[:-5]
                for rel in root.findall(f"{relns}Relationship"):
                    if rel.attrib.get("TargetMode")=="External":errors.append(f"external relationship: {n} -> {rel.attrib.get('Target')}")
                    else:
                        target=rel.attrib.get('Target','').split('#',1)[0]
                        resolved=posixpath.normpath(posixpath.join(str(source.parent),target)).lstrip('/')
                        if resolved not in names:errors.append(f'missing relationship target: {n} -> {target}')
    except Exception as e:errors.append(f"invalid PPTX package: {e}")

def collect(prs,errors):
    ag=[]; task=[]; suh=[]; suh_seq=[]; summaries={title:[] for title in SUMMARY_HEADERS}
    for slide in prs.slides:
        for shape in slide.shapes:
            if not shape.has_table or len(shape.table.columns)!=6:continue
            for row in list(shape.table.rows)[2:]:
                cell=row.cells[4]
                color={'مكتمل':'DCE6F2','مكتملة':'DCE6F2','على المخطط':'EBF1DE'}.get(norm(cell.text))
                if color:
                    fill=cell._tc.get_or_add_tcPr().find('{http://schemas.openxmlformats.org/drawingml/2006/main}solidFill')
                    actual=fill.find('{http://schemas.openxmlformats.org/drawingml/2006/main}srgbClr') if fill is not None else None
                    if actual is None or actual.get('val')!=color:errors.append('detail status fill differs from approved template')
    for si,sl in enumerate(prs.slides,1):
        for sh in sl.shapes:
            if not getattr(sh,"has_table",False):continue
            t=sh.table; rr=rows(t); cols=len(t.columns)
            if cols==4 and rr and rr[0]==AGENDA_HEADERS:
                ag.extend(tuple(x) for x in rr[1:] if any(x))
            elif cols==6 and len(rr)>=2 and rr[1][:2]==["#","المهمة"]:
                if not any(x.startswith("تفاصيل المهام (") for x in rr[0]):errors.append(f"slide {si} task table missing title row")
                if rr[1]!=TASK_HEADERS:errors.append(f"slide {si} task headers missing/corrupt")
                m=re.fullmatch(r"تفاصيل المهام \((.+)\)",rr[0][0] if rr else ""); sec=norm(m.group(1)) if m else ""
                for x in rr[2:]:
                    if x and isnum(x[0]):task.append((sec,*x[:6]))
            elif cols==6 and len(rr)>=2 and rr[1][:2]==["#","اسم المشروع"]:
                if not rr[0][0]:errors.append(f"slide {si} Suhail table missing sector/title row")
                if rr[1]!=SUHAIL_HEADERS:errors.append(f"slide {si} Suhail headers missing/corrupt")
                sec=re.sub(r"\s*\(تابع\)\s*$","",rr[0][0]).strip()
                if not suh_seq or suh_seq[-1]!=sec:suh_seq.append(sec)
                for x in rr[2:]:
                    if x and isnum(x[0]):suh.append((sec,*x[:6]))
            elif cols==4 and rr:
                title=next((x for x in rr[0] if x in SUMMARY_HEADERS),None)
                if title:
                    if len(rr)<2 or rr[1]!=SUMMARY_HEADERS[title]:errors.append(f"slide {si} {title} headers missing/corrupt")
                    for x in rr[2:]:
                        if x and isnum(x[-1]):summaries[title].append(tuple(x))
                    if not any(x and isnum(x[-1]) for x in rr[2:]):
                        errors.append(f'slide {si} retains empty summary table {title}')
            is_ag=cols==4 and rr and rr[0]==AGENDA_HEADERS; want=AGENDA_PT if is_ag else TABLE_PT
            for rw in t.rows:
                for cell in rw.cells:
                    for p in cell.text_frame.paragraphs:
                        for run in p.runs:
                            if not run.text.strip():continue
                            if run.font.name not in ALLOWED:errors.append(f"slide {si} unapproved font {run.font.name!r}: {run.text[:30]!r}")
                            sz=run.font.size.pt if run.font.size else None
                            if sz is None or abs(sz-want)>.01:errors.append(f"slide {si} wrong table font size {sz!r}; expected {want:g}")
    if ag and ag[-1][1]!="ملخص الاجتماع":errors.append("agenda does not end at ملخص الاجتماع")
    if len(ag)!=len(set(ag)):errors.append("agenda contains duplicate rows")
    if len(suh_seq)!=len(set(suh_seq)):
        errors.append(f"Suhail sector continuation is non-contiguous: {suh_seq!r}")
    return ag,task,suh,summaries

def cmp(label,exp,act,errors):
    ce=collections.Counter(canon_record(x) for x in exp); ca=collections.Counter(canon_record(x) for x in act)
    if ce==ca:
        if [canon_record(x) for x in exp]!=[canon_record(x) for x in act]:
            errors.append(f'{label}: source record order changed')
        return
    miss=list((ce-ca).elements()); extra=list((ca-ce).elements())
    if miss:errors.append(f"{label}: missing {len(miss)} current-input record(s); first={miss[0]!r}")
    if extra:errors.append(f"{label}: stale/extra {len(extra)} record(s); first={extra[0]!r}")


CHART_NS={'c':'http://schemas.openxmlformats.org/drawingml/2006/chart'}
REL_NS='{http://schemas.openxmlformats.org/package/2006/relationships}'
RID='{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id'


def source_metrics(path, sheet, statuses, kpi_cells):
    """Read the final Excel summary independently of the report builder."""
    with open(path,'rb') as stream:
        wb=load_workbook(stream,data_only=True)
        ws=wb[sheet]
        header,columns=header_row(ws,{'القطاع',*statuses},ws.max_row)
        if header is None:raise ValueError(f'{sheet}: sector status header missing')
        records=[]
        for row in range(header+1,ws.max_row+1):
            sector=norm(ws.cell(row,columns['القطاع']).value)
            if not sector:break
            vals=[ws.cell(row,columns[s]).value for s in statuses]
            if any(v is not None and (not isinstance(v,(int,float)) or v<0 or int(v)!=v) for v in vals):
                raise ValueError(f'{sheet}: invalid sector counts at row {row}')
            records.append([sector,*[v if v else None for v in vals]])
        kpis=[ws[cell].value for cell in kpi_cells]
        totals=[sum(r[i+1] or 0 for r in records) for i in range(len(statuses))]
        if kpis!=[sum(totals),*totals]:raise ValueError(f'{sheet}: KPI and sector counts disagree')
        wb.close()
        return records,kpis


def formula_values(wb,formula):
    sheet,area=formula.rsplit('!',1)
    sheet=sheet.strip("'").replace("''", "'")
    c1,r1,c2,r2=range_boundaries(area.replace('$',''))
    if c1!=c2 and r1!=r2:raise ValueError(f'non-linear chart range: {formula}')
    return [wb[sheet].cell(r,c).value for r in range(r1,r2+1) for c in range(c1,c2+1)]


def cached_values(ref):
    cache=next((x for x in ref if x.tag.endswith('Cache')),None)
    if cache is None:raise ValueError('chart cache missing')
    count=cache.find('c:ptCount',CHART_NS)
    if count is None:raise ValueError('chart cache count missing')
    values=[None]*int(count.get('val'))
    seen=set()
    for point in cache.findall('c:pt',CHART_NS):
        idx=int(point.get('idx'))
        if idx in seen or not 0<=idx<len(values):raise ValueError('invalid chart cache index')
        seen.add(idx); values[idx]=point.findtext('c:v',namespaces=CHART_NS)
    return values


def chart_metrics(path, chart_name, records, statuses, errors):
    """Require source Excel, chart caches and referenced embedded cells to agree."""
    try:
        records=records or [[None]*(len(statuses)+1)]
        with zipfile.ZipFile(path) as z:
            root=ET.fromstring(z.read(chart_name))
            rel_path=posixpath.join(posixpath.dirname(chart_name),'_rels',posixpath.basename(chart_name)+'.rels')
            rels=ET.fromstring(z.read(rel_path))
            external=root.find('c:externalData',CHART_NS)
            if external is None:raise ValueError('embedded chart workbook reference missing')
            rel=next((x for x in rels if x.get('Id')==external.get(RID)),None)
            if rel is None or rel.get('TargetMode')=='External':raise ValueError('chart workbook is not embedded')
            book=posixpath.normpath(posixpath.join(posixpath.dirname(chart_name),rel.get('Target'))).lstrip('/')
            wb=load_workbook(io.BytesIO(z.read(book)),data_only=False)
            series=root.findall('.//c:ser',CHART_NS)
            if len(series)!=len(statuses):raise ValueError('wrong number of chart series')
            for i,ser in enumerate(series):
                for tag,expected in [('tx',[statuses[i]]),('cat',[r[0] for r in records]),('val',[r[i+1] for r in records])]:
                    holder=ser.find('c:'+tag,CHART_NS)
                    if holder is None:raise ValueError(f'series {i}: missing {tag}')
                    ref=next((x for x in holder if x.tag.endswith('Ref')),None)
                    if ref is None:raise ValueError(f'series {i}: missing {tag} reference')
                    formula=ref.findtext('c:f',namespaces=CHART_NS)
                    actual=cached_values(ref)
                    cells=formula_values(wb,formula)
                    normalize=lambda seq:[norm(v) if v is not None else '' for v in seq]
                    if normalize(actual)!=normalize(expected):errors.append(f'{chart_name} series {i} {tag}: cache differs from source')
                    if normalize(cells)!=normalize(expected):errors.append(f'{chart_name} series {i} {tag}: embedded workbook differs from source')
            gap=root.find('.//c:dispBlanksAs',CHART_NS)
            if gap is None or gap.get('val')!='gap':errors.append(f'{chart_name}: blanks must display as gaps')
            wb.close()
    except Exception as exc:errors.append(f'{chart_name}: {exc}')


def verify_metrics(path, prs, tasks_path, suhail_path, errors):
    specs=[('ملخص المهام',tasks_path,['مكتملة','على المخطط','متأخر','معلق'],
            ['B2','B3','B4','B5','B6'],[21,23,25,19,18],'ppt/charts/chart1.xml'),
           ('ملخص مشاريع سهيل',suhail_path,['مكتملة','على المخطط','متأخر','لم تبدأ'],
            ['A3','B3','C3','D3','E3'],[5,8,11,17,13],'ppt/charts/chart2.xml')]
    for title,source,statuses,cells,ids,chart in specs:
        try:
            records,kpis=source_metrics(source,title,statuses,cells)
            details=read_tasks(source) if title=='ملخص المهام' else read_suhail(source)
            source_counts=collections.Counter()
            for row in details:
                sector=row[3] if title=='ملخص المهام' else row[0]
                status=row[5]
                status={'مكتمل':'مكتملة','متأخرة':'متأخر','معلقة':'معلق'}.get(status,status)
                if status not in statuses:raise ValueError(f'unknown detail status {status}')
                source_counts[(sector,status)]+=1
            chart_counts={(norm(row[0]),status):int(row[i+1] or 0) for row in records for i,status in enumerate(statuses)}
            if any(source_counts[key]!=chart_counts.get(key,0) for key in set(source_counts)|set(chart_counts)):
                errors.append(f'{title}: sector summary differs from detail records')
            chart_metrics(path,chart,[row for row in records if any(float(value or 0)>0 for value in row[1:])],statuses,errors)
            candidates=[s for s in prs.slides if any(sh.has_chart for sh in s.shapes)
                        and any(sh.has_text_frame and title in sh.text for sh in s.shapes)]
            if len(candidates)!=1:raise ValueError('expected exactly one summary chart slide')
            def walk(shapes):
                for shape in shapes:
                    yield shape
                    if hasattr(shape,'shapes'):yield from walk(shape.shapes)
            shapes={sh.shape_id:sh for sh in walk(candidates[0].shapes)}
            for sid,value in zip(ids,kpis):
                if sid not in shapes or norm(shapes[sid].text)!=norm(value):
                    errors.append(f'{title}: KPI shape {sid} differs from source')
        except Exception as exc:errors.append(f'{title}: {exc}')

def main():
    ap=argparse.ArgumentParser(); ap.add_argument("--report",required=True); ap.add_argument("--topics",required=True); ap.add_argument("--tasks",required=True); ap.add_argument("--suhail",required=True); a=ap.parse_args()
    for p in [a.report,a.topics,a.tasks,a.suhail]:
        if not os.path.isfile(p):raise SystemExit(f"missing input: {p}")
    exp_ag=read_agenda(a.topics); exp_task=read_tasks(a.tasks); exp_suh=read_suhail(a.suhail)
    task_titles=list(SUMMARY_HEADERS)[:4];suhail_titles=list(SUMMARY_HEADERS)[4:]
    exp_summaries={**read_summary(a.tasks,"ملخص المهام",task_titles),**read_summary(a.suhail,"ملخص مشاريع سهيل",suhail_titles)}
    report=Path(a.report)
    with tempfile.TemporaryDirectory(prefix="mm-gate-") as td:
        tmp=Path(td)/report.name; shutil.copy2(report,tmp)
        prs=Presentation(str(tmp)); normalize_typography(prs); prs.save(str(tmp))
        errors=[]; package(tmp,errors); prs2=None
        try:prs2=Presentation(str(tmp)); ag,task,suh,summaries=collect(prs2,errors)
        except Exception as e:errors.append(f"PowerPoint reopen failed: {e}");ag=task=suh=[];summaries={}
        if ag!=exp_ag:errors.append(f"agenda mismatch: expected={exp_ag!r}; actual={ag!r}")
        cmp("task details",exp_task,task,errors); cmp("Suhail details",exp_suh,suh,errors)
        for title,expected in exp_summaries.items():cmp(title,expected,summaries.get(title,[]),errors)
        if prs2 is not None:verify_metrics(tmp,prs2,a.tasks,a.suhail,errors)
        print("REPORT_SHA256",sha(tmp)); print("TOPICS_SHA256",sha(a.topics)); print("TASKS_SHA256",sha(a.tasks)); print("SUHAIL_SHA256",sha(a.suhail))
        if errors:
            print(f"QA FAILED: {len(errors)} blocking issue(s)")
            for e in errors:print("ERROR:",e)
            return 1
        shutil.copy2(tmp,report)
    print("PROGRAMMATIC QA PASSED - render and inspect these final bytes before delivery")
    return 0
if __name__=="__main__":raise SystemExit(main())
