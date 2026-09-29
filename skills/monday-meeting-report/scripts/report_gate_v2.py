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
import argparse, collections, hashlib, os, re, shutil, tempfile, zipfile
from pathlib import Path, PurePosixPath
from xml.etree import ElementTree as ET
from openpyxl import load_workbook
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
    wb=load_workbook(path,read_only=True,data_only=False); ws=wb["Sheet1"]
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
    wb=load_workbook(path,read_only=True,data_only=False); ws=wb["تفاصيل المهام"]
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
    wb=load_workbook(path,read_only=True,data_only=False); ws=wb["تفاصيل مشاريع سهيل"]
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
    ws=load_workbook(path,read_only=True,data_only=False)[sheet]
    result={title:[] for title in titles}
    for r in range(1,ws.max_row):
        heading=norm(ws.cell(r,1).value)
        title=next((t for t in titles if heading==t or heading.endswith(': '+t)),None)
        if title is None:continue
        expected=list(reversed(SUMMARY_HEADERS[title]))
        got=[norm(ws.cell(r+1,c).value) for c in range(1,5)]
        if got!=expected:continue
        for rr in range(r+2,ws.max_row+1):
            first=norm(ws.cell(rr,1).value)
            if not isnum(first):break
            result[title].append(tuple(norm(ws.cell(rr,c).value) for c in range(4,0,-1)))
    return result

def package(path,errors):
    try:
        with zipfile.ZipFile(path) as z:
            bad=z.testzip(); names=set(z.namelist())
            if bad:errors.append(f"ZIP CRC failure: {bad}")
            relns="{http://schemas.openxmlformats.org/package/2006/relationships}"
            for n in names:
                if n.endswith((".xml",".rels")):
                    try:ET.fromstring(z.read(n))
                    except Exception as e:errors.append(f"malformed XML {n}: {e}")
                if not n.endswith(".rels"):continue
                root=ET.fromstring(z.read(n)); p=PurePosixPath(n)
                source=PurePosixPath("") if p.name==".rels" and str(p.parent)=="_rels" else p.parent.parent/p.name[:-5]
                for rel in root.findall(f"{relns}Relationship"):
                    if rel.attrib.get("TargetMode")=="External":errors.append(f"external relationship: {n} -> {rel.attrib.get('Target')}")
    except Exception as e:errors.append(f"invalid PPTX package: {e}")

def collect(prs,errors):
    ag=[]; task=[]; suh=[]; suh_seq=[]; summaries={title:[] for title in SUMMARY_HEADERS}
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
    if ce==ca:return
    miss=list((ce-ca).elements()); extra=list((ca-ce).elements())
    if miss:errors.append(f"{label}: missing {len(miss)} current-input record(s); first={miss[0]!r}")
    if extra:errors.append(f"{label}: stale/extra {len(extra)} record(s); first={extra[0]!r}")

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
        errors=[]; package(tmp,errors)
        try:prs2=Presentation(str(tmp)); ag,task,suh,summaries=collect(prs2,errors)
        except Exception as e:errors.append(f"PowerPoint reopen failed: {e}");ag=task=suh=[];summaries={}
        if ag!=exp_ag:errors.append(f"agenda mismatch: expected={exp_ag!r}; actual={ag!r}")
        cmp("task details",exp_task,task,errors); cmp("Suhail details",exp_suh,suh,errors)
        for title,expected in exp_summaries.items():cmp(title,expected,summaries.get(title,[]),errors)
        print("REPORT_SHA256",sha(tmp)); print("TOPICS_SHA256",sha(a.topics)); print("TASKS_SHA256",sha(a.tasks)); print("SUHAIL_SHA256",sha(a.suhail))
        if errors:
            print(f"QA FAILED: {len(errors)} blocking issue(s)")
            for e in errors:print("ERROR:",e)
            return 1
        shutil.copy2(tmp,report)
    print("QA PASSED - report repaired and safe to deliver")
    return 0
if __name__=="__main__":raise SystemExit(main())
