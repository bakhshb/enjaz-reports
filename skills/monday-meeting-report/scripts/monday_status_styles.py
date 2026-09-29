"""Approved Monday status styles, read from the actual master examples."""
from copy import deepcopy
from pathlib import Path
import os
import shutil
import uuid
from pptx import Presentation
from pptx.oxml.ns import qn

MASTER=Path(__file__).resolve().parents[1]/'assets/monday-meeting-master.pptx'
ALIASES={'مكتمل':'مكتملة','متأخرة':'متأخر','معلقة':'معلق','لم يبدأ':'لم تبدأ'}
TASK_STATUSES={'مكتملة','على المخطط','متأخر','معلق'}
PROJECT_STATUSES={'مكتملة','على المخطط','متأخر','لم تبدأ'}
SUPPORTED=TASK_STATUSES | PROJECT_STATUSES

def canonical(text):
    text=' '.join(text.split())
    return ALIASES.get(text,text)

def signature(cell):
    properties=cell._tc.get_or_add_tcPr()
    def xml(element):
        if element is None:return None
        return (element.tag,tuple(sorted(element.attrib.items())),tuple(xml(c) for c in element))
    fill=next((c for c in properties if c.tag in {qn('a:'+n) for n in ('solidFill','noFill','gradFill','pattFill')}),None)
    runs=[]
    for paragraph in cell.text_frame.paragraphs:
        for run in paragraph.runs:
            if not run.text.strip():continue
            cs=run._r.find('./'+qn('a:rPr')+'/'+qn('a:cs'))
            runs.append((run.font.name,cs.get('typeface') if cs is not None else run.font.name,
                         run.font.size.pt if run.font.size else None,bool(run.font.bold),bool(run.font.italic),
                         str(run.font.underline),str(paragraph.alignment),
                         paragraph._p.get_or_add_pPr().get('rtl','0')))
    return (xml(fill),properties.get('anchor','t'),tuple(sorted(set(runs),key=repr)))

def examples(template=MASTER):
    result={}
    for slide in Presentation(template).slides:
        for shape in slide.shapes:
            if not shape.has_table or len(shape.table.columns)!=6:continue
            table=shape.table
            if len(table.rows)<3 or table.cell(1,4).text.strip()!='الحالة':continue
            for row in list(table.rows)[2:]:
                cell=row.cells[4];status=canonical(cell.text)
                if not status:continue
                if status not in SUPPORTED:raise ValueError('Unsupported template detail status: '+status)
                if status in result and signature(cell)!=signature(result[status]):
                    raise ValueError('Conflicting approved template examples for status: '+status)
                result.setdefault(status,cell)
    return result

def apply(cell,example):
    # Copy the approved status cell's fill, borders, alignment and text formatting.
    for tag in ('a:txBody','a:tcPr'):
        previous=cell._tc.find(qn(tag))
        replacement=deepcopy(example._tc.find(qn(tag)))
        if previous is not None:cell._tc.replace(previous,replacement)
        else:cell._tc.append(replacement)


def publish(source,destination):
    """Replace atomically, inheriting destination-directory access on Windows.

    Moving directly out of TemporaryDirectory would retain its owner-only ACL.
    Copy the validated bytes to a sibling with normal inherited permissions first.
    """
    destination=Path(destination)
    staged=destination.with_name('.monday-publish-'+uuid.uuid4().hex+'.pptx')
    try:
        with Path(source).open('rb') as reader,staged.open('xb') as writer:
            shutil.copyfileobj(reader,writer)
        os.replace(staged,destination)
    finally:
        staged.unlink(missing_ok=True)
