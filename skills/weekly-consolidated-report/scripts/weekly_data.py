"""Weekly-only template, transaction schema and approved status mapping."""
import math
import re
from collections import Counter
from copy import deepcopy
from pathlib import Path
from openpyxl import load_workbook
from pptx.table import _Cell
from pptx.dml.color import RGBColor
from pptx import Presentation
from report_common import pptx_helpers as helpers
from report_common import notices, status_styles, validation
from report_common.table_typography import DEFAULT_INSET
from report_common.header_heights import TABLE_GAP
from report_common.validation import field_key

MASTER=Path(__file__).resolve().parents[1]/'assets/weekly-report-master.pptx'
MASTER_HASH='173d48ace27661a4900cc1d6428826e41515d6d4a764a26385171701ead5d96a'
# Each summary (tasks, Suhail, transactions, digital experience) stays on its one
# slide. When its measured 11 pt tables do not fit, the table text is reduced in
# half-point steps down to the last size here; when even that does not fit, the
# rows that do not fit move to an appendix slide and the build records a notice.
SUMMARY_SIZES=(11,10.5,10,9.5,9,8.5,8)
CELL_INSETS=2*DEFAULT_INSET  # top+bottom cell margins; they do not scale with the font
PT=12700  # EMU per point; layout files store row heights in points

def emu(points):
    return math.ceil(points*PT)
COVER_DATE_ID=19  # cover text box holding the report date
TASK_DETAIL_TITLE='تفاصيل المهام ({})'

def task_source(title):
    """The task source named in a «تفاصيل المهام (…)» table title, or None."""
    found=re.fullmatch(re.escape(TASK_DETAIL_TITLE).replace(r'\{\}','(.*)'),title)
    return found.group(1) if found else None

TRANS_FIELDS=('رقم المعاملة','موضوع المعاملة','الجهة الوارد منها المعاملة','تاريخ إنشاء المعاملة','تاريخ الإنجاز المخطط')
# User confirmed on 2026-09-29 that Monday status colors apply to weekly too.
# Weekly typography, margins, alignment and borders still come from weekly cells.
STATUS_FILLS=status_styles.FILLS
# A status the template has no style for keeps its own text and takes this gray style.
NEUTRAL_STATUS={'task':'معلق','project':'لم تبدأ'}
# How the notices name each summary slide: tasks, Suhail, transactions, digital experience.
SUMMARY_LABELS=('ملخص المهام','ملخص مشاريع سهيل','ملخص المعاملات','ملخص التجربة الرقمية')

def approved_styles(template=MASTER,kind='task'):
    result={}
    presentation=template if hasattr(template,'slides') else Presentation(template)
    for slide in helpers.detail_pattern_slides(presentation,kind):
        for shape in helpers.detail_tables(slide,kind):
            for row in list(shape.table.rows)[2:]:
                cell=row.cells[4];status=status_styles.canonical(cell.text)
                if not status:continue
                if status not in STATUS_FILLS:raise ValueError('Unknown weekly master status: '+status)
                if status in result and status_styles.signature(cell)!=status_styles.signature(result[status]):
                    raise ValueError('Conflicting weekly status examples: '+status)
                result.setdefault(status,cell)
    if 'مكتملة' not in result:raise ValueError('Weekly master missing base status-cell example')
    for status,color in STATUS_FILLS.items():
        if status in result:
            fill=result[status].fill
            if fill.type is None or str(fill.fore_color.rgb)!=color:
                raise ValueError('Weekly status example conflicts with approved mapping: '+status)
        else:
            cell=_Cell(deepcopy(result['مكتملة']._tc),result['مكتملة']._parent)
            cell.fill.solid();cell.fill.fore_color.rgb=RGBColor.from_string(color)
            helpers.set_cell_text(cell,status)
            result[status]=cell
    return result

def summary_columns(title,headers,is_task):
    last='طلب الدعم' if title=='طلبات الدعم' else 'التحديث' if title=='أبرز التحديثات' else 'التحدي' if title=='التحديات' else 'ملاحظات'
    source=['#','المهمة' if is_task else 'المشروع','القطاع',last]
    # The approved suspended-task pattern labels its notes destination التحديث.
    # It still receives the independent suspended-task notes, never project updates.
    keys=[last if title=='المهام المعلقة' and h=='التحديث' else h for h in headers]
    if sorted(keys)!=sorted(source):raise ValueError('Unexpected summary headers: '+title)
    return [source.index(h) for h in keys]

def set_cell_exact(cell,value):
    text=helpers.val(value).replace('\r\n','\n')
    helpers.set_cell_text(cell,text)
    count=len(text.split('\n'))
    while len(cell.text_frame.paragraphs)>count:
        paragraph=cell.text_frame.paragraphs[-1]._p
        paragraph.getparent().remove(paragraph)
    while len(cell.text_frame.paragraphs)<count:cell.text_frame.add_paragraph()

def transaction_data(path):
    chart,kpis=validation.source_metrics(path,'المعاملات',['مكتملة','على المخطط','متأخر'],['B2','B3','B4','B5'])
    wb=load_workbook(path,data_only=True)
    try:
        ws=wb['المعاملات']
        expected=['رقم المعاملة','موضوع المعاملة','القطاع','الجهة الوارد منها المعاملة','تاريخ إنشاء المعاملة','تاريخ الإنجاز المخطط']
        header=None
        for r in range(1,ws.max_row+1):
            got={field_key(ws.cell(r,c).value):c for c in range(1,ws.max_column+1) if ws.cell(r,c).value is not None}
            if all(field_key(s) in got for s in expected):header=r;columns=got;break
        if header is None:raise ValueError('Transactions detail header missing')
        records=[]
        for r in range(header+1,ws.max_row+1):
            values=[ws.cell(r,columns[field_key(s)]).value for s in expected]
            if values[0] in (None,'','لا يوجد'):continue
            records.append([helpers.val(v) for v in values])
        # A workbook that disagrees with itself is the user's to settle: both
        # parts are transferred as written and the difference is reported.
        if len(records)!=kpis[3]:
            notices.warn(f'ملف المعاملات: مؤشر المتأخرة ({kpis[3]}) يختلف عن عدد صفوف جدول المعاملات المتأخرة ({len(records)}). نقلت الاثنين كما وردا.','source')
        actual=Counter(row[2] for row in records)
        counts={str(row[0]):int(row[3] or 0) for row in chart}
        differing=sorted(s for s in actual.keys()|counts.keys() if actual[s]!=counts.get(s,0))
        if differing:
            notices.warn('ملف المعاملات: عدد المعاملات المتأخرة في جدول القطاعات يختلف عن صفوف الجدول التفصيلي للقطاعات: '+'، '.join(differing)+'. نقلت الاثنين كما وردا.','source')
        return kpis,chart,records
    finally:wb.close()

def scaled_height(height,size):
    """Upper bound for a measured 11 pt row set at `size` pt.

    Smaller text wraps into the same number of lines or fewer, and each line is
    `size/11` as tall, so the text part of the row shrinks at least in that
    ratio. The cell insets stay the same.
    """
    text=max(0,height-CELL_INSETS)
    return CELL_INSETS+math.ceil(text*size/11)

def measured_rows(body,smaller):
    """Body heights (EMU) at a summary size: measured at that size when the layout file has them, else the 11 pt bound."""
    def at(size):
        if size==11:return body
        found=smaller.get(str(size))
        return [emu(h) for h in found] if found and len(found)==len(body) else [scaled_height(h,size) for h in body]
    return at

def fit_summary(tables,available,gap=TABLE_GAP,sizes=SUMMARY_SIZES):
    """Largest of `sizes` at which every table fits `available` (EMU) with `gap` between them.

    `tables` is a list of ``(fixed, at)``: title/header heights at 11 pt and a
    function giving the body heights at a size (see `measured_rows`). Returns
    ``(size, heights)`` with one list of row heights per table, or None when
    no size fits; the caller then moves the overflow to an appendix slide.
    """
    room=available-gap*max(0,len(tables)-1)
    for size in sizes:
        heights=[[scaled_height(h,size) for h in fixed]+at(size) for fixed,at in tables]
        if sum(map(sum,heights))<=room:return size,heights
    return None
