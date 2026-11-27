"""Weekly-only template, transaction schema and approved status mapping."""
import math
import re
from collections import Counter, namedtuple
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
# The four summaries (tasks, Suhail, transactions, digital experience) each stay
# on their one slide, and share one look. When every summary fits at 11 pt with
# the template margins nothing changes. When any does not, every summary table
# takes the compact body-cell margins the template uses in its detail tables;
# the three standard summaries then share the largest size, not below
# SUMMARY_FLOOR, at which each fits; a table that still does not fit moves whole
# to an appendix slide. The digital summary alone may go smaller, down to
# DIGITAL_SUMMARY_FLOOR, and its rows are squeezed past that. The ladder is also
# the set of sizes a layout file is measured at.
SUMMARY_SIZES=(11,10.5,10,9.5,9,8.5,8)
SUMMARY_FLOOR=10
DIGITAL_SUMMARY_FLOOR=8
CELL_INSETS=2*DEFAULT_INSET  # top+bottom cell margins of the summary tables; they do not scale with the font
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

def size_key(size):
    """How a size is written as a key of the layout file."""
    return f'{size:g}'

def compact_margin(prs):
    """The margin the template uses in the body cells of its task detail tables (all four sides of a summary body cell take it).

    Read from the first task detail pattern table, the left margin of its first
    non-status body cell, so a template with a different value changes the
    summaries with it.
    """
    slides=helpers.detail_pattern_slides(prs,'task')
    shapes=helpers.detail_tables(slides[0],'task') if slides else []
    rows=list(shapes[0].table.rows)[2:] if shapes else []
    if not rows:raise ValueError('Missing weekly detail patterns')
    return rows[0].cells[0].margin_left

def set_body_margins(table,margin,first=2):
    """Give every cell of the body rows (from row index `first`) `margin` on all four sides; the title and header rows keep theirs."""
    for row in list(table.rows)[first:]:
        for cell in row.cells:
            cell.margin_left=cell.margin_right=cell.margin_top=cell.margin_bottom=margin

def body_insets(table,first=2):
    """Top plus bottom margin of the first body row of `table`."""
    cell=list(table.rows)[first].cells[0]
    return cell.margin_top+cell.margin_bottom

def scaled_height(height,size,insets=CELL_INSETS,new_insets=None):
    """Upper bound for a measured 11 pt row set at `size` pt.

    Smaller text wraps into the same number of lines or fewer, and each line is
    `size/11` as tall, so the text part of the row shrinks at least in that
    ratio. The cell insets (`insets`, top plus bottom) stay the same, or become
    `new_insets` when the margins are reduced: a smaller margin never makes a
    row taller.
    """
    text=max(0,height-insets)
    return (insets if new_insets is None else new_insets)+math.ceil(text*size/11)

def measured_rows(body,compact,insets=CELL_INSETS,compact_insets=None):
    """``at(size, compact)``: body heights (EMU) of a table at a summary size.

    `body` is the measured 11 pt height of each row with the template margins.
    With compact margins the heights come from the layout file's `compact`
    measurements (points, by size) when it has them for these rows; an older
    layout file, or one from the PowerShell measurer, has none, and the bound
    computed from `body` is used instead. Without compact margins only 11 pt is
    ever asked for.
    """
    def usable(found):
        return (isinstance(found,list) and len(found)==len(body)
                and all(isinstance(h,(int,float)) and math.isfinite(h) and h>0 for h in found))
    def at(size,is_compact=False):
        if not is_compact:
            return list(body) if size==11 else [scaled_height(h,size,insets) for h in body]
        found=(compact or {}).get(size_key(size))
        if usable(found):return [emu(h) for h in found]
        return [scaled_height(h,size,insets,insets if compact_insets is None else compact_insets) for h in body]
    return at

# One summary as the planner sees it: its display label, its tables as
# ``(fixed, at)`` (title/header heights at 11 pt and a function giving the body
# heights at a size, see `measured_rows`), the height available to them (EMU)
# and whether it is the digital-experience one.
Summary=namedtuple('Summary','label tables available digital',defaults=(False,))
# The decision for the whole deck: whether every summary table takes the
# compact margins, the size shared by the three standard summaries, the labels
# that did not fit at 11 pt with template margins, which summaries fit at the
# shared look (a standard one that does not moves tables to an appendix), the
# digital summary's own size (None without one) and whether it still needs its
# rows squeezed.
SummaryPlan=namedtuple('SummaryPlan','compact size tight fits digital_size squeeze')

def summary_fits(summary,size,compact,gap=TABLE_GAP):
    """Whether all tables of `summary` fit its space at `size` with `gap` between them."""
    room=summary.available-gap*max(0,len(summary.tables)-1)
    used=sum(sum(scaled_height(h,size) for h in fixed)+sum(at(size,compact)) for fixed,at in summary.tables)
    return used<=room

def _ladder(floor,ceiling=11):
    return tuple(s for s in SUMMARY_SIZES if floor<=s<=ceiling) or (ceiling,)

def floors(min_pt=None):
    """(standard floor, digital floor): the defaults, or one explicit `min_pt` for both."""
    return (SUMMARY_FLOOR,DIGITAL_SUMMARY_FLOOR) if min_pt is None else (min_pt,min_pt)

def plan_summaries(summaries,overflow='shrink',floor=SUMMARY_FLOOR,digital_floor=DIGITAL_SUMMARY_FLOOR):
    """Decide margins and sizes of all summaries together, so they look the same.

    1. Every summary fits at 11 pt with the template margins: nothing changes.
    2. Otherwise every summary takes the compact margins.
    3. Each standard summary finds the largest size from 11 down to `floor` at
       which it fits; the shared size is the smallest of them. With
       ``overflow='appendix'`` the size stays 11. A summary that does not fit at
       the floor sets the shared size to the floor and is marked as not fitting
       (its overflow goes to an appendix).
    4. The digital summary takes the shared size when it fits there, else its
       table alone goes down to `digital_floor`, and past it `squeeze` is set.
    """
    digital=next((s for s in summaries if s.digital),None)
    standard=[s for s in summaries if not s.digital]
    tight=tuple(s.label for s in summaries if not summary_fits(s,11,False))
    if not tight:
        return SummaryPlan(False,11,(),{s.label:True for s in summaries},11 if digital else None,False)
    sizes=(11,) if overflow=='appendix' else _ladder(floor)
    best=[next((z for z in sizes if summary_fits(s,z,True)),None) for s in standard]
    shared=sizes[-1] if None in best else min(best,default=11)
    fits={s.label:summary_fits(s,shared,True) for s in standard}
    digital_size=None;squeeze=False
    if digital:
        options=_ladder(digital_floor,shared)
        digital_size=next((z for z in options if summary_fits(digital,z,True)),None)
        squeeze=digital_size is None
        if squeeze:digital_size=options[-1]
        fits[digital.label]=digital_size==shared
    return SummaryPlan(True,shared,tight,fits,digital_size,squeeze)

def announce_plan(plan,shown=None):
    """The deck-wide notice (kind «space») for the margins and size `plan` chose; none when everything fits as the template has it."""
    if not plan.tight:return
    shown=shown or {}
    names='، '.join(shown.get(label,label) for label in plan.tight)
    if plan.size==11:
        notices.warn(f'{names}: لم تتسع الجداول في الشريحة بخط 11. تم تقليص هوامش خلايا الجداول في جميع شرائح الملخص ليبقى شكلها موحدا، ويبقى الخط 11. المقترح: اختصار النص في الملف.','space')
    else:
        notices.warn(f'{names}: لم تتسع الجداول في الشريحة بخط 11. تم تقليص هوامش خلايا الجداول وتصغير الخط إلى {plan.size:g} في جميع شرائح الملخص ليبقى شكلها موحدا. المقترح: اختصار النص في الملف.','space')
