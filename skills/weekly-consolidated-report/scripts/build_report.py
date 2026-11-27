# /// script
# requires-python = ">=3.12"
# dependencies = ["openpyxl>=3.1", "python-pptx>=1.0.2", "lxml>=5.3", "pdfplumber>=0.11"]
# ///
"""Generate a validated weekly report.

With ``--measure`` (the delivery path) one command builds a first candidate,
measures every row through LibreOffice, rebuilds with those heights and
validates the result. The build stops only when it cannot produce the report;
every best assumption it makes is written to ``<output>.notices.json`` for the
final reply. Rendered layout review remains required.
"""
import argparse, copy, datetime as dt, json, math, subprocess, sys, tempfile, zipfile
from copy import deepcopy
from pathlib import Path
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.oxml.ns import qn
sys.path.insert(0,str(Path(__file__).resolve().parents[2]))
sys.path.insert(0,str(Path(__file__).resolve().parent))
from report_common import pptx_helpers as builder
from report_common.pptx_helpers import content_bottom
from report_common import notices, status_styles, reflow
from report_common.validation import sha
from report_common.render import abar_installed
from report_common import header_heights as shared_heights
from report_common.header_heights import MIN_ROW_HEIGHT
from report_common.table_typography import set_table_sizes, set_table_directions, apply_table
from report_common.package_tools import reachable_parts
from weekly_data import transaction_data, field_key, MASTER, MASTER_HASH, approved_styles, set_cell_exact, summary_columns, TABLE_GAP, emu, fit_summary, measured_rows, scaled_height, SUMMARY_SIZES, COVER_DATE_ID, TASK_DETAIL_TITLE, NEUTRAL_STATUS, SUMMARY_LABELS

from update_format import format_update
from report_common.update_format import format_table_updates
from weekly_artwork import place_table_copy, prepare_navigation

def remove(shape):shape._element.getparent().remove(shape._element)

# Header heights live in report_common/header_heights.py with the reasons for
# each value. The summary task tables are auto-height in the masters, and the
# transactions header wraps to two lines, so both get the shared fixed heights.
TASK_HEADER_HEIGHT=shared_heights.TASK_HEADER
TRANSACTION_HEADER_HEIGHT=shared_heights.TWO_LINE_HEADER
TRANSACTION_TITLE='المعاملات المتأخرة'
SUMMARY_CONTINUATION_TOP=900000

def header_heights(pattern):
    """Approved title/header heights; auto-height rows are summary task tables."""
    heights=[row.height or TASK_HEADER_HEIGHT for row in list(pattern.table.rows)[:2]]
    if builder.table_title(pattern)==TRANSACTION_TITLE:heights[1]=TRANSACTION_HEADER_HEIGHT
    return heights

def usable_measurement(title,measured,records):
    """The layout-file entry for `title`, or None (with a notice) when it does not describe these records."""
    if measured is None:
        notices.warn(f'لا يوجد قياس لجدول «{title}» في ملف القياس؛ استخدمت الارتفاع التقديري له. أعد القياس قبل التسليم.','layout')
        return None
    heights=measured.get('header',[])+measured.get('body',[])
    valid=(len(measured.get('header',[]))==2 and len(measured.get('body',[]))==len(records)
           and all(isinstance(h,(int,float)) and math.isfinite(h) and h>0 for h in heights))
    if not valid:
        notices.warn(f'قياس جدول «{title}» لا يطابق سجلاته الحالية؛ استخدمت الارتفاع التقديري له. أعد القياس قبل التسليم.','layout')
        return None
    return measured



def flow(prs,seed,sections,summary=False,measurements=None,bullet_style=None,sizes=None,
         overflow='shrink',min_pt=SUMMARY_SIZES[-1],label=None):
    """Lay out titled tables from the top of `seed` down to the footer.

    Detail tables continue on copies of `seed`. A summary built with
    measurements stays on its one slide: when its 11 pt tables do not fit, the
    table text shrinks to the largest size, not below `min_pt`, whose measured
    rows fit, and `sizes` records each shrunk table so the size survives the
    final typography pass. When even `min_pt` does not fit, or when `overflow`
    is ``appendix``, the rows that do not fit move to an appendix slide after
    the summary at 11 pt. Every such decision is recorded as a notice; the
    build never stops for lack of space. Without measurements the placement is
    provisional and only feeds the measurement.
    """
    seed._element.set('showMasterSp','1')
    patterns={builder.table_title(sh):sh for sh in builder.tables(seed)}
    top=min(sh.top for sh in patterns.values())
    bottom=content_bottom(prs,seed)
    label=label or 'الملخص'
    # Continuations preserve layout/navigation, but not the source KPI/chart.
    def keep(sh):
        outside_summary = sh.top + sh.height <= 700000 or sh.top >= bottom
        return (not sh.has_table and not sh.has_chart
                and not getattr(sh,'has_text_frame',False)
                and (not summary or outside_summary))
    current=seed;y=top;targets=[seed]
    # Summary continuations (provisional pass and appendix) start under the header buttons.
    start=SUMMARY_CONTINUATION_TOP if summary else top
    for sh in list(builder.tables(seed)):remove(sh)
    def place(title,pattern,chunk,row_heights,size=11):
        nonlocal y
        sh=place_table_copy(current,pattern._element);sh.top=y
        builder.fill_titled_table(sh,chunk,title=title)
        for row,values in zip(list(sh.table.rows)[2:],chunk):
            for cell,value in zip(row.cells,values):set_cell_exact(cell,value)
        if summary:
            format_table_updates(sh.table,bullet_style,paragraph_items=True)
        for row,height in zip(sh.table.rows,row_heights):row.height=height
        sh.height=sum(row.height for row in sh.table.rows)
        y+=sh.height+TABLE_GAP
        if size!=11 and sizes is not None:sizes.append((sh,size))
        return sh
    prepared=[]
    for title,records,pattern in sections:
        if not records:continue
        measured=usable_measurement(title,measurements.get(title),records) if measurements is not None else None
        header=[emu(h) for h in measured['header']] if measured else header_heights(pattern)
        base=builder._body_base_height(pattern.table,2)
        body=[emu(h) for h in measured['body']] if measured else [base*builder.row_units(row,pattern.table,11) for row in records]
        at=measured_rows(body,(measured or {}).get('sizes',{}))
        prepared.append((title,records,pattern,header,body,bool(measured),at))
    fitted=summary and measurements is not None
    first_size=11
    if fitted and prepared:
        # One slide per summary: every table below the indicators and chart, at
        # the largest size (11 pt first) whose measured rows fit.
        allowed=tuple(s for s in SUMMARY_SIZES if s>=min_pt) or (SUMMARY_SIZES[0],)
        names='، '.join(p[0] for p in prepared)
        if overflow=='appendix':
            found=fit_summary([(p[3],p[6]) for p in prepared],bottom-top,sizes=(11,))
            if found is None:
                notices.warn(f'{label}: الجداول ({names}) لا تتسع في شريحة الملخص بخط 11، فنقلت كل جدول لا يتسع كاملا إلى شريحة ملحق بعدها كما طُلب، دون قطع جدول.','space')
        else:
            found=fit_summary([(p[3],p[6]) for p in prepared],bottom-top,sizes=allowed)
            if found is None:
                first_size=allowed[-1]
                notices.warn(f'{label}: الجداول ({names}) لا تتسع في شريحة واحدة حتى بخط {first_size:g}، فأبقيت في الشريحة الجداول التي تتسع كاملة بهذا الخط ونقلت كل جدول لا يتسع كاملا إلى شريحة ملحق بعدها بخط 11 دون قطعه. '
                             'البديل: اختصار صياغة الخلايا الطويلة ثم إعادة البناء.','space')
            elif found[0]!=11:
                notices.warn(f'{label}: صغرت خط الجداول ({names}) إلى {found[0]:g} لتبقى في شريحة واحدة. '
                             'البدائل: اختصار الصياغة، أو نقل الزائد إلى شريحة ملحق (--summary-overflow appendix)، أو رفع الحد الأدنى للخط (--summary-min-pt).','space')
        if found is not None:
            size,heights=found
            for (title,records,pattern,*_),row_heights in zip(prepared,heights):
                place(title,pattern,records,row_heights,size)
            return targets
    for title,records,pattern,header,body,measured,at in prepared:
        chunk=[];chunk_heights=[]
        def rows_at(size):
            return [scaled_height(h,size) for h in header],at(size)
        # Only the summary slide itself may hold shrunk rows; an appendix is at 11 pt.
        page_size=first_size if current is seed else 11
        fixed,heights=rows_at(page_size) if fitted else (header,body)
        index=0
        if summary:
            # A summary table is never cut: when it does not fit under the tables above it,
            # the whole table moves to a fresh slide (the appendix). Only a table taller than
            # a whole slide is split below, and then its header rows repeat on each part.
            if y+sum(fixed)+sum(heights)<=bottom:
                place(title,pattern,records,fixed+heights,page_size);continue
            fixed11,heights11=rows_at(11) if fitted else (header,body)
            if sum(fixed11)+sum(heights11)<=bottom-start:
                current=builder.clone_slide(prs,seed,after=current,shape_predicate=keep)
                targets.append(current);y=start
                page_size=11
                place(title,pattern,records,fixed11+heights11,11);continue
            if y>(top if current is seed else start):
                # Taller than a slide: it starts on a fresh slide so its first part is not squeezed under another table.
                current=builder.clone_slide(prs,seed,after=current,shape_predicate=keep)
                targets.append(current);y=start
                page_size=11;fixed,heights=fixed11,heights11
        while index<len(records):
            height=heights[index]
            room=bottom-(top if current is seed else start)-sum(fixed)
            if height>room:
                # Never stop for one oversized record: it takes a whole page and is flagged.
                if measured:notices.warn(f'سجل في «{title}» أطول من مساحة شريحة كاملة؛ وضعته في شريحة بمفرده وقد يتجاوز النص حدها. اختصر نصه أو قسّمه.','space')
                height=max(room,MIN_ROW_HEIGHT)
            if y+sum(fixed)+sum(chunk_heights)+height>bottom and (chunk or y>(top if current is seed else start)):
                if chunk:place(title,pattern,chunk,fixed+chunk_heights,page_size)
                current=builder.clone_slide(prs,seed,after=current,shape_predicate=keep)
                targets.append(current);y=start
                chunk=[];chunk_heights=[]
                page_size=11
                fixed,heights=rows_at(11) if fitted else (header,body)
                continue
            chunk.append(records[index]);chunk_heights.append(height);index+=1
        if chunk:place(title,pattern,chunk,fixed+chunk_heights,page_size)
    return targets

DIGIT_WIDTH=320040  # one 36 pt bold Abar digit plus its share of the box insets (0.35 in)

def fit_number(box):
    """Widen a KPI box around its centre so a long number stays on one line.

    The masters size these boxes for three digits; «1032» wrapped to «103 / 2»
    over the label below it.
    """
    need=len(box.text_frame.text.strip())*DIGIT_WIDTH
    if need>box.width:
        box.left-= (need-box.width)//2;box.width=need

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
    show_all_labels(chart)
    return records

def show_all_labels(chart):
    """Label every bar segment.

    The master hides labels on individual points (``c:dLbl`` with
    ``c:delete``) chosen for its sample data, and turns them off for the
    delayed and suspended series altogether.  Once the data is replaced
    those switches land on arbitrary bars, so every series gets the
    template's visible label style with no per-point overrides.
    """
    series=chart.findall('.//'+qn('c:ser'))
    pattern=None
    for ser in series:
        labels=ser.find(qn('c:dLbls'))
        if labels is not None and labels.find(qn('c:showVal')) is not None and labels.find(qn('c:showVal')).get('val')=='1':
            pattern=deepcopy(labels);break
    if pattern is None:raise ValueError('chart has no visible data-label pattern')
    for point in pattern.findall(qn('c:dLbl')):pattern.remove(point)
    delete=pattern.find(qn('c:delete'))
    if delete is not None:pattern.remove(delete)
    for ser in series:
        old=ser.find(qn('c:dLbls'))
        fresh=deepcopy(pattern)
        if old is not None:old.addprevious(fresh);ser.remove(old)
        else:
            # Schema order: dLbls precedes trendline/errBars/cat/val.
            anchor=next((ser.find(qn(t)) for t in ('c:trendline','c:errBars','c:cat','c:val') if ser.find(qn(t)) is not None),None)
            anchor.addprevious(fresh)

def reason(exc):
    """The text of `exc` when it is Arabic, else a plain Arabic phrase; notices are never English."""
    text=str(exc)
    return text if any('؀'<=ch<='ۿ' for ch in text) else 'الملف ناقص أو تالف أو غير مطابق للصيغة المتوقعة'

def read_measurements(args,inputs):
    """(tables, digital) from the --measurements file; ValueError (or a KeyError/TypeError for a malformed file) when it cannot be used with these inputs."""
    measured=json.loads(Path(args.measurements).read_text(encoding='utf-8-sig'))
    hashes=measured['hashes']
    for key,path in inputs:
        if sha(path)!=hashes[key].lower():raise ValueError('ملف القياس قديم: تغيّر '+{'tasks':'ملف المهام','suhail':'ملف مشاريع سهيل','transactions':'ملف المعاملات','template':'القالب'}[key])
    if args.model=='digital':
        for key,path in [('digital_pdf',Path(args.digital_pdf)),('digital_data',Path(args.digital_data))]:
            if sha(path)!=hashes.get(key,'').lower():raise ValueError('ملف القياس قديم: تغيّر '+('ملف PDF للنضج الرقمي' if key=='digital_pdf' else 'ملف بيانات النضج الرقمي'))
    elif 'digital_pdf' in hashes or 'digital_data' in hashes:
        raise ValueError('قياسات النضج الرقمي لا تصلح للنموذج القياسي')
    if measured.get('typography')!='11pt-suhail-paragraph-updates-v3':raise ValueError('ملف القياس قديم: تغير تنسيق تحديثات سهيل')
    tables=measured['tables']
    digital=None
    if args.model=='digital':
        digital=measured.get('digital')
        if not digital:raise ValueError('ملف القياس لا يحوي جداول النضج الرقمي')
    return tables,digital

def build_candidate(args):
    if args.model=='digital':
        from digital_maturity import MASTER as DIGITAL_MASTER, MASTER_HASH as DIGITAL_HASH, load_reviewed
        digital_data=load_reviewed(args.digital_data,args.digital_pdf)
        bundled=Path(DIGITAL_MASTER);expected_hash=DIGITAL_HASH
    else:
        digital_data=None;bundled=MASTER;expected_hash=MASTER_HASH
    template=Path(args.template) if args.template else bundled
    if not args.template and sha(template)!=expected_hash:
        raise ValueError('Bundled weekly template checksum mismatch')
    tasks=Path(args.tasks);projects=Path(args.suhail);transactions=Path(args.transactions)
    measurements=None;digital_measurements=None;sizes=[]
    if args.measurements:
        try:
            measurements,digital_measurements=read_measurements(args,[('tasks',tasks),('suhail',projects),('transactions',transactions),('template',template)])
        except (ValueError,KeyError,TypeError,AttributeError,OSError) as exc:
            notices.warn(f'تعذر استخدام ملف القياس ({reason(exc)}). أهملته وبنيت التقرير بارتفاعات تقديرية، وهو غير جاهز للتسليم قبل إعادة القياس بـ --measure.','layout')
            measurements=None;digital_measurements=None;args.measurements=None  # the build is then provisional
    tk,tc,ts,td=builder.task_data(tasks);pk,pc,ps,pd=builder.suhail_data(projects);xk,xc,xd=transaction_data(transactions)
    prs=Presentation(template);set_table_sizes(prs);slides=list(prs.slides)
    if digital_data is not None:
        from digital_maturity import populate
        populate(prs,digital_data,digital_measurements,sizes,min_pt=args.summary_min_pt)
    # Discover roles first; IDs are checked within the identified summary role.
    summary=[builder.find_slide_with_text(prs,label) for label in ('إجمالي المهام','إجمالي مشاريع سهيل','إجمالي المعاملات')]
    task_patterns=builder.detail_pattern_slides(prs,'task')
    project_patterns=builder.detail_pattern_slides(prs,'project')
    if not task_patterns or not project_patterns:raise ValueError('Missing weekly detail patterns')
    styles={kind:approved_styles(prs,kind) for kind in ('task','project')}
    bullet_style=next((deepcopy(p._p.get_or_add_pPr()) for sl in project_patterns
                       for sh in builder.detail_tables(sl,'project') for row in sh.table.rows
                       for cell in row.cells for p in cell.text_frame.paragraphs
                       if p._p.find('./'+qn('a:pPr')+'/'+qn('a:buChar')) is not None),None)

    date=dt.datetime.now().astimezone().date()
    builder.set_text(builder.shape_by_id(slides[0],COVER_DATE_ID),f'{date.day} {builder.MONTHS[date.month-1]} {date.year}')
    chart_specs=[]
    for sl,kpis,ids,records,statuses,chart_name in [
        (summary[0],tk,[11,14,16,7,6],tc,['مكتملة','على المخطط','متأخر','معلق'],'chart1.xml'),
        (summary[1],pk,[10,12,14,8,7],pc,['مكتملة','على المخطط','متأخر','لم تبدأ'],'chart2.xml'),
        (summary[2],xk,[24,29,31,26],xc,['مكتملة','على المخطط','متأخر'],'chart3.xml')]:
        for sid,value in zip(ids,kpis):
            box=builder.shape_by_id(sl,sid);builder.set_text(box,value);fit_number(box)
        ordered=update_chart(next(sh for sh in sl.shapes if sh.has_chart),records,statuses)
        chart_specs.append((chart_name,ordered,statuses))
    space=dict(overflow=args.summary_overflow,min_pt=args.summary_min_pt)
    for sl,sections in [(summary[0],ts),(summary[1],ps)]:
        # Map source header order to the approved template header order.
        seq=[]
        for title,records in sections.items():
            pattern=next(sh for sh in builder.tables(sl) if builder.table_title(sh)==title)
            headers=[c.text.strip() for c in pattern.table.rows[1].cells]
            columns=summary_columns(title,headers,sl==summary[0])
            seq.append((title,[[row[i] for i in columns] for row in records],pattern))
        flow(prs,sl,seq,True,measurements,bullet_style,sizes,label=SUMMARY_LABELS[0 if sl==summary[0] else 1],**space)
    tframe=next(sh for sh in builder.tables(summary[2]))
    headings=[field_key(c.text) for c in tframe.table.rows[1].cells]
    xheaders=[field_key(s) for s in ['رقم المعاملة','موضوع المعاملة','القطاع','الجهة الوارد منها المعاملة','تاريخ إنشاء المعاملة','تاريخ الإنجاز المخطط']]
    flow(prs,summary[2],[(TRANSACTION_TITLE,[[row[xheaders.index(h)] for h in headings] for row in xd],tframe)],True,measurements,sizes=sizes,label=SUMMARY_LABELS[2],**space)
    generated={};unused=[]
    for kind,patterns,sections in [('task',task_patterns,td),('project',project_patterns,pd)]:
        seed=patterns[0];pattern=builder.detail_tables(seed,kind)[0]
        # Remove any additional sample tables before filling a single flowing stream.
        seq=[((TASK_DETAIL_TITLE.format(title) if kind=='task' else title),[row[:6] for row in rows],pattern) for title,rows in sections]
        generated[kind]=flow(prs,seed,seq,measurements=measurements)
        if not any(rows for _,rows in sections):generated[kind]=[]
        for sl in generated[kind]:
            for sh in builder.tables(sl,6):
                for row in list(sh.table.rows)[2:]:
                    cell=row.cells[4];status=status_styles.canonical(cell.text)
                    allowed=status_styles.TASK_STATUSES if kind=='task' else status_styles.PROJECT_STATUSES
                    style=status_styles.style_key(status,allowed,styles[kind],NEUTRAL_STATUS[kind])
                    if style!=status:
                        # A status the template has no style for keeps its text and takes the neutral fill.
                        notices.warn(f'حالة غير معروفة «{status}» في «{builder.table_title(sh)}»؛ أبقيت نصها كما ورد بتعبئة رمادية محايدة، ولم تدخل في ألوان الحالات المعتمدة.','status')
                    status_styles.apply(cell,styles[kind][style])
                    set_cell_exact(cell,status)
                    if kind=='project':
                        format_update(row.cells[5],bullet_style)
        unused.extend(patterns[1:])
        if not generated[kind]:unused.append(seed)
    for slide in unused:builder.remove_slide(prs,slide)
    # Repoint all internal slide links from removed examples to their surviving section.
    retained={sl.part for sl in prs.slides}
    for part in list(prs.part.package.iter_parts()):
        for rel in list(part.rels.values()):
            if rel.is_external or not rel.reltype.endswith('/slide') or rel.target_part in retained:continue
            fallback=generated['task'] if rel.target_part in {slide.part for slide in task_patterns} else generated['project']
            if fallback:rel._target=fallback[0].part
            else:
                for el in list(part._element.iter()) if hasattr(part,'_element') else []:
                    if el.get(qn('r:id'))==rel.rId and el.tag in (qn('a:hlinkClick'),qn('a:hlinkMouseOver')):el.getparent().remove(el)
                part.drop_rel(rel.rId)
    for i,sl in enumerate(prs.slides,1):
        for sh in sl.shapes:
            if sh._element.find('.//'+qn('a:fld')) is not None and getattr(sh,'has_text_frame',False):
                for t in sh._element.findall('.//'+qn('a:t')):t.text=str(i)
    prepare_navigation(prs)
    set_table_sizes(prs)
    # Summary tables that had to shrink to stay on their one slide keep that size.
    for shape,size in sizes:apply_table(shape.table,size)
    set_table_directions(prs)
    deck=Path(args.output);prs.save(deck)
    prune_package(deck)
    return deck

def prune_package(deck):
    """Rewrite the saved file keeping only the parts reachable from the package root."""
    with tempfile.TemporaryDirectory() as td:
        with zipfile.ZipFile(deck) as z:z.extractall(td)
        included=reachable_parts(td)
        with zipfile.ZipFile(deck,'w',zipfile.ZIP_DEFLATED) as z:
            for name in sorted(included):z.write(Path(td)/name,name)

# Gate findings about the builder's own layout. They are soft in the gate, but a build repairs them once.
LAYOUT_FINDINGS=('Tables overlap on slide ','Table frame reaches footer artwork on slide ')

def layout_slides(warnings):
    """{finding kind: sorted slide numbers} for the overlap and footer findings in `warnings`."""
    found={}
    for text in warnings:
        for prefix in LAYOUT_FINDINGS:
            if text.startswith(prefix):
                digits=''.join(ch for ch in text[len(prefix):].split(';')[0] if ch.isdigit())
                if digits:found.setdefault(prefix,set()).add(int(digits))
    return {k:sorted(v) for k,v in found.items()}

def repair_layout(deck,slide_numbers):
    """Give each table on the named slides the height of its rows and stack overlapping tables again with the approved gap.

    Row heights, text and typography are untouched; only frame heights and
    tops change, in the order the tables already have on the slide.
    """
    prs=Presentation(deck)
    for number in slide_numbers:
        slide=prs.slides[number-1]
        stack=builder.tables(slide)
        if stack:reflow.restack(slide,min(sh.top for sh in stack),content_bottom(prs,slide),TABLE_GAP,tighten=False)
    prs.save(deck)
    prune_package(deck)

def parse_args(argv=None):
    parser=argparse.ArgumentParser(description=__doc__)
    for name in ('tasks','suhail','transactions','output'):parser.add_argument('--'+name,required=True)
    parser.add_argument('--template',help='Explicitly approved alternative template')
    parser.add_argument('--model',choices=('standard','digital'),default='standard')
    parser.add_argument('--digital-pdf',help='Current digital-maturity PDF, required for the digital model')
    parser.add_argument('--digital-data',help='Visually reviewed JSON extracted from the same PDF')
    parser.add_argument('--measurements',help='Native row heights from a rendered candidate with these exact inputs')
    parser.add_argument('--measure',action='store_true',help='Delivery path: build, measure every row through LibreOffice, rebuild with the measured heights')
    parser.add_argument('--summary-overflow',choices=('shrink','appendix'),default='shrink',
                        help='A summary that does not fit its slide: shrink its table text (default), or keep 11 pt and move the overflow to an appendix slide')
    parser.add_argument('--summary-min-pt',type=float,default=SUMMARY_SIZES[-1],
                        help='Smallest summary table text when shrinking (default 8); past it the overflow moves to an appendix slide')
    return parser.parse_args(argv)

def measure_candidate(args,temporary,bundled):
    """Build the provisional candidate and return the path of its layout file, or None with a notice."""
    import measure_layout
    first=copy.copy(args);first.measurements=None;first.output=str(Path(temporary)/'provisional.pptx')
    build_candidate(first)
    try:
        if not abar_installed():
            notices.warn('خط «Abar Mid» غير مثبت على هذا الجهاز؛ القياس تم بخط بديل وقد تختلف ارتفاعات الصفوف في باوربوينت. ثبّت الخط وأعد البناء قبل التسليم.','layout')
        result=measure_layout.measure(first.output,args.tasks,args.suhail,args.transactions,args.template or bundled,
                                      args.digital_pdf,args.digital_data)
    except (RuntimeError,OSError,subprocess.SubprocessError,ValueError,KeyError) as exc:
        notices.warn(f'تعذر قياس الصفوف عبر LibreOffice ({reason(exc)}). التقرير مبني بارتفاعات تقديرية وليس جاهزا للتسليم قبل قياسه.','layout')
        return None
    layout=Path(temporary)/'provisional.layout.json'
    layout.write_text(json.dumps(result,ensure_ascii=False,indent=1),encoding='utf-8')
    return layout

def main(argv=None):
    args=parse_args(argv)
    if args.model=='digital' and not (args.digital_pdf and args.digital_data):
        raise ValueError('Digital model requires --digital-pdf and --digital-data')
    if args.model=='standard' and (args.digital_pdf or args.digital_data):
        raise ValueError('Digital-maturity inputs require --model digital')
    output=Path(args.output).resolve()
    if args.model=='digital':
        from digital_maturity import MASTER as DIGITAL_MASTER
        bundled=DIGITAL_MASTER
    else:bundled=MASTER
    protected=[Path(args.template) if args.template else bundled,Path(args.tasks),Path(args.suhail),Path(args.transactions)]
    if args.model=='digital':protected.extend((Path(args.digital_pdf),Path(args.digital_data)))
    if args.measurements:protected.append(Path(args.measurements))
    if any(output==p.resolve() or (output.exists() and p.exists() and output.samefile(p)) for p in protected):
        raise ValueError('Output must not overwrite the template or an input workbook')
    if args.measure and args.measurements:
        raise ValueError('Use --measure or --measurements, not both')
    from validate_report import validate, arabic_finding
    output.parent.mkdir(parents=True,exist_ok=True)
    notices.reset()
    notices.clear_stale(output)
    with tempfile.TemporaryDirectory(prefix='.weekly-build-',dir=output.parent) as temporary:
        candidate=copy.copy(args);candidate.output=str(Path(temporary)/'report.pptx')
        layout=measure_candidate(args,temporary,bundled) if args.measure else None
        if layout is not None:candidate.measurements=str(layout)
        build_candidate(candidate)
        result=validate(candidate.output,args.tasks,args.suhail,args.transactions,args.template or bundled,
                        args.model,args.digital_pdf,args.digital_data,strict=False)
        if layout_slides(result['warnings']):
            # Repair once, then validate again; only what remains is reported, never a stop.
            repair_layout(candidate.output,sorted({n for v in layout_slides(result['warnings']).values() for n in v}))
            result=validate(candidate.output,args.tasks,args.suhail,args.transactions,args.template or bundled,
                            args.model,args.digital_pdf,args.digital_data,strict=False)
        left=layout_slides(result['warnings'])
        for number in left.get(LAYOUT_FINDINGS[0],[]):
            notices.warn(f'الشريحة {number}: جدولان متداخلان وبقيا كذلك بعد محاولة إصلاح واحدة. افتح هذه الشريحة وراجعها قبل التسليم، وأخبر المستخدم بها.','layout')
        for number in left.get(LAYOUT_FINDINGS[1],[]):
            notices.warn(f'الشريحة {number}: جدول يصل إلى منطقة التذييل وبقي كذلك بعد محاولة إصلاح واحدة. افتح هذه الشريحة وراجعها قبل التسليم، وأخبر المستخدم بها.','layout')
        for finding in result['warnings']:
            if not finding.startswith(LAYOUT_FINDINGS):notices.warn('فحص التقرير: '+arabic_finding(finding),'check')
        status_styles.publish(candidate.output,output)
        if layout is not None:status_styles.publish(layout,output.with_suffix('.layout.json'))
    measured=bool(candidate.measurements)
    if not measured:
        notices.warn('هذا البناء بارتفاعات تقديرية (بدون قياس). لا يُسلَّم قبل إعادة البناء مع --measure.','layout')
    notices.save(output)
    print(str(output)+': '+('measured build' if measured else 'PROVISIONAL build, not for delivery')
          +'; programmatic checks passed; render and visual review required')
    return 0

if __name__=='__main__':raise SystemExit(main())
