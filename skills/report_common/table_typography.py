"""User-approved PowerPoint table direction, font size and column alignment.

The same pass keeps a table sound when the user edits it by hand in
PowerPoint: every paragraph ends with complete end-of-paragraph properties
(size and Abar), so text typed into an emptied cell or an inserted row keeps
the table's typography instead of the theme default, and every row carries
its own row ID (rows copied from one pattern row shared one).
"""
from copy import deepcopy
from pptx.util import Pt
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn

TABLE_PT = 11
DEFAULT_INSET = 45720  # DrawingML default top/bottom cell margin (0.05 in)
# Every approved master writes its table text in Abar Mid. A run without its
# own typeface falls back to the theme font (Aptos/Arial), so every run gets it.
TABLE_FONT = 'Abar Mid'
FONT_SLOTS = ('latin', 'ea', 'cs')
_AFTER_FONTS = ('sym', 'hlinkClick', 'hlinkMouseOver', 'rtl', 'extLst')
AGENDA_HEADERS = ['م','جدول الأعمال','المسؤول','المدة الزمنية (بالدقيقة)']

def reverse_table_columns(table):
    """Mirror the grid and merged groups together, retaining field formatting."""
    grid = table._tbl.tblGrid
    for col in reversed(list(grid)):
        grid.append(col)
    for row in table.rows:
        groups = []
        cells = list(row._tr.tc_lst)
        i = 0
        while i < len(cells):
            span = int(cells[i].get('gridSpan', '1'))
            groups.append(cells[i:i + span])
            i += span
        for group in reversed(groups):
            for cell in group:
                row._tr.insert_element_before(cell, 'a:extLst')
                props = cell.find(qn('a:tcPr'))
                if props is not None:
                    left, right = props.find(qn('a:lnL')), props.find(qn('a:lnR'))
                    for node in (left, right):
                        if node is not None: props.remove(node)
                    if left is not None:
                        left.tag = qn('a:lnR'); props.insert(0,left)
                    if right is not None:
                        right.tag = qn('a:lnL'); props.insert(0,right)

def rtl_table(table):
    """Convert LTR storage to RTL without changing column content or widths."""
    if table._tbl.tblPr.get('rtl') not in ('1', 'true'):
        reverse_table_columns(table)
    table._tbl.tblPr.set('rtl', '1')

def set_font(props, name=TABLE_FONT):
    """Give a run-properties element `name` for Latin, East Asian and Arabic text, in schema order."""
    for slot in FONT_SLOTS:
        for old in props.findall(qn('a:' + slot)):
            props.remove(old)
    anchor = next((child for child in props if child.tag in {qn('a:' + t) for t in _AFTER_FONTS}), None)
    for slot in FONT_SLOTS:
        node = props.makeelement(qn('a:' + slot), {'typeface': name})
        if anchor is None:
            props.append(node)
        else:
            anchor.addprevious(node)


def end_paragraph_properties(paragraph):
    """Give `paragraph` (an ``a:p``) end-of-paragraph properties copied from its last run.

    PowerPoint formats text typed into an empty cell, and the cells of a row
    the user inserts, from these properties; without them it falls back to
    the theme font at 18 pt.
    """
    end = paragraph.find(qn('a:endParaRPr'))
    if end is None:
        runs = [run.find(qn('a:rPr')) for run in paragraph.findall(qn('a:r'))]
        runs = [props for props in runs if props is not None]
        end = deepcopy(runs[-1]) if runs else paragraph.makeelement(qn('a:endParaRPr'), {'lang': 'ar-SA'})
        end.tag = qn('a:endParaRPr')
        for link in end.findall(qn('a:hlinkClick')) + end.findall(qn('a:hlinkMouseOver')):
            end.remove(link)
        paragraph.append(end)
    return end


def set_cell_size(cell, font_pt=TABLE_PT):
    for run in cell._tc.iter(qn('a:r')):
        if run.find(qn('a:rPr')) is None:
            run.insert(0, run.makeelement(qn('a:rPr'), {}))
    for paragraph in cell._tc.iter(qn('a:p')):
        end_paragraph_properties(paragraph)
    # A list first: set_font rebuilds each node's children, and changing the tree
    # while walking it made lxml skip the nodes that follow (the end-of-paragraph
    # properties kept the master's 10 pt, which is what typed text then used).
    for node in list(cell._tc.iter(qn('a:rPr'), qn('a:defRPr'), qn('a:endParaRPr'))):
        node.set('sz', str(round(font_pt * 100)))
        set_font(node)
    for paragraph in cell.text_frame.paragraphs:
        paragraph.font.size = Pt(font_pt)
        for run in paragraph.runs:
            run.font.size = Pt(font_pt)

def column_alignment(header):
    header = ' '.join(header.replace('أ','ا').replace('إ','ا').split())
    if header in {'القطاع','تاريخ الانجاز المخطط','تاريخ البداية','تاريخ النهاية','الحالة','حالة المشروع'}:
        return PP_ALIGN.CENTER
    if header in {'المهمة','المشروع','اسم المشروع','ملاحظات','الملاحظات','ما تم حتى تاريخه','ما تم حتى آخر تحديث','التحديث','التحديثات','طلب الدعم','طلبات الدعم','التحدي','التحديات'}:
        return PP_ALIGN.RIGHT
    return None

def apply_table(table, font_pt=TABLE_PT):
    headers = [c.text for c in table.rows[1].cells] if len(table.rows)>1 else []
    for i,row in enumerate(table.rows):
        for j,cell in enumerate(row.cells):
            set_cell_size(cell,font_pt)
            alignment = column_alignment(headers[j]) if i>=1 and j<len(headers) else None
            if alignment is not None:
                for paragraph in cell.text_frame.paragraphs:
                    paragraph.alignment = alignment
                    paragraph._p.get_or_add_pPr().set('rtl','1')

def all_tables(prs):
    def walk(shapes):
        for shape in shapes:
            if shape.has_table: yield shape.table
            if hasattr(shape,'shapes'): yield from walk(shape.shapes)
    for slide in prs.slides: yield from walk(slide.shapes)

def table_font_size(table, agenda_pt=TABLE_PT):
    return agenda_pt if [c.text.strip() for c in table.rows[0].cells]==AGENDA_HEADERS else TABLE_PT


def set_table_sizes(prs, agenda_pt=TABLE_PT):
    """Apply the approved size and per-column alignment, preserving merged titles."""
    for table in all_tables(prs): apply_table(table,table_font_size(table,agenda_pt))

ROW_ID = '{http://schemas.microsoft.com/office/drawing/2014/main}rowId'


def unique_row_ids(table):
    """Give every row of `table` its own Office row ID.

    Rows are built by copying one pattern row, ID included. PowerPoint tracks
    rows by this ID when the user inserts, deletes or moves one, so the
    copies get fresh values; the first row that holds a value keeps it.
    """
    nodes = [node for row in table._tbl.tr_lst for node in row.iter(ROW_ID)]
    used = {node.get('val') for node in nodes}
    seen = set()
    fresh = 10000
    for node in nodes:
        value = node.get('val')
        if value in seen or not value:
            while str(fresh) in used:
                fresh += 1
            value = str(fresh)
            used.add(value)
            node.set('val', value)
        seen.add(value)


def set_table_directions(prs):
    """Final table pass of every builder: RTL direction and unique row IDs."""
    for table in all_tables(prs):
        rtl_table(table)
        unique_row_ids(table)

def expected_size(table, reduced, default=TABLE_PT):
    """The size `table` must use: `default`, or the one reduced size it uses throughout when `reduced` grants it."""
    allowed = reduced.get(table._tbl)
    if not allowed:
        return default
    used = {run.font.size.pt for row in table.rows for cell in row.cells
            for p in cell.text_frame.paragraphs for run in p.runs
            if run.text.strip() and run.font.size is not None}
    return used.pop() if len(used) == 1 and used <= set(allowed) else default


def check_table_typography(prs, errors, agenda_pt=TABLE_PT, reduced=None):
    """Check direction, alignment and size of every table.

    ``reduced`` maps a table element (``table._tbl``) to the sizes it may use
    instead of the approved size: a summary that had to shrink to stay on one
    slide. Such a table must still use one size throughout.
    """
    reduced = reduced or {}
    for table in all_tables(prs):
        if table._tbl.tblPr.get('rtl') not in ('1', 'true'):
            errors.append('Table direction must be RTL')
        font_pt=expected_size(table,reduced,table_font_size(table,agenda_pt))
        headers = [c.text for c in table.rows[1].cells] if len(table.rows)>1 else []
        for i,row in enumerate(table.rows):
            for j,cell in enumerate(row.cells):
                expected = column_alignment(headers[j]) if i>=1 and j<len(headers) else None
                for p in cell.text_frame.paragraphs:
                    if expected is not None and p.alignment != expected:
                        errors.append('Table column alignment differs: '+headers[j])
                    for run in p.runs:
                        if run.text and (run.font.size is None or run.font.size.pt!=font_pt):
                            errors.append(f'Table font size must be {font_pt} pt')
                        props = run._r.find(qn('a:rPr'))
                        if run.text.strip() and (props is None or any(
                                (props.find(qn('a:' + slot)) is None or props.find(qn('a:' + slot)).get('typeface') != TABLE_FONT)
                                for slot in ('latin', 'cs'))):
                            errors.append(f'Table text must use {TABLE_FONT}')
