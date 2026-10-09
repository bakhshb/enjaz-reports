"""User-approved table font size, without changing other template styling."""
from pptx.util import Pt
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn

TABLE_PT = 11

def set_cell_size(cell):
    for node in cell._tc.iter():
        if node.tag in {qn('a:rPr'), qn('a:defRPr'), qn('a:endParaRPr')}:
            node.set('sz', str(TABLE_PT * 100))
    for paragraph in cell.text_frame.paragraphs:
        paragraph.font.size = Pt(TABLE_PT)
        for run in paragraph.runs:
            run.font.size = Pt(TABLE_PT)

def column_alignment(header):
    header = ' '.join(header.replace('أ','ا').replace('إ','ا').split())
    if header in {'القطاع','تاريخ الانجاز المخطط','تاريخ البداية','تاريخ النهاية','الحالة','حالة المشروع'}:
        return PP_ALIGN.CENTER
    if header in {'المهمة','المشروع','اسم المشروع','ملاحظات','الملاحظات','ما تم حتى تاريخه','ما تم حتى آخر تحديث','التحديث','التحديثات','طلب الدعم','طلبات الدعم','التحدي','التحديات'}:
        return PP_ALIGN.RIGHT
    return None

def apply_table(table):
    headers = [c.text for c in table.rows[1].cells] if len(table.rows)>1 else []
    for i,row in enumerate(table.rows):
        for j,cell in enumerate(row.cells):
            set_cell_size(cell)
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

def set_table_sizes(prs):
    """Apply the approved size and per-column alignment, preserving merged titles."""
    for table in all_tables(prs): apply_table(table)

def check_table_typography(prs, errors):
    for table in all_tables(prs):
        headers = [c.text for c in table.rows[1].cells] if len(table.rows)>1 else []
        for i,row in enumerate(table.rows):
            for j,cell in enumerate(row.cells):
                expected = column_alignment(headers[j]) if i>=1 and j<len(headers) else None
                for p in cell.text_frame.paragraphs:
                    if expected is not None and p.alignment != expected:
                        errors.append('Table column alignment differs: '+headers[j])
                    for run in p.runs:
                        if run.text and (run.font.size is None or run.font.size.pt!=TABLE_PT):
                            errors.append('Table font size must be 11 pt')
