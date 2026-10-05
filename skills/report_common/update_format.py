"""Source-preserving Suhail update text, shared by report generators."""
import re
from copy import deepcopy
from pptx.oxml.xmlchemy import OxmlElement
from pptx.oxml.ns import qn

DATE = re.compile(r'^(\s*تاريخ التحديث\s*[:：]?\s*(?:\d{1,4}[/.-]\d{1,2}[/.-]\d{1,4}|\d{1,2}\s+[^\W\d_]+(?:\s+[^\W\d_]+)?\s+\d{4})(?:\s*[هـم])?)(?=\s|$|[،؛:])')
MARKER = re.compile(r'^\s*(?:[•●▪◦*-]|\d+[.)\-]|[٠-٩]+[.)\-])\s+')

def update_parts(text, paragraph_items=False):
    lines = text.replace('\r\n', '\n').replace('\v', '\n').split('\n')
    # Weekly uses explicit markers; Monday uses source update paragraphs.
    if paragraph_items:
        lines = list(map(str.strip, lines))
        while lines and not lines[0]: lines.pop(0)
        while lines and not lines[-1]: lines.pop()
        if not lines: lines = ['']
    marked = [i for i, line in enumerate(lines) if MARKER.match(line)]
    if paragraph_items and lines and DATE.match(lines[0]):
        marked = [i for i, line in enumerate(lines) if i > 0 and line.strip()]
    bullets = set(marked) if len(marked) > 1 else set()
    result = []
    for i, line in enumerate(lines):
        bullet = i in bullets
        value = MARKER.sub('', line, count=1) if bullet or paragraph_items else line
        date = DATE.match(value) if i == 0 else None
        result.append((value, date.end() if date else 0, bullet))
    return result

def display_text(text, paragraph_items=False):
    return '\n'.join(value for value, _, _ in update_parts(text, paragraph_items))

def format_update(cell, bullet_style=None, paragraph_items=False):
    parts = update_parts(cell.text, paragraph_items)
    frame = cell.text_frame
    base = next((deepcopy(r._r.get_or_add_rPr()) for p in frame.paragraphs for r in p.runs), None)
    paragraph_style = deepcopy(frame.paragraphs[0]._p.get_or_add_pPr())
    frame.clear()
    for i, (value, bold_end, bullet) in enumerate(parts):
        p = frame.paragraphs[0] if i == 0 else frame.add_paragraph()
        old = p._p.find(qn('a:pPr'))
        if old is not None: p._p.remove(old)
        props = deepcopy(paragraph_style)
        for node in list(props):
            if node.tag in {qn('a:buNone'), qn('a:buChar'), qn('a:buAutoNum'), qn('a:buBlip')}:
                props.remove(node)
        if bullet:
            if bullet_style is None:
                raise ValueError('Approved bullet style is required for Suhail list items')
            for attr in ('marL', 'marR', 'indent'):
                if bullet_style.get(attr) is not None: props.set(attr, bullet_style.get(attr))
        marker = OxmlElement('a:buChar' if bullet else 'a:buNone')
        if bullet: marker.set('char', '•')
        # Bullets precede tabLst/defRPr/extLst in DrawingML schema order.
        at = next((n for n, el in enumerate(props) if el.tag in {qn('a:tabLst'), qn('a:defRPr'), qn('a:extLst')}), len(props))
        props.insert(at, marker)
        p._p.insert(0, props)
        for text, bold in ((value[:bold_end], True), (value[bold_end:], False)):
            if not text: continue
            run = p.add_run(); run.text = text
            if base is not None: run._r.insert(0, deepcopy(base))
            run.font.bold = bold
