"""Monday status examples from its approved master."""
from pathlib import Path
import sys, os
sys.path.insert(0,str(Path(__file__).resolve().parents[2]))
from report_common.status_styles import TASK_STATUSES, PROJECT_STATUSES, canonical, signature, apply, publish, style_key
from report_common.status_styles import examples as template_examples
from report_common.table_typography import set_cell_size
from pptx.enum.text import PP_ALIGN
# Re-exported for the Monday builder, gate and tests (tests patch status_styles.os).
__all__=['NEUTRAL_STATUS','style_status','TASK_STATUSES','PROJECT_STATUSES','canonical','signature','apply','publish','examples','MASTER','os']
# A status the master has no style for keeps its own text and takes the neutral approved fill.
NEUTRAL_STATUS={'task':'معلق','suhail':'لم تبدأ'}
def style_status(status,kind,approved):
    """The approved style to use for ``status``: itself when the master styles it, else the neutral one."""
    allowed=TASK_STATUSES if kind=='task' else PROJECT_STATUSES
    return style_key(status,allowed,approved,NEUTRAL_STATUS[kind])
MASTER=Path(__file__).resolve().parents[1]/'assets/monday-meeting-master.pptx'
def examples(template=MASTER):
    styles=template_examples(template)
    for cell in styles.values():
        set_cell_size(cell)
        for p in cell.text_frame.paragraphs:p.alignment=PP_ALIGN.CENTER
    return styles
