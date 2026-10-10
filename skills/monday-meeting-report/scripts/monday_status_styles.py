"""Monday status examples from its approved master."""
from pathlib import Path
import sys, os
sys.path.insert(0,str(Path(__file__).resolve().parents[2]))
from report_common.status_styles import ALIASES, TASK_STATUSES, PROJECT_STATUSES, SUPPORTED, canonical, signature, apply, publish
from report_common.status_styles import examples as template_examples
from report_common.table_typography import set_cell_size
from pptx.enum.text import PP_ALIGN
MASTER=Path(__file__).resolve().parents[1]/'assets/monday-meeting-master.pptx'
def examples(template=MASTER):
    styles=template_examples(template)
    for cell in styles.values():
        set_cell_size(cell)
        for p in cell.text_frame.paragraphs:p.alignment=PP_ALIGN.CENTER
    return styles
