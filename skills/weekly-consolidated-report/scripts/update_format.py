"""Compatibility entrypoint for the shared Suhail update formatter."""
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parents[2]))
from functools import partial
from report_common import update_format as shared
DATE, MARKER = shared.DATE, shared.MARKER
# Independent source paragraphs following the date are update items, as in Monday.
update_parts = partial(shared.update_parts, paragraph_items=True)
display_text = partial(shared.display_text, paragraph_items=True)
format_update = partial(shared.format_update, paragraph_items=True)
