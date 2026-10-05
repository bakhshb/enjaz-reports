"""Compatibility entrypoint for the shared Suhail update formatter."""
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parents[2]))
from report_common.update_format import DATE, MARKER, update_parts, display_text, format_update
