#!/usr/bin/env python3
"""Verify the complete bundled master and copy it to the working template path."""
import hashlib
import shutil
from pathlib import Path

EXPECTED = '86a1ed90844059c88d9da10a33c81cd14902a6b99c02cfa37661f527e774ab4a'
root = Path(__file__).resolve().parents[1]
source = root / 'assets' / 'monday-meeting-master.pptx'
if hashlib.sha256(source.read_bytes()).hexdigest() != EXPECTED:
    raise SystemExit('Bundled template checksum mismatch')
out = root / 'template' / 'monday-meeting-master.pptx'
out.parent.mkdir(parents=True, exist_ok=True)
shutil.copyfile(source, out)
print(out)
