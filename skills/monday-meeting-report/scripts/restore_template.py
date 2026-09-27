#!/usr/bin/env python3
"""Verify the complete bundled master and copy it to the working template path."""
import hashlib
import shutil
from pathlib import Path

EXPECTED = '6f66a2e16d370e24e26b208b6b75f8b036f573430cc735d40e2f102fddf59db3'
root = Path(__file__).resolve().parents[1]
source = root / 'assets' / 'monday-meeting-master.pptx'
if hashlib.sha256(source.read_bytes()).hexdigest() != EXPECTED:
    raise SystemExit('Bundled template checksum mismatch')
out = root / 'template' / 'monday-meeting-master.pptx'
out.parent.mkdir(parents=True, exist_ok=True)
shutil.copyfile(source, out)
print(out)
