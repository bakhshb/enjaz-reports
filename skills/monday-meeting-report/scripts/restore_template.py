#!/usr/bin/env python3
"""Verify the complete bundled master and copy it to the working template path."""
import hashlib
import shutil
from pathlib import Path

EXPECTED = '0a91475077b86cb2c0d8e2089eb8016ea665cf8c0eb86dc655bf45f40b7eb69e'
root = Path(__file__).resolve().parents[1]
source = root / 'assets' / 'MM_W38.pptx'
if hashlib.sha256(source.read_bytes()).hexdigest() != EXPECTED:
    raise SystemExit('Bundled template checksum mismatch')
out = root / 'template' / 'MM_W38.pptx'
out.parent.mkdir(parents=True, exist_ok=True)
shutil.copyfile(source, out)
print(out)
