# /// script
# requires-python = ">=3.12"
# dependencies = ["pypdf>=5", "python-pptx>=1.0.2"]
# ///
"""Read-only PDF font and page-count gate; visual layout review is separate."""
import argparse
import hashlib
import json
import re
from pathlib import Path
from pypdf import PdfReader
from pptx import Presentation


def check(root):
    results = []
    for deck in sorted(root.rglob('*.pptx')):
        pdf = deck.with_suffix('.pdf')
        errors, fonts, arabic_runs = [], set(), 0
        if not pdf.exists():
            results.append({'file': str(deck), 'errors': ['Missing PDF']})
            continue
        reader = PdfReader(pdf)
        if len(reader.pages) != len(Presentation(deck).slides):
            errors.append('PDF page count differs from PowerPoint')
        for number, page in enumerate(reader.pages, 1):
            def visit(text, cm, tm, font, size):
                nonlocal arabic_runs
                if not text.strip():
                    return
                name = str(font.get('/BaseFont', '') if font else '')
                fonts.add(name)
                if re.search(r'[\u0600-\u06ff]', text):
                    arabic_runs += 1
                    if 'Abar' not in name:
                        errors.append(f'Page {number}: Arabic text uses {name}: {text[:40]!r}')
            page.extract_text(visitor_text=visit)
        if not arabic_runs:
            errors.append('No readable Arabic text; font acceptance cannot be established')
        results.append({
            'file': str(deck), 'pdf': str(pdf), 'pages': len(reader.pages),
            'arabic_runs': arabic_runs, 'fonts': sorted(fonts), 'errors': errors,
            'pptx_sha256': hashlib.sha256(deck.read_bytes()).hexdigest(),
            'pdf_sha256': hashlib.sha256(pdf.read_bytes()).hexdigest(),
        })
    if not results:
        raise ValueError('No presentations found')
    return results


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('root', type=Path)
    args = parser.parse_args()
    results = check(args.root)
    (args.root / 'pdf-font-acceptance.json').write_text(
        json.dumps(results, ensure_ascii=False, indent=2), encoding='utf-8')
    for row in results:
        print(('FAIL' if row['errors'] else 'PASS') + ' ' + row['file'])
        print('\n'.join(row['errors']))
    raise SystemExit(any(row['errors'] for row in results))
