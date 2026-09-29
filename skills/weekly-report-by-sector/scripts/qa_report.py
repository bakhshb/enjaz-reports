# -*- coding: utf-8 -*-
"""Verify the generated per-sector decks.

    python qa_report.py out                # content checks only
    python qa_report.py out --render qa    # also renders one composite image per deck

Content checks (the mandatory ones from the brief):
  * every table row in a deck belongs to that sector
  * total rows == the sector's rows in the master
  * no empty table anywhere
  * no other sector's name appears in any deck
"""
import argparse, glob, json, os, shutil, subprocess, sys, tempfile, zipfile
from pptx import Presentation
from pptx.oxml.ns import qn

SOFFICE = shutil.which('soffice')

def cell_text(c):
    s = ''
    for p in c.text_frame.paragraphs:
        for ch in p._p:
            if ch.tag == qn('a:r'): s += ch.find(qn('a:t')).text or ''
            elif ch.tag == qn('a:br'): s += '\n'
        s += '\n'
    return s.strip()

FOOTER_Y = 12687759

def check(outdir):
    man = json.load(open(os.path.join(outdir, 'manifest.json'), encoding='utf-8'))
    names = {m['sector'] for m in man}
    ok = True
    for it in man:
        path = it.get('out') or os.path.join(outdir, 'التقرير الأسبوعي - %s.pptx' % it['sector'])
        prs = Presentation(path)
        rows, titles, empty, geometry, blob, sector_cells = 0, [], [], [], '', []
        for s in prs.slides:
            for sh in s.shapes:
                if sh.has_table:
                    rr = [[cell_text(c) for c in r.cells] for r in sh.table.rows]
                    titles.append(' '.join(x for x in rr[0] if x)[:40])
                    if abs(sum(r.height for r in sh.table.rows) - sh.height) > 2:
                        geometry.append(titles[-1])
                    body = len(rr) - 2
                    rows += body
                    if body == 0: empty.append(titles[-1])
                    if len(rr) > 1 and 'القطاع' in rr[1]:
                        sector_col = rr[1].index('القطاع')
                        sector_cells.extend(
                            r[sector_col] for r in rr[2:] if len(r) > sector_col
                        )
                    for r in rr:
                        blob += ' | '.join(r)
                elif sh.has_text_frame:
                    blob += sh.text_frame.text + ' | '
        over = []
        for si, sld in enumerate(prs.slides, 1):
            for sh in sld.shapes:
                if sh.has_table and sh.top is not None and sh.top + sh.height > FOOTER_Y:
                    over.append('slide %d' % si)
        expected = sum(it['n'])
        sector_blob = ' | '.join(sector_cells)
        leaks = sorted(n for n in names - {it['sector']} if n in sector_blob)
        good = rows == expected and not leaks and not empty and not over and not geometry
        ok &= good
        print(('OK   ' if good else 'CHECK') +
              ' | %-42s rows=%-3d expected=%-3d tables=%d' % (it['sector'], rows, expected, len(titles)))
        if leaks: print('        leaked sectors:', leaks)
        if empty: print('        empty tables:', empty)
        if geometry: print('        inconsistent table heights:', geometry)
        if over: print('        table runs into the footer on:', sorted(set(over)))
        print('        ' + '  //  '.join(titles))
    print('\nALL OK' if ok else '\nISSUES FOUND — fix before delivering')
    return ok

def render(outdir, qadir):
    from PIL import Image
    os.makedirs(qadir, exist_ok=True)
    for f in sorted(glob.glob(os.path.join(outdir, '*.pptx'))):
        base = os.path.splitext(os.path.basename(f))[0]
        shutil.copy(f, 't.pptx')
        if os.path.exists('t.pdf'): os.remove('t.pdf')
        if not SOFFICE:
            raise RuntimeError('soffice is required to render QA previews')
        pdf_dir = tempfile.mkdtemp(prefix='qa-pdf-')
        profile_dir = tempfile.mkdtemp(prefix='qa-soffice-')
        subprocess.run([SOFFICE, '-env:UserInstallation=file://' + profile_dir,
                        '--headless', '--convert-to', 'pdf',
                        '--outdir', pdf_dir, 't.pptx'],
                       capture_output=True, check=True)
        pdf_path = os.path.join(pdf_dir, 't.pdf')
        subprocess.run(['pdftoppm', '-jpeg', '-r', '60', pdf_path, 'pg'], check=True)
        pages = sorted(glob.glob('pg-*.jpg'), key=lambda x: int(x.rsplit('-', 1)[1].split('.')[0]))
        ims = [Image.open(p) for p in pages]
        w = sum(i.width for i in ims); h = max(i.height for i in ims)
        c = Image.new('RGB', (w, h), 'white'); x = 0
        for i in ims: c.paste(i, (x, 0)); x += i.width
        c.save(os.path.join(qadir, base + '.jpg'), quality=80)
        for p in pages: os.remove(p)
        os.remove('t.pptx')
        shutil.rmtree(pdf_dir, ignore_errors=True)
        shutil.rmtree(profile_dir, ignore_errors=True)
        print('rendered', base)

def validate(outdir, master=None):
    ok = True
    for f in sorted(glob.glob(os.path.join(outdir, '*.pptx'))):
        try:
            with zipfile.ZipFile(f) as package:
                bad = package.testzip()
                if bad: raise ValueError('damaged package member: ' + bad)
            Presentation(f)
            print('VALID |', os.path.basename(f))
        except (OSError, ValueError, KeyError, zipfile.BadZipFile) as exc:
            ok = False
            print('INVALID |', os.path.basename(f), '|', exc)
    return ok

if __name__ == '__main__':
    ap = argparse.ArgumentParser()
    ap.add_argument('outdir')
    ap.add_argument('--render', metavar='QADIR')
    ap.add_argument('--master')
    a = ap.parse_args()
    valid = validate(a.outdir, a.master)
    good = check(a.outdir) if valid else False
    if a.render: render(a.outdir, a.render)
    sys.exit(0 if good else 1)
