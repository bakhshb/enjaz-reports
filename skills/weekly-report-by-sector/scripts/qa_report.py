# -*- coding: utf-8 -*-
"""Verify the generated per-sector decks.

    python qa_report.py out --master MASTER.pptx
    python qa_report.py out --master MASTER.pptx --render qa

Content checks (the mandatory ones from the brief):
  * every table row in a deck belongs to that sector
  * total rows == the sector's rows in the master
  * no empty table anywhere
  * no other sector's name appears in any deck
"""
import argparse, glob, hashlib, json, os, posixpath, shutil, subprocess, sys, tempfile, zipfile
from collections import defaultdict
from pathlib import Path
from xml.etree import ElementTree as ET
from pptx import Presentation
from pptx.oxml.ns import qn
import split_report as splitter
from report_common.table_typography import check_table_typography

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

def record_text(value):
    return str(value or '').replace('\r\n','\n').replace('\v','\n').strip()


TASK_FIELDS=('المهمة','القطاع','ملاحظات')
PROJECT_FIELDS=('اسم المشروع','تاريخ البداية','تاريخ النهاية','الحالة','ما تم حتى تاريخه')
FIELDS={**{t:TASK_FIELDS for t in splitter.TASK_TITLE.values()},
        **{t:PROJECT_FIELDS for t in splitter.PROJ_TITLE.values()},
        'المهام على المخطط':('المهمة','القطاع','تاريخ الإنجاز المخطط','الحالة','ملاحظات'),
        'طلبات الدعم':('المهمة','القطاع','طلب الدعم'),
        'أبرز التحديثات':('المشروع','القطاع','التحديث'),
        'التحديات':('المشروع','القطاع','التحدي'),
        'المعاملات المتأخرة':splitter.TRANS_FIELDS}


def expected_records(master,transactions=None,aliases=None,only=None):
    src=Presentation(master)
    tasks,projects,trans,updates,support,challenges=splitter.read_data(src,splitter.locate(src))
    splitter.attach_transaction_sectors(trans,transactions)
    alias_path=aliases or Path(__file__).resolve().parents[1]/'references'/'sector-aliases.json'
    mapping={splitter.norm(k):v for k,v in json.loads(Path(alias_path).read_text(encoding='utf-8')).items()}
    canonical=lambda s:mapping.get(splitter.norm(s),splitter.norm(s))
    result=defaultdict(lambda:defaultdict(list))
    for item in tasks:
        sector=canonical(item['sector']); title=splitter.TASK_TITLE[item['status']]
        values = (item['text'], sector, item['planned_date'], item['status'], item['note']) if item['status']=='على المخطط' else (item['text'], sector, item['note'])
        result[sector][title].append(tuple(map(record_text,values)))
    for item in projects:
        sector=canonical(item['sector']); title=splitter.PROJ_TITLE[item['status']]
        result[sector][title].append(tuple(map(record_text,(item['name'],item['start'],item['end'],item['status'],item['note']))))
    for items,title,name in [(updates,'أبرز التحديثات','project'),(support,'طلبات الدعم','name'),(challenges,'التحديات','name')]:
        for item in items:
            sector=canonical(item['sector'])
            result[sector][title].append(tuple(map(record_text,(item[name],sector,item['text']))))
    for item in trans:
        sector=canonical(item['sector']); headers=[splitter.field_key(h) for h in item['head']]
        result[sector]['المعاملات المتأخرة'].append(tuple(record_text(item['row'][headers.index(splitter.field_key(h))]) for h in splitter.TRANS_FIELDS))
    return {sector:dict(tables) for sector,tables in result.items() if not only or only in sector}


def actual_records(prs):
    result=defaultdict(list)
    for slide in prs.slides:
        for shape in slide.shapes:
            if not shape.has_table:continue
            rr=[[cell_text(c) for c in row.cells] for row in shape.table.rows]
            title=splitter.table_title(rr)
            if title not in FIELDS:raise ValueError(f'Unexpected output table: {title}')
            if len(rr)<3:raise ValueError(f'Empty output table: {title}')
            headers=[splitter.field_key(h) for h in rr[1]]
            for row in rr[2:]:
                values=[]
                for field in FIELDS[title]:
                    key=splitter.field_key(field)
                    if key not in headers and field=='ملاحظات' and 'التحديث' in headers:key='التحديث'
                    if key not in headers:raise ValueError(f'{title}: missing column {field}')
                    value=record_text(row[headers.index(key)])
                    if field=='الحالة':value=splitter.norm_status(value)
                    values.append(value)
                result[title].append(tuple(values))
    return dict(result)


def check(outdir,master=None,transactions=None,aliases=None,only=None):
    if not master:
        print('FAILED: --master is required for source-record verification')
        return False
    expected_by_sector=expected_records(master,transactions,aliases,only)
    man = json.loads((Path(outdir)/'manifest.json').read_text(encoding='utf-8'))
    names = {m['sector'] for m in man}
    ok = names==set(expected_by_sector) and len(names)==len(man)
    listed={Path(m.get('out') or Path(outdir)/('التقرير الأسبوعي - %s.pptx'%m['sector'])).resolve() for m in man}
    if listed!={p.resolve() for p in Path(outdir).glob('*.pptx')}:
        print('Output files differ from manifest');ok=False
    if not ok:print('Sector set mismatch:',sorted(names),'expected',sorted(expected_by_sector))
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
        source_tables=expected_by_sector.get(it['sector'],{})
        expected = sum(len(records) for records in source_tables.values())
        record_errors=[]
        check_table_typography(prs,record_errors)
        record_errors.extend(identity_errors(prs,it['sector'],source_tables,master))
        try:
            actual=actual_records(prs)
            for title in set(actual)|set(source_tables):
                if actual.get(title,[])!=source_tables.get(title,[]):
                    record_errors.append(f'{title}: records/status/order differ from master')
        except ValueError as exc:record_errors.append(str(exc))
        sector_blob = ' | '.join(sector_cells)
        leaks = sorted(n for n in names - {it['sector']} if n in sector_blob)
        good = rows == expected and not leaks and not empty and not over and not geometry and not record_errors
        ok &= good
        print(('OK   ' if good else 'CHECK') +
              ' | %-42s rows=%-3d expected=%-3d tables=%d' % (it['sector'], rows, expected, len(titles)))
        if leaks: print('        leaked sectors:', leaks)
        if empty: print('        empty tables:', empty)
        if geometry: print('        inconsistent table heights:', geometry)
        if over: print('        table runs into the footer on:', sorted(set(over)))
        for error in record_errors:print('        '+error)
        print('        SHA256',hashlib.sha256(Path(path).read_bytes()).hexdigest())
        print('        ' + '  //  '.join(titles))
    print('\nALL OK' if ok else '\nISSUES FOUND — fix before delivering')
    return ok


def identity_errors(prs,sector,records,master):
    """Check cover and approved summary identifiers independently of the manifest."""
    def walk(shapes):
        for shape in shapes:
            yield shape
            if hasattr(shape,'shapes'):yield from walk(shape.shapes)
    errors=[]
    if not any(getattr(sh,'has_text_frame',False) and sh.text.strip()==sector
               for sh in walk(prs.slides[0].shapes)):
        errors.append('Cover sector differs from source')
    source=Presentation(master)
    specs=[('إجمالي المهام',[11,14,16,7,6],['المهام المكتملة','المهام على المخطط','المهام المتأخرة','المهام المعلقة']),
           ('إجمالي مشاريع سهيل',[10,12,14,8,7],['المشاريع المكتملة','المشاريع على المخطط','المشاريع المتأخرة','المشاريع لم تبدأ'])]
    for label,ids,titles in specs:
        # The fixed identifiers apply only to the approved template's KPI pattern.
        seeds=[sl for sl in source.slides if any(getattr(sh,'has_text_frame',False) and sh.text.strip()==label for sh in sl.shapes)]
        if not seeds or not set(ids).issubset({sh.shape_id for sh in seeds[0].shapes}):continue
        candidates=[sl for sl in prs.slides if any(getattr(sh,'has_text_frame',False) and sh.text.strip()==label for sh in sl.shapes)]
        counts=[len(records.get(title,[])) for title in titles]
        if sum(counts) and not candidates:errors.append(label+': missing KPI slide')
        for slide in candidates:
            shapes={sh.shape_id:sh for sh in slide.shapes}
            for sid,expected in zip(ids,[sum(counts),*counts]):
                if sid not in shapes or getattr(shapes[sid],'text','').strip()!=str(expected):
                    errors.append(f'{label}: KPI {sid} differs from source')
    trans=len(records.get('المعاملات المتأخرة',[]))
    if trans:
        candidates=[sl for sl in prs.slides if any(getattr(sh,'has_text_frame',False) and sh.text.strip()=='إجمالي المعاملات المتأخرة' for sh in sl.shapes)]
        if not candidates:errors.append('Missing delayed-transactions KPI')
        for slide in candidates:
            nums=[sh for sh in slide.shapes if getattr(sh,'has_text_frame',False) and sh.shape_id==24]
            if len(nums)!=1 or nums[0].text.strip()!=str(trans):errors.append('Transaction KPI differs from source')
    return errors

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
    files=sorted(glob.glob(os.path.join(outdir, '*.pptx')))
    if not files:
        # check() still proves that the source and manifest both have zero sectors.
        print('No output decks; source/manifest must confirm an empty result')
        return True
    for f in files:
        try:
            with zipfile.ZipFile(f) as package:
                bad = package.testzip()
                if bad: raise ValueError('damaged package member: ' + bad)
                names=set(package.namelist())
                if len(names)!=len(package.namelist()):raise ValueError('duplicate package members')
                for name in names:
                    if name.endswith(('.xml','.rels')):ET.fromstring(package.read(name))
                    if name.endswith('.xml'):
                        relpath=posixpath.join(posixpath.dirname(name),'_rels',posixpath.basename(name)+'.rels')
                        ids={r.get('Id') for r in ET.fromstring(package.read(relpath))} if relpath in names else set()
                        for el in ET.fromstring(package.read(name)).iter():
                            for key,value in el.attrib.items():
                                if key.startswith('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}') and value and value not in ids:
                                    raise ValueError(f'missing relationship ID: {name} -> {value}')
                    if not name.endswith('.rels'):continue
                    source=posixpath.join(posixpath.dirname(posixpath.dirname(name)),posixpath.basename(name)[:-5])
                    for rel in ET.fromstring(package.read(name)):
                        if rel.get('TargetMode')=='External':
                            raise ValueError('external relationship: '+name)
                        target=posixpath.normpath(posixpath.join(posixpath.dirname(source),rel.get('Target','').split('#',1)[0])).lstrip('/')
                        if target not in names:raise ValueError('missing relationship target: '+target)
                presentation=ET.fromstring(package.read('ppt/presentation.xml'))
                ids={el.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
                     for el in presentation.findall('.//{http://schemas.openxmlformats.org/presentationml/2006/main}sldId')}
                linked={posixpath.normpath(posixpath.join('ppt',rel.get('Target',''))).lstrip('/')
                        for rel in ET.fromstring(package.read('ppt/_rels/presentation.xml.rels'))
                        if rel.get('Id') in ids}
            packaged={n for n in names if n.startswith('ppt/slides/slide') and n.endswith('.xml')}
            if linked!=packaged:raise ValueError('unlisted slide parts remain in the package')
            print('VALID |', os.path.basename(f))
        except (OSError, ValueError, KeyError, ET.ParseError, zipfile.BadZipFile) as exc:
            ok = False
            print('INVALID |', os.path.basename(f), '|', exc)
    return ok

if __name__ == '__main__':
    ap = argparse.ArgumentParser()
    ap.add_argument('outdir')
    ap.add_argument('--render', metavar='QADIR')
    ap.add_argument('--master')
    ap.add_argument('--transactions')
    ap.add_argument('--aliases')
    ap.add_argument('--only')
    a = ap.parse_args()
    valid = validate(a.outdir, a.master)
    good = check(a.outdir,a.master,a.transactions,a.aliases,a.only) if valid else False
    if a.render and good: render(a.outdir, a.render)
    sys.exit(0 if good else 1)
