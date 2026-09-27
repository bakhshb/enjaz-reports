# -*- coding: utf-8 -*-
"""Split the ministry-wide weekly report (تقرير المهام) into one deck per جهة.

Usage:
    python split_report.py MASTER.pptx --outdir out [--aliases aliases.json] [--no-refine]

The master deck is the single source of both data and design. Nothing here
re-designs anything: every table, row, KPI card and title bar is cloned from a
shape that already exists in the master, and only the text, the fills that encode
status, and the geometry are changed.
"""
import argparse, json, math, os, re, shutil, subprocess, sys, zipfile, pathlib, tempfile
from copy import deepcopy

from pptx import Presentation
from pptx.util import Emu
from lxml import etree

A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
def q(t):  return '{%s}%s' % (A, t)
def qp(t): return '{%s}%s' % (P, t)
EMU = 914400
SOFFICE = shutil.which('soffice')

# ---- layout constants taken from the master deck -------------------------
CX, CW   = 474245, 6623050      # content column: x and width of the KPI card
TOP_Y    = 1780255              # first table starts where the comparison chart was
GAP      = 200000               # vertical gap between stacked tables
KPI_ZONE = 1700000              # everything below this y on a section slide is body
NAV_XS   = (5741129, 5741130, 3970297, 2199462)   # nav tabs that stay
NAV_DROP_X = 428628                                # the التفاصيل tab
NAV_SHIFT  = 1012345            # shift left so the 3 remaining tabs are page-centred
NAV_TOPS   = (-1, -2, -3, 137425, 194575)

# ---- palette (read off the master's KPI cards) ---------------------------
STATUS_FILL = {'مكتملة':'DCE6F2', 'على المخطط':'EBF1DE', 'متأخرة':'F2DCDB',
               'لم تبدأ':'F2F2F2', 'معلقة':'F2F2F2'}
ACCENT = {'متأخرة':'87452B', 'مكتملة':'174766', 'على المخطط':'0A8062',
          'معلقة':'A6A6A6', 'لم تبدأ':'A6A6A6', 'تحديثات':'D9A86A'}

TASK_ORDER = ['متأخرة', 'معلقة', 'مكتملة', 'على المخطط']
TASK_TITLE = {'متأخرة':'المهام المتأخرة', 'مكتملة':'المهام المكتملة',
              'على المخطط':'المهام على المخطط', 'معلقة':'المهام المعلقة'}
PROJ_ORDER = ['متأخرة', 'على المخطط', 'لم تبدأ', 'مكتملة']
PROJ_TITLE = {'متأخرة':'المشاريع المتأخرة', 'على المخطط':'المشاريع على المخطط',
              'لم تبدأ':'المشاريع لم تبدأ', 'مكتملة':'المشاريع المكتملة'}

# =========================== small helpers ================================
def norm(s):
    return ' '.join(str(s).split())

def norm_status(s):
    s = norm(s)
    if s.startswith('مكتمل'): return 'مكتملة'
    if s.startswith('على المخطط'): return 'على المخطط'
    if s.startswith('متأخر'): return 'متأخرة'
    if s.startswith('لم تبدأ') or s.startswith('لم يبدأ'): return 'لم تبدأ'
    if s.startswith('معلق'): return 'معلقة'
    return s

def cell_text(c):
    """Cell text with <a:br/> preserved as newlines."""
    out = []
    for para in c.text_frame.paragraphs:
        s = ''
        for ch in para._p:
            if ch.tag == q('r'):   s += ch.find(q('t')).text or ''
            elif ch.tag == q('br'): s += '\n'
        out.append(s)
    return '\n'.join(out).strip()

def cell_text_el(tc):
    out = []
    for p in tc.iter(q('p')):
        sfrag = ''
        for ch in p:
            if ch.tag == q('r'): sfrag += ch.find(q('t')).text or ''
            elif ch.tag == q('br'): sfrag += '\n'
        out.append(sfrag)
    return '\n'.join(out).strip()

def table_rows(shape):
    return [[cell_text(c) for c in r.cells] for r in shape.table.rows]

def first_rpr(tc):
    for r in tc.iter(q('r')):
        rpr = r.find(q('rPr'))
        if rpr is not None: return deepcopy(rpr)
    for e in tc.iter(q('endParaRPr')):
        rpr = deepcopy(e); rpr.tag = q('rPr'); return rpr
    return None

def set_cell(tc, text, algn=None, fallback_rpr=None, bold_date_line=False):
    txb = tc.find(q('txBody'))
    rpr = first_rpr(tc)
    if rpr is None and fallback_rpr is not None: rpr = deepcopy(fallback_rpr)
    ps = txb.findall(q('p'))
    ppr = None
    if ps:
        e = ps[0].find(q('pPr'))
        if e is not None: ppr = deepcopy(e)
    for p in ps: txb.remove(p)
    p = etree.SubElement(txb, q('p'))
    if ppr is None:
        ppr = etree.Element(q('pPr')); ppr.set('rtl', '1')
    if algn:
        ppr.set('algn', algn); ppr.set('rtl', '1')
    p.append(ppr)
    text = re.sub('[\u064B-\u0652]', '', str(text))
    for i, ln in enumerate(str(text).split('\n')):
        if i:
            br = etree.SubElement(p, q('br'))
            if rpr is not None: br.append(deepcopy(rpr))
        r = etree.SubElement(p, q('r'))
        run_rpr = deepcopy(rpr) if rpr is not None else etree.Element(q('rPr'))
        if bold_date_line:
            is_date_line = (i == 0 and re.fullmatch(
                r'\s*تاريخ التحديث\s+\d{1,2}\s+\S+\s+\d{4}\s*', ln) is not None)
            run_rpr.set('b', '1' if is_date_line else '0')
        r.append(run_rpr)
        etree.SubElement(r, q('t')).text = ln

def set_fill(tc, hexcol):
    tcpr = tc.find(q('tcPr'))
    if tcpr is None: return
    for tag in ('solidFill', 'noFill', 'gradFill'):
        for e in tcpr.findall(q(tag)): tcpr.remove(e)
    sf = etree.SubElement(tcpr, q('solidFill'))
    etree.SubElement(sf, q('srgbClr')).set('val', hexcol)

def clean_title_row(tr, accent):
    """Borderless title row: only the small coloured accent bar survives.

    PowerPoint silently drops a whole <a:tcPr> whose children are out of schema
    order and falls back to a black default border, so rebuild it in order:
    lnL, lnR, lnT, lnB, then the fill.
    """
    for i, tc in enumerate(tr.findall(q('tc'))):
        tcpr = tc.find(q('tcPr'))
        if tcpr is None: tcpr = etree.SubElement(tc, q('tcPr'))
        attrs = dict(tcpr.attrib)
        for ch in list(tcpr): tcpr.remove(ch)
        tcpr.attrib.clear(); tcpr.attrib.update(attrs)
        for side in ('lnL', 'lnR', 'lnT', 'lnB'):
            ln = etree.SubElement(tcpr, q(side))
            ln.set('cap', 'flat'); ln.set('cmpd', 'sng'); ln.set('algn', 'ctr')
            if side == 'lnL' and i == 0 and accent:
                ln.set('w', '57150')
                sf = etree.SubElement(ln, q('solidFill'))
                etree.SubElement(sf, q('srgbClr')).set('val', accent)
            else:
                ln.set('w', '12700')
                etree.SubElement(ln, q('noFill'))
            etree.SubElement(ln, q('prstDash')).set('val', 'solid')
        etree.SubElement(tcpr, q('noFill'))

def expand_title_row(tr, ncols):
    """Re-merge a cloned title row across ncols columns.

    A <a:tc> added after <a:extLst> breaks the schema order and PowerPoint then
    ignores hMerge, leaving unmerged cells with black borders — so insert before it.
    """
    tcs = tr.findall(q('tc'))
    tcs[0].set('gridSpan', str(ncols))
    for extra in ('hMerge', 'vMerge', 'rowSpan'):
        tcs[0].attrib.pop(extra, None)
    filler = deepcopy(tcs[1]) if len(tcs) > 1 else deepcopy(tcs[0])
    filler.attrib.pop('gridSpan', None); filler.set('hMerge', '1')
    ext = tr.find(q('extLst'))
    while len(tr.findall(q('tc'))) > ncols:
        tr.remove(tr.findall(q('tc'))[-1])
    while len(tr.findall(q('tc'))) < ncols:
        f = deepcopy(filler)
        ext.addprevious(f) if ext is not None else tr.append(f)
    for tc in tr.findall(q('tc'))[1:]:
        tc.attrib.pop('gridSpan', None); tc.set('hMerge', '1')
    return tr

def gf_tbl(gf): return gf.find('.//' + q('tbl'))

def tbl_colwidths(tbl):
    return [int(g.get('w')) for g in tbl.find(q('tblGrid')).findall(q('gridCol'))]

def rescale(tbl, new_w):
    gcs = tbl.find(q('tblGrid')).findall(q('gridCol'))
    old = [int(g.get('w')) for g in gcs]; tot = sum(old)
    new = [int(round(w * new_w / tot)) for w in old]
    new[-1] += new_w - sum(new)
    for g, w in zip(gcs, new): g.set('w', str(w))

def gf_set_pos(gf, x, y, cx, cy):
    xf = gf.find(qp('xfrm'))
    xf.find(q('off')).set('x', str(int(x))); xf.find(q('off')).set('y', str(int(y)))
    xf.find(q('ext')).set('cx', str(int(cx))); xf.find(q('ext')).set('cy', str(int(cy)))

def next_id(spTree):
    ids = [int(e.get('id')) for e in spTree.iter(qp('cNvPr'))]
    return max(ids) + 1 if ids else 2

def add_frame(spTree, gf, name):
    gf = deepcopy(gf)
    nv = gf.find(qp('nvGraphicFramePr') + '/' + qp('cNvPr'))
    nv.set('id', str(next_id(spTree))); nv.set('name', name)
    spTree.append(gf)
    return gf

def make_table(tmpl_gf, title, rows, algns, status_col=None, title_tr=None, accent=None,
               bold_update_col=None):
    gf = deepcopy(tmpl_gf)
    tbl = gf_tbl(gf)
    trs = tbl.findall(q('tr'))
    ncols = len(tbl_colwidths(tbl))
    body = deepcopy(trs[2])
    fb = first_rpr(body.findall(q('tc'))[0])
    if title_tr is not None:
        newt = expand_title_row(deepcopy(title_tr), ncols)
        tbl.replace(trs[0], newt); ttr = newt
    else:
        ttr = trs[0]
    tcs_t = ttr.findall(q('tc'))
    tcell = next((tc for tc in tcs_t if tc.get('gridSpan')), tcs_t[0])
    clean_title_row(ttr, accent)
    set_cell(tcell, title, algn='r', fallback_rpr=fb)
    for tr in tbl.findall(q('tr'))[2:]: tbl.remove(tr)
    for row in rows:
        if isinstance(row, dict):
            tr = deepcopy(row['tr']); tcs = tr.findall(q('tc'))
            set_cell(tcs[0], row['num'], algn='ctr', fallback_rpr=first_rpr(tcs[0]))
            if status_col is not None:
                set_fill(tcs[status_col],
                         STATUS_FILL.get(norm_status(cell_text_el(tcs[status_col])), 'F2F2F2'))
            tbl.append(tr)
            continue
        tr = deepcopy(body); tcs = tr.findall(q('tc'))
        for ci, val in enumerate(row):
            if ci >= len(tcs): break
            set_cell(tcs[ci], val, algn=algns[ci], fallback_rpr=fb,
                     bold_date_line=(ci == bold_update_col))
        if status_col is not None:
            set_fill(tcs[status_col], STATUS_FILL.get(norm_status(row[status_col]), 'F2F2F2'))
        tr.set('h', '152400')
        tbl.append(tr)
    return gf

# ---- first-pass height estimate (refine_positions corrects it later) -----
CPI, LINE, PAD = 15.5, 0.145, 0.12
def est_rows(rows, widths):
    if rows and isinstance(rows[0], dict):
        return 0.58 * EMU + sum(int(r['tr'].get('h', 152400)) for r in rows)
    return est_table(rows, widths)
def est_table(rows, widths):
    h = 0.28 + 0.30
    for r in rows:
        lines = 1
        for txt, w in zip(r, widths):
            cpl = max(6, (w / EMU - 0.12) * CPI)
            n = sum(max(1, math.ceil(len(seg) / cpl)) for seg in str(txt).split('\n'))
            lines = max(lines, n)
        h += PAD + lines * LINE
    return h * EMU

# ======================= reading the master deck ==========================
def find_tables(slide, must_contain):
    hits = []
    for sh in slide.shapes:
        if not sh.has_table: continue
        rows = table_rows(sh)
        head = ' '.join(rows[1]) if len(rows) > 1 else ''
        if all(k in head for k in must_contain): hits.append((sh, rows))
    return hits

def locate(prs):
    """Identify slide roles by content, not by fixed slide numbers."""
    slides = list(prs.slides)
    visible = [i for i, s in enumerate(slides) if s._element.get('show') != '0']
    thanks = next((i for i in reversed(visible)
                   if any(sh.has_text_frame and 'شكرا' in sh.text_frame.text for sh in slides[i].shapes)),
                  visible[-1])
    role = {'cover': visible[0], 'thanks': thanks,
            'tasks': None, 'projects': None, 'trans': None,
            'task_details': [], 'proj_details': [], 'trans_slides': []}
    for i in visible:
        s = slides[i]
        has_chart = any(sh.has_chart for sh in s.shapes)
        if find_tables(s, ['رقم المعاملة']):
            role['trans_slides'].append(i)
            if role['trans'] is None: role['trans'] = i
        if has_chart and find_tables(s, ['المشروع', 'التحديث']) and not find_tables(s, ['المهمة']):
            role['projects'] = i
        elif has_chart and find_tables(s, ['المهمة']):
            role['tasks'] = i
        elif not has_chart and find_tables(s, ['اسم المشروع']):
            role['proj_details'].append(i)
        elif not has_chart and find_tables(s, ['المهمة']):
            role['task_details'].append(i)
    missing = [k for k in ('tasks', 'projects') if role[k] is None]
    if missing:
        raise SystemExit('Could not locate section slide(s): %s — check the master layout' % missing)
    if role['trans'] is None:
        print('  note: no delayed-transactions table (رقم المعاملة) in the master this week')
    return role

def read_data(prs, role):
    slides = list(prs.slides)
    tasks, projs, trans, ups = [], [], [], []
    for i in role['task_details']:
        for sh, rows in find_tables(slides[i], ['المهمة']):
            for r in rows[2:]:
                if r[0].strip(): tasks.append({'text': r[1], 'sector': r[2],
                                               'status': norm_status(r[4]),
                                               'note': r[5] if len(r) > 5 else ''})
    for i in role['proj_details']:
        for sh, rows in find_tables(slides[i], ['اسم المشروع']):
            sector = rows[0][0]
            trs = sh._element.find('.//' + q('tbl')).findall(q('tr'))
            for ri, r in enumerate(rows[2:], start=2):
                if r[0].strip():
                    projs.append({'sector': sector, 'name': r[1], 'start': r[2],
                                  'end': r[3], 'status': norm_status(r[4]), 'note': r[5],
                                  'tr': deepcopy(trs[ri])})
    for i in role['trans_slides']:
        for sh, rows in find_tables(slides[i], ['رقم المعاملة']):
            for r in rows[2:]:
                if r[0].strip(): trans.append(r)
    for sh, rows in find_tables(slides[role['projects']], ['التحديث']):
        for r in rows[2:]:
            if r[0].strip(): ups.append({'project': r[1], 'sector': r[2], 'text': r[3]})
    return tasks, projs, trans, ups

def kpi_shapes(slide):
    """Pair each KPI number box with its label box by horizontal proximity."""
    nums, labels = [], []
    for sh in slide.shapes:
        if not sh.has_text_frame or sh.top is None or sh.top >= KPI_ZONE: continue
        t = norm(sh.text_frame.text)
        if not t: continue
        cx = sh.left + sh.width / 2
        (nums if t.replace(',', '').isdigit() else labels).append((cx, sh, t))
    out = {}
    for cx, sh, t in nums:
        if not labels: continue
        _, lsh, lt = min(labels, key=lambda L: abs(L[0] - cx))
        key = 'إجمالي' if 'إجمالي' in lt else norm_status(lt)
        out[key] = {'num': sh, 'label': lsh}
    return out

def set_text(shape, value):
    done = False
    for p in shape.text_frame.paragraphs:
        for r in p.runs:
            if not done: r.text = str(value); done = True
            else: r.text = ''
    if not done: shape.text_frame.text = str(value)

# ============================ deck assembly ===============================
DY = 700000   # room opened under the cover title for the sector line

def fix_cover(prs, sector):
    slide = prs.slides[0]
    grp = slide.shapes[0]._element
    xf = grp.find(qp('grpSpPr')).find(q('xfrm'))
    off, ext, chext = xf.find(q('off')), xf.find(q('ext')), xf.find(q('chExt'))
    sps = [ch for ch in grp if ch.tag in (qp('sp'), qp('cxnSp'), qp('pic'), qp('graphicFrame'))]
    title_el = max(sps, key=lambda e: _font_size(e))
    date_el = next((e for e in sps if e is not title_el and e.find('.//' + q('t')) is not None), None)
    txf = _xfrm(title_el)
    base_y = int(txf.find(q('off')).get('y')) + int(txf.find(q('ext')).get('cy'))
    for ch in sps:
        if ch is title_el: continue
        x = _xfrm(ch)
        if x is None: continue
        o = x.find(q('off'))
        if o is not None: o.set('y', str(int(o.get('y')) + DY))
    ext.set('cy', str(int(ext.get('cy')) + DY))
    chext.set('cy', str(int(chext.get('cy')) + DY))
    off.set('y', str(int(off.get('y')) - DY // 2))

    new = deepcopy(date_el)
    new.find('.//' + qp('cNvPr')).set('id', '900')
    new.find('.//' + qp('cNvPr')).set('name', 'SectorName')
    xf2 = _xfrm(new)
    xf2.find(q('off')).set('x', str(int(xf.find(q('chOff')).get('x'))))
    xf2.find(q('off')).set('y', str(base_y + 30000))
    xf2.find(q('ext')).set('cx', str(int(chext.get('cx'))))
    xf2.find(q('ext')).set('cy', '620000')
    bp = new.find('.//' + q('bodyPr'))
    for af in bp.findall(q('spAutoFit')): bp.remove(af)
    bp.set('anchor', 'ctr'); bp.set('wrap', 'square')
    first = True
    for p in new.findall('.//' + q('p')):
        for r in p.findall(q('r')):
            if first:
                r.find(q('t')).text = sector
                rpr = r.find(q('rPr')); rpr.set('sz', '1600')
                for f in rpr.findall(q('solidFill')): rpr.remove(f)
                sf = etree.Element(q('solidFill'))
                etree.SubElement(sf, q('srgbClr')).set('val', 'FFFFFF')
                rpr.insert(0, sf); first = False
            else:
                p.remove(r)
    grp.append(new)

def _xfrm(el):
    for holder in (qp('spPr'), qp('grpSpPr')):
        h = el.find(holder)
        if h is not None:
            x = h.find(q('xfrm'))
            if x is not None: return x
    return None

def _font_size(sp):
    return max([int(r.get('sz', 0)) for r in sp.iter(q('rPr'))] + [0])

def strip_and_centre_nav(prs):
    for layout in prs.slide_masters[0].slide_layouts:
        for sh in list(layout.shapes):
            if sh.left == NAV_DROP_X and sh.top in NAV_TOPS:
                layout.shapes._spTree.remove(sh._element)
        for sh in layout.shapes:
            if sh.left in NAV_XS and sh.top in NAV_TOPS:
                sh.left = Emu(sh.left - NAV_SHIFT)

def clear_body(slide):
    """Drop the cross-sector chart, its card and the template table."""
    spTree = slide.shapes._spTree
    for sh in list(slide.shapes):
        if sh.top is not None and sh.top >= KPI_ZONE:
            spTree.remove(sh._element)

def single_kpi(slide, count, label_text):
    """Transactions: keep only the gold total, recentred, relabelled."""
    kp = kpi_shapes(slide)
    keep = kp.get('إجمالي')
    if keep is None: return
    keep_ids = {keep['num'].shape_id, keep['label'].shape_id}
    spTree = slide.shapes._spTree
    for sh in list(slide.shapes):
        if sh.top is not None and sh.top < KPI_ZONE:
            if sh.width and sh.width > 5000000:
                continue
            if sh.shape_id in keep_ids: continue
            spTree.remove(sh._element)
    set_text(keep['num'], count)
    set_text(keep['label'], label_text)
    keep['num'].left = Emu(int(CX + (CW - keep['num'].width) / 2))
    keep['label'].width = Emu(2600000)
    keep['label'].left = Emu(int(CX + (CW - 2600000) / 2))
    for p in keep['label'].text_frame.paragraphs:
        pPr = p._p.find(q('pPr'))
        if pPr is None:
            pPr = etree.Element(q('pPr')); p._p.insert(0, pPr)
        pPr.set('algn', 'ctr'); pPr.set('rtl', '1')

def build(master, outdir, aliases, do_refine=True):
    os.makedirs(outdir, exist_ok=True)
    full_dir = os.path.join(outdir, '_full'); os.makedirs(full_dir, exist_ok=True)

    src = Presentation(master)
    role = locate(src)
    tasks, projs, trans, ups = read_data(src, role)

    canon = {norm(k): v for k, v in aliases.items()}
    def C(name): return canon.get(norm(name), norm(name))

    sectors = {}
    def B(n): return sectors.setdefault(n, {'tasks': [], 'projs': [], 'trans': [], 'ups': []})
    for t in tasks: B(C(t['sector']))['tasks'].append(t)
    for p in projs: B(C(p['sector']))['projs'].append(p)
    for r in trans: B(C(r[3]))['trans'].append(r)
    for u in ups:   B(C(u['sector']))['ups'].append(u)

    sl = list(src.slides)
    def frame(i, keyword):
        return deepcopy(find_tables(sl[i], keyword)[0][0]._element)
    TMPL_TASK  = frame(role['tasks'], ['المهمة'])
    TMPL_UPD   = frame(role['projects'], ['التحديث'])
    TMPL_TRANS = frame(role['trans'], ['رقم المعاملة']) if role['trans'] is not None else None
    upd_task = find_tables(sl[role['tasks']], ['المهمة', 'التحديث'])
    TMPL_TASK_UPD = deepcopy(upd_task[0][0]._element) if upd_task else None
    UPD_TASK_TITLE = norm(upd_task[0][1][0][0]) if upd_task else ''
    TMPL_PROJ  = frame(role['proj_details'][-1], ['اسم المشروع'])
    GOLD_TITLE = deepcopy(gf_tbl(TMPL_UPD).findall(q('tr'))[0])
    rescale(gf_tbl(TMPL_PROJ), CW)

    manifest = []
    for sec in sorted(sectors, key=lambda s: (0 if s.startswith('وكالة') else 1, s)):
        d = sectors[sec]
        if not (d['tasks'] or d['projs'] or d['trans']): continue
        prs = Presentation(master)
        strip_and_centre_nav(prs)
        fix_cover(prs, sec)
        keep = [role['cover']]
        if d['tasks']: keep.append(role['tasks'])
        if d['projs']: keep.append(role['projects'])
        if d['trans']: keep.append(role['trans'])
        keep.append(role['thanks'])

        if d['tasks']:
            s = prs.slides[role['tasks']]
            kp = kpi_shapes(s)
            cnt = {k: sum(1 for t in d['tasks'] if t['status'] == k) for k in TASK_ORDER}
            for key, val in [('إجمالي', len(d['tasks']))] + list(cnt.items()):
                if key in kp: set_text(kp[key]['num'], val)
            clear_body(s)
            widths = tbl_colwidths(gf_tbl(TMPL_TASK)); y = TOP_Y
            for st in TASK_ORDER:
                sel = [x for x in d['tasks'] if x['status'] == st]
                if not sel: continue
                if TMPL_TASK_UPD is not None and UPD_TASK_TITLE == TASK_TITLE[st]:
                    rows = [[str(i + 1), t['text'], sec, t['note'].strip()] for i, t in enumerate(sel)]
                    gf = make_table(TMPL_TASK_UPD, TASK_TITLE[st], rows, ['ctr', 'r', 'ctr', 'r'],
                                    accent=ACCENT[st])
                    h = est_table(rows, tbl_colwidths(gf_tbl(TMPL_TASK_UPD)))
                else:
                    rows = [[str(i + 1), t['text'], sec] for i, t in enumerate(sel)]
                    gf = make_table(TMPL_TASK, TASK_TITLE[st], rows, ['ctr', 'r', 'ctr'],
                                    accent=ACCENT[st])
                    h = est_table(rows, widths)
                gf_set_pos(add_frame(s.shapes._spTree, gf, 'tbl_' + st), CX, y, CW, h)
                y += h + GAP

        if d['projs']:
            s = prs.slides[role['projects']]
            kp = kpi_shapes(s)
            cnt = {k: sum(1 for p in d['projs'] if p['status'] == k) for k in PROJ_ORDER}
            for key, val in [('إجمالي', len(d['projs']))] + list(cnt.items()):
                if key in kp: set_text(kp[key]['num'], val)
            clear_body(s)
            y = TOP_Y
            if d['ups']:
                rows = [[str(i + 1), u['project'], sec, u['text']] for i, u in enumerate(d['ups'])]
                w = tbl_colwidths(gf_tbl(TMPL_UPD))
                gf = make_table(TMPL_UPD, 'أبرز التحديثات', rows, ['ctr', 'r', 'ctr', 'r'],
                                accent=ACCENT['تحديثات'], bold_update_col=3)
                h = est_table(rows, w)
                gf_set_pos(add_frame(s.shapes._spTree, gf, 'tbl_updates'), CX, y, CW, h)
                y += h + GAP
            wp = tbl_colwidths(gf_tbl(TMPL_PROJ))
            for st in PROJ_ORDER:
                rows = [{'num': str(i + 1), 'tr': p['tr']}
                        for i, p in enumerate([x for x in d['projs'] if x['status'] == st])]
                if not rows: continue
                gf = make_table(TMPL_PROJ, PROJ_TITLE[st], rows,
                                ['ctr', 'r', 'ctr', 'ctr', 'ctr', 'r'],
                                status_col=4, title_tr=GOLD_TITLE, accent=ACCENT[st])
                h = est_rows(rows, wp)
                gf_set_pos(add_frame(s.shapes._spTree, gf, 'proj_' + st), CX, y, CW, h)
                y += h + GAP

        if d['trans']:
            s = prs.slides[role['trans']]
            tframe = find_tables(s, ['رقم المعاملة'])[0][0]
            tx, ty, tw = tframe.left, tframe.top, tframe.width
            clear_body(s)
            single_kpi(s, len(d['trans']), 'إجمالي المعاملات المتأخرة')
            rows = [[f'{i+1:02d}', r[1], r[2], sec, r[4], r[5], r[6]]
                    for i, r in enumerate(d['trans'])]
            w = tbl_colwidths(gf_tbl(TMPL_TRANS))
            gf = make_table(TMPL_TRANS, 'المعاملات المتأخرة', rows,
                            ['ctr', 'ctr', 'r', 'ctr', 'r', 'ctr', 'ctr'],
                            accent=ACCENT['متأخرة'])
            gf_set_pos(add_frame(s.shapes._spTree, gf, 'tbl_trans'),
                       tx, ty, tw, est_table(rows, w))

        path = os.path.join(full_dir, 'تقرير الإنجاز الأسبوعي - %s.pptx' % sec)
        prs.save(path)
        manifest.append({'sector': sec, 'file': path, 'keep': keep,
                         'edited': [i for i in (role['tasks'], role['projects'], role['trans'])
                                    if i in keep],
                         'n': [len(d['tasks']), len(d['projs']), len(d['trans']), len(d['ups'])],
                         'nslides': len(sl)})
    return manifest

# ================== pass 2: measure, then reflow ==========================
def _bands(page):
    rs = [r for r in page.rects if 20 < r['width'] < 560 and r['height'] < 200 and r['top'] > 60]
    rs.sort(key=lambda r: r['top'])
    gs = []
    for r in rs:
        if gs and r['top'] - gs[-1][1] < 10: gs[-1][1] = max(gs[-1][1], r['bottom'])
        else: gs.append([r['top'], r['bottom']])
    return gs

def to_pdf(pptx_path, work='_m'):
    shutil.copy(pptx_path, work + '.pptx')
    if os.path.exists(work + '.pdf'): os.remove(work + '.pdf')
    if not SOFFICE:
        raise RuntimeError('soffice is required for the refinement pass')
    pdf_dir = tempfile.mkdtemp(prefix='soffice-pdf-')
    profile_dir = tempfile.mkdtemp(prefix='soffice-profile-')
    try:
        subprocess.run([SOFFICE, '-env:UserInstallation=file://' + profile_dir,
                        '--headless', '--convert-to', 'pdf',
                        '--outdir', pdf_dir, work + '.pptx'],
                       capture_output=True, check=True)
        shutil.move(os.path.join(pdf_dir, os.path.basename(work) + '.pdf'),
                    work + '.pdf')
    finally:
        shutil.rmtree(pdf_dir, ignore_errors=True)
        shutil.rmtree(profile_dir, ignore_errors=True)
    os.remove(work + '.pptx')
    return work + '.pdf'

SAFETY = 1.12
OFFPAGE = 40_000_000

def _row_bands(page, x0, x1, y0):
    """Row rectangles of the one table rendered on this page, top to bottom (points)."""
    rs = [r for r in page.rects if r['x0'] > x0 - 3 and r['x1'] < x1 + 3
          and r['top'] > y0 - 3 and r['height'] > 2]
    bands = []
    for r in sorted(rs, key=lambda r: (r['top'], r['bottom'])):
        if bands and r['top'] < bands[-1][1] - 1:
            bands[-1][1] = max(bands[-1][1], r['bottom'])
        else:
            bands.append([r['top'], r['bottom']])
    return bands

def _frames(slide):
    return sorted([sh for sh in slide.shapes if sh.has_table], key=lambda s: s.top)

def refine_positions(manifest):
    """Measure every row and write the real height back, so a table's declared
    geometry matches what it occupies. Stacked tables overlap each other in the
    render, which would corrupt the measurement, so each pass sends every table
    but the one being measured off the page."""
    import pdfplumber
    for item in manifest:
        prs = Presentation(item['file'])
        nmax = max((len(_frames(prs.slides[si])) for si in item['edited']), default=0)
        heights = {}
        for k in range(nmax):
            probe = Presentation(item['file'])
            for si in item['edited']:
                for j, sh in enumerate(_frames(probe.slides[si])):
                    if j != k: sh.top = Emu(OFFPAGE + j * 100000)
            probe.save('_probe.pptx')
            pdf_path = to_pdf('_probe.pptx')
            with pdfplumber.open(pdf_path) as pdf:
                for si in item['edited']:
                    fr = _frames(prs.slides[si])
                    if k >= len(fr): continue
                    sh = fr[k]
                    ntr = len(sh._element.find('.//' + q('tbl')).findall(q('tr')))
                    bands = _row_bands(pdf.pages[si], sh.left / EMU * 72,
                                       (sh.left + sh.width) / EMU * 72, sh.top / EMU * 72)
                    if len(bands) != ntr - 1:
                        print('  ! row-band mismatch on slide %d of %s (%d bands, %d rows)'
                              % (si + 1, item['sector'], len(bands), ntr))
                        continue
                    title_h = max(0.0, bands[0][0] - sh.top / EMU * 72)
                    heights[(si, k)] = [int(h / 72 * EMU * SAFETY)
                                        for h in [title_h] + [b[1] - b[0] for b in bands]]
            os.remove(pdf_path); os.remove('_probe.pptx')
        for si in item['edited']:
            fr = _frames(prs.slides[si])
            y = fr[0].top if fr else 0
            for k, sh in enumerate(fr):
                hs = heights.get((si, k))
                if hs:
                    for tr, h in zip(sh._element.find('.//' + q('tbl')).findall(q('tr')), hs):
                        tr.set('h', str(h))
                    sh.height = Emu(sum(hs))
                sh.top = Emu(int(y))
                y += sh.height + GAP
        prs.save(item['file'])

# ============ pass 3: drop unused slides at the XML level =================
def shrink_media(wd):
    """The three full-bleed backgrounds are ~9 MB of the master. Re-encode them at
    the same pixel size — no visible change, and a sector deck lands around 2.7 MB."""
    from PIL import Image
    for f in pathlib.Path(wd, 'ppt', 'media').glob('*'):
        if f.suffix.lower() not in ('.jpg', '.jpeg') or f.stat().st_size < 500_000: continue
        im = Image.open(f)
        buf = __import__('io').BytesIO()
        im.save(buf, 'JPEG', quality=85, optimize=True)
        if buf.tell() < f.stat().st_size: f.write_bytes(buf.getvalue())

def slim(manifest, outdir):
    """python-pptx renumbers slide parts on save and collides after deletion,
    so remove the sldId entries directly and let clean.py garbage-collect."""
    final = []
    for item in manifest:
        wd = os.path.join(outdir, '_work')
        shutil.rmtree(wd, ignore_errors=True)
        with zipfile.ZipFile(item['file']) as z: z.extractall(wd)
        pres = pathlib.Path(wd) / 'ppt' / 'presentation.xml'
        xml = pres.read_text(encoding='utf-8')
        m = re.search(r'(<p:sldIdLst>)(.*?)(</p:sldIdLst>)', xml, re.S)
        ids = re.findall(r'<p:sldId\b[^>]*/>', m.group(2))
        keep = set(item['keep'])
        xml = xml[:m.start()] + m.group(1) + ''.join(s for i, s in enumerate(ids) if i in keep) \
              + m.group(3) + xml[m.end():]
        pres.write_text(xml, encoding='utf-8')
        # Unreferenced slide parts are harmless; dangling relationships are removed below.
        # back-references from layouts to deleted slides, and the nav hyperlinks
        for rels in pathlib.Path(wd).rglob('_rels/*.rels'):
            base = rels.parent.parent
            txt = rels.read_text(encoding='utf-8'); orig = txt
            for tag in re.findall(r'<Relationship\b[^>]*/>', txt):
                if 'TargetMode="External"' in tag: continue
                tm = re.search(r'Target="([^"]+)"', tag)
                if tm and not (base / tm.group(1)).resolve().exists():
                    txt = txt.replace(tag, '')
            if txt != orig: rels.write_text(txt, encoding='utf-8')
        for rels in pathlib.Path(wd).rglob('_rels/*.rels'):
            part = rels.parent.parent / rels.name[:-5]
            if not part.exists() or part.suffix != '.xml': continue
            valid = set(re.findall(r'Id="([^"]+)"', rels.read_text(encoding='utf-8')))
            txt = part.read_text(encoding='utf-8'); orig = txt
            def drop(mm):
                rid = re.search(r'r:id="([^"]*)"', mm.group(0))
                return '' if (rid and rid.group(1) and rid.group(1) not in valid) else mm.group(0)
            txt = re.sub(r'<a:hlinkClick\b[^>]*/>', drop, txt)
            txt = re.sub(r'<a:hlinkClick\b[^>]*>.*?</a:hlinkClick>', drop, txt, flags=re.S)
            if txt != orig: part.write_text(txt, encoding='utf-8')
        shrink_media(wd)
        out = os.path.abspath(os.path.join(outdir, os.path.basename(item['file'])))
        if os.path.exists(out): os.remove(out)
        subprocess.run(['zip', '-Xqr', out, '.'], cwd=wd, check=True)
        shutil.rmtree(wd, ignore_errors=True)
        item['out'] = out
        final.append(item)
    shutil.rmtree(os.path.join(outdir, '_full'), ignore_errors=True)
    return final

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('master')
    ap.add_argument('--outdir', default='out')
    ap.add_argument('--aliases', default=os.path.join(
        os.path.dirname(os.path.abspath(__file__)), '..', 'references', 'sector-aliases.json'))
    ap.add_argument('--no-refine', action='store_true')
    ap.add_argument('--only', default=None)
    a = ap.parse_args()
    aliases = json.load(open(a.aliases, encoding='utf-8')) if os.path.exists(a.aliases) else {}
    man = build(a.master, a.outdir, aliases)
    if a.only: man = [m for m in man if a.only in m['sector']]
    if not a.no_refine:
        refine_positions(man)
    man = slim(man, a.outdir)
    json.dump(man, open(os.path.join(a.outdir, 'manifest.json'), 'w'),
              ensure_ascii=False, indent=1)
    for it in man:
        print('%-45s tasks=%d projects=%d transactions=%d updates=%d slides=%d'
              % (it['sector'], *it['n'], len(it['keep'])))

if __name__ == '__main__':
    main()
