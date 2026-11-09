import os, glob as _glob
HERE = os.path.dirname(os.path.abspath(__file__))
WORK = os.environ.get('RGC_WORK', os.path.join(HERE, 'work'))
BASE = os.environ.get('RGC_BASE', '/mnt/user-data/uploads/torre-control-ccr')
import datetime as _dt
CORTE = _dt.datetime.fromisoformat(os.environ.get('RGC_CORTE', '2026-11-03'))
"""Diagrama por objetivo al estilo de las RGC anteriores: indicadores -> línea base -> meta -> producto crítico -> adquisiciones -> riesgos."""
import re, math
from pptx.util import Emu, Pt
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor

PX = 12192000 / 1067
ICONS = os.path.join(HERE, 'icons')
FONT = 'Public Sans'
C_IND, C_VAL, C_HDR, C_BOX, C_BOXL, C_LINE, C_TXT = '0E6B78', '5B9FCB', '3A6E8A', 'EAEAEA', 'B4C3D6', 'A6A6A6', '1F2A36'
PROB = {'Alta': '00B050', 'Media': 'FFC000', 'Baja': 'FF0000'}
MAX_IND, MAX_PROD = 7, 5
T0, B0 = 182, 484          # vertical band for the rows
X_IND, W_IND = 62, 222
X_BASE, X_META, W_VAL = 300, 368, 58
X_PROD, W_PROD = 474, 168
X_ADQ, W_ADQ = 660, 164
X_RSK, W_RSK = 840, 165

E = lambda v: Emu(int(v * PX))

def _nostyle(sh):
    st = sh._element.find('{http://schemas.openxmlformats.org/presentationml/2006/main}style')
    if st is not None: sh._element.remove(st)

def box(s, x, y, w, h, fill, text='', size=9, color='FFFFFF', bold=False, line=None, align=PP_ALIGN.CENTER,
        radius=0.18, shape=MSO_SHAPE.ROUNDED_RECTANGLE, anchor=MSO_ANCHOR.MIDDLE, pars=None, inset=4):
    sh = s.shapes.add_shape(shape, E(x), E(y), E(w), E(h)); _nostyle(sh)
    if shape == MSO_SHAPE.ROUNDED_RECTANGLE: sh.adjustments[0] = radius
    if fill: sh.fill.solid(); sh.fill.fore_color.rgb = RGBColor.from_string(fill)
    else: sh.fill.background()
    if line: sh.line.color.rgb = RGBColor.from_string(line); sh.line.width = Pt(1)
    else: sh.line.fill.background()
    tf = sh.text_frame; tf.word_wrap = True; tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = E(inset); tf.margin_top = tf.margin_bottom = E(2)
    for i, t in enumerate(pars if pars is not None else ([text] if text else [])):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        if i: p.space_before = Pt(3)
        r = p.add_run(); r.text = t
        f = r.font; f.size = Pt(size); f.bold = bold; f.name = FONT; f.color.rgb = RGBColor.from_string(color)
    return sh

def text(s, x, y, w, h, runs, size=10, align=PP_ALIGN.LEFT):
    tb = s.shapes.add_textbox(E(x), E(y), E(w), E(h)); tf = tb.text_frame; tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    p = tf.paragraphs[0]; p.alignment = align
    for t, b, c in runs:
        r = p.add_run(); r.text = t; r.font.size = Pt(size); r.font.bold = b; r.font.name = FONT
        r.font.color.rgb = RGBColor.from_string(c)
    return tb

def line(s, x1, y1, x2, y2, w=2.5):
    c = s.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, E(x1), E(y1), E(x2), E(y2)); _nostyle(c)
    c.line.color.rgb = RGBColor.from_string(C_LINE); c.line.width = Pt(w)
    return c

def circle(s, x, y, d, fill):
    sh = s.shapes.add_shape(MSO_SHAPE.OVAL, E(x), E(y), E(d), E(d)); _nostyle(sh)
    sh.fill.solid(); sh.fill.fore_color.rgb = RGBColor.from_string(fill)
    sh.line.color.rgb = RGBColor.from_string('FFFFFF'); sh.line.width = Pt(0.75)
    return sh

def icon(s, estado, x, y, d):
    try: return s.shapes.add_picture(f'{ICONS}/{estado}.png', E(x), E(y), E(d), E(d))
    except FileNotFoundError: return None

def fit(t, w, h, size):
    """Shorten text so it fits a w x h box at the given point size."""
    cpl = max(8, int((w - 12) / (size * 1.333 * 0.46)))
    lines = max(1, int((h - 2) / (size * 1.333 * 1.17)))
    t = (t or '').strip()
    out, cur = [], ''
    for word in t.split():
        if len(cur) + len(word) + (1 if cur else 0) <= cpl: cur = (cur + ' ' + word).strip()
        else: out.append(cur); cur = word
    out.append(cur)
    if len(out) <= lines: return t
    keep = ' '.join(out[:lines])
    return keep[:max(0, len(keep) - 2)].rsplit(' ', 1)[0] + '…'

def nlines(t, w, size):
    cpl = max(8, int((w - 12) / (size * 1.333 * 0.46))); n, cur = 1, ''
    for word in (t or '').split():
        if len(cur) + len(word) + (1 if cur else 0) <= cpl: cur = (cur + ' ' + word).strip()
        else: n += 1; cur = word
    return n

def val(v):
    if v is None or v == '': return '—'
    if isinstance(v, float):
        s = f'{v:,.3f}'.rstrip('0').rstrip('.')
        return s.replace(',', 'X').replace('.', ',').replace('X', '.')
    if isinstance(v, int): return f'{v:,}'.replace(',', '.')
    return str(v).replace('.', ',')

UNIT = re.compile(r'\s*\[([^\]]+)\]\s*$')

def plan(S):
    """One page per chunk of <=7 indicators (and <=5 products) per objective."""
    pages = []
    objs = list(dict.fromkeys(r['obj_id'] for r in S['res']))
    for o in objs:
        rs = [r for r in S['res'] if r['obj_id'] == o]
        n = math.ceil(len(rs) / MAX_IND); size = math.ceil(len(rs) / n)
        chunks = [rs[i:i + size] for i in range(0, len(rs), size)]
        sub = []
        for ch in chunks:
            ids = {r['ind_id'] for r in ch}
            ps = [p for p in S['prods'] if set(S['prod_inds'].get(p['prod'], [])) & ids]
            # order products by the first indicator they serve, so lines cross as little as possible
            first = {p['prod']: min(i for i, r in enumerate(ch) if r['ind_id'] in S['prod_inds'].get(p['prod'], [])) for p in ps}
            ps.sort(key=lambda p: first[p['prod']])
            for i in range(0, max(1, len(ps)), MAX_PROD): sub.append((ch, ps[i:i + MAX_PROD]))
        for k, (ch, ps) in enumerate(sub): pages.append(dict(obj=o, inds=ch, prods=ps, k=k, n=len(sub)))
    return pages

def adq_text(S, p):
    if p.get('adq'): return p['adq']
    procs = [c for c in S['crit'] if c['prod'] == p['prod'] and c.get('nombre')]
    if procs: return ' / '.join(c['nombre'] for c in procs[:2])
    gone = [c for c in S['crit'] if c['prod'] == p['prod'] and c.get('eliminado')]
    if gone: return gone[0]['eliminado']['nombre'] + (' y otros' if len(gone) > 1 else '') + ' (retirado del plan de adquisiciones)'
    return 'No aplica (producto logrado)' if p.get('estado') == 'Logrado' else 'Sin adquisición crítica registrada'

def risk_list(S, p):
    rs = [x['riesgo'] for x in S['prod_adq'].get(p['prod'], []) if x.get('critica') == 'Sí' and x.get('riesgo')]
    if not rs and p.get('riesgo'): rs = [p['riesgo']]
    out = []
    for r in rs:
        for part in re.split(r'\s+/\s+', r):
            if part.strip() and part.strip() not in out: out.append(part.strip())
    return out

def draw(s, S, page, set_title):
    inds, prods = page['inds'], page['prods']
    ALLP = [q['prod'] for q in S['prods']]; ALLA = [adq_text(S, q) for q in S['prods']]
    obj = inds[0]['obj']
    suf = f" ({page['k'] + 1} de {page['n']})" if page['n'] > 1 else ''
    lead = f"Objetivo específico {page['obj']}: "
    body = obj.strip()
    CPL = 74
    if len(lead) + len(body) + len(suf) > 2 * CPL:
        body = body[:2 * CPL - len(lead) - len(suf) - 2].rsplit(' ', 1)[0] + '…'
    set_title(lead, f"{body}{suf}")
    dy = 24 if len(lead) + len(body) + len(suf) > CPL else 0
    T = T0 + dy
    # units shared by the page go in the header
    units = {(UNIT.search(r['ind'] or '') or [None, ''])[1] for r in inds}
    unit = units.pop() if len(units) == 1 else ''
    # header band
    box(s, 290, 136 + dy, 715, 36, C_HDR, radius=0.25)
    hdr = [(X_BASE, W_VAL, ['Línea base'] + ([f'[{unit}]'] if unit else [])), (X_META, W_VAL, ['Meta'] + ([f'[{unit}]'] if unit else [])),
           (X_PROD, W_PROD, ['Producto crítico']), (X_ADQ, W_ADQ, ['Adquisiciones / procesos críticos']), (X_RSK, W_RSK, ['Riesgos'])]
    for x, w, ps in hdr:
        tb = box(s, x - 6, 136 + dy, w + 12, 36, None, pars=ps, size=9 if len(ps) == 1 else 8, bold=True, inset=1)
        if len(ps) > 1:
            tb.text_frame.paragraphs[1].runs[0].font.bold = False; tb.text_frame.paragraphs[1].space_before = Pt(0)
    # indicator rows
    n = len(inds); band = B0 - T; slot = band / n; top = T
    hi = min(44, slot - 6)
    yc = {}
    for i, r in enumerate(inds):
        yc[r['ind_id']] = top + (i + 0.5) * slot
    m = max(1, len(prods)); pslot = band / m
    ph = min(72, pslot - 10)
    pyc = {p['prod']: top + (j + 0.5) * pslot for j, p in enumerate(prods)}
    # lines first (behind boxes)
    for r in inds:
        c = yc[r['ind_id']]
        line(s, X_IND + W_IND, c, X_META + W_VAL, c)
        for p in prods:
            if r['ind_id'] in S['prod_inds'].get(p['prod'], []):
                line(s, X_META + W_VAL, c, X_PROD, pyc[p['prod']])
    for p in prods:
        c = pyc[p['prod']]
        line(s, X_PROD + W_PROD, c, X_ADQ, c)
        if risk_list(S, p): line(s, X_ADQ + W_ADQ, c, X_RSK, c)
    # indicator, base and meta boxes
    for r in inds:
        y = yc[r['ind_id']] - hi / 2
        name = f"{r['ind_id']} {UNIT.sub('', r['ind'] or '')}"
        fs = 8 if len(name) > 62 else 9
        box(s, X_IND, y, W_IND, hi, C_IND, fit(name, W_IND - 6, hi, fs), size=fs, align=PP_ALIGN.LEFT, inset=10)
        circle(s, X_IND - 6, y + hi - 12, 15, PROB.get(r['prob'], 'BFBFBF'))
        vh = min(30, hi)
        box(s, X_BASE, yc[r['ind_id']] - vh / 2, W_VAL, vh, C_VAL, val(r['base']), size=9.5)
        box(s, X_META, yc[r['ind_id']] - vh / 2, W_VAL, vh, C_VAL, val(r['meta']), size=9.5)
    # product, procurement and risk boxes
    for p in prods:
        c = pyc[p['prod']]; y = c - ph / 2
        fs = 10 if ph >= 50 else 9
        box(s, X_PROD, y, W_PROD, ph, C_BOX, fit(compact(p['prod'], ALLP, capacity(W_PROD, ph, fs)), W_PROD, ph, fs), size=fs, color=C_TXT, line=C_BOXL, radius=0.15)
        icon(s, p['estado'], X_PROD + W_PROD - 12, y + ph - 12, 17)
        afs = 8
        box(s, X_ADQ, y, W_ADQ, ph, C_BOX, fit(compact(adq_text(S, p), ALLA, capacity(W_ADQ, ph, afs)), W_ADQ, ph, afs), size=afs, color=C_TXT, line=C_BOXL,
            align=PP_ALIGN.LEFT, radius=0.15, inset=6)
        rk = risk_list(S, p)
        if rk:
            rfs = 8 if len(rk) == 1 else 7.5
            lh = rfs * 1.333 * 1.17
            need = sum(nlines(t, W_RSK, rfs) for t in rk) * lh + 6 * (len(rk) - 1) + 10
            rh = min(pslot - 4, max(ph, need)); ry = c - rh / 2
            if need <= rh + 1: pars = rk
            else: pars = [fit(' · '.join(rk), W_RSK, rh, rfs)]
            box(s, X_RSK, ry, W_RSK, rh, C_BOX, pars=pars, size=rfs, color=C_TXT,
                line=C_BOXL, align=PP_ALIGN.LEFT, radius=0.12, inset=6)
    if not prods:
        box(s, X_PROD, (T0 + B0) / 2 - 25, W_PROD, 50, C_BOX, 'Sin producto crítico vinculado', size=9, color=C_TXT, line=C_BOXL)
    legend(s)

def legend(s):
    y, h = 497, 26
    box(s, 62, y, 318, h, 'F2F2F2', line='7F7F7F', radius=0.5)
    text(s, 72, y, 150, h, [('“Probabilidad” de logro:', True, C_TXT)], size=9.5)
    x = 214
    for lab, col in PROB.items():
        circle(s, x, y + 6, 14, col); text(s, x + 18, y, 46, h, [(lab, False, C_TXT)], size=9.5); x += 56
    box(s, 396, y, 609, h, 'F2F2F2', line='7F7F7F', radius=0.5)
    x = 410
    for lab in ('Pausado', 'Cancelado', 'Retrasado', 'En tiempo', 'Logrado'):
        icon(s, lab, x, y + 5, 16); text(s, x + 21, y, 90, h, [(lab, False, C_TXT)], size=9.5); x += 119

# ---------- distinguishing short names ----------
GENERIC = {'conclusión', 'diseño', 'construcción', 'contratación', 'adquisición', 'suministro', 'servicios', 'servicio',
           'rehabilitación', 'ampliación', 'equipamiento', 'supervisión', 'consultoría', 'elaboración', 'implementación', 'compra'}
CONN = {'de', 'del', 'la', 'el', 'los', 'las', 'en', 'y', 'para', 'por', 'a', 'al', 'e', 'con'}
def _cap(w): return bool(w) and (w[0].isupper() or w[0].isdigit() or w[0] == '#')
def _core(ws):
    if len(ws) > 1 and _cap(ws[0]) and (re.search(r'[#\d]', ws[1]) or ws[1][0].isupper()): start = 0
    else: start = next((i for i in range(1, len(ws)) if ws[i][0].isupper()), None)
    if start is None: return []
    out = [ws[start]]; i = start + 1
    while i < len(ws):
        if _cap(ws[i]) and not ws[i - 1].endswith((',', ':', ';')): out.append(ws[i]); i += 1
        elif ws[i] in ('de', 'y') and i + 1 < len(ws) and ws[i + 1][0].isupper(): out += ws[i:i + 2]; i += 2
        else: break
    return out
def _loc(ws):
    idx = [i for i, w in enumerate(ws[:-1]) if w == 'en' and ws[i + 1][0].isupper()]
    return ws[idx[-1] + 1:] if idx else []
def _cut(s, n):
    s = (s or '').strip()
    return s if len(s) <= n else s[:n - 1].rsplit(' ', 1)[0].rstrip(',;:') + '…'
def _base(n):
    return re.sub(r'\s*\((previsto|prevista)\)\s*$', '', (n or '').strip()).rstrip('. ')
def compact(name, others, limit, generic=False):
    """Shorten `name` to <= limit chars; when plain truncation would make it look like another name
    of the set, keep the words that tell them apart (e.g. 'Centro Cívico Pérez Zeledón')."""
    name = (name or '').strip()
    if len(name) <= limit: return name
    m = re.search(r'\s*\((previsto|prevista)\)\s*$', name); suf = ' (prev.)' if m else ''
    base = _base(name)
    plain = _cut(base, limit - len(suf)) + suf if len(base) + len(suf) > limit else base + suf
    rivals = [_base(o) for o in others if o and _base(o) != base]
    head = lambda x: _cut(x, limit - len(suf))
    amb = [o for o in rivals if head(o) == head(base) or (len(head(base)) > 12 and head(o)[:len(head(base)) - 1] == head(base)[:-1])]
    ws = base.replace('–', ' – ').split()
    def _wp(o):
        ow = o.replace('–', ' – ').split(); k = 0
        while k < min(len(ws), len(ow)) and ws[k] == ow[k]: k += 1
        return k
    amb += [o for o in rivals if o not in amb and _wp(o) >= 6]
    if not amb:
        if generic and ws[0].lower().rstrip(',') in GENERIC:
            core = _core(ws)
            if core and len(core) > 1:
                loc = [w for w in _loc(ws) if w not in core]
                cand = ' '.join(core + loc)
                if len(cand) + len(suf) <= limit: return cand + suf
                if len(cand) <= limit: return cand
        return plain
    def prefix(o):
        ow = o.replace('–', ' – ').split(); k = 0
        while k < min(len(ws), len(ow)) and ws[k] == ow[k]: k += 1
        return k
    pre = max(prefix(o) for o in amb)
    if pre < 3: return plain
    dist = ws[pre:]
    if dist and re.match(r'^\d+\.\d+', dist[0]): return plain
    while dist and dist[0].lower().rstrip(',') in CONN: dist = dist[1:]
    if not dist: return plain
    core = _core(ws[:pre])
    if not core or core[0].lower().rstrip(',') in GENERIC:
        cand = ' '.join(dist).replace(' – ', '–'); cand = cand[0].upper() + cand[1:]
    else: cand = ' '.join(core + dist).replace(' – ', '–')
    if len(cand) < 12:
        tail = ' '.join(dist)
        cand = _cut(' '.join(ws[:pre]), limit - len(suf) - len(tail) - 1) + ' ' + tail
    if len(cand) + len(suf) <= limit: return cand + suf
    return _cut(cand, limit)
def compact_all(names, limit, generic=False):
    out = [compact(n, names, limit, generic) for n in names]
    for i, o in enumerate(out):
        if out.count(o) > 1 and names.count(names[i]) == 1: out[i] = _cut(names[i], limit)
    return out
def capacity(w, h, size):
    cpl = max(8, int((w - 12) / (size * 1.333 * 0.46))); lines = max(1, int((h - 2) / (size * 1.333 * 1.17)))
    return cpl * lines - 4
