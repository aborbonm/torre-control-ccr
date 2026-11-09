import os, glob as _glob
HERE = os.path.dirname(os.path.abspath(__file__))
WORK = os.environ.get('RGC_WORK', os.path.join(HERE, 'work'))
BASE = os.environ.get('RGC_BASE', '/mnt/user-data/uploads/torre-control-ccr')
import datetime as _dt
import json as _json
_AG = os.path.join(WORK, 'agenda.json')
CORTE = _dt.datetime.fromisoformat(os.environ.get('RGC_CORTE') or (_json.load(open(_AG)).get('corte') if os.path.exists(_AG) else '2026-11-03'))
import sys as _s; _s.path.insert(0, HERE); import diag
import sys, json, re, copy, math, subprocess, shutil, zipfile, os, datetime as dt
from pptx import Presentation
from pptx.util import Emu

G = WORK
TEMPLATE = os.environ.get('RGC_TEMPLATE', os.path.join(HERE, 'template_RGC.pptx'))
SKILL = os.environ.get('PPTX_SCRIPTS') or next(iter(sorted(_glob.glob('/mnt/skills/public/pptx/scripts') + _glob.glob('/root/.claude/skills/**/pptx/scripts', recursive=True))), '')
DATA = json.load(open(f'{G}/data.json')); CI = json.load(open(f'{G}/ci.json'))
PMR = json.load(open(f'{G}/pmr.json')); OPS = json.load(open(f'{G}/ops.json'))
RAW_NOTE = ('\n\nDatos del Raw Data del Portal del Cliente: si algo no parece estar bien (montos, fechas, estados, enmiendas o '
            'vinculación de procesos), debe corregirse en el Portal del Cliente para que esta lámina lo presente correctamente '
            'la próxima vez que se genere la presentación.')
RAW_SLIDES = {5, 8, 11, 14, 15}
MANUAL = {k: {a: b for a, b in v.items() if not a.startswith('_')} for k, v in json.load(open(os.path.join(WORK, 'agenda.json'))).get('manual', {}).items()}
PX = 12192000 / 1067
NS = {'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'}
MES = dict(ENE=1, FEB=2, MAR=3, ABR=4, MAY=5, JUN=6, JUL=7, AGO=8, SEP=9, OCT=10, NOV=11, DIC=12)

AGENDA = json.load(open(os.path.join(WORK, 'agenda.json')))
META = AGENDA['meta']

# ---------- formatting ----------
def num(x, d=1):
    s = f'{x:,.{d}f}'
    return s.replace(',', 'X').replace('.', ',').replace('X', '.')
def mm(x, d=1): return f'USD {num(x / 1e6, d)} MM'
def usd(x): return 'USD ' + (num(x, 0) if x >= 1000 else num(x, 0))
def pct(x): return f'{round(x)} %'
def pdate(s):
    try:
        d, m, y = s.split('/'); return dt.datetime(int(y), MES[m.upper()[:3]], int(d))
    except Exception: return None
def short(s, n):
    s = (s or '').strip()
    return s if len(s) <= n else s[:n - 1].rsplit(' ', 1)[0] + '…'
def clean_num(v):
    if isinstance(v, (int, float)):
        return num(v, 2).rstrip('0').rstrip(',') if isinstance(v, float) else num(v, 0)
    return str(v) if v is not None else '—'

# ---------- low level ----------
def shp(slide, sid):
    for s in slide.shapes:
        if s.shape_id == sid: return s
    raise KeyError(sid)
def set_par(par, text):
    runs = par.runs
    if not runs:
        epr = par._p.find('a:endParaRPr', NS)
        r = par.add_run(); r.text = text
        if epr is not None:
            rpr = r._r.get_or_add_rPr()
            for k, v in epr.attrib.items(): rpr.set(k, v)
            for ch in epr: rpr.append(copy.deepcopy(ch))
        return
    runs[0].text = text
    for r in runs[1:]: r._r.getparent().remove(r._r)
def set_text(slide, sid, *pars):
    tf = shp(slide, sid).text_frame
    ps = tf.paragraphs
    for i, t in enumerate(pars):
        if t is None: continue
        if i < len(ps): set_par(ps[i], t)
def set_run(slide, sid, p, r, text):
    shp(slide, sid).text_frame.paragraphs[p].runs[r].text = text
def set_cell(cell, text):
    tf = cell.text_frame
    for extra in tf.paragraphs[1:]: extra._p.getparent().remove(extra._p)
    set_par(tf.paragraphs[0], text)
def set_w(slide, sid, frac, full):
    s = shp(slide, sid); s.width = Emu(max(1, int(full * PX * max(0.0, min(1.0, frac)))))
def remove(slide, sid):
    s = shp(slide, sid); s._element.getparent().remove(s._element)

def table_rows(gf, data, tmpl_idx=1, keep_header=True):
    """Replace data rows of table graphicFrame with `data` (list of lists or (list, style_row_idx))."""
    tbl = gf._element.find('.//a:tbl', NS)
    trs = tbl.findall('a:tr', NS)
    templates = [copy.deepcopy(t) for t in trs]
    for t in trs[1:]: tbl.remove(t)
    for item in data:
        vals, ti = (item if isinstance(item, tuple) else (item, tmpl_idx))
        tr = copy.deepcopy(templates[min(ti, len(templates) - 1)])
        tbl.append(tr)
    t = gf.table
    for ri, item in enumerate(data, 1):
        vals = item[0] if isinstance(item, tuple) else item
        for ci, v in enumerate(vals): set_cell(t.cell(ri, ci), v)
def est_h(gf, data, fs=12):
    """Estimate rendered table height in px."""
    cw = [c.width / PX for c in gf.table.columns]
    charw = fs * 0.50 * 1.11; lh = fs * 1.4 * 1.11
    def rh(vals):
        return max(math.ceil(max(1, len(str(v))) * charw / max(20, w - 12)) for v, w in zip(vals, cw)) * lh + 11
    hdr = [''.join(c.text_frame.text) for c in gf.table.rows[0].cells]
    return rh(hdr), [rh(item[0] if isinstance(item, tuple) else item) for item in data]
def paginate(gf, data, avail, fs=12, maxrows=99):
    h0, hs = est_h(gf, data, fs)
    pages, cur, tot = [], [], h0
    for item, h in zip(data, hs):
        if cur and (tot + h > avail or len(cur) >= maxrows):
            pages.append(cur); cur, tot = [], h0
        cur.append(item); tot += h
    pages.append(cur)
    return pages
def fit_bg(slide, gf, bg_sid, data, fs=12):
    h0, hs = est_h(gf, data, fs)
    h = h0 + sum(hs)
    try:
        b = shp(slide, bg_sid); b.height = Emu(int((h + 2) * PX))
    except KeyError: pass
    return gf.top / PX + h
def font_size(gf, pt):
    for r in gf._element.iter('{%s}rPr' % NS['a']): r.set('sz', str(int(pt * 100)))
    for r in gf._element.iter('{%s}endParaRPr' % NS['a']): r.set('sz', str(int(pt * 100)))
    for p in gf._element.iter('{%s}spcPts' % NS['a']): p.set('val', str(int(pt * 140)))

# ---------- data helpers ----------
def build(op):
    d = DATA[op]; ci = CI.get(op); m = META[op]
    loans = ci['loans'] if ci else []
    MON = {v: k for k, v in MES.items()}
    def ddm(x):
        try: d_, m_, y_ = x.split('/'); return f'{d_}/{MON[int(m_)]}/{y_}'
        except Exception: return None
    for L in loans:
        o = OPS.get(L['prestamo'])
        if not o: continue
        L['ops'] = o
        if o.get('disb') is not None: L['desemb'] = o['disb']; L['pendiente'] = o['undisb']
        if L.get('anual') and o.get('ytd') is not None: L['anual_desemb'] = o['ytd']; L['anual_pend'] = max(0, L['anual'] - o['ytd'])
        if ddm(o.get('exp') or ''): L['ultimo'] = ddm(o['exp'])
        if o.get('min_pend') is not None: L['min_desemb'] = o['min_pend']
        if L.get('antic') and o.get('ant'): L['antic'][0] = o['ant']; L['antic'][1] = ddm(o['ant_date']) or L['antic'][1]
    # manual overrides confirmed by the team (not yet in CI/OPS1)
    for L in loans:
        mo = MANUAL.get(L['prestamo'])
        if mo:
            L.update(mo)
            if pdate(L['ultimo']):
                L['pend'] = max(0, round((pdate(L['ultimo']) - CORTE).days / 30.44)); L['plazo'] = L['transc'] + L['pend']
    res = d.get('resultados', []); prods = d.get('productos', []); pr = d.get('prod_res', [])
    acu = d.get('acuerdos', []); crit = d.get('crit', []); venc = d.get('vencidos', [])
    tot = lambda k: sum(l[k] or 0 for l in loans)
    prod_inds = {}
    for x in pr: prod_inds.setdefault(x['prod'], []).append(x['ind_id'])
    prod_adq = {}
    for x in d.get('prod_adq', []):
        prod_adq.setdefault(x['prod'], []).append(x)
    S = dict(pmr=PMR.get(op, {}), op=op, d=d, ci=ci, m=m, loans=loans, res=res, prods=prods, prod_inds=prod_inds, prod_adq=prod_adq,
             acu=acu, crit=crit, venc=venc,
             vig=tot('vigente'), des=tot('desemb'), pend=tot('pendiente'), anual=tot('anual'), anual_des=tot('anual_desemb'), anual_pend=tot('anual_pend'))
    return S

def pmr_class(d):
    h = [c for y, c in d.get('pmr_hist', []) if c]
    return h[-1] if h else None

# ---------- slide fillers ----------
def f_cover(s, S, pg):
    m, d, ci = S['m'], S['d'], S['ci']
    title = m.get('nombre') or d.get('nombre')
    set_text(s, 4, title)
    pt = 40 if len(title) <= 45 else 34 if len(title) <= 90 else 30
    for r in shp(s, 4).text_frame.paragraphs[0].runs: r.font.size = Emu(pt * 12700)
    lines = math.ceil(len(title) * pt * 1.11 * 0.55 / 952)
    t4 = shp(s, 4); t4.top = Emu(int(150 * PX))
    shp(s, 5).top = Emu(int((150 + lines * pt * 1.11 * 1.5 + 20) * PX))
    set_text(s, 5, m['linea'])
    # three columns re-laid out so long names fit
    for sid, x, w in ((6, 71, 250), (7, 71, 250), (8, 341, 380), (9, 341, 380), (10, 741, 200), (11, 741, 200)):
        sh = shp(s, sid); sh.left = Emu(int(x * PX)); sh.width = Emu(int(w * PX))
    set_text(s, 7, m['je']); set_text(s, 9, ci['oe'] if ci else '[Organismo ejecutor]'); set_text(s, 11, '03/nov/2026')

def f_general(s, S, pg):
    loans = S['loans']; d = S['d']
    set_text(s, 6, mm(S['vig'], 0 if S['vig'] >= 1e8 else 1) if S['vig'] >= 1e6 else usd(S['vig']))
    if len(shp(s, 6).text_frame.text) >= 10:
        for r in shp(s, 6).text_frame.paragraphs[0].runs: r.font.size = Emu(24 * 12700)
    set_text(s, 7, S['m']['tipo'])
    if loans and loans[0]['plazo']:
        tp = loans[0]['transc'] / loans[0]['plazo'] * 100
        set_text(s, 10, pct(tp)); set_w(s, 12, tp / 100, 193)
    else:
        set_text(s, 10, 'n.d.'); set_w(s, 12, 0, 193)
    dp = S['des'] / S['vig'] * 100 if S['vig'] else 0
    set_text(s, 15, pct(dp)); set_w(s, 17, dp / 100, 193)
    set_text(s, 20, S['ci']['pmr'] if S['ci'] else 'n.d.')
    cls = pmr_class(d)
    set_text(s, 21, f'{cls} (ciclo 2025)' if cls else 'Sin clasificación PMR')
    if S['anual']:
        set_text(s, 45, pct(S['anual_des'] / S['anual'] * 100)); set_text(s, 47, f"{mm(S['anual_des'])} de {mm(S['anual'])}")
    else:
        set_text(s, 45, 'Sin meta'); set_text(s, 47, 'Proyección anual 2026 en cero en CI')
    acu = S['acu']
    if acu:
        cum = sum(1 for a in acu if a['estado'] == 'Cumplidos')
        vig = sum(1 for a in acu if a['estado'] == 'Vigente'); pa = sum(1 for a in acu if 'aprobaci' in str(a['estado']))
        set_text(s, 53, f'{cum} de {len(acu)}'); set_text(s, 55, f"{vig} {'vigente' if vig == 1 else 'vigentes'} · {pa} {'pendiente' if pa == 1 else 'pendientes'} de aprobación")
    else:
        set_text(s, 53, 'n.a.'); set_text(s, 55, 'Sin compromisos de la I RGC 2026')
    res = S['res']
    if res:
        mb = sum(1 for r in res if r['prob'] in ('Media', 'Baja'))
        set_text(s, 61, f'{mb} de {len(res)}')
    else: set_text(s, 61, 'n.d.')
    pm = S['pmr'].get('prod_match') or {}
    logros = [f"Logrado: {diag.compact(k, list(pm), 40)} (PMR 100 %)" for k, v in pm.items() if v['pct'] >= 100]
    logros += [f"Logrado: {short(p['prod'], 58)}" for p in S['prods'] if p['estado'] == 'Logrado' and p['prod'] not in pm]
    logros = list(dict.fromkeys(logros))
    for sid, t in zip((25, 27, 29), logros[:3]): set_text(s, sid, t)
    ret = sum(1 for p in S['prods'] if p['estado'] == 'Retrasado')
    set_text(s, 63, (f"{ret} {'producto crítico retrasado' if ret == 1 else 'productos críticos retrasados'}") if S['prods'] else 'Sin productos críticos registrados')

def f_estado(s, S, pg):
    L = S['loans'][pg]; op = S['op']
    kind = 'Donación' if ('GR' in L['prestamo'] or 'ATN' in L['prestamo']) else 'Préstamo'
    set_text(s, 3, f"{kind} {L['prestamo']}: plazo y avance financiero")
    for sid, k in ((6, 'aprob'), (7, 'efect'), (8, 'eleg'), (9, 'primer')): set_text(s, sid, None, L[k])
    orig = ''
    if L['ext'] and pdate(L['ultimo']):
        u = pdate(L['ultimo']); mo = u.month - L['ext']; y = u.year + (mo - 1) // 12; mo = (mo - 1) % 12 + 1
        orig = f" · original {u.day:02d}/{[k for k, v in MES.items() if v == mo][0]}/{y}"
    set_text(s, 10, None, f"{L['ultimo']} (vigente)", f"Extensión: {L['ext']} meses{orig}")
    set_text(s, 12, f"Plazo de ejecución · {L['plazo']} meses")
    tp = L['transc'] / L['plazo'] if L['plazo'] else 0
    set_w(s, 14, tp, 298)
    p15 = shp(s, 15).text_frame.paragraphs[0].runs
    p15[1].text = f"{L['transc']} meses ({round(tp * 100)} %)"; p15[3].text = f"{L['pend']} meses"
    set_text(s, 16, f"Avance financiero · {mm(L['vigente']) if L['vigente'] >= 1e6 else usd(L['vigente'])}")
    set_w(s, 18, L['pct'] / 100, 298)
    tf = shp(s, 19).text_frame
    tf.paragraphs[0].runs[1].text = f"{mm(L['desemb']) if L['vigente'] >= 1e6 else usd(L['desemb'])} ({L['pct']} %)"
    tf.paragraphs[1].runs[1].text = mm(L['pendiente']) if L['vigente'] >= 1e6 else usd(L['pendiente'])
    h = [(y, c) for y, c in S['d'].get('pmr_hist', []) if c][-4:]
    ab = {'Satisfactorio': 'S', 'Alerta': 'A', 'Problema': 'P'}
    set_text(s, 20, None, ' · '.join(f'{y} {ab.get(c, c)}' for y, c in h)  if h else 'Sin historial PMR')
    a = L['amb']
    set_text(s, 21, 'Ambiental y social (categoría · riesgo · desempeño)', f'{a[0]} · {a[1]} · {a[2]}' if a else 'No aplica para esta operación')
    if L.get('antic'):
        an = L['antic']
        set_text(s, 24, None, f"{mm(an[0])} · {an[1]}", f"Plan {an[2]} días: vence {an[3]}")
        set_text(s, 25, None, f"{mm(L['justif'][0])} · {L['justif'][1]}")
        set_text(s, 26, None, f"{mm(L['saldo'])} · {L['pct_just']} %")
    else:
        set_text(s, 24, None, 'Sin anticipo vigente', ''); set_text(s, 25, None, '—'); set_text(s, 26, None, '—')
    set_text(s, 27, None, mm(L['min_desemb']) if L['min_desemb'] >= 1e5 else usd(L['min_desemb']))
    hb = L.get('habilitado')
    set_text(s, 28, 'Nuevo anticipo: ' + ('Habilitado' if hb else 'No habilitado' if hb is False else 'n.d.'))
    set_text(s, 29, 'Fuente: Cartera Inteligente, informe extendido al 03/11/2026 (p. 1 y 2)')

def f_financiera(s, S, pg):
    loans = S['loans']; d = S['d']
    L0 = loans[0]
    ult = pdate(L0['ultimo']); prim = pdate(L0['primer'])
    years_left = (ult - CORTE).days / 365.25 if ult else None
    years_done = (CORTE - prim).days / 365.25 if prim else None
    pend = S['pend']; des = S['des']
    per = 'año' if (years_left or 0) >= 1 else 'mes'; f = 1 if per == 'año' else 12
    if years_left is not None and 0 < years_left < 0.25 and pend:
        set_text(s, 7, mm(pend)); set_text(s, 8, f'por desembolsar antes del {L0["ultimo"]}'); S["short_term"] = f"{mm(pend)} por desembolsar antes del {L0['ultimo']}"
        set_text(s, 9, 'Plazo vigente por vencer: ritmo no comparable con el histórico'); req = None
        years_done = None
    elif years_left and years_left > 0 and pend:
        req = pend / years_left / f; set_text(s, 7, f'{mm(req)}/{per}')
    else:
        req = None; set_text(s, 7, 'Sin saldo o plazo vencido' if not pend else 'n.d.')
    if years_done and years_done > 0.5 and des:
        hist = des / years_done / f
        set_text(s, 9, f'Histórico: {mm(hist)}/{per}')
        set_text(s, 8, f'≈ {num(req / hist, 1)} veces el histórico' if req else 'Sin ritmo requerido')
        if req: S['ratio'] = req / hist
    elif not (years_left is not None and 0 < years_left < 0.25 and pend):
        set_text(s, 9, 'Histórico: sin desembolsos suficientes'); set_text(s, 8, 'Sin histórico comparable')
    # component table
    gf = shp(s, 16)
    comps = []
    for L in loans:
        for c in L['comp']: comps.append(c)
    contr = d.get('contratado_comp', {})
    def match(name):
        import difflib
        best, bv = None, 0
        for k in contr:
            if k == 'Otros costos': continue
            r = difflib.SequenceMatcher(None, name.lower(), k.lower()).ratio()
            if r > bv: best, bv = k, r
        return contr[best] if best and bv > 0.6 else None
    rows = []; tj = tt = 0
    for name, j, p in comps:
        total = j + p; c = match(name)
        porpagar = f'{num(max(0, c - j) / 1e6)}' if c is not None else '[__]'
        porcomp = f'{num(max(0, total - max(c, j)) / 1e6)}' if c is not None else '[__]'
        rows.append([short(name, 60), num(j / 1e6), porpagar, porcomp, num(total / 1e6)])
        tj += j; tt += total
    if rows:
        rows.append(([ 'Total', num(tj / 1e6), '[__]', '[__]', num(tt / 1e6)], 4))
    else:
        rows = [['Sin información por componente en CI', '—', '—', '—', '—']]
    table_rows(gf, rows); fit_bg(s, gf, 15, rows)
    adq = [L['adq'] for L in loans if L.get('adq')]
    if adq and adq[0]['contratado'] >= des:
        porc = max(0, S['vig'] - sum(a['contratado'] for a in adq))
        set_text(s, 19, f"{pct(porc / S['vig'] * 100)} · {mm(porc)}")
    else:
        set_text(s, 19, '[__] % · USD [__] MM')
    set_text(s, 22, f"{pct(pend / S['vig'] * 100) if S['vig'] else 'n.d.'} · {mm(pend)}")

def chart_into(s, S, box_sid, remove_sids):
    for sid in remove_sids: remove(s, sid)
    box = shp(s, box_sid)
    codes = [L['prestamo'][:4] for L in S['loans'] if os.path.exists(f"{G}/charts/{L['prestamo'][:4]}.png")]
    if not codes: return False
    bx, by, bw, bh = box.left, box.top, box.width, box.height
    pad = int(10 * PX)
    n = len(codes)
    for i, c in enumerate(codes):
        w = (bw - pad * (n + 1)) // n; h = int(w * 812 / 1224)
        if h > bh - 2 * pad: h = bh - 2 * pad; w = int(h * 1224 / 812)
        x = bx + pad + i * ((bw - pad) // n); y = by + (bh - h) // 2
        s.shapes.add_picture(f'{G}/charts/{c}.png', x, y, w, h)
    return True

def f_desemb(s, S, pg):
    ok = chart_into(s, S, 4, (5, 6, 7))
    if S['anual']:
        set_text(s, 10, mm(S['anual']))
        set_text(s, 13, f"{mm(S['anual_des'])} · {pct(S['anual_des'] / S['anual'] * 100)}")
    else:
        set_text(s, 10, 'Sin meta registrada'); set_text(s, 13, f"{mm(S['anual_des'])}")
    set_text(s, 12, 'Desembolsado al 03/nov/2026')
    prox = [L['proximo'] for L in S['loans'] if L['proximo'] not in ('N/A', '-')]
    set_text(s, 16, f"{mm(S['anual_pend'])} · {prox[0] if prox else 'sin fecha'}")
    if S['anual'] and S['anual_des'] >= S['anual']:
        set_text(s, 19, 'Meta 2026 cumplida según Cartera Inteligente. [Confirmar si se prevén desembolsos adicionales antes del cierre].')
    else:
        set_text(s, 19, '[Se cumple la meta / Se reduce a USD __ MM, porque …]. Si no se cumple: ver la ruta al siguiente desembolso.')

def f_ruta(s, S, pg):
    Ls = [L for L in S['loans'] if L.get('antic')]
    L = Ls[0] if Ls else S['loans'][0]
    if len(S['loans']) > 1: set_text(s, 4, f"Solo para operaciones que no han alcanzado la meta de desembolsos 2026 · {L['prestamo']}")
    set_text(s, 7, mm(L['min_desemb']) if L['min_desemb'] >= 1e5 else usd(L['min_desemb']))
    if L.get('justif'):
        set_text(s, 9, f"Justificado al {L['justif'][1]}"); set_text(s, 10, mm(L['justif'][0]))
    o = L.get('ops') or {}
    if L.get('antic'): set_text(s, 6, f"Último anticipo ({L['antic'][1]})"); set_text(s, 7, mm(L['antic'][0]))
    if o: set_text(s, 12, 'Brecha por justificar (OPS1)')
    set_text(s, 13, (mm(L['min_desemb']) if L['min_desemb'] >= 1e5 else usd(L['min_desemb'])) if o else '[Reporte financiero]')
    if o.get('deadline') and re.match(r'\d', o['deadline']):
        set_text(s, 15, 'Vence plan financiero'); set_text(s, 16, o['deadline'])
    set_text(s, 19, 'Fuente: OPS1 (resumen operativo y financiero) y Cartera Inteligente; hitos y pagos: Unidad Ejecutora')

def f_proy(s, S, pg):
    pj = S['d'].get('proy2027'); gf = shp(s, 5)
    if pj:
        rows = []
        for k in range(4):
            allc = [x[0] for t in pj['top'] for x in t]
            top = '; '.join(diag.compact(x[0], allc, 34, True) for x in pj['top'][k]) if pj['top'][k] else '—'
            rows.append([f'T{k + 1}', num(pj['q'][k] / 1e6), top, '[Supuesto]'])
        totq = sum(pj['q'])
        rows.append(([ 'Total*', num(totq / 1e6), f"{pct(totq / S['pend'] * 100) if S['pend'] else 'n.d.'} del saldo por desembolsar", ''], 5))
        table_rows(gf, rows); bottom = fit_bg(s, gf, 4, rows, 12)
        shp(s, 6).top = Emu(int(min(bottom + 24, 500) * PX))
        saldo = max(0, S['pend'] - S['anual_pend'] - totq)
        s10 = shp(s, 10).text_frame.paragraphs[0].runs; s10[1].text = mm(saldo)
    else:
        set_text(s, 6, '* Sin estimado: la operación no tiene Raw Data del plan de adquisiciones en el Portal del Cliente. La Unidad Ejecutora completa la proyección.')
    ult = [L['ultimo'] for L in S['loans']]
    shp(s, 9).text_frame.paragraphs[0].runs[1].text = ' y '.join(dict.fromkeys(ult)) if ult else '[dd/mmm/aaaa]'
    u = pdate(ult[0]) if ult else None
    if u and u < dt.datetime(2027, 12, 31):
        set_text(s, 11, f'No: el plazo vence el {ult[0]}')
    elif u:
        set_text(s, 11, 'Sí, con el plazo vigente')

def estado_acu(a):
    e = a['estado']
    if e == 'Cumplidos': return 'Cumplido'
    if 'aprobaci' in str(e): return 'En proceso (pendiente de aprobación)'
    v = pdate(str(a['venc']).replace('/', '/').upper()) if a['venc'] else None
    return 'Pendiente (vencido)' if v and v < CORTE else 'En proceso'
def acu_rows(S):
    rows = []
    for i, a in enumerate(S['acu'], 1):
        st = estado_acu(a)
        com = a['coment'] if a['coment'] not in (None, '-', '') else None
        if st == 'Cumplido':
            causa, acc = '—', '—'
        else:
            causa = '[Unidad Ejecutora]' + (f' · Último comentario en la Torre: {short(com, 240)}' if com else '')
            nf = a['nueva'] if a['nueva'] not in (None, '-') else None
            acc = '[Acción]' + (f' · nueva fecha {nf}' if nf else ' · [nueva fecha]')
        rows.append([str(i), a['desc'] or '', (a['resp'] or '').title(), a['venc'] or '', st, causa, acc])
    return rows

def res_rows(S):
    rows = []; last = None
    for r in S['res']:
        if r['obj_id'] != last:
            rows.append(([f"OE{r['obj_id']}: {r['obj']}", '', '', '', '', '', '', ''], 1)); last = r['obj_id']
        ind = r['ind'] or ''; unit = '[__]'
        mt = re.search(r'\[([^\]]+)\]\s*$', ind)
        if mt: unit = mt.group(1); ind = ind[:mt.start()].strip()
        unit = {'porcentaje': '%', 'tasa': 'Tasa'}.get(unit, unit)
        pm = S['pmr'].get('ind_match', {}).get(r['ind_id'])
        va, av = '[UE] · sin medición en PMR', '[__] %'
        if pm:
            va = f"{pm['actual']} ({pm['year']}, PMR)"
            try:
                b_, m_, a_ = (float(str(x).replace('%', '').replace(',', '.')) for x in (r['base'], r['meta'], pm['actual']))
                if m_ != b_: av = pct(max(0, (a_ - b_) / (m_ - b_) * 100))
            except Exception: pass
        rows.append(([f"{r['ind_id']} {ind}", unit, clean_num(r['base']), va, clean_num(r['meta']), av, r['prob'] or '—',
                      '[UE]' if r['prob'] in ('Media', 'Baja') else '—'], 2))
    return rows

def prod_rows(S):
    rows = []
    for p in S['prods']:
        inds = ', '.join(sorted(set(S['prod_inds'].get(p['prod'], [])), key=lambda x: [(0, int(t), '') if t.isdigit() else (1, 0, t) for t in re.split(r'\.', x)])) or '—'
        risks = [p['riesgo']] if p['riesgo'] else []
        risks += [x['riesgo'] for x in S['prod_adq'].get(p['prod'], []) if x['riesgo'] and x['critica'] == 'Sí']
        risks = list(dict.fromkeys(risks))
        atr = p['estado'] in ('Retrasado', 'Pausado', 'Cancelado')
        pm = (S['pmr'].get('prod_match') or {}).get(p['prod'])
        af = f"{pm['pct']} % ({clean_num(pm['a'])}/{clean_num(pm['pa'])})" if pm else ('[PMR] %' if S['pmr'].get('outputs') else 'n.d. en PMR')
        rows.append([p['prod'], inds, af, p['estado'] or '—', '; '.join(risks) or ('[Unidad Ejecutora]' if atr else '—'),
                     '[Unidad Ejecutora]' if atr else '—', '[UE] · [fecha]' if atr else '—'])
    return rows

def adq_rows(S):
    rows = []; seen = set()
    for c in S['crit']:
        if c['pid'] in seen: continue
        seen.add(c['pid'])
        if not c.get('nombre'):
            el = c.get('eliminado')
            if el:
                rows.append([f"{c['pid']} · {diag.compact(el['nombre'], [x.get('nombre') or '' for x in S['crit']], 58)}", diag.compact(c['prod'], [q['prod'] for q in S['prods']], 30), '—', '—',
                             f"Retirado del plan ({el['fecha']})", '—', short('Comentario en el Portal: ' + el['com'], 80), '[UE]']); continue
            rows.append([f"{c['pid']} (no figura en el Raw Data)", diag.compact(c['prod'], [q['prod'] for q in S['prods']], 30), '—', '—', '—', '—', short(c['riesgo'], 80) if c['riesgo'] else '[UE]', '[UE]']); continue
        f = f"{c['firma']} ({c['firma_tipo']})" if c['firma'] else '—'
        rows.append([f"{c['pid']} · {diag.compact(c['nombre'], [x.get('nombre') for x in S['crit']], 58)}", diag.compact(c['prod'], [q['prod'] for q in S['prods']], 30), c['metodo'], mm(c['monto'], 2).replace('USD ', ''), c['estado'], f,
                     short(c['riesgo'], 80) if c['riesgo'] else '[UE]', '[UE]'])
    return rows

def f_compromisos(s, S, pg, rows):
    gf = shp(s, 5); set_widths(gf, ACU_W); table_rows(gf, rows); font_size(gf, 11); bottom = fit_bg(s, gf, 4, rows, 11)
    shp(s, 6).top = Emu(int(max(min(bottom + 14, 512), 380) * PX))
    set_text(s, 7, 'Fuente: Torre de Control CCR (tab Acuerdos: acuerdos de la I RGC 2026 y último comentario registrado)')
def f_resultados(s, S, pg, rows, fs):
    gf = shp(s, 5); table_rows(gf, rows)
    if fs != 12: font_size(gf, fs)
    bottom = fit_bg(s, gf, 4, rows, fs)
    shp(s, 6).top = Emu(int(512 * PX))
def f_detalle(s, S, pg, rows):
    gf = shp(s, 5); set_widths(gf, DET_W); table_rows(gf, rows); font_size(gf, 10); bottom = fit_bg(s, gf, 4, rows, 10)
    shp(s, 6).top = Emu(int(512 * PX))
def set_widths(gf, fr):
    tot = sum(c.width for c in gf.table.columns)
    for c, f in zip(gf.table.columns, fr): c.width = Emu(int(tot * f))
ACU_W = [0.04, 0.27, 0.10, 0.09, 0.10, 0.26, 0.14]
ADQ_W = [0.25, 0.14, 0.06, 0.08, 0.11, 0.11, 0.17, 0.08]
DET_W = [0.24, 0.10, 0.08, 0.09, 0.21, 0.16, 0.12]
def f_adq(s, S, pg, rows):
    adq = [L['adq'] for L in S['loans'] if L.get('adq')]
    if adq:
        a = adq[0]
        set_text(s, 6, a['clasif']); set_text(s, 9, f"{a['exp_n']} · {usd(a['exp_usd'])}")
        set_text(s, 12, f"{num(a['contratado'] / 1e6)} / {num(a['pendiente'] / 1e6)} MM")
    if rows:
        gf = shp(s, 14); set_widths(gf, ADQ_W); table_rows(gf, rows); font_size(gf, 10); fit_bg(s, gf, 13, rows, 10)
def f_contratos(s, S, pg):
    v = S['venc']
    set_text(s, 6, f"{len(v)} · {mm(sum(x['monto'] for x in v))}")
    set_text(s, 9, '[Unidad Ejecutora]')
    gf = shp(s, 14)
    rows = [[x['cid'], diag.compact(x['nombre'], [y['nombre'] for y in v], 70), 'Vencido', x['fin'], mm(x['monto'], 2).replace('USD ', ''), '[Enmienda de plazo o cierre]', '[UE]'] for x in v[:3]]
    if not rows: rows.append(['—', 'Sin contratos vencidos en el Raw Data', '—', '—', '—', '—', '—'])
    table_rows(gf, rows, tmpl_idx=1)
    tbl = gf._element.find('.//a:tbl', NS)
    font_size(gf, 11)
    bottom = fit_bg(s, gf, 13, rows, 11)
    if len(v) > 3:
        tb = s.shapes.add_textbox(gf.left, Emu(int(468 * PX)), gf.width, Emu(int(24 * PX)))
        p = tb.text_frame.paragraphs[0]; r = p.add_run()
        r.text = f"Se muestran los 3 de mayor monto de {len(v)} contratos con plazo vencido y estado en ejecución en el Raw Data (USD {num(sum(x['monto'] for x in v) / 1e6)} MM en total)."
        r.font.size = Emu(int(12 * 12700)); r.font.name = 'Public Sans'
        from pptx.dml.color import RGBColor
        r.font.color.rgb = RGBColor(0x4A, 0x55, 0x68)
def extension(S):
    """One-line finding on whether an extension looks necessary."""
    Ls = [L for L in S['loans'] if pdate(L.get('ultimo') or '') and L['pendiente'] > 0]
    if not S['loans']: return 'Extensión: no aplica (operación sin datos de ejecución)'
    if not Ls:
        if all(L['pendiente'] <= 0 for L in S['loans']): return 'Extensión: no requerida para desembolsar (100 % desembolsado); confirmar cierre de actividades'
        return 'Extensión: no aplica aún (sin fecha de último desembolso)'
    L = min(Ls, key=lambda L: pdate(L['ultimo']))
    exp = pdate(L['ultimo']); months = (exp - CORTE).days / 30.4; pp = L['pendiente'] / L['vigente'] * 100 if L['vigente'] else 0
    post = [f for f in S['d'].get('fins', []) if dt.datetime.strptime(f['fin'], '%d/%m/%Y') > exp]
    pmr_txt = (S['pmr'].get('status', '') + ' ' + S['pmr'].get('likelihood', '')).lower()
    pmr_flag = bool(re.search(r'prórroga|prorroga|extensi[oó]n del plazo|ampliaci[oó]n del plazo|requerir[aá] (una )?extensi', pmr_txt)) and L['prestamo'] not in MANUAL
    reasons = []
    if months < 12 and pp > 15: reasons.append(f'{round(pp)} % sin desembolsar')
    if post: reasons.append(f"{len(post)} {'contrato termina' if len(post) == 1 else 'contratos terminan'} después")
    if (S.get('ratio') or 0) > 2 and not reasons: reasons.append('ritmo requerido muy superior al histórico')
    if pmr_flag: reasons.append('PMR prevé prórroga')
    if reasons:
        return f"Extensión: parece necesaria. Vence {L['ultimo']}; " + '; '.join(reasons)
    return f"Extensión: no parece necesaria; {max(0, round(months))} meses de plazo y {round(pp)} % por desembolsar"

def f_hallazgos(s, S, pg):
    res = S['res']; prods = S['prods']
    if res:
        b = sum(1 for r in res if r['prob'] == 'Baja'); set_text(s, 8, f'{b} de {len(res)} indicadores con probabilidad baja')
    else: set_text(s, 8, 'Sin indicadores registrados en la Torre')
    ret = sum(1 for p in prods if p['estado'] == 'Retrasado'); pau = sum(1 for p in prods if p['estado'] == 'Pausado')
    pl = lambda n, a, b: f'{n} {a if n == 1 else b}'
    set_text(s, 9, f"{pl(ret, 'producto retrasado', 'productos retrasados')} y {pl(pau, 'pausado', 'pausados')}" if prods else '[Productos cancelados o pausados]')
    if S['crit']:
        sc = sum(1 for c in S['crit'] if c.get('estado') in ('Previsto', 'Proceso en Curso'))
        set_text(s, 13, f'{sc} de {len(set(c["pid"] for c in S["crit"]))} críticos sin contrato')
    adq = [L['adq'] for L in S['loans'] if L.get('adq')]
    if adq: set_text(s, 14, f"{adq[0]['exp_n']} procesos con fechas expiradas")
    if S['d'].get('raw'):
        v = S['venc']; set_text(s, 18, (f"{len(v)} contratos vencidos" if len(v) != 1 else '1 contrato vencido') if v else 'Sin contratos vencidos')
    if S.get('ratio'): set_text(s, 23, f"Ritmo requerido ≈ {num(S['ratio'], 1)} veces el histórico")
    if S['anual']:
        set_text(s, 24, f"Desembolsos 2026: {pct(S['anual_des'] / S['anual'] * 100)} de la meta")
    from pptx.util import Pt
    t23, t24 = shp(s, 23), shp(s, 24)
    lines = []
    if S.get('ratio'): lines.append(f"Ritmo requerido ≈ {num(S['ratio'], 1)} veces el histórico")
    elif S.get('short_term'): lines.append(S['short_term'])
    if S['anual']: lines.append(f"Desembolsos 2026: {pct(S['anual_des'] / S['anual'] * 100)} de la meta")
    lines.append(extension(S))
    ext = copy.deepcopy(t24._element); t24._element.addnext(ext)
    ids = [int(x) for x in s._element.xpath('//p:cNvPr/@id')]
    ext.xpath('./p:nvSpPr/p:cNvPr')[0].set('id', str(max(ids) + 1))
    te = [x for x in s.shapes if x._element is ext][0]
    boxes = [t23, t24, te]
    y = 190
    for sh in boxes: set_par(sh.text_frame.paragraphs[0], '')
    for sh, t in zip(boxes[3 - len(lines):] if len(lines) < 3 else boxes, lines):
        set_par(sh.text_frame.paragraphs[0], t)
        h = math.ceil(len(t) / 33) * 14 + 4
        sh.top = Emu(int(y * PX)); sh.height = Emu(int(h * PX)); y += h + 4
    for sh in boxes:
        for par in sh.text_frame.paragraphs:
            for r in par.runs: r.font.size = Pt(10)
            lsp = par._p.find('a:pPr/a:lnSpc', NS)
            if lsp is not None: lsp.getparent().remove(lsp)

# ---------- notes number remapping ----------
def remap_text(t, M):
    def rng(n):
        lo, hi = M.get(n, (n, n)); return str(lo) if lo == hi else f'{lo} a {hi}'
    def rep(mo):
        word, seq = mo.group(1), mo.group(2)
        parts = re.split(r'( a | o | y |, )', seq); out = []; k = 0
        while k < len(parts):
            n = int(parts[k])
            if k + 2 < len(parts) and parts[k + 1] == ' a ':
                b = int(parts[k + 2]); out.append(f"{M.get(n, (n, n))[0]} a {M.get(b, (b, b))[1]}"); k += 3
            else: out.append(rng(n)); k += 1
            if k < len(parts): out.append(parts[k]); k += 1
        res = ''.join(out)
        if word.lower() == 'lámina' and (' a ' in res): word = word + 's'
        return f'{word} {res}'
    return re.sub(r'([Ll]áminas?) (\d+(?:(?: a | o | y |, )\d+)*)', rep, t)

# ---------- main ----------
def generate(op, outpath):
    S = build(op)
    tmp = f'{G}/tmp_{op}'; shutil.rmtree(tmp, ignore_errors=True)
    zipfile.ZipFile(TEMPLATE).extractall(tmp)
    prs0 = Presentation(TEMPLATE)
    # page plans
    plan = {i: 1 for i in range(1, 22)}
    plan[4] = max(1, len(S['loans']))
    acur = acu_rows(S); resr = res_rows(S); prr = prod_rows(S); adr = adq_rows(S)
    s0 = list(prs0.slides)
    pages = {}
    if acur:
        g9 = shp(s0[8], 5); set_widths(g9, ACU_W); pages[9] = paginate(g9, acur, 350, 11)
    if resr:
        fs = 10 if len(resr) > 14 else 12
        pg10 = paginate(shp(s0[9], 5), resr, 360 if fs == 10 else 345, fs)
        for i in range(len(pg10) - 1):
            if pg10[i] and isinstance(pg10[i][-1], tuple) and pg10[i][-1][1] == 1:
                pg10[i + 1].insert(0, pg10[i].pop())
        pages[10] = [x for x in pg10 if x]; pages['fs10'] = fs
    if prr:
        g13 = shp(s0[12], 5); set_widths(g13, DET_W); pages[13] = paginate(g13, prr, 335, 10)
    if adr:
        g14 = shp(s0[13], 14); set_widths(g14, ADQ_W); pages[14] = paginate(g14, adr, 290, 10, maxrows=3)
    objs = list(dict.fromkeys(r['obj_id'] for r in S['res']))
    dpages = diag.plan(S); plan[11] = max(1, len(dpages))
    # matrix pages
    inds = [r['ind_id'] for r in S['res']]
    by_obj = len(inds) > 10
    linked = {x for p in S['prods'] for x in S['prod_inds'].get(p['prod'], [])}
    if by_obj:
        cols = [[r['ind_id'] for r in S['res'] if r['obj_id'] == o] for o in objs]
        cols = [c for c in cols if set(c) & linked]
    else: cols = [[x] for x in inds if x in linked]
    colpages = [cols[i:i + 5] for i in range(0, len(cols), 5)] or [[]]
    mpages = []
    for cols in colpages:
        allids = set(x for c in cols for x in c)
        prs_ = [p for p in S['prods'] if set(S['prod_inds'].get(p['prod'], [])) & allids]
        for i in range(0, max(1, len(prs_)), 8): mpages.append((cols, prs_[i:i + 8]))
    plan[12] = max(1, len(mpages))
    for k in (9, 10, 13, 14):
        if k in pages: plan[k] = len(pages[k])
    # duplicate slides
    created = {}
    for k, n in plan.items():
        for _ in range(n - 1):
            out = subprocess.run([sys.executable, f'{SKILL}/add_slide.py', tmp, f'slide{k}.xml', '--after', f'slide{k}.xml'],
                                 capture_output=True, text=True, cwd=SKILL)
            mt = re.search(r'(slide\d+\.xml)', out.stdout.split('Created')[-1])
            created[mt.group(1)] = k
    pres_xml = open(f'{tmp}/ppt/presentation.xml').read(); rels_xml = open(f'{tmp}/ppt/_rels/presentation.xml.rels').read()
    rid2t = dict(re.findall(r'Id="(rId\d+)"[^>]*Target="slides/(slide\d+\.xml)"', rels_xml))
    rid2t.update({a: b for b, a in re.findall(r'Target="slides/(slide\d+\.xml)"[^>]*Id="(rId\d+)"', rels_xml)})
    order = [rid2t[r] for r in re.findall(r'<p:sldId [^>]*r:id="(rId\d+)"', pres_xml)]
    tidx = [created.get(n, int(re.search(r'(\d+)', n).group(1))) for n in order]
    pre = f'{G}/pre_{op}.pptx'
    if os.path.exists(pre): os.remove(pre)
    subprocess.run(f'cd "{tmp}" && zip -qXr "{pre}" .', shell=True, check=True)
    prs = Presentation(pre)
    slides = list(prs.slides)
    # map template number -> (first, last) position
    M = {}
    for pos, k in enumerate(tidx, 1):
        lo, hi = M.get(k, (pos, pos)); M[k] = (min(lo, pos), max(hi, pos))
    seen = {}
    notes_src = {i + 1: s.notes_slide.notes_text_frame.text for i, s in enumerate(s0) if s.has_notes_slide}
    hide = set()
    if not S['loans']: hide |= {3, 4, 5, 6, 7, 8}
    if S['loans'] and S['des'] == 0 and not S['anual']: hide |= {5, 6, 7}
    if S['anual'] and S['anual_des'] >= S['anual'] - 10 and all((L['anual_desemb'] >= L['anual'] - 10) for L in S['loans']): hide.add(7)
    if not S['anual'] and not any(L.get('antic') for L in S['loans']): hide.add(7)
    if not S['anual'] and not any(os.path.exists(f"{G}/charts/{L['prestamo'][:4]}.png") for L in S['loans']): hide.add(6)
    if not S['acu']: hide.add(9)
    if not S['res']: hide |= {10, 11, 12}
    if not S['prods']: hide.add(13)
    if not S['crit'] and not any(L.get('adq') for L in S['loans']): hide.add(14)
    if not S['d'].get('raw'): hide.add(15)
    if not S['ci'] and not S['res'] and not S['prods']: hide |= {3, 13, 14, 15}  # sin CI ni datos en la Torre
    for pos, (s, k) in enumerate(zip(slides, tidx), 1):
        pg = seen.get(k, 0); seen[k] = pg + 1
        # footer number and notes
        for sh in s.shapes:
            if sh.has_text_frame:
                for par in sh.text_frame.paragraphs:
                    t = ''.join(r.text for r in par.runs).strip()
                    if re.fullmatch(r'\[CR-L_+\] · \d+', t): set_par(par, f'{op} · {pos}')
        if k in notes_src:
            ns = s.notes_slide
            if ns.notes_text_frame is None:
                orig = s0[k - 1].notes_slide
                new_sp = copy.deepcopy(orig._element.cSld.spTree)
                old_sp = ns._element.cSld.spTree
                for ch in list(old_sp): old_sp.remove(ch)
                for ch in list(new_sp): old_sp.append(ch)
                if ns.notes_text_frame is None: print('WARN notes', k, [x.name for x in ns.shapes])
            s.notes_slide.notes_text_frame.text = remap_text(notes_src[k], M) + (RAW_NOTE if k in RAW_SLIDES else '')
        if k in hide or k >= 19: s._element.set('show', '0')
        if k in hide: continue
        n = plan[k]
        def suffix():
            if n > 1 and k not in (4, 11):
                t = shp(s, 3).text_frame.paragraphs[0]; set_par(t, ''.join(r.text for r in t.runs) + f' ({pg + 1} de {n})')
        if k == 1: f_cover(s, S, pg)
        elif k == 3: f_general(s, S, pg)
        elif k == 4: f_estado(s, S, pg) if S['loans'] else None
        elif k == 5: f_financiera(s, S, pg)
        elif k == 6: f_desemb(s, S, pg)
        elif k == 7: f_ruta(s, S, pg)
        elif k == 8: f_proy(s, S, pg)
        elif k == 9: f_compromisos(s, S, pg, pages[9][pg]); suffix()
        elif k == 10: f_resultados(s, S, pg, pages[10][pg], pages['fs10']); suffix()
        elif k == 11: f_diagrama(s, S, pg, dpages[pg])
        elif k == 12: f_matriz(s, S, pg, mpages[pg][0], by_obj, objs, mpages[pg][1]); suffix()
        elif k == 13: f_detalle(s, S, pg, pages[13][pg]); suffix()
        elif k == 14: f_adq(s, S, pg, pages.get(14, [[]])[pg]); suffix()
        elif k == 15: f_contratos(s, S, pg)
        elif k == 17: f_hallazgos(s, S, pg)
        elif k >= 19: f_guide(s, M)
    for sl in prs.slides:
        for bp in sl._element.iter('{%s}bodyPr' % NS['a']):
            for af in bp.findall('a:normAutofit', NS): bp.remove(af)
    prs.save(outpath)
    shutil.rmtree(tmp, ignore_errors=True); os.remove(pre)
    return M, hide

def f_guide(s, M):
    gf = [sh for sh in s.shapes if sh.has_table][0]
    for row in list(gf.table.rows)[1:]:
        c = row.cells[0]; t = c.text_frame.text
        mt = re.match(r'Lámina (\d+) · (.*)', t)
        if mt:
            k = int(mt.group(1)); lo, hi = M.get(k, (k, k))
            set_cell(c, (f'Lámina {lo}' if lo == hi else f'Láminas {lo} a {hi}') + ' · ' + mt.group(2))
        rest = remap_text(c.text_frame.text, M)
    for row in list(gf.table.rows)[1:]:
        for c in list(row.cells)[1:]:
            t = c.text_frame.text; nt = remap_text(t, M)
            if nt != t: set_cell(c, nt)

def f_diagrama(s, S, pg, page):
    for sid in range(4, 26):
        try: remove(s, sid)
        except KeyError: pass
    t = shp(s, 3)
    def set_title(a, b):
        p = t.text_frame.paragraphs[0]; set_par(p, a)
        r0 = p.runs[0]; r1 = copy.deepcopy(r0._r); r0._r.addnext(r1)
        from pptx.dml.color import RGBColor
        from pptx.util import Pt
        r0.font.color.rgb = RGBColor.from_string('5BB4E5'); p.runs[1].text = b
        for r in p.runs: r.font.size = Pt(20)
        lsp = p._p.find('a:pPr/a:lnSpc', NS)
        if lsp is not None: lsp.getparent().remove(lsp)
        t.height = Emu(int(56 * PX))
    diag.draw(s, S, page, set_title)

def f_matriz(s, S, pg, cols, by_obj, objs, plist):
    res = S['res']
    b = sum(1 for r in res if r['prob'] == 'Baja')
    set_text(s, 5, f'{b} de {len(res)}')
    r10 = shp(s, 10).text_frame.paragraphs[0].runs
    cnt = {p: sum(1 for r in res if r['prob'] == p) for p in ('Alta', 'Media', 'Baja')}
    r10[1].text = f" {cnt['Alta']} · "; r10[3].text = f" {cnt['Media']} · "; r10[5].text = f" {cnt['Baja']}"
    prob = {r['ind_id']: r['prob'] for r in res}
    gf = shp(s, 12)
    # header labels
    hdr = []
    for c in cols:
        if by_obj:
            o = [r['obj_id'] for r in res if r['ind_id'] == c[0]][0]
            nb = sum(1 for x in c if prob.get(x) == 'Baja')
            hdr.append(f'OE{o} ({len(c)} ind., {nb} baja)')
        else:
            hdr.append(f"{c[0]} ({prob.get(c[0], '—')})")
    t = gf.table
    # drop indicator columns this page does not use, so no column is left empty
    used = [ci for ci, c in enumerate(cols) if any(set(S['prod_inds'].get(p['prod'], [])) & set(c) for p in plist)]
    cols = [cols[ci] for ci in used]; hdr = [hdr[ci] for ci in used]
    tbl = t._tbl; grid = tbl.tblGrid; gcs = list(grid.iterchildren('{%s}gridCol' % NS['a']))
    total = sum(int(g.get('w')) for g in gcs); n0 = len(gcs)
    keep = [0] + [1 + i for i in range(len(cols))] + [n0 - 1]
    for ci in reversed(range(n0)):
        if ci not in keep:
            grid.remove(gcs[ci])
            for tr in tbl.tr_lst: tr.remove(tr.tc_lst[ci])
    gcs = list(grid.iterchildren('{%s}gridCol' % NS['a']))
    w0, wl = int(gcs[0].get('w')), int(gcs[-1].get('w'))
    mid = (total - w0 - wl) // max(1, len(gcs) - 2)
    for g in gcs[1:-1]: g.set('w', str(mid))
    ncol = len(t.columns)
    for ci in range(1, ncol - 1):
        set_cell(t.cell(0, ci), hdr[ci - 1] if ci - 1 < len(hdr) else '')
    allids = set(x for c in cols for x in c)
    rows = []
    color_map = []
    for p in plist:
        pid = set(S['prod_inds'].get(p['prod'], []))
        vals = [diag.compact(p['prod'], [q['prod'] for q in S['prods']], 70)]
        for c in cols: vals.append(p['estado'] if pid & set(c) else '')
        while len(vals) < ncol - 1: vals.append('')
        vals.append(str(len(pid)))
        rows.append(vals)
    if not rows: rows = [['Sin productos vinculados a estos indicadores'] + [''] * (ncol - 1)]
    table_rows(gf, rows)
    COL = {'En tiempo': '2E7D5B', 'Retrasado': 'B4531A', 'Cancelado': 'B42318', 'Pausado': '5B6675', 'Logrado': '1F6B4A'}
    from pptx.dml.color import RGBColor
    for ri, vals in enumerate(rows, 1):
        for ci, v in enumerate(vals):
            for par in t.cell(ri, ci).text_frame.paragraphs:
                for r in par.runs:
                    if v in COL: r.font.color.rgb = RGBColor.from_string(COL[v])
                    else: r.font.color.rgb = RGBColor.from_string('12304F')
    bottom = fit_bg(s, gf, 11, rows)
    shp(s, 13).top = Emu(int(512 * PX))

if __name__ == '__main__':
    os.makedirs(f'{G}/out', exist_ok=True)
    for op in (sys.argv[1:] or AGENDA['operaciones']):
        M, hide = generate(op, f'{G}/out/{op}.pptx')
        print(op, 'ok', 'hidden:', sorted(hide))
