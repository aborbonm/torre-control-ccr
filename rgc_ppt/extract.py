import os, glob as _glob
HERE = os.path.dirname(os.path.abspath(__file__))
WORK = os.environ.get('RGC_WORK', os.path.join(HERE, 'work'))
BASE = os.environ.get('RGC_BASE', '/mnt/user-data/uploads/torre-control-ccr')
import datetime as _dt
import json as _json
_AG = os.path.join(WORK, 'agenda.json')
CORTE = _dt.datetime.fromisoformat(os.environ.get('RGC_CORTE') or (_json.load(open(_AG)).get('corte') if os.path.exists(_AG) else '2026-11-03'))
"""Extract structured data per operation from Torre, agreement_report, PMR_Historico and Raw Data."""
import openpyxl, json, glob, re, datetime as dt
from collections import defaultdict

AGENDA = json.load(open(os.path.join(WORK, 'agenda.json')))
OPS = AGENDA['operaciones']
RAW = {}  # se completa con los préstamos de datos_ops que tienen Raw Data en clientportal/

def rows(wb, s):
    return [r for r in wb[s].iter_rows(values_only=True) if any(c is not None for c in r)]
def D(x):
    if isinstance(x, dt.datetime): return x
    if not x: return None
    try: return dt.datetime.fromisoformat(str(x)[:10])
    except Exception: return None

out = {op: {} for op in OPS}
wb = openpyxl.load_workbook(f'{BASE}/torre_de_control_CCR.xlsx', read_only=True, data_only=True)

# datos_ops
for r in rows(wb, 'datos_ops')[1:]:
    op = r[1]
    if op in out:
        out[op].setdefault('prestamos', []).append(r[2]); out[op]['nombre'] = r[4]; out[op]['corto'] = r[3]
# equipo
eq = rows(wb, 'equipo'); hdr = [c for c in eq[2]]
for r in eq[3:]:
    if r[1] == 'Jefe de Equipo':
        for k, op in enumerate(hdr):
            if op in out and r[k]: out[op]['je'] = r[k]
# resultados
for r in rows(wb, 'resultados')[2:]:
    if r[0] in out:
        out[r[0]].setdefault('resultados', []).append(dict(obj_id=str(r[1]), obj=r[2], ind_id=str(r[3]), ind=r[4], base=r[5], meta=r[6], prob=r[7]))
# productos criticos
for r in rows(wb, 'productos criticos')[1:]:
    if r[1] in out:
        out[r[1]].setdefault('productos', []).append(dict(prod=r[2], estado=r[3], adq=r[4], riesgo=r[5]))
# productos_resultados
for r in rows(wb, 'productos_resultados')[2:]:
    if r[0] in out:
        out[r[0]].setdefault('prod_res', []).append(dict(ind_id=str(r[1]), prod=r[2]))
# productos_adquisiciones
for r in rows(wb, 'productos_adquisiciones')[2:]:
    if r[0] in out:
        out[r[0]].setdefault('prod_adq', []).append(dict(prod=r[1], pid=r[2], critica=r[3], riesgo=r[4]))

# Raw Data disponible por operación
for op in OPS:
    for p in out[op].get('prestamos', []):
        code = str(p).replace('/', '_')
        if os.path.exists(f'{BASE}/clientportal/Raw Data - {code}.xlsx'): RAW.setdefault(op, []).append(code)

# agreements
ar = openpyxl.load_workbook(f'{BASE}/agreement_report.xlsx', read_only=True, data_only=True)
R = list(ar['Sheet1'].iter_rows(values_only=True))
hi = [i for i, r in enumerate(R) if r and 'Proyecto' in [str(c) for c in r]][0]
H = {str(h): i for i, h in enumerate(R[hi]) if h}
for r in R[hi + 1:]:
    op = r[H['Proyecto']]
    if op in out:
        g = lambda k: r[H[k]]
        out[op].setdefault('acuerdos', []).append(dict(
            estado=g('Estado'), tema=g('Tema crítico'), desc=g('Descripción del acuerdo'), venc=g('Fecha de Vencimiento'),
            resp=g('Responsable'), coment=g('Último comentario ingresado'), nueva=g('Nueva Fecha de Vencimiento'), revision=g('Revisión')))

# PMR history
ph = openpyxl.load_workbook(f'{BASE}/PMR_Historico.xlsx', read_only=True, data_only=True)
for op in OPS:
    if op in ph.sheetnames:
        hist = []
        for r in list(ph[op].iter_rows(values_only=True))[2:]:
            m = re.search(r'(20\d\d)', str(r[0] or ''))
            if m: hist.append((int(m.group(1)), r[10]))
        out[op]['pmr_hist'] = sorted(hist)

# Raw data
def durm(met, cat):
    met = str(met); cat = str(cat)
    if 'Consultor Individual' in met or cat == 'Consultores Individuales': return 12
    if 'Pública Internacional' in met: return 24
    if 'Pública Nacional' in met or 'Sistema Nacional' in met: return 18 if cat == 'Obras' else 12
    if 'Comparación de Precios' in met: return 12 if cat == 'Obras' else 9
    if '(SBCC)' in met: return 24
    return 18
def abbr(met):
    met = str(met)
    for k, v in (('Pública Internacional', 'LPI'), ('Pública Nacional', 'LPN'), ('Licitación Limitada', 'LL'), ('una sola etapa', 'LPI'),
                 ('Consultor Individual (SD)', 'SD-CI'), ('Selección Directa de Firmas', 'SD'), ('Auditoría Externa', 'SD')):
        if k in met: return v
    m = re.search(r'\(([A-Z0-9]+)\)', met)
    if m: return m.group(1)
    if 'Comparación de Precios' in met: return 'CP'
    if 'Sistema Nacional' in met: return 'SN'
    if 'Contratación Directa' in met: return 'CD'
    return met[:12]
Q = [(dt.datetime(2027, 1, 1), dt.datetime(2027, 4, 1)), (dt.datetime(2027, 4, 1), dt.datetime(2027, 7, 1)),
     (dt.datetime(2027, 7, 1), dt.datetime(2027, 10, 1)), (dt.datetime(2027, 10, 1), dt.datetime(2028, 1, 1))]
def share(a, b, y0, y1):
    if not a or not b or b <= a: return 0
    return max(0, (min(b, y1) - max(a, y0)).days) / (b - a).days

for op, codes in RAW.items():
    plan = {}; ctr = defaultdict(list); enm = defaultdict(list); hit = defaultdict(dict); cna = defaultdict(list); adj = {}; pcom = {}
    for code in codes:
        w = openpyxl.load_workbook(f'{BASE}/clientportal/Raw Data - {code}.xlsx', read_only=True, data_only=True)
        def tab(s):
            it = w[s].iter_rows(values_only=True); h = next(it)
            return [dict(zip(h, r)) for r in it if any(c is not None for c in r)]
        for d in tab('Plan de Adquisiciones'): plan[d['ID del Proceso']] = d
        for d in tab('Contratos'):
            if d['ID del Proceso']: ctr[d['ID del Proceso']].append(d)
        for d in tab('Enmiendas'): enm[d['ID del Contrato']].append(d)
        for d in tab('Hitos'): hit[d['ID del Proceso']][d['Nombre del Hito']] = d
        if 'Procesos Comentarios Públicos' in w.sheetnames:
            for d in tab('Procesos Comentarios Públicos'):
                k = d.get('ID del Proceso'); f = D(d.get('Fecha del Comentario'))
                if k and (k not in pcom or (f and pcom[k]['f'] and f > pcom[k]['f'])):
                    pcom[k] = dict(f=f, nombre=d.get('Nombre del Proceso'), com=re.sub(r'&nbsp;|\s+', ' ', str(d.get('Comentario') or '')).strip())
        for d in tab('Participantes y Evaluación'):
            if d.get('Resultado de la Evaluación') == 'Adjudicado' and d.get('Precio Evaluado (USD)'): adj[d['ID del Proceso']] = float(d['Precio Evaluado (USD)'])
        import os
        fca = f'{BASE}/clientportal/Contracts and Amendments - {code}.xlsx'
        if os.path.exists(fca):
            wca = openpyxl.load_workbook(fca, read_only=True, data_only=True)
            it = wca['Contratos'].iter_rows(values_only=True); hh = next(it)
            for r in it:
                d = dict(zip(hh, r))
                if d.get('ID del Contrato'): cna[d['ID del Contrato']].append(d)
    def firma(pid):
        h = hit[pid].get('Contrato Firmado')
        if not h: return None, None
        for k, lab in (('Fecha Real', 'real'), ('Fecha Re-estimada', 're-estimada'), ('Fecha Estimada', 'estimada')):
            if D(h.get(k)): return D(h[k]), lab
        return None, None
    # contracts (with amendments)
    contratos = []
    for pid, cs in ctr.items():
        p = plan.get(pid, {}); pct = float(p.get('% Monto BID') or 100) / 100
        for c in cs:
            a = D(c['Fecha de Inicio']); b = D(c['Fecha de Terminación']); m = float(c['Monto del Contrato (USD)'] or 0); mb = m * pct
            for e in sorted(enm.get(c['ID del Contrato'], []), key=lambda e: str(e.get('Fecha de firma de la Enmienda'))):
                b = D(e['Nueva Fecha de Terminación']) or b
                if e.get('Monto Total Acumulado (USD)'): m = float(e['Monto Total Acumulado (USD)'])
                if e.get('Monto Total Acumulado BID (USD)'): mb = float(e['Monto Total Acumulado BID (USD)'])
                else: mb = m * pct
            contratos.append(dict(pid=pid, cid=c['ID del Contrato'], nombre=c['Nombre del Contrato'], estado=c['Estado del Contrato'],
                                  inicio=a, fin=b, monto=m, monto_bid=mb, comp=p.get('Componente'), adjud=c.get('Nombre del Adjudicado')))
    # contracts registered only in the new Contracts and Amendments module (may be multi-currency)
    rates = []
    for cid, rows_ in cna.items():
        usd = sum(float(x['Total moneda'] or 0) for x in rows_ if x['Tipo de Moneda'] == 'USD')
        crc = sum(float(x['Total moneda'] or 0) for x in rows_ if x['Tipo de Moneda'] == 'CRC')
        a = adj.get(rows_[0]['ID del Proceso'])
        if crc and a and a > usd: rates.append(crc / (a - usd))
    RATE = sorted(rates)[len(rates) // 2] if rates else 500.0
    for cid, rows_ in cna.items():
        x0 = rows_[0]; pid = x0['ID del Proceso']; p = plan.get(pid, {})
        tot = bid = 0.0
        for x in rows_:
            fx = 1.0 if x['Tipo de Moneda'] == 'USD' else (1 / RATE if x['Tipo de Moneda'] == 'CRC' else None)
            if fx is None: continue
            tot += float(x['Total moneda'] or 0) * fx; bid += float(x['Monto BID'] or 0) * fx
        contratos.append(dict(pid=pid, cid=(re.search(r'CR-[A-Z]\d{4}-P\S+', cid) or [cid])[0], nombre=x0['Nombre del contrato'], estado=x0['Estado del Contrato'],
                              inicio=D(x0['Fecha de inicio']), fin=D(x0['Fecha de fin']), monto=tot, monto_bid=bid,
                              comp=p.get('Componente'), adjud=x0.get('Nombre de adjudicatario')))
        ctr[pid].append(x0)
    out[op]['ca_rate'] = RATE
    # overdue contracts
    venc = [c for c in contratos if 'jecuci' in str(c['estado']) and c['fin'] and c['fin'] < CORTE]
    out[op]['vencidos'] = [dict(cid=c['cid'], nombre=c['nombre'], fin=c['fin'].strftime('%d/%m/%Y'), monto=c['monto'], estado=c['estado']) for c in sorted(venc, key=lambda c: -c['monto'])]
    # 2027 projection by quarter
    qt = [0.0] * 4; contrib = [defaultdict(float) for _ in range(4)]
    for c in contratos:
        if str(c['estado']).startswith(('Final', 'Termin', 'Rescind', 'Cancel', 'Completado')): continue
        for k, (y0, y1) in enumerate(Q):
            v = c['monto_bid'] * share(c['inicio'], c['fin'], y0, y1)
            if v: qt[k] += v; contrib[k][c['nombre']] += v
    for pid, p in plan.items():
        if ctr.get(pid) or p['Estado del Proceso'] not in ('Previsto', 'Proceso en Curso'): continue
        a, _ = firma(pid)
        if not a: continue
        b = a + dt.timedelta(days=30 * durm(p['Método de Adquisición'], p['Categoría']))
        pct = float(p.get('% Monto BID') or 100) / 100; m = float(p['Monto Estimado (USD)'] or 0) * pct
        for k, (y0, y1) in enumerate(Q):
            v = m * share(a, b, y0, y1)
            if v: qt[k] += v; contrib[k][p['Nombre del Proceso'] + ' (previsto)'] += v
    fins = []
    for c in contratos:
        if str(c['estado']).startswith(('Final', 'Termin', 'Rescind', 'Cancel', 'Completado')) or not c['fin']: continue
        fins.append(dict(nombre=c['nombre'], fin=c['fin'].strftime('%d/%m/%Y'), monto=c['monto_bid'], tipo='contrato'))
    for pid, p in plan.items():
        if ctr.get(pid) or p['Estado del Proceso'] not in ('Previsto', 'Proceso en Curso'): continue
        a, _ = firma(pid)
        if not a: continue
        b = a + dt.timedelta(days=30 * durm(p['Método de Adquisición'], p['Categoría']))
        fins.append(dict(nombre=p['Nombre del Proceso'], fin=b.strftime('%d/%m/%Y'), monto=float(p['Monto Estimado (USD)'] or 0), tipo='previsto'))
    out[op]['fins'] = fins
    out[op]['proy2027'] = dict(q=qt, top=[sorted(cb.items(), key=lambda x: -x[1])[:2] for cb in contrib])
    # contracted by component (BID)
    comp = defaultdict(float)
    for c in contratos:
        if str(c['estado']).startswith(('Rescind', 'Cancel')): continue
        comp[str(c['comp'])] += c['monto_bid']
    out[op]['contratado_comp'] = dict(comp)
    # critical processes joined
    crit = []
    for pa in out[op].get('prod_adq', []):
        if pa['critica'] != 'Sí': continue
        p = plan.get(pa['pid'])
        if not p:
            pc_ = pcom.get(pa['pid'])
            crit.append(dict(pid=pa['pid'], prod=pa['prod'], nombre=None, riesgo=pa['riesgo'],
                             eliminado=dict(nombre=pc_['nombre'], fecha=pc_['f'].strftime('%d/%m/%Y') if pc_['f'] else '', com=pc_['com']) if pc_ else None)); continue
        f, lab = firma(pa['pid'])
        monto = sum(c['monto'] for c in contratos if c['pid'] == pa['pid']) or float(p['Monto del Contrato (USD)'] or p['Monto Estimado (USD)'] or 0)
        crit.append(dict(pid=pa['pid'], prod=pa['prod'], nombre=p['Nombre del Proceso'], metodo=abbr(p['Método de Adquisición']),
                         monto=monto, estado=p['Estado del Proceso'], firma=f.strftime('%d/%m/%Y') if f else None, firma_tipo=lab,
                         riesgo=pa['riesgo']))
    out[op]['crit'] = crit
    out[op]['raw'] = True

json.dump(out, open(os.path.join(WORK, 'data.json'), 'w'), ensure_ascii=False, indent=1, default=str)
for op in OPS:
    d = out[op]
    print(op, d.get('prestamos'), d.get('je'), 'res', len(d.get('resultados', [])), 'prod', len(d.get('productos', [])),
          'acu', len(d.get('acuerdos', [])), 'crit', len(d.get('crit', [])), 'venc', len(d.get('vencidos', [])),
          'proy', [round(x / 1e6, 2) for x in d.get('proy2027', {}).get('q', [])])
