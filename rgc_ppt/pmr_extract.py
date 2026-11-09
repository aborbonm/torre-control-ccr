import os, glob as _glob
HERE = os.path.dirname(os.path.abspath(__file__))
WORK = os.environ.get('RGC_WORK', os.path.join(HERE, 'work'))
BASE = os.environ.get('RGC_BASE', '/mnt/user-data/uploads/torre-control-ccr')
import datetime as _dt
import json as _json
_AG = os.path.join(WORK, 'agenda.json')
CORTE = _dt.datetime.fromisoformat(os.environ.get('RGC_CORTE') or (_json.load(open(_AG)).get('corte') if os.path.exists(_AG) else '2026-11-03'))
"""Extract from PMR Operational Reports (text), PMR simulators and OPS1 reports -> gen/pmr.json, gen/ops.json"""
import re, json, glob, difflib, subprocess, datetime as dt
import openpyxl

G = WORK
TXT = os.path.join(WORK, 'pmr_txt')
DATA = json.load(open(f'{G}/data.json'))
NUMRE = re.compile(r'^-?[\d.,]+%?$')

def norm(s):
    s = (s or '').lower()
    s = re.sub(r'\[[^\]]*\]', '', s)
    s = re.sub(r'^[a-z]?\.?[\d.]+\s+', '', s)
    s = s.replace('proyecto ', '').replace('–', '-')
    return re.sub(r'[^a-z0-9áéíóúñ]+', ' ', s).strip()

STOP = set('de la el en los las del y a para por con e al su sus o un una programa'.split())
def toks(s): return [t[:6] for t in norm(s).split() if t not in STOP and (len(t) > 2 or t.isdigit())]
def score(a, b):
    ta, tb = toks(a), set(toks(b))
    if not ta: return 0
    cov = sum(1 for t in ta if t in tb) / len(ta)
    dig = [t for t in ta if t.isdigit()]
    if dig and not all(t in tb for t in dig): cov *= 0.5
    prec = sum(1 for t in tb if t in set(ta)) / max(1, len(tb))
    return 0.55 * cov + 0.2 * prec + 0.25 * difflib.SequenceMatcher(None, norm(a), norm(b)).ratio()
def best(name, cands, key, thr=0.55):
    sc = [(score(name, key(c)), i) for i, c in enumerate(cands)]
    if not sc: return None
    r, i = max(sc)
    return cands[i] if r >= thr else None

def parse_indicators(lines):
    out = []
    stop = next((i for i, l in enumerate(lines) if 'OUTPUTS PHYSICAL PROGRESS' in l), len(lines))
    hdr = None
    i = 0
    while i < stop:
        l = lines[i]
        if 'Indicator' in l and 'Unit of Measure' in l:
            hdr = l; i += 1; continue
        m = re.match(r'^\s*(\d+\.\d+)\s{2,}(\S.*)$', l)
        if hdr and m and re.search(r'\sP\s', l):
            upos = hdr.index('Unit of Measure')
            ppos = re.search(r'\sP\s', l).start() + 1
            name = l[m.start(2):upos].strip()
            # years from header after "Baseline Year"
            yrs = re.findall(r'(EOP \d{4}|\d{4})', hdr[hdr.index('Baseline Year') + 13:]) if 'Baseline Year' in hdr else []
            pv = l[ppos + 1:].split()
            base_seg = l[upos:ppos].split()
            av = None; j = i + 1
            while j < min(i + 6, stop):
                lj = lines[j]
                ma = re.search(r'\sA\s', lj)
                if ma and abs(ma.start() + 1 - ppos) < 6:
                    av = lj[ma.start() + 2:].split()
                    extra = lj[:upos].strip()
                    if extra: name += ' ' + extra
                    break
                extra = lj[:upos].strip()
                if extra and not extra.startswith('Details'): name += ' ' + extra
                j += 1
            unit = ' '.join(x for x in base_seg if not NUMRE.match(x))
            nums = [x for x in base_seg if NUMRE.match(x)]
            base = nums[0] if nums else None
            vals = pv[-len(yrs):] if yrs and len(pv) >= len(yrs) else pv
            act = (av or [])[-len(yrs):] if yrs and av and len(av) >= len(yrs) else (av or [])
            actual = None
            ys = [y for y in yrs if not y.startswith('EOP')]
            for y, a in zip(ys, act[:len(ys)]):
                if a not in ('-', ''): actual = (a, y)
            out.append(dict(pid=m.group(1), name=re.sub(r'\s+', ' ', name), unit=unit, base=base, actual=actual,
                            eop=vals[-1] if vals else None))
            i = j + 1; continue
        i += 1
    return out

def section(lines, start, ends):
    try: a = next(i for i, l in enumerate(lines) if start in l)
    except StopIteration: return ''
    b = next((i for i in range(a + 1, len(lines)) if any(e in lines[i] for e in ends)), len(lines))
    return '\n'.join(lines[a + 1:b])

pmr = {}
os.makedirs(TXT, exist_ok=True)
for pdf in glob.glob(f'{BASE}/pmr reports/*.pdf'):
    first = subprocess.run(['pdftotext', '-l', '1', pdf, '-'], capture_output=True, text=True).stdout
    m = re.search(r'CR-[A-Z]\d{4}', first)
    if m: subprocess.run(['pdftotext', '-layout', pdf, f'{TXT}/{m.group(0)}.txt'])
for f in glob.glob(f'{TXT}/*.txt'):
    op = f.split('/')[-1][:-4]
    if op not in DATA: continue
    lines = open(f).read().split('\n')
    inds = parse_indicators(lines)
    st = section(lines, 'Status of the Implementation', ['Likelihood of Achieving'])
    st = re.sub(r'\s+', ' ', ' '.join(x for x in st.split('\n') if not re.fullmatch(r'\s*(EOP \d{4}|-)?\s*', x))).strip()
    lk = section(lines, 'Likelihood of Achieving Development Objectives', ['Lesson Learned'])
    lk = re.sub(r'\s+', ' ', lk).strip()
    pmr[op] = dict(inds=inds, status=st, likelihood=lk)

# match indicators to Torre
for op, P in pmr.items():
    res = DATA[op].get('resultados', [])
    m = {}
    for r in res:
        b = best(r['ind'], P['inds'], lambda c: c['name'])
        if b and b['actual']: m[r['ind_id']] = dict(actual=b['actual'][0], year=b['actual'][1], pmr_name=b['name'])
    P['ind_match'] = m

# simulators: physical progress per output
for f in glob.glob(f'{BASE}/pmr reports/PMR_SIMULATOR*.xlsx'):
    if '_v2' in f: continue
    op = re.search(r'CR-[A-Za-z]\d{4}', f).group(0).upper()
    wb = openpyxl.load_workbook(f, read_only=True, data_only=True)
    ws = wb['Physical progress']
    outs = {}; cur = None
    for r in ws.iter_rows(min_row=4, values_only=True):
        if r[0] is None: continue
        nums = [c for c in r if isinstance(c, (int, float))]
        lab = r[4] if len(r) > 4 else None
        if lab == 'P' and isinstance(r[1], str):
            outs[r[0]] = dict(name=r[1].strip(), unit=r[2], pa=None, a=None)
        elif lab == 'P(a)' and r[0] in outs: outs[r[0]]['pa'] = nums[-1] if nums else None
        elif lab == 'A' and r[0] in outs: outs[r[0]]['a'] = nums[-1] if nums else None
    sim = [o for o in outs.values() if o['pa']]
    for o in sim: o['pct'] = round(100 * (o['a'] or 0) / o['pa'])
    pmr.setdefault(op, dict(inds=[], status='', likelihood='', ind_match={}))['outputs'] = sim
    pm = {}
    for p in DATA[op].get('productos', []):
        b = best(p['prod'], sim, lambda c: c['name'], 0.5)
        if b: pm[p['prod']] = dict(pct=b['pct'], a=b['a'], pa=b['pa'], unit=b['unit'], name=b['name'])
    pmr[op]['prod_match'] = pm
json.dump(pmr, open(f'{G}/pmr.json', 'w'), ensure_ascii=False, indent=1)

# OPS1
def D(x):
    if isinstance(x, dt.datetime): return x.strftime('%d/%m/%Y')
    try: return dt.datetime.strptime(str(x), '%b-%d-%Y').strftime('%d/%m/%Y')
    except Exception: return str(x)
def F(x):
    try: return float(str(x).replace(',', '').replace('%', ''))
    except Exception: return None
ops = {}
for f in glob.glob(f'{BASE}/OPS reports/*.xlsx'):
    wb = openpyxl.load_workbook(f, read_only=True, data_only=True)
    cells = [c for r in wb['Summary'].iter_rows(values_only=True) for c in r if c not in (None, '')]
    kv = {}
    for k, v in zip(cells, cells[1:]):
        if isinstance(k, str) and k.strip().endswith(':'): kv[k.strip()[:-1]] = v
    loan = kv.get('Operation Number'); proj = str(kv.get('Project Number', ''))[:8]
    tx = list(wb['Oper Transactions (OPS10)'].iter_rows(min_row=6, values_only=True))
    ext = [t for t in tx if len(t) > 2 and t[2] == 'EXT']
    ops[loan] = dict(proj=proj, status=kv.get('Operation Status'), disb=F(kv.get('Disbursed')), undisb=F(kv.get('Undisbursed')),
                     ytd=F(kv.get('Disbursed YTD')), approved=F(kv.get('Curr. Approved')),
                     exp=D(kv.get('Current Disbursement Exp. Date')), exp_orig=D(kv.get('Original Disbursement  Exp. Date')),
                     ant=F(kv.get('Last Advance of Funds (ANT) Amount')), ant_date=D(kv.get('Last Advance of Funds (ANT) Date')),
                     pct_just=kv.get('Advance Justification Percentage'), min_pend=F(kv.get('Minimum Amount Pending Justification')),
                     deadline=D(kv.get('Financial Planning Period Deadline')), n_ext=len(ext))
json.dump(ops, open(f'{G}/ops.json', 'w'), ensure_ascii=False, indent=1)
for op, P in pmr.items():
    print(op, 'inds', len(P['inds']), 'match', len(P['ind_match']), 'outputs', len(P.get('outputs', [])), 'pm', len(P.get('prod_match', {})), 'st', len(P['status']))
for k, v in ops.items(): print(k, v)
