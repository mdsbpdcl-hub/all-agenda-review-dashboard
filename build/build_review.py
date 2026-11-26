"""Build review.json: every agenda tagged to Master IDs, totalled to each level, with district review items."""
import json, re, sys, collections, datetime, openpyxl, difflib
import os
BASE = os.environ.get('AAR_BASE') or os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE)
import tag
from tag import key, M, DIVS, CIRCLES, DISTS
D = os.path.join(BASE, 'data') + '/'
ST = os.path.join(BASE, 'static') + '/'
WB = {}
def _revive(v):
    if isinstance(v, dict) and '__dt' in v: return datetime.datetime.fromisoformat(v['__dt'])
    return v
class _Sheet:
    def __init__(self, title, rows): self.title, self._r = title, rows
    def iter_rows(self, min_row=1, max_row=None, values_only=True):
        for r in self._r[min_row - 1:max_row]: yield tuple(_revive(v) for v in r)
class _Book:
    def __init__(self, d): self.worksheets = [_Sheet(t, r) for t, r in d.items()]; self._m = {w.title: w for w in self.worksheets}
    def __getitem__(self, t): return self._m[t]
    def __contains__(self, t): return t in self._m
    @property
    def sheetnames(self): return list(self._m)
def book(f):
    if f not in WB:
        if os.path.exists(D + f): WB[f] = openpyxl.load_workbook(D + f, read_only=True, data_only=True)
        else: WB[f] = _Book(json.load(open(D + f.replace('.xlsx', '.json'))))   # yesterday's copy, saved as JSON
    return WB[f]
def rows_st(f, t, hdr=1):
    wb = _Book(json.load(open(ST + f.replace('.xlsx', '.json'))))
    it = wb[t].iter_rows(min_row=hdr, values_only=True)
    h = [str(c).strip() if c is not None else '' for c in next(it)]
    for r in it:
        yield {h[i]: r[i] for i in range(min(len(h), len(r)))}
def rows(f, t, hdr=1):
    it = book(f)[t].iter_rows(min_row=hdr, values_only=True)
    h = [str(c).strip() if c is not None else '' for c in next(it)]
    for r in it:
        yield {h[i]: r[i] for i in range(min(len(h), len(r)))}
def num(v):
    if v is None or v == '': return 0.0
    if isinstance(v, (int, float)): return float(v)
    s = str(v).replace(',', '').replace('%', '').strip()
    try: return float(s)
    except ValueError: return 0.0
def txt(v): return '' if v is None else str(v).strip()

SECS = {x['id']: x for x in M['sections']}
SUBS = {x['id']: x for x in M['subdivs']}
SS = M['substations']
# ---------------------------------------------------------------- district of each unit
sub_dt = collections.defaultdict(collections.Counter)
for s in SS.values():
    if s.get('sd') and s.get('dt'): sub_dt[s['sd']][s['dt']] += 1
def div_dt(dv): return tag.DIV_DIST.get(dv, '')
def subdiv_dt(sd):
    if sub_dt.get(sd): return sub_dt[sd].most_common(1)[0][0]
    return div_dt(SUBS[sd]['div']) if sd in SUBS else ''
def sec_dt(se):
    x = SECS.get(se)
    if not x: return ''
    if x.get('dt'): return x['dt']
    if x.get('sd'): return subdiv_dt(x['sd'])
    return div_dt(x['div'])

# ---------------------------------------------------------------- hierarchy
H = {'circles': {}, 'districts': {}, 'divisions': {}, 'subdivs': {}, 'sections': {}, 'zones': M['zones']}
for cid, (d, n) in CIRCLES.items():
    H['circles'][cid] = dict(n=(n.title() if n.isupper() else n).replace('Pesu', 'PESU'), d=d, z=M['zone_of_circle'].get(cid, ''))
for did, (d, cid, n) in DIVS.items():
    if 'OLD' in n: continue
    H['divisions'][did] = dict(n=n, d=d, c=cid, t=div_dt(did))
for sd, x in SUBS.items():
    H['subdivs'][sd] = dict(n=x['std'], d=x['discom'], v=x['div'], t=subdiv_dt(sd))
for se, x in SECS.items():
    H['sections'][se] = dict(n=x['std'], d=x['discom'], v=x['div'], s=x.get('sd', ''), t=sec_dt(se))
share = collections.defaultdict(dict)
for dt, link in M['dist_link'].items():
    tot = collections.Counter()
    for dv, n in link.items(): share[dt][dv] = n
for dtid, nm in DISTS.items():
    link = share.get(dtid, {})
    divs = sorted(link, key=lambda x: -link[x])
    dsc = collections.Counter(DIVS[x][0] for x in divs)
    cir = collections.Counter()
    for x in divs: cir[DIVS[x][1]] += link[x]
    H['districts'][dtid] = dict(n=nm, d=dsc.most_common(1)[0][0] if dsc else '', c=[c for c, _ in cir.most_common()],
                                v=[[x, link[x]] for x in divs])
# circle order (meeting order): by DISCOM then circle name; districts under the circle that serves most of them
for dtid, x in H['districts'].items(): x['pc'] = x['c'][0] if x['c'] else ''

# ---------------------------------------------------------------- officers
OFF = collections.defaultdict(list)
def offer(uid, post, name, mob, status='', vacant_word='Vacant'):
    if not uid: return
    if not name and vacant_word != 'Vacant': return
    m = re.sub(r'\.0$', '', txt(mob))
    for o in OFF[uid]:
        if o['p'] == post and (o['n'] == (name or vacant_word) or not name): return
    OFF[uid].append(dict(p=post, n=name or vacant_word, m=m, s=status))
for r in rows_st('hr.xlsx', 'Officers'):
    cad, ut = txt(r.get('Cadre')), txt(r.get('Unit Type'))
    nm = (txt(r.get('Title')) + ' ' + txt(r.get('Officer Name'))).strip() if txt(r.get('Post Status')) == 'Filled' else ''
    mob = r.get('Directory Mobile (post)') or r.get('Mobile (as listed)')
    cid = tag.circle('SBPDCL', txt(r.get('Circle'))) if r.get('Circle') else None
    did, _ = tag.division('SBPDCL', cid, txt(r.get('Division'))) if r.get('Division') else (None, '')
    supply = 'SUPPLY' in (txt(r.get('Post Title')) + txt(r.get('Wing / Department'))).upper()
    if cad == 'ESE' and 'Circle' in ut and supply: offer(cid, 'ESE', nm, mob)
    elif cad == 'EEE' and 'Supply Division' in ut and not r.get('Sub-Division') and supply: offer(did, 'EEE', nm, mob)
    elif cad == 'AEE' and 'Supply Division' in ut and r.get('Sub-Division') and did:
        sd, _ = tag.unit('Subdivision', 'SBPDCL', DIVS[did][1], did, txt(r['Sub-Division'])); offer(sd, 'AEE', nm, mob)
    elif cad == 'JEE' and 'Section' in ut and r.get('Section') and did:
        se, _ = tag.unit('Section', 'SBPDCL', DIVS[did][1], did, txt(r['Section'])); offer(se, 'JEE', nm, mob)
for r in rows_st('contacts.xlsx', 'NBPDCL'):
    post, lvl = txt(r.get('Post / Designation')), txt(r.get('Hierarchy Level'))
    if post not in ('ESE', 'EEE', 'AEE', 'JEE'): continue
    nm, mob = txt(r.get('Officer / Employee Name')), r.get('Mobile Number')
    nm = nm or ''
    cid = tag.circle('NBPDCL', txt(r.get('Circle')))
    did, _ = tag.division('NBPDCL', cid, txt(r.get('Division'))) if r.get('Division') else (None, '')
    W = 'Name not listed'
    if post == 'ESE' and not r.get('Division'): offer(cid, 'ESE', nm, mob, vacant_word=W)
    elif post == 'EEE' and lvl == 'Division': offer(did, 'EEE', nm, mob, vacant_word=W)
    elif post == 'AEE' and lvl == 'Subdivision' and did:
        sd, _ = tag.unit('Subdivision', 'NBPDCL', DIVS[did][1], did, txt(r.get('Subdivision'))); offer(sd, 'AEE', nm, mob, vacant_word=W)
    elif post == 'JEE' and lvl == 'Section' and did:
        se, _ = tag.unit('Section', 'NBPDCL', DIVS[did][1], did, txt(r.get('Section'))); offer(se, 'JEE', nm, mob, vacant_word=W)
# EEE names/mobiles from the division social-media contact list (fills NBPDCL gaps)
_eee = os.path.join(BASE, 'static', 'eee_contacts.json')
if os.path.exists(_eee):
    for e in json.load(open(_eee, encoding='utf-8')):
        d, n = e.get('did'), (e.get('name') or '').strip()
        if not d or not n: continue
        L = OFF[d]; cur = [o for o in L if o['p'] == 'EEE']
        if d.startswith('DV-NB-'):
            for o in cur: L.remove(o)
            L.insert(0, dict(p='EEE', n=n, m=txt(e.get('mobile')), s=''))
        elif not cur:
            L.insert(0, dict(p='EEE', n=n, m=txt(e.get('mobile')), s=''))
OFF = {k: v[:3] for k, v in OFF.items()}

# ---------------------------------------------------------------- agenda engine
AG = {}
def agenda(key_, title, src, asof, kpis, note='', key=None):
    AG[key_] = dict(t=title, src=src, asof=asof, kpis=kpis, note=note, key=key or [], L={}, items=collections.defaultdict(list))
    return AG[key_]
def put(a, level, uid, base):
    if not uid: return
    L = a['L'].setdefault(level, {})
    cur = L.setdefault(uid, collections.Counter())
    for k, v in base.items(): cur[k] += v
def rollup(a, from_level, to_level, parent_of):
    for uid, b in list(a['L'].get(from_level, {}).items()):
        p = parent_of(uid)
        if p: put(a, to_level, p, b)
def finish(a, f):
    for lv, units in a['L'].items():
        for uid, b in units.items():
            b2 = dict(b); b2.update(f(b))
            units[uid] = {kk: (round(v, 2) if isinstance(v, float) else v) for kk, v in b2.items()}
def pct(a, b): return 100.0 * a / b if b else None
UNRES = collections.Counter()
def need(x, what):
    if not x: UNRES[what] += 1
    return x
sec_par = lambda se: H['sections'].get(se, {}).get('s') or None
sec_div = lambda se: H['sections'].get(se, {}).get('v') or None
sub_div = lambda sd: H['subdivs'].get(sd, {}).get('v') or None
div_cir = lambda dv: H['divisions'].get(dv, {}).get('c') or (DIVS[dv][1] if dv in DIVS else None)
sec_dist = lambda se: H['sections'].get(se, {}).get('t') or None
sub_dist = lambda sd: H['subdivs'].get(sd, {}).get('t') or None
div_dist = lambda dv: H['divisions'].get(dv, {}).get('t') or None

# ================================================================ 1. PMSG-MBY (Regular)
a = agenda('pmsg', 'PMSG-MBY (Regular)', 'PMSGDataUploadSheet-Bihar', '', [
    dict(k='inst', l='Installed', f='int', b='high', primary=True),
    dict(k='l7', l='Installed, last 7 days', f='int', b='high'),
    dict(k='l30', l='Installed, last 30 days', f='int', b='high'),
    dict(k='pend', l='Installation pending', f='int', b='low'),
    dict(k='wd', l='Waiting on DISCOM', f='int', b='low'),
    dict(k='ia', l='Inspection approved', f='int', b='high'),
    dict(k='app', l='Applications received', f='int', b='high'),
    dict(k='a7', l='Applications, last 7 days', f='int', b='high', lv=['district']),
    dict(k='tgt', l='Target', f='int', b='', lv=['district']),
    dict(k='pt', l='Installed, % of target', f='pct', b='high', lv=['district']),
    dict(k='pa', l='Installed, % of applications', f='pct', b='high')],
    key=['app', 'a7', 'inst', 'l7', 'pend', 'wd', 'tgt'])
for r in rows('pmsg_regular.xlsx', 'District KPIs'):
    if txt(r.get('DISCOM')) not in tag.DC: continue
    dt = tag.district(txt(r.get('District')))
    put(a, 'district', dt, dict(tgt=num(r['Target']), app=num(r['Applications']), inst=num(r['Installed']), l7=num(r['Installed, last 7 days']), l30=num(r['Installed, last 30 days']), ia=num(r['Inspection approved']),
                               pend=num(r['Installation pending']), wd=num(r['Waiting on DISCOM'])))
for r in rows('pmsg_regular.xlsx', 'Division KPIs'):
    d = txt(r.get('DISCOM'))
    if d not in tag.DC: continue
    cid = tag.circle(d, txt(r.get('Circle'))); did, _ = tag.division(d, cid, txt(r.get('Division')))
    put(a, 'division', need(did, 'pmsg div'), dict(app=num(r['Applications']), inst=num(r['Installed']), l7=num(r['Installed, last 7 days']), l30=num(r['Installed, last 30 days']), ia=num(r['Inspection approved']),
                                pend=num(r['Installation pending']), wd=num(r['Waiting on DISCOM'])))
for r in rows('pmsg_regular.xlsx', 'Sub-division KPIs'):
    d = txt(r.get('DISCOM'))
    if d not in tag.DC: continue
    did, _ = tag.division(d, None, txt(r.get('Division')))
    if not did: continue
    sd, _ = tag.unit('Subdivision', d, DIVS[did][1], did, re.sub(r'\s+[RU]$', '', txt(r.get('Sub-division'))))
    put(a, 'subdivision', need(sd, 'pmsg sub'), dict(app=num(r['Applications']), inst=num(r['Installed']), l7=num(r['Installed, last 7 days']), l30=num(r['Installed, last 30 days']), ia=num(r['Inspection approved']),
                                pend=num(r['Installation pending']), wd=num(r['Waiting on DISCOM'])))
rollup(a, 'division', 'circle', div_cir)
for r in rows('pmsg_regular.xlsx', 'Recent progress'):
    if txt(r.get('DISCOM')) not in tag.DC: continue
    dt = tag.district(txt(r.get('District')))
    if dt: put(a, 'district', dt, dict(a7=num(r['Last 7 days: Applications'])))
finish(a, lambda b: dict(pt=pct(b['inst'], b.get('tgt', 0)), pa=pct(b['inst'], b['app'])))
for r in rows('pmsg_regular.xlsx', 'Headline KPIs'):
    a['items']['_head'].append([txt(r.get('KPI')), txt(r.get('Bihar')), txt(r.get('SBPDCL')), txt(r.get('NBPDCL'))])
a['asof'] = '24 Nov 2026'

# ================================================================ 2. PMSG-MBY (ULA)
a = agenda('ula', 'PMSG-MBY (ULA)', 'Daily Progress Report of PM Surya Ghar Yojna', '', [
    dict(k='inv', l='Inverters installed', f='int', b='high', primary=True),
    dict(k='str', l='Structures installed', f='int', b='high'),
    dict(k='pan', l='Panels installed', f='int', b='high'),
    dict(k='y', l='Inverters installed yesterday', f='int', b='high'),
    dict(k='upd', l='Updated on portal', f='int', b='high'),
    dict(k='ca', l='Consumer accepted', f='int', b='high'),
    dict(k='pi', l='Pending for inspection', f='int', b='low'),
    dict(k='pi7', l='Inspection pending over 7 days', f='int', b='low'),
    dict(k='tgt', l='Target', f='int', b=''),
    dict(k='pt', l='Inverters installed, % of target', f='pct', b='high'),
    dict(k='pp', l='Updated on portal, % of installed', f='pct', b='high')],
    key=['str', 'pan', 'inv', 'y', 'upd', 'ca', 'pi', 'pi7'])
ws = book('pmsg_ula.xlsx')['CMDDashboard']
hdr1 = [txt(c) for c in next(ws.iter_rows(min_row=1, max_row=1, values_only=True))]
ula_date = re.search(r'Installed (?:Yesterday|Today) \((\d\d)\.(\d\d)\.(\d{4})\)', ' '.join(hdr1))
agency = ''
for r in ws.iter_rows(min_row=3, values_only=True):
    if r[0]: agency = txt(r[0])
    dn = txt(r[1])
    if not dn or 'total' in dn.lower(): continue
    dt = tag.district(dn)
    if not dt: continue
    put(a, 'district', dt, dict(tgt=num(r[2]), str=num(r[10]), pan=num(r[11]), inv=num(r[12]), y=num(r[9]), upd=num(r[13]), ca=num(r[14]), pi=num(r[18])))
    a['items'][dt].append(dict(k='ula', ag=agency, tgt=num(r[2]), str=num(r[10]), pan=num(r[11]), inv=num(r[12]), upd=num(r[13]), ca=num(r[14]), pi=num(r[18])))
for r in book('pmsg_ula.xlsx')['Inspections'].iter_rows(min_row=2, values_only=True):
    dt = tag.district(txt(r[0])) if r and r[0] else None
    if dt and dt in a['L'].get('district', {}): put(a, 'district', dt, dict(pi7=num(r[5])))
finish(a, lambda b: dict(pt=pct(b['inv'], b['tgt']), pp=pct(b['upd'], b['inv'])))
a['asof'] = datetime.date(int(ula_date.group(3)), int(ula_date.group(2)), int(ula_date.group(1))).strftime('%d %b %Y') if ula_date else ''

# ================================================================ 3. 1912 complaints
a = agenda('c1912', '1912 Complaint Analytics', 'Customer Call Centre - 1912', '', [
    dict(k='y', l='Complaints yesterday', f='int', b='low', primary=True),
    dict(k='ys', l='Supply complaints yesterday', f='int', b='low'),
    dict(k='yb', l='Billing complaints yesterday', f='int', b='low'),
    dict(k='sla', l='Resolved within SLA, %', f='pct', b='high'),
    dict(k='pb', l='Pending beyond SLA', f='int', b='low'),
    dict(k='rec', l='Complaints received (this month)', f='int', b=''),
    dict(k='pen', l='Pending now', f='int', b='low'),
    dict(k='l7', l='Received, last 7 days', f='int', b=''),
    dict(k='sp', l='Supply complaints pending', f='int', b='low', lv=['division', 'circle', 'district']),
    dict(k='spb', l='Supply pending beyond SLA', f='int', b='low', lv=['division', 'circle', 'district']),
    dict(k='sr', l='Supply complaints received', f='int', b='', lv=['division', 'circle', 'district']),
    dict(k='bp', l='Billing complaints pending', f='int', b='low', lv=['division', 'circle', 'district']),
    dict(k='bpb', l='Billing pending beyond SLA', f='int', b='low', lv=['division', 'circle', 'district']),
    dict(k='br', l='Billing complaints received', f='int', b='', lv=['division', 'circle', 'district'])],
    note='Month to date. Units with fewer than 30 complaints are left out of rankings.', key=['y', 'ys', 'yb', 'sp', 'spb', 'bp', 'bpb', 'sla'])
meta = {txt(r['key']): txt(r['value']) for r in rows('c1912.xlsx', 'DB_META')}
a['asof'] = meta.get('cutoff_text', '')
for r in rows('c1912.xlsx', 'DB_CIRCLE'):
    d = txt(r['discom']); cid = tag.circle(d, txt(r['circle']))
    put(a, 'circle', need(cid, '1912 circle'), dict(rec=num(r['received']), pen=num(r['pending']), pb=num(r['pend_beyond']), slaw=num(r['sla']) * num(r['received']), l7=num(r['l7'])))
for r in rows('c1912.xlsx', 'DB_DIVISION'):
    d = txt(r['discom']); cid = tag.circle(d, txt(r['circle'])); did, _ = tag.division(d, cid, txt(r['division']))
    put(a, 'division', need(did, '1912 div'), dict(rec=num(r['received']), pen=num(r['pending']), pb=num(r['pend_beyond']), slaw=num(r['sla']) * num(r['received']), l7=num(r['l7'])))
for r in rows('c1912.xlsx', 'DB_SECTION'):
    d = txt(r['discom']); cid = tag.circle(d, txt(r['circle'])); did, _ = tag.division(d, cid, txt(r['division']))
    se, _ = tag.unit('Section', d, cid, did, txt(r['section']))
    if not se: UNRES['1912 sec'] += 1; continue
    put(a, 'section', se, dict(rec=num(r['received']), pen=num(r['pending']), pb=num(r['pend_beyond']), slaw=num(r['sla']) * num(r['received']), l7=num(r['l7'])))
YK = {None: 'y', 'Supply Related': 'ys', 'Bill Related': 'yb'}
for r in rows('c1912.xlsx', 'DB_GEO'):
    lv_, k_ = txt(r['lvl']), YK.get(r['cat'] if r['cat'] else None, '-')
    if k_ == '-' or lv_ not in ('Circle', 'Division', 'Section'): continue
    d = txt(r['discom']); cid = tag.circle(d, txt(r['cir']))
    if lv_ == 'Circle': put(a, 'circle', cid, {k_: num(r['y'])}); continue
    did, _ = tag.division(d, cid, txt(r['div']))
    if lv_ == 'Division': put(a, 'division', did, {k_: num(r['y'])}); continue
    se, _ = tag.unit('Section', d, cid, did, txt(r['sec']))
    if se and se in a['L'].get('section', {}): put(a, 'section', se, {k_: num(r['y'])})
rollup(a, 'section', 'subdivision', sec_par); rollup(a, 'section', 'district', sec_dist)
catd = collections.defaultdict(collections.Counter)
for r in rows('c1912.xlsx', 'DB_DIVCAT'):
    c = txt(r['category']); pre = {'Supply Related': 's', 'Bill Related': 'b'}.get(c)
    if not pre: continue
    d = txt(r['discom']); did, _ = tag.division(d, None, txt(r['division']))
    if not did: UNRES['1912 divcat'] += 1; continue
    catd[did].update({pre + 'r': num(r['received']), pre + 'p': num(r['pending']), pre + 'pb': num(r['pend_beyond'])})
cc, cdt = collections.defaultdict(collections.Counter), collections.defaultdict(collections.Counter)
for did, b in catd.items():
    put(a, 'division', did, b)
    if div_cir(did): cc[div_cir(did)].update(b)
    if div_dist(did): cdt[div_dist(did)].update(b)
for k_, b in cc.items(): put(a, 'circle', k_, b)
for k_, b in cdt.items(): put(a, 'district', k_, b)
finish(a, lambda b: dict(sla=(b['slaw'] / b['rec']) if b.get('rec') else None, _n=b.get('rec', 0)))
for r in rows('c1912.xlsx', 'DB_REPEATS'):
    d = txt(r['discom']); did, _ = tag.division(d, None, txt(r['division']))
    se, _ = tag.unit('Section', d, DIVS[did][1] if did else None, did, txt(r['section'])) if did else (None, '')
    dt = sec_dist(se) if se else (div_dist(did) if did else None)
    if dt: a['items'][dt].append(dict(k='repeat', v=did, s=se, area=txt(r['area']), sub=txt(r['subtype']), n=num(r['complaints']), days=num(r['days']), pen=num(r['pending'])))
for r in rows('c1912.xlsx', 'DB_KPI'):
    if txt(r['period_key']) == 'MTD': a['items']['_head'].append([txt(r['scope']), num(r['received']), num(r['pending']), num(r['sla']), num(r['pend_beyond'])])

# ================================================================ 4. Outage
a = agenda('outage', 'Bihar Outage Analytics', 'Daily Outage Analytics - Bihar', '', [
    dict(k='aof', l='Average outage per feeder, minutes', f='dec0', b='low', primary=True),
    dict(k='o4', l='Feeders out over 4 hours', f='int', b='low'),
    dict(k='aif', l='Interruptions per feeder', f='dec1', b='low'),
    dict(k='fo', l='Feeders with outage, %', f='pct', b='low'),
    dict(k='tf', l='Total feeders', f='int', b='')], note='One day of feeder meter data.')
for d in ('SBPDCL', 'NBPDCL'):
    for r in rows('outage.xlsx', f'{d} Summary'):
        if txt(r.get('Discom')) == d:
            dt_ = r.get('Data Up To'); a['asof'] = dt_.strftime('%d %b %Y') if hasattr(dt_, 'strftime') else txt(dt_)
            a['items']['_head'].append([d, num(r['Total Feeders']), num(r['Feeders With Outage']), num(r['Avg Outage Min per Feeder']), num(r['Feeders Over 4 Hrs']), num(r['Over 24 Hrs Left Out'])])
    for r in rows('outage.xlsx', f'{d} Circle'):
        cid = tag.circle(d, txt(r['Circle']))
        put(a, 'circle', need(cid, 'out circle'), dict(tf=num(r['Total Feeders']), fwo=num(r['Feeders With Outage']), om=num(r['Outage Min']), it=num(r['Interruptions']), o4=num(r['Feeders Over 4 Hrs'])))
    for r in rows('outage.xlsx', f'{d} Division'):
        cid = tag.circle(d, txt(r['Circle'])); did, _ = tag.division(d, cid, txt(r['Division']))
        put(a, 'division', need(did, 'out div'), dict(tf=num(r['Total Feeders']), fwo=num(r['Feeders With Outage']), om=num(r['Outage Min']), it=num(r['Interruptions']), o4=num(r['Feeders Over 4 Hrs'])))
    for r in rows('outage.xlsx', f'{d} Substation'):
        sid = f"SS-{tag.DC[d]}-{txt(r.get('Substation Code'))}"
        if sid in SS and txt(r.get('Big Enough To Rank')) == 'Yes':
            put(a, 'substation', sid, dict(fwo=num(r['Feeders With Outage']), om=num(r['Outage Min']), it=num(r['Interruptions']), o4=num(r['Feeders Over 4 Hrs'])))
    for r in rows('outage.xlsx', f'{d} Over 24 Hrs Left Out'):
        cid = tag.circle(d, txt(r['Circle'])); did, _ = tag.division(d, cid, txt(r['Division']))
        dt = div_dist(did) if did else None
        for s_id, s in SS.items():
            if s['d'] == d and s['dv'] == did and tag.sskey(s['name']) == tag.sskey(r['Substation']) and s.get('dt'): dt = s['dt']; break
        if dt: a['items'][dt].append(dict(k='o24', v=did, ss=txt(r['Substation']), f=txt(r['Feeder']), ty=txt(r['Feeder Type']), h=num(r['Outage Hrs'])))
rollup(a, 'division', 'district', div_dist)
finish(a, lambda b: dict(aof=b['om'] / b['tf'] if b.get('tf') else (b['om'] / b['fwo'] if b.get('fwo') else None),
                         aif=b['it'] / b['tf'] if b.get('tf') else (b['it'] / b['fwo'] if b.get('fwo') else None),
                         fo=pct(b['fwo'], b['tf']) if b.get('tf') else None))

# ================================================================ 5. Revenue
a = agenda('revenue', 'Revenue Collection Analytics', 'Revenue Collection Monitoring', '', [
    dict(k='coll', l='Collection, ₹ lakh', f='dec1', b='high', primary=True),
    dict(k='ass', l='Current assessment, ₹ lakh', f='dec1', b=''),
    dict(k='pay', l='Paying consumers', f='int', b='high'),
    dict(k='act', l='Active consumers', f='int', b=''),
    dict(k='ce', l='Collection, % of current assessment', f='pct', b='high'),
    dict(k='pp', l='Paying consumers, % of active', f='pct', b='high')], note='Month to date. Sections with assessment under ₹1 lakh are left out of rankings.', key=['coll', 'ass', 'pay', 'act'])
for d in ('SBPDCL', 'NBPDCL'):
    allr = list(rows('revenue.xlsx', f'{d}_Sections'))
    last = max(txt(r['Period']) for r in allr if r.get('Period'))
    a['asof'] = datetime.date.fromisoformat(last[:10]).strftime('%d %b %Y')
    for r in allr:
        if txt(r['Period']) != last: continue
        cid = tag.circle(d, txt(r['Circle'])); did, _ = tag.division(d, cid, txt(r['Division']))
        se, _ = tag.unit('Section', d, cid, did, txt(r['Section']))
        if not se: UNRES['rev sec'] += 1; continue
        put(a, 'section', se, dict(ass=num(r['Current Assessment']), coll=num(r['Collection']), act=num(r['Active Consumers']), pay=num(r['Paying Consumers'])))
rollup(a, 'section', 'subdivision', sec_par); rollup(a, 'section', 'division', sec_div)
rollup(a, 'division', 'circle', div_cir); rollup(a, 'section', 'district', sec_dist)
finish(a, lambda b: dict(ce=pct(b['coll'], b['ass']), pp=pct(b['pay'], b['act']), _n=b['ass']))

# ================================================================ 6. NSC
a = agenda('nsc', 'NSC and Load Change', 'NSC & Load Change Tracking - Dashboard', '', [
    dict(k='pen', l='Applications pending', f='int', b='low', primary=True),
    dict(k='w15', l='Completed within 15 days, % of completed', f='pct', b='high'),
    dict(k='reg', l='Applications registered', f='int', b=''),
    dict(k='amisp', l='Pending at AMISP', f='int', b='low'),
    dict(k='lcp', l='Load change pending', f='int', b='low', lv=['division', 'circle', 'district'])],
    note='New service connections up to 59 kW, since 1 Jan 2020. Sections with fewer than 50 completed are left out of rankings.', key=['pen', 'amisp', 'lcp'])
for r in rows('nsc.xlsx', 'Analysis Summary'):
    if txt(r.get('Period To')): a['asof'] = txt(r.get('Period To')).replace('-', ' ')
for tb in ['North - NSC <19 kW', 'South - NSC <19 kW', 'North - NSC 19-59 kW', 'South - NSC 19-59 kW']:
    for r in rows('nsc.xlsx', tb):
        d = txt(r.get('Discom'))
        if d not in tag.DC: continue
        cid = tag.circle(d, txt(r['Circle'])); did, _ = tag.division(d, cid, txt(r['Division']))
        se, _ = tag.unit('Section', d, cid, did, txt(r['Section']))
        if not se: UNRES['nsc sec'] += 1; continue
        pen = sum(num(r[c]) for c in ['Pending For Document Verification', 'On-Hold Document Verification', 'Pending For Site Inspection',
                                       'On-Hold Site Inspection', 'Pending For Infra Installation', 'Pending At AMISP'])
        put(a, 'section', se, dict(reg=num(r['Registered']), c15=num(r['Completed Within 15 Days']), cb=num(r['Completed Beyond 15 Days']), pen=pen, amisp=num(r['Pending At AMISP'])))
for tb in ['North - Load Change', 'South - Load Change']:
    for r in rows('nsc.xlsx', tb):
        d = txt(r.get('Discom'))
        if d not in tag.DC: continue
        cid = tag.circle(d, txt(r['Circle'])); did, _ = tag.division(d, cid, txt(r['Division']))
        lcp = num(r['Total Registered']) - num(r['Cancelled']) - num(r['Service Released']) - num(r['Meter Change Not Required'])
        put(a, 'lc_div', did, dict(lcp=max(lcp, 0)))
rollup(a, 'section', 'subdivision', sec_par); rollup(a, 'section', 'division', sec_div)
for did, b in a['L'].pop('lc_div', {}).items(): put(a, 'division', did, b)
rollup(a, 'division', 'circle', div_cir)
dist_tmp = collections.defaultdict(collections.Counter)
for se, b in a['L']['section'].items():
    dt = sec_dist(se)
    if dt:
        for k in ('reg', 'c15', 'cb', 'pen', 'amisp'): dist_tmp[dt][k] += b[k]
for did, b in a['L']['division'].items():
    dt = div_dist(did)
    if dt: dist_tmp[dt]['lcp'] += b.get('lcp', 0)
for dt, b in dist_tmp.items(): put(a, 'district', dt, b)
finish(a, lambda b: dict(w15=pct(b.get('c15', 0), b.get('c15', 0) + b.get('cb', 0)), _n=b.get('c15', 0) + b.get('cb', 0)))

# ================================================================ 7. Utility shifting (SBPDCL)
a = agenda('utility', 'Pending Utility Shifting', 'Pending Utility Shifting - SBPDCL', '', [
    dict(k='np', l='Works not yet served an estimate', f='int', b='low', primary=True),
    dict(k='tot', l='Utility shifting works', f='int', b='')], note='SBPDCL only.')
upd = []
for r in rows('utility.xlsx', 'UtilityShifting'):
    if not r.get('Circle'): continue
    cid = tag.circle('SBPDCL', txt(r['Circle']))
    parts = [p.strip() for p in re.split(r'[\n,]+', txt(r.get('Division'))) if p.strip()]
    dids = [tag.division('SBPDCL', cid, p)[0] for p in parts]
    pend_at = txt(r.get('Pendency At'))
    served = 'SERVED' in pend_at.upper()
    for did in dids:
        if not did: continue
        put(a, 'division', did, dict(tot=1 / len(dids), np=(0 if served else 1) / len(dids)))
    dt = div_dist(dids[0]) if dids and dids[0] else None
    if r.get('Updated on'): upd.append(txt(r['Updated on']))
    if dt: a['items'][dt].append(dict(k='util', v=dids[0], ag=txt(r.get('Agency')), p=txt(r.get('Name of Project /\nChainage') or r.get('Name of Project / Chainage'))[:220],
                                     at=pend_at, rem=txt(r.get('SBPDCL Remarks'))[:300], served=served))
rollup(a, 'division', 'circle', div_cir); rollup(a, 'division', 'district', div_dist)
finish(a, lambda b: {})
a['asof'] = 'updated up to 19 Nov 2026' if upd else ''

# ================================================================ 8. PSS projects (SBPDCL)
a = agenda('pss', 'Project Monitoring - PSS', 'StatePlan - RDSS PSS Project Dashboard - SBPDCL', '', [
    dict(k='lw', l='RDSS PSS line work done, average %', f='pct', b='high', primary=True),
    dict(k='n', l='RDSS PSS under construction', f='int', b=''),
    dict(k='late', l='RDSS PSS past expected date or blocked', f='int', b='low'),
    dict(k='sn', l='State Plan PSS', f='int', b='', lv=['district']),
    dict(k='li', l='PSS with land issues', f='int', b='low')], note='SBPDCL only.', key=['li', 'n', 'late', 'sn'])
LANDOK = re.compile(r'rec(ei|ie)ved|transferred|handed', re.I)
today = datetime.datetime.now(datetime.timezone(datetime.timedelta(hours=5, minutes=30))).date()
for r in rows('pss_projects.xlsx', 'RDSSPSS'):
    if not txt(r.get('PSS ID')): continue
    cid = tag.circle('SBPDCL', txt(r['Circle'])); did, _ = tag.division('SBPDCL', cid, txt(r['Division']))
    if not did: did, _ = tag.division('SBPDCL', cid, 'Sasaram')
    dt = tag.district(txt(r['District']))
    lw = num(r.get('Line Work Overall %')); lw = lw * 100 if lw <= 1.0001 else lw
    exp = r.get('Expected Completion'); land = txt(r.get('Land Status'))
    late = (hasattr(exp, 'date') and exp.date() < today) or ('court' in land.lower()) or ('pending' in land.lower()) or not exp
    li = 1 if land and not LANDOK.search(land) else 0
    b = dict(n=1, lwsum=lw, late=1 if late else 0, li=li)
    put(a, 'division', did, b); put(a, 'district', dt, b)
    if dt: a['items'][dt].append(dict(k='rdss', v=did, id=txt(r['PSS ID']), nm=txt(r['Name of PSS']), tkc=txt(r.get('TKC')), lw=round(lw),
                                      exp=exp.strftime('%d %b %Y') if hasattr(exp, 'strftime') else txt(exp), land=land, dly=num(r.get('Delay vs Schedule (days)'))))
for r in rows('pss_projects.xlsx', 'StatePlanPSS'):
    if not txt(r.get('PSS-ID')): continue
    for p in txt(r['District']).split(','):
        dt = tag.district(p.strip())
        put(a, 'district', dt, dict(sn=1))
        fp = num(r.get('Financial % of LOA')); fp = fp * 100 if 0 < fp <= 1.0001 else fp
        tc = r.get('Target Completion')
        rem = txt(r.get('Remarks')); rem = '' if rem.startswith('PSS Detail') else rem
        land = ''
        if not r.get('Land NOC Date'): land = 'Land NOC not received'
        if re.search(r'land|noc|court|encroach|dispute', rem, re.I): land = rem
        if land: put(a, 'district', dt, dict(li=1))
        if dt: a['items'][dt].append(dict(k='state', nm=txt(r['PSS Name']), ag=txt(r.get('Agency')), fp=round(fp) if fp else None,
                                          tc=tc.strftime('%b %Y') if hasattr(tc, 'strftime') else txt(tc), rem=rem if rem != land else '', land=land))
rollup(a, 'division', 'circle', div_cir)
finish(a, lambda b: dict(lw=b['lwsum'] / b['n'] if b.get('n') else None))

# ================================================================ 9. Electrocution (SBPDCL)
a = agenda('electro', 'Pending Electrocution Cases', 'Pending Electrocution Cases - SBPDCL', '', [
    dict(k='sp', l='Cases still pending', f='int', b='low', primary=True),
    dict(k='cp', l='Cases received', f='int', b=''),
    dict(k='ap', l='Approved', f='int', b='high')], note='SBPDCL only.')
for r in rows('electrocution.xlsx', 'Pending Cases'):
    if not txt(r.get('Circle')): continue
    cid = tag.circle('SBPDCL', txt(r['Circle'])); did, _ = tag.division('SBPDCL', cid, txt(r['Division']))
    if not did: continue
    put(a, 'division', did, dict(sp=num(r['Cases Still Pending']), cp=num(r['Cases Pending']), ap=num(r['Approved'])))
    if num(r['Cases Still Pending']) > 0 and div_dist(did):
        a['items'][div_dist(did)].append(dict(k='electro', v=did, n=num(r['Cases Still Pending']), why=txt(r.get('Reason for Pendency')), rem=txt(r.get('Status/Remarks'))[:300]))
rollup(a, 'division', 'circle', div_cir); rollup(a, 'division', 'district', div_dist)
finish(a, lambda b: {})

# ================================================================ 10. Off-grid electrification
a = agenda('offgrid', 'OffGrid Electrification', 'OffGrid Electrification - Analytics', '', [
    dict(k='hh', l='Households to be electrified', f='int', b=''),
    dict(k='done', l='Households connected', f='int', b='high'),
    dict(k='dp', l='Households connected, %', f='pct', b='high', primary=True),
    dict(k='noc', l='NOCs and clearances pending', f='int', b='low')], key=['noc', 'hh', 'done'])
NOCW = re.compile(r'\bNOCs?\b|clearance|permission|wildlife|NBWL|SBWL|forest', re.I)
PENDW = re.compile(r'pending|awaited|under process|EDS|required|forward|return|not yet|to be', re.I)
ALLD = {v.lower(): k_ for k_, v in DISTS.items()}
ALLD.update({'west champaran': 'DT-25', 'east champaran': 'DT-27', 'purnea': tag.district('Purnia'), 'kaimur': tag.district('Kaimur')})
last = dict(proj='', st='', iss='')
GRP = collections.defaultdict(set); _p = ''
for r in book('offgrid.xlsx')['EnergyAccounting'].iter_rows(min_row=4, values_only=True):
    if txt(r[1]) == 'Total': break
    if txt(r[1]): _p = txt(r[1])
    GRP[_p].update(x for x in (tag.district(q.strip()) for q in re.split(r'&|,| and ', txt(r[2])) if q.strip()) if x)
for r in book('offgrid.xlsx')['EnergyAccounting'].iter_rows(min_row=4, values_only=True):
    if txt(r[1]) == 'Total': break
    if txt(r[1]): last.update(proj=txt(r[1]), st='', iss='')
    if txt(r[7]): last['st'] = txt(r[7])
    if txt(r[9]): last['iss'] = txt(r[9])
    ds = [tag.district(p.strip()) for p in re.split(r'&|,| and ', txt(r[2])) if p.strip()]
    ds = [x for x in ds if x]
    text = (last['st'] + '\n' + last['iss'])
    lines = [re.sub(r'^[\s▪*•-]+', '', x).strip() for x in re.split(r'\n|(?<=\.)\s+(?=[A-Z▪*])', text)]
    lines = [x for x in lines if len(x) > 20 and NOCW.search(x) and PENDW.search(x) and not re.search(r'\b(obtained|approved|cleared by)\b', x, re.I) or (len(x) > 20 and NOCW.search(x) and re.search(r'pending', x, re.I))]
    for dt in ds:
        mine = []
        for x in dict.fromkeys(lines):
            named = {v for k_, v in ALLD.items() if v and re.search(r'\b' + re.escape(k_) + r'\b', x.lower())}
            named &= GRP[last['proj']]
            if not named or dt in named: mine.append(x)
        put(a, 'district', dt, dict(hh=num(r[4]) / len(ds), done=num(r[5]) / len(ds), v=num(r[3]) / len(ds), noc=0))
        for x in mine:
            if not any(i.get('txt') == x[:2000] for i in a['items'][dt]): a['items'][dt].append(dict(k='noc', proj=last['proj'], txt=x[:2000]))
        a['items'][dt].append(dict(k='offgrid', proj=last['proj'], vill=num(r[3]), hh=num(r[4]), done=num(r[5]), st=last['st'][:500], iss=last['iss'][:400]))
for dt, b in a['L']['district'].items(): b['noc'] = sum(1 for i in a['items'].get(dt, []) if i['k'] == 'noc')
finish(a, lambda b: dict(dp=pct(b['done'], b['hh'])))

# ================================================================ 11. PMSG DM feedback (open, not test)
a = agenda('dmfb', 'PMSG DM Feedback', 'Energy Department - DM Feedback', '25 Nov 2026', [
    dict(k='open', l='Open issues from the DM', f='int', b='low', primary=True)])
DMFB = [('PMSG/FB/0008', '08-11-2026', 'Araria', 'Other', '', 'National portal problem', '18', 'Loan applications already disbursed by banks do not show in the "Applications Status Master Report" on the national portal, so vendors cannot be monitored.'),
        ('PMSG/FB/0009', '08-11-2026', 'Araria', 'Bank', 'State Bank of India', 'Sanction pending too long, Disbursal pending after sanction, Branch not responding', '11', 'SBI Kankhudia'),
        ('PMSG/FB/0011', '08-11-2026', 'Araria', 'Other', '', 'National portal problem', '', 'The 45-day wait before a consumer can change a non-responding vendor should be cut to 25 days.'),
        ('PMSG/FB/0014', '09-11-2026', 'Pashchim Champaran', 'Vendor', 'Eritech', 'Installation delayed, Vendor not available in area', '400', 'Many applicants chose this vendor but the vendor has not started work.'),
        ('PMSG/FB/0015', '09-11-2026', 'Sheikhpura', 'Bank', '', 'Sanction pending too long, CIBIL or credit score, Branch not responding, Other', '', 'Banks apply different CIBIL cut-offs; sanctions are slow; inter-branch Aadhaar verification fails; senior citizens are refused on age; repeated branch visits; rejections without checking with the DISCOM; margin money is hard for consumers.'),
        ('PMSG/FB/0016', '10-11-2026', 'Sheikhpura', 'Vendor', 'Edha Enterprises, Chhaya Solar, Diwakar Solutions, Usha Enterprises', 'Installation delayed, Documents or portal upload pending', '', 'Vendors short of technical staff; site visits and installation delayed after selection; documents late, so installations are not marked on the portal.')]
_fbp = os.path.join(D, 'dm_feedback.json')
if os.path.exists(_fbp):
    _v = json.load(open(_fbp)); _v = _v.get('values', _v) if isinstance(_v, dict) else _v
    _h = [str(x).strip().lower() for x in _v[0]]
    _g = lambda r, n: (str(r[_h.index(n)]).strip() if n in _h and _h.index(n) < len(r) else '')
    DMFB = []
    for r in _v[1:]:
        st = _g(r, 'status') or 'Open'
        if not _g(r, 'ref no') or re.match(r'(resolved|closed|not valid|invalid|disposed|rejected|duplicate|withdrawn)', st, re.I) or _g(r, 'closed on'): continue
        if re.search(r'hidden|hide|exclude|^no$', _g(r, 'in bulletin'), re.I): continue
        DMFB.append((_g(r, 'ref no'), _g(r, 'submitted on')[:10], _g(r, 'district'), _g(r, 'category'), _g(r, 'bank / vendor'), _g(r, 'issue type(s)'), _g(r, 'cases affected'), _g(r, 'details')[:600]))
a['asof'] = datetime.datetime.now(datetime.timezone(datetime.timedelta(hours=5, minutes=30))).strftime('%d %b %Y')
for ref, on, dn, cat, bv, it, n, det in DMFB:
    dt = tag.district(dn)
    put(a, 'district', dt, dict(open=1))
    a['items'][dt].append(dict(k='dmfb', ref=ref, on=on, cat=cat, bv=bv, it=it, n=n, det=det))
finish(a, lambda b: {})

# ================================================================ 12. CMD Cell monitoring sheet – district mentions
a = agenda('cmd', 'CMD Cell Monitoring Sheet', 'Energy Department - Daily Monitoring Sheet', '', [], note='Lines in the CMD Cell agenda tabs that name a district.')
DNAMES = {}
for dtid, nm in DISTS.items():
    for v in set([nm, nm.replace(' (Bhabua)', ''), nm.replace(' Ji', '')]): DNAMES[v.lower()] = dtid
DNAMES.update({'east champaran': 'DT-27', 'west champaran': 'DT-25', 'purnea': DISTS and tag.district('Purnia'), 'kaimur': tag.district('Kaimur'),
               'bhabhua': tag.district('Kaimur'), 'motihari': 'DT-27', 'bettiah': 'DT-25', 'ara': tag.district('Bhojpur'), 'chhapra': tag.district('Saran'), 'hajipur': tag.district('Vaishali')})
DNAMES = {k: v for k, v in DNAMES.items() if v}
skip = re.compile(r'^(Sheet\d*|Sync_Log|Morning Brief|Reporting_Agenda|Copy of.*)$')
agenda_titles = {}
for r in book('cmd_cell.xlsx')['Reporting_Agenda'].iter_rows(min_row=3, values_only=True):
    if r[1]: agenda_titles[txt(r[1]).lower()] = txt(r[2])[:140]
seen = set()
for ws in book('cmd_cell.xlsx').worksheets:
    if skip.match(ws.title): continue
    for r in ws.iter_rows(values_only=True):
        line = ' | '.join(txt(c) for c in r if c not in (None, ''))
        if len(line) < 25: continue
        low = line.lower()
        hits = {v for k, v in DNAMES.items() if re.search(r'\b' + re.escape(k) + r'\b', low)}
        if len(hits) > 2: continue          # lists of many districts are coverage tables, not district issues
        if len(re.findall(r'[A-Za-z]{3,}', line)) < 10: continue   # mostly-number table rows
        for dt in hits:
            sig_ = (dt, line[:120])
            if sig_ in seen: continue
            seen.add(sig_)
            a['items'][dt].append(dict(k='cmd', tab=re.sub(r'^\d+[a-z]?[_.]\s*', '', ws.title).strip(), txt=re.sub(r'\s+', ' ', line)[:2000]))
for dt in list(a['items']):
    if dt.startswith('DT'):
        a['items'][dt] = a['items'][dt][:12]
        put(a, 'district', dt, dict(n=len(a['items'][dt])))
finish(a, lambda b: {})
mb = list(book('cmd_cell.xlsx')['Morning Brief'].iter_rows(values_only=True))
a['items']['_head'] = [[txt(r[0]), txt(r[1]), txt(r[2])] for r in mb[8:16] if r and r[0] and r[1]]
a['asof'] = 'read on ' + datetime.datetime.now(datetime.timezone(datetime.timedelta(hours=5, minutes=30))).strftime('%d %b %Y')

# ================================================================ 13. NFMS and RDSS (state / DISCOM level)
nf = json.load(open(os.path.join(BASE, 'ext', 'nfms_latest.json')))
nrows = nf['rows']
def nfrank(field, level, name_match):
    pool = [r for r in nrows if r['level'] == level and isinstance(r.get(field), (int, float))]
    pool.sort(key=lambda r: -r[field])
    for i, r in enumerate(pool, 1):
        if name_match(r): return dict(v=r[field], rank=i, of=len(pool))
    return None
NF = {}
for lab, fld in [('Hours of supply (excl. agriculture)', 'hos_excl_agri'), ('Rural feeders, hours', 'rural'), ('Urban feeders, hours', 'urban')]:
    NF[lab] = dict(Bihar=nfrank(fld, 'State', lambda r: 'BIHAR' in r['state'].upper()),
                   SBPDCL=nfrank(fld, 'DISCOM', lambda r: 'SOUTH BIHAR' in (r['name'] + r.get('discom_label', '')).upper() or 'SBPDCL' in (r['name'] + r.get('discom_label', '')).upper()),
                   NBPDCL=nfrank(fld, 'DISCOM', lambda r: 'NORTH BIHAR' in (r['name'] + r.get('discom_label', '')).upper() or 'NBPDCL' in (r['name'] + r.get('discom_label', '')).upper()),
                   India=next((r[fld] for r in nrows if r['level'] == 'National'), None))
rd = json.load(open(os.path.join(BASE, 'ext', 'rdss_latest.json')))
RD = {}
for k, p in rd['progress'].items():
    if not p.get('bihar_discoms'): continue
    RD[p['label']] = dict(unit=p['unit'], done=p['done'], discoms=p['bihar_discoms'], india=p.get('india'))
atc = rd['annual']['atc']
RD['_atc'] = dict(years=atc['years'], discoms=[x for x in atc['discoms'] if re.search(r'BIHAR|SBPDCL|NBPDCL', json.dumps(x).upper())], india=atc['india'])
STATE = dict(nfms=dict(asof=nf['snapshot_date'], v=NF), rdss=dict(asof=rd['fetched_at'][:10], v=RD))

# ================================================================ 14. District pending issues (daily briefing summary)
_pib = json.load(open(os.path.join(BASE, 'pib', 'summary.json')))
_b = _pib['pending_issues_bulletin']
_rd = _b.get('report_date') or (_pib.get('metadata') or {}).get('report_snapshot_date') or ''
a = agenda('dpi', 'District pending issues', 'Pending Issues Bulletin', datetime.date.fromisoformat(_rd[:10]).strftime('%d %b %Y') if _rd else '', [
    dict(k='n', l='Pending issues', f='int', b='low', primary=True)])
DPI_ALIAS = {'west champaran': 'Pashchim Champaran', 'east champaran': 'Purbi Champaran', 'purnea': 'Purnia', 'gaya': 'Gaya Ji', 'kaimur': 'Kaimur'}
for it in _b['issues']:
    for dn in it.get('districts') or []:
        dt = tag.district(DPI_ALIAS.get(dn.lower(), dn)) or tag.district(dn)
        if not dt: UNRES['dpi district'] += 1; continue
        put(a, 'district', dt, dict(n=1))
        a['items'][dt].append(dict(k='dpi', t=it.get('title', ''), s=it.get('summary', ''), pr=it.get('project', ''), ro=it.get('responsible_office', ''),
                                   pw=', '.join(it.get('pending_with') or []), on=it.get('as_of_label', ''), dl=it.get('deadline') or ''))
finish(a, lambda b: {})

# ---------------------------------------------------------------- pending issues: one list per district
PEND = collections.defaultdict(list)
for akey, kinds in [('dpi', ['dpi']), ('dmfb', ['dmfb']), ('utility', ['util']), ('electro', ['electro']), ('pss', ['rdss', 'state']), ('offgrid', ['noc'])]:
    for dt, its in AG[akey]['items'].items():
        if not dt.startswith('DT'): continue
        for it in its:
            if it['k'] not in kinds: continue
            if it['k'] == 'util' and it['served']: continue
            if it['k'] == 'rdss' and not (it['land'] or it['lw'] < 30): continue
            if it['k'] == 'state' and not it.get('land'): continue
            PEND[dt].append(dict(it, src=akey))
for dt, its in AG['outage']['items'].items():
    if dt.startswith('DT'):
        for it in its: PEND[dt].append(dict(it, src='outage'))

_news = json.load(open(os.path.join(BASE, 'news', 'news_index.json')))
NEWS = dict(date=_news['date'], paper=_news['paper'], by={})
for c in _news['clips']:
    if c['dt']: NEWS['by'].setdefault(c['dt'], []).append(dict(f=c['f'], w=c['w'], h=c['h'], p=c['p'], place=c['place'], hl=c.get('hl', '')))
_now = datetime.datetime.now(datetime.timezone(datetime.timedelta(hours=5, minutes=30)))
out = dict(generated=_now.strftime('%d %b %Y, %H:%M'),
           H=H, OFF=OFF, AG={k: dict(v, items=dict(v['items'])) for k, v in AG.items()}, STATE=STATE, PEND=dict(PEND), NEWS=NEWS,
           SS={k: dict(n=v['name'], v=v['dv'], t=v.get('dt', '')) for k, v in SS.items()})
json.dump(out, open(os.path.join(BASE, 'review.json'), 'w'), ensure_ascii=False, separators=(',', ':'), default=str)
print('unresolved', dict(UNRES))
for k, v in AG.items():
    print(k, v['asof'], {lv: len(u) for lv, u in v['L'].items()}, 'items in', len([d for d in v['items'] if d.startswith('DT')]), 'districts')
print('officers', collections.Counter(o['p'] for v in OFF.values() for o in v))
print('pending districts', len(PEND), sum(len(v) for v in PEND.values()))
import os; print('size KB', os.path.getsize(os.path.join(BASE, 'review.json')) // 1024)
