"""Resolve any name from any agenda sheet to a Location Master ID."""
import json, re, sys, difflib, collections
import os
BASE = os.environ.get('AAR_BASE') or os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE)
from norm import key

M = json.load(open(os.path.join(BASE, 'master.json')))
DIVS, CIRCLES, DISTS = M['divisions'], M['circles'], M['districts']
SIG = re.compile(r'(EAST|WEST|NORT|SOUT|URBAN|RURAL|TOWN|\d+)')
def sg(k): return tuple('URBAN' if x == 'TOWN' else x for x in SIG.findall(k))
DC = {'SBPDCL': 'SB', 'NBPDCL': 'NB'}

# circle names written as the district / old names in some sheets
CIRCLE_ALIAS = {'BOJPUR': 'ARA', 'ROHTAS': 'SASARAM', 'NALANDA': 'BIHARSARIF', 'BIHARSARIFNALANDA': 'BIHARSARIF',
                'PATNAEAST': 'PESUEAST', 'PATNAWEST': 'PESUWEST', 'GAIAJI': 'GAIA', 'DEHARIONSONE': 'SASARAM'}
DIST_ALIAS = {'EASTCAMPARAN': 'PURBICAMPARAN', 'WESTCAMPARAN': 'PASCIMCAMPARAN', 'KAIMUR': 'KAIMURBABUA',
              'BABUA': 'KAIMURBABUA', 'PURNEA': 'PURNEA', 'SASARAM': 'ROHTAS', 'MOTIHARI': 'PURBICAMPARAN',
              'BETIAH': 'PASCIMCAMPARAN', 'CAPRA': 'SARAN', 'HAJIPUR': 'VAISALI', 'ARA': 'BOJPUR', 'GAIA': 'GAIA'}

par, dsc = {}, {}
for x in M['subdivs'] + M['sections']:
    par[x['id']] = x['div']; dsc[x['id']] = x['discom']
IDX = collections.defaultdict(dict)
for k, v in M['variants'].items():
    lv, mid = k.split('||')
    if lv == 'Circle': d = CIRCLES[mid][0]
    elif lv == 'Division': d = DIVS[mid][0]
    else: d = dsc.get(mid, '')
    for raw in v:
        IDX[lv][(d, par.get(mid, ''), key(raw))] = mid
for x in M['sections'] + M['subdivs']:
    lv = 'Section' if x['id'].startswith('SE') else 'Subdivision'
    IDX[lv][(x['discom'], x['div'], key(x['std']))] = x['id']
DKEY = {key(v): i for i, v in DISTS.items()}

def discom_of(raw):
    r = str(raw).upper()
    if 'SB' in r or 'SOUTH' in r: return 'SBPDCL'
    if 'NB' in r or 'NORTH' in r: return 'NBPDCL'
    return None

def circle(d, raw):
    if not raw: return None
    k = key(raw); k = CIRCLE_ALIAS.get(k, k)
    if d is None:
        for dd in DC:
            h = IDX['Circle'].get((dd, '', k))
            if h: return h
        return None
    return IDX['Circle'].get((d, '', k))

def district(raw):
    if not raw: return None
    k = key(raw); k = DIST_ALIAS.get(k, k)
    if k in DKEY: return DKEY[k]
    m = difflib.get_close_matches(k, list(DKEY), 1, 0.86)
    return DKEY[m[0]] if m else None

def division(d, cid, raw):
    if not raw: return None, ''
    k = key(raw)
    h = IDX['Division'].get((d, '', k))
    if h and (cid is None or DIVS[h][1] == cid or 'OLD' in DIVS[h][2]): return h, 'exact'
    pool = {kk: i for (dd, pp, kk), i in IDX['Division'].items() if dd == d and (cid is None or DIVS[i][1] == cid)}
    if k in pool: return pool[k], 'exact'
    m = [c for c in difflib.get_close_matches(k, list(pool), 3, 0.8) if sg(c) == sg(k)]
    if m: return pool[m[0]], 'close'
    if h: return h, 'other circle'          # e.g. Dak Bungalow listed under PESU East
    m = re.match(r'^(.*\S)\s+([EWUR])$', str(raw).strip(), flags=re.I)
    if m:
        full = {'E': 'EAST', 'W': 'WEST', 'U': 'URBAN', 'R': 'RURAL'}[m.group(2).upper()]
        return division(d, cid, m.group(1) + ' (' + full + ')')
    allp = {kk: i for (dd, pp, kk), i in IDX['Division'].items() if dd == d}
    m2 = [c for c in difflib.get_close_matches(k, list(allp), 2, 0.8) if sg(c) == sg(k)]
    if m2: return allp[m2[0]], 'other circle'
    return None, ''

def _within(lv, d, p, raw, thr=0.84):
    k = key(raw)
    h = IDX[lv].get((d, p, k))
    if h: return h, 'exact'
    pool = {kk: i for (dd, pp, kk), i in IDX[lv].items() if dd == d and pp == p}
    m = [c for c in difflib.get_close_matches(k, list(pool), 3, thr) if sg(c) == sg(k)]
    if m: return pool[m[0]], 'close'
    base = re.sub(r'(EAST|WEST|NORT|SOUT|URBAN|RURAL|TOWN|\d+)', '', key(DIVS[p][2])) if p in DIVS else ''
    for kk, i in pool.items():
        for a, b in ((k, kk), (kk, k)):
            if len(base) >= 4 and a.startswith(base) and b and a[len(base):] == b: return i, 'division name left out'
        if len(k) >= 10 and kk.startswith(k): return i, 'name cut short'      # NSC cuts names at 12 characters
    return None, ''

def unit(lv, d, cid, did, raw):
    """Subdivision / Section: inside its division, else elsewhere in the same circle."""
    if not raw: return None, ''
    if did:
        h, how = _within(lv, d, did, raw)
        if h: return h, how
    for dv, (dd, cc, nm) in DIVS.items():
        if dd == d and cc == cid and dv != did:
            h, how = _within(lv, d, dv, raw, thr=0.9)
            if h: return h, 'same circle, other division'
    return None, ''

def sskey(x):
    x = re.sub(r'\(.*?\)', '', str(x))
    x = re.sub(r'^\s*(PSS|GSS|ESS|33/11\s*KV|33/11KV|33\s*KV)[\s\-:.]*', '', x, flags=re.I)
    return key(x)
SS_BY_CIRCLE = collections.defaultdict(dict)
for sid, s in M['substations'].items():
    if not s['dv']: continue
    c = DIVS[s['dv']][1]
    SS_BY_CIRCLE[c][sskey(s['name'])] = sid
    for raw in M['variants'].get('Substation||' + sid, {}): SS_BY_CIRCLE[c][sskey(raw)] = sid
def substation(cid, raw):
    k = sskey(raw); pool = SS_BY_CIRCLE.get(cid, {})
    if k in pool: return pool[k], 'exact'
    m = [x for x in difflib.get_close_matches(k, list(pool), 2, 0.82) if sg(x) == sg(k)]
    if m: return pool[m[0]], 'close'
    return None, ''

# primary district of every division / subdivision / section
DIV_DIST = {}
for dt, link in M['dist_link'].items():
    for dv, n in link.items():
        if n > DIV_DIST.get(dv, ('', 0))[1]: DIV_DIST[dv] = (dt, n)
DIV_DIST = {k: v[0] for k, v in DIV_DIST.items()}
