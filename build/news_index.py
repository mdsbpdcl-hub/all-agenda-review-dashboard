"""docs/news/clips.json (saved by the Apps Script updater) -> news/news_index.json with each clipping's district."""
import json, os, sys
BASE = os.path.dirname(os.path.abspath(__file__)); sys.path.insert(0, BASE)
import tag
from make_news import AL
src = json.load(open(os.path.join(BASE, '..', 'docs', 'news', 'clips.json')))
hl_path = os.path.join(BASE, 'headlines', (src.get('paper_id') or 'x') + '.json')
hl = json.load(open(hl_path)) if os.path.exists(hl_path) else {}
out = []
for c in src['clips']:
    nm = c.get('place') or ''
    dt = (tag.district(AL.get(nm.lower(), nm)) or tag.district(nm)) if nm else None
    out.append(dict(c, dt=dt, hl=c.get('hl') or hl.get(str(c.get('p')), '')))
os.makedirs(os.path.join(BASE, 'news'), exist_ok=True)
json.dump(dict(date=src['date'], paper=src.get('paper', 'BSPHCL News Tracking'), clips=out), open(os.path.join(BASE, 'news', 'news_index.json'), 'w'), ensure_ascii=False)
print(len(out), 'clippings,', sum(1 for o in out if o['dt']), 'tagged; untagged places:', sorted({o['place'] for o in out if not o['dt']}))
