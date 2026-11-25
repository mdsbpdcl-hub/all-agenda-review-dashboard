"""Turn the news dashboard's prepared clippings into news/clip_*.jpg and news/news_index.json.
Usage: python3 -I make_news.py <folder with the manifest and part files> "<paper date, e.g. 24 Nov 2026>" [headlines.json]
headlines.json (optional): {"<page number>": "<headline>"}"""
import json, base64, os, sys, glob, shutil
BASE = os.environ.get('AAR_BASE') or os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE)
import tag
AL = {'biharsharif': 'Nalanda', 'bihar sharif': 'Nalanda', 'bhabhua': 'Kaimur', 'ara': 'Bhojpur', 'arrah': 'Bhojpur', 'motihari': 'Purbi Champaran',
      'bettiah': 'Pashchim Champaran', 'betiya': 'Pashchim Champaran', 'betia': 'Pashchim Champaran', 'bagaha': 'Pashchim Champaran', 'chhapra': 'Saran', 'chapra': 'Saran', 'hajipur': 'Vaishali', 'purnea': 'Purnia',
      'gaya': 'Gaya Ji', 'sasaram': 'Rohtas', 'dehri': 'Rohtas', 'jehanabad': 'Jehanabad', 'east champaran': 'Purbi Champaran', 'west champaran': 'Pashchim Champaran'}
if __name__ == '__main__':
    src, date = sys.argv[1], sys.argv[2]
    hl = json.load(open(sys.argv[3])) if len(sys.argv) > 3 and os.path.exists(sys.argv[3]) else {}
    man = json.load(open(glob.glob(os.path.join(src, '*.manifest.json'))[0]))
    clips = []
    for part in man['parts']:
        clips += json.load(open(os.path.join(src, part)))['clips']

    out_dir = os.path.join(BASE, 'news'); shutil.rmtree(out_dir, ignore_errors=True); os.makedirs(out_dir)
    out = []
    for n, c in enumerate(clips):
        nm = str(c.get('d') or '').strip()
        dt = (tag.district(AL.get(nm.lower(), nm)) or tag.district(nm)) if nm else None
        f = 'news/clip_%02d_%d.jpg' % (int(c.get('p') or 0), n)
        open(os.path.join(BASE, f), 'wb').write(base64.b64decode(str(c['i']).split(',')[-1]))
        out.append(dict(dt=dt, place=nm, p=c.get('p'), w=c.get('w'), h=c.get('h'), f=f, hl=hl.get(str(c.get('p')), '')))
    json.dump(dict(date=date, paper='BSPHCL News Tracking', clips=out), open(os.path.join(out_dir, 'news_index.json'), 'w'), ensure_ascii=False)
    print(len(out), 'clippings,', sum(1 for o in out if o['dt']), 'tagged to a district; untagged places:', sorted({o['place'] for o in out if not o['dt']}))
