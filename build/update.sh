#!/usr/bin/env bash
# Runs in GitHub Actions: fetch sources, make news, build, copy to docs/, refresh fallback copies.
set -euo pipefail
cd "$(dirname "$0")"
python3 fetch.py
DATE=$(python3 -c "import json;print(json.load(open('news_raw/paper.json')).get('date') or '')")
if [ -n "$DATE" ]; then
  PAPER=$(python3 -c "import json;print(json.load(open('news_raw/paper.json'))['paper'])")
  HL="headlines/$PAPER.json"   # optional Hindi headlines for that paper
  python3 make_news.py news_raw "$DATE" "$HL"
else
  mkdir -p news && cp ../docs/news/news_index.json news/news_index.json   # keep yesterday's clippings
fi
python3 build_review.py
# publish
if [ -n "$DATE" ]; then rm -rf ../docs/news && cp -r news ../docs/news; fi
cp review.json ../docs/review.json
# refresh yesterday's copies for the next run's fallback
python3 - <<'PY'
import json, os, subprocess, shutil
S = json.load(open('sources.json')); rep = json.load(open('fetch_report.json'))
for name in S['sheets_to_xlsx']:
    if name in rep['fresh']:
        subprocess.run(['python3', 'xlsx2json.py', 'data/' + name, 'fallback/' + name.replace('.xlsx', '.json')] + S['fallback_sheets'].get(name, []), check=True)
for src, dst in (('pib/summary.json', 'summary.json'), ('ext/nfms_latest.json', 'nfms_latest.json'), ('ext/rdss_latest.json', 'rdss_latest.json')):
    if os.path.exists(src): shutil.copy(src, 'fallback/' + dst)
PY
