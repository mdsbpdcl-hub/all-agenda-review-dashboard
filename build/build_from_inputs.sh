#!/usr/bin/env bash
# GitHub Actions: build the review from the inputs the Apps Script updater saved in build/data_in/ and docs/news/.
set -euo pipefail
cd "$(dirname "$0")"
mkdir -p data pib ext news
for f in data_in/*.json; do
  n=$(basename "$f")
  case "$n" in
    summary.json) cp "$f" pib/summary.json ;;
    *) cp "$f" "data/$n" ;;        # the build reads data/<name>.json when there is no .xlsx
  esac
done
[ -f pib/summary.json ] || cp fallback/summary.json pib/summary.json
for k in nfms rdss; do
  url=$(python3 -c "import json;print(json.load(open('sources.json'))['${k}_latest'])")
  curl -sf "$url" -o "ext/${k}_latest.json" && cp "ext/${k}_latest.json" "fallback/${k}_latest.json" || cp "fallback/${k}_latest.json" "ext/${k}_latest.json"
done
if [ -f ../docs/news/clips.json ]; then python3 news_index.py; else cp ../docs/news/news_index.json news/news_index.json; fi
python3 build_review.py
cp review.json ../docs/review.json
cp news/news_index.json ../docs/news/news_index.json
