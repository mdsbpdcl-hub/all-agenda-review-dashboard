# Updating the review without the GitHub key (Claude scheduled task)

Use this until the `GOOGLE_SA_KEY` secret is set; after that the GitHub workflow does the same job by itself.

1. Clone `mdsbpdcl-hub/all-agenda-review-dashboard` (add_repo with push access) and `cd build`. `pip install openpyxl --break-system-packages -q`.
2. For each entry in `sources.json` -> `sheets_to_xlsx`: `mcp__Google_Drive__download_file_content` (fileId, exportMimeType
   `application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`), copy to `build/data/<name>`. Retry once; if it still fails,
   copy `fallback/<name w/o .xlsx>.json` to `data/` instead and note it.
3. `mcp__Google_Sheets__get_values` on `dm_feedback_sheet` -> write the JSON result to `data/dm_feedback.json`.
   Download `pending_issues_json` -> `pib/summary.json` (fallback `fallback/summary.json`).
   `curl -sf` the two NFMS/RDSS urls -> `ext/nfms_latest.json`, `ext/rdss_latest.json` (fallback copies in `fallback/`).
4. News: search Drive `parentId = '<news_prepared_folder>'`, newest `*.manifest.json` (`prep_<paperId>_<v>.manifest.json`);
   download it and all its parts into `news_raw/`; get the paper's title with get_file_metadata (date is in it).
   If `headlines/<paperId>.json` does not exist, run `python3 make_news.py news_raw "<dd Mon yyyy>"`, look at the clippings
   (contact sheets, 9 per image), and write `headlines/<paperId>.json` = {"<page>": "<Hindi headline as printed>"}.
   Then `python3 make_news.py news_raw "<dd Mon yyyy>" headlines/<paperId>.json`.
   If the newest paper is the one already in `../docs/news/news_index.json`, skip this step and copy that file to `news/news_index.json`.
5. `python3 build_review.py` (no traceback; every agenda shows a recent date).
6. `cp review.json ../docs/review.json`; if news was made: `rm -rf ../docs/news && cp -r news ../docs/news`.
   Refresh fallbacks: for each fresh source `python3 xlsx2json.py data/<name> fallback/<name w/o .xlsx>.json <fallback_sheets...>`;
   copy `pib/summary.json`, `ext/*.json` into `fallback/`.
7. Commit `docs build/fallback build/headlines` ("Update review figures <date time>") and push to main (`git -c push.negotiate=true push`).
8. Report in 2-3 lines with SendUserMessage: time, each agenda's "as on" date, any source that used yesterday's copy.
