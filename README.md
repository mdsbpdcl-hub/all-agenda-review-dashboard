# All-Agenda Review Dashboard

Circle-, district-, division-, sub-division- and section-wise review of Energy Department agendas
(SBPDCL and NBPDCL), built on the common Location Master IDs.

- `master_updates/` – rows added to the Location Master (Lookup, Substation, Section, TRW) on 25 Nov 2026.

## Location Tagger (Apps Script)

`apps-script/LocationTagger.gs` goes into the **DistrictWiseReviewSheet-CMDLeveDashboard** spreadsheet
(Extensions › Apps Script › paste › Save, then reload the sheet). It adds a **Location Master** menu:

1. **Add new master rows from GitHub** – appends the rows in `master_updates/` to Lookup, Substation and Section.
   Rows whose key or ID is already in the tab are skipped, so it is safe to run again.
2. **Tag all dashboard sheets now** – for every tab listed in `SOURCES`, writes a companion tab
   `IDs – <tab>` in that dashboard's own spreadsheet: source row, names as written, Master IDs, and a status.
   The dashboard's data tab itself is not changed. Names that could not be matched are listed in the
   master's **Tagging Gaps** tab; a run summary goes to **Tagging Log**.
3. **Tag automatically every morning** – daily at about 10:15 IST.

The Google account that runs it needs edit access to each dashboard spreadsheet in `SOURCES`.


## All-Agenda Review (meeting dashboard)

- `docs/` - the review page (`index.html`), its figures (`review.json`) and the day's district news clippings (`news/`).
- `build/` - the build: `fetch.py` (downloads every source with a Google service account), `make_news.py`, `build_review.py`,
  `update.sh` (runs it all), `fallback/` (yesterday's copy of each source, used when a sheet cannot be read), `headlines/` (optional Hindi headlines per paper).
- `.github/workflows/update-review.yml` - runs `build/update.sh` at 10:30 and 12:00 IST and commits the new figures.
  Needs the repository secret `GOOGLE_SA_KEY` (service account JSON key); every source in `build/sources.json` must be shared with that account as Viewer.
- `apps-script/review-web-app/` - the web app that serves the page (paste `Code.gs`, deploy as a web app).
