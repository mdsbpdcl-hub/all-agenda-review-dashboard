/**
 * All-Agenda Review — Apps Script web app
 * Serves the review page from the GitHub repo (docs/ folder). The figures are refreshed in the repo
 * at 10:30 and 12:00 IST by GitHub Actions, so this script never needs editing for daily updates.
 *
 * Deploy: Deploy > New deployment > Web app
 *   Execute as: Me    ·    Who has access: Anyone   (change ALLOWED_EMAILS below to restrict later)
 */
const CONFIG = {
  RAW_BASE: 'https://raw.githubusercontent.com/mdsbpdcl-hub/all-agenda-review-dashboard/main/docs/',
  TITLE: 'All-Agenda Review',
  CACHE_SECONDS: 300,                 // page shell cached 5 minutes; figures are always read fresh by the page
  // Leave empty for open access. To restrict: list Gmail IDs here, then redeploy with
  // Execute as: "User accessing the web app" and Who has access: "Anyone with Google account"
  // (when it runs as you, Google does not tell the script who the visitor is).
  ALLOWED_EMAILS: []
};

function doGet() {
  if (CONFIG.ALLOWED_EMAILS.length) {
    let who = '';
    try { who = (Session.getActiveUser().getEmail() || '').toLowerCase(); } catch (e) { who = ''; }
    if (CONFIG.ALLOWED_EMAILS.map(String).map(s => s.toLowerCase()).indexOf(who) < 0) {
      return HtmlService.createHtmlOutput(
        '<div style="font:16px Nunito,Segoe UI,sans-serif;max-width:420px;margin:15vh auto;padding:24px;border:1px solid #e7eaf0;border-radius:12px;text-align:center">' +
        '<b>This review page is not open to this account.</b><br><br>Ask the MD office to add your Gmail ID.</div>')
        .setTitle(CONFIG.TITLE);
    }
  }
  return HtmlService.createHtmlOutput(pageHtml_())
    .setTitle(CONFIG.TITLE)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** The page from the repo, with its figures and clippings pointed at the repo. */
function pageHtml_() {
  const cache = CacheService.getScriptCache();
  const hit = cache.get('page');
  if (hit) return hit;
  const res = UrlFetchApp.fetch(CONFIG.RAW_BASE + 'index.html?t=' + Date.now(), { muteHttpExceptions: true });
  if (res.getResponseCode() !== 200) {
    return '<div style="font:16px sans-serif;max-width:420px;margin:15vh auto;text-align:center">Could not load the review page. Please try again in a minute.</div>';
  }
  let html = res.getContentText('UTF-8');
  // relative addresses (review.json, news/...) resolve to the repo
  html = html.replace(/<head>/i, '<head><base href="' + CONFIG.RAW_BASE + '" target="_top">');
  // always read the newest figures (no stale copy)
  html = html.replace("fetch('review.json'", "fetch('review.json?t=' + Date.now()");
  try { if (html.length < 90000) cache.put('page', html, CONFIG.CACHE_SECONDS); } catch (e) { /* too large to cache: fine */ }
  return html;
}

/** Run once from the editor to check the repo can be reached. */
function testFetch() {
  const r = UrlFetchApp.fetch(CONFIG.RAW_BASE + 'review.json', { muteHttpExceptions: true });
  Logger.log('review.json: HTTP ' + r.getResponseCode() + ', ' + r.getContent().length + ' bytes');
}
