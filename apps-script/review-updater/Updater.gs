/**
 * All-Agenda Review — daily updater (no Claude needed)
 *
 * Runs as you at about 10:15 and 11:45 IST. It reads the dashboard sheets the review needs, the DM feedback
 * sheet, the pending-issues summary file and the newest news clippings, and saves them in one commit to the
 * GitHub repo. GitHub then rebuilds the review page by itself (about 2 minutes) and the web app shows it.
 *
 * One-time setup
 *  1. Project Settings > Script properties > Add: GITHUB_TOKEN = a GitHub fine-grained token for the repo
 *     mdsbpdcl-hub/all-agenda-review-dashboard with "Contents: Read and write".
 *  2. Run updateNow once and allow access. Check the log says "Saved to GitHub".
 *  3. Run installTriggers once. (Run removeTriggers to stop.)
 */
const CFG = {
  REPO: 'mdsbpdcl-hub/all-agenda-review-dashboard',
  BRANCH: 'main',
  TZ: 'Asia/Kolkata',
  RUN_AT: [[10, 15], [11, 45]],          // hour, minute (IST)
  // name: [spreadsheet id, tabs it needs]  (no tabs = every tab)
  SHEETS: {
    'c1912':         ['1MyjzcCz-IUs1GNEtKQHHohRwtG-FEhpFirb4uAWrguU', ['DB_CIRCLE', 'DB_DIVCAT', 'DB_DIVISION', 'DB_GEO', 'DB_KPI', 'DB_META', 'DB_REPEATS', 'DB_SECTION']],
    'outage':        ['19cDJxYkrnSSUBB30-IMztkg-9sCtZhxd0aD4iFo29rA', ['SBPDCL Summary', 'SBPDCL Circle', 'SBPDCL Division', 'SBPDCL Substation', 'SBPDCL Over 24 Hrs Left Out',
                                                                       'NBPDCL Summary', 'NBPDCL Circle', 'NBPDCL Division', 'NBPDCL Substation', 'NBPDCL Over 24 Hrs Left Out']],
    'revenue':       ['1nbkHaVm0eNDoz7xL5vRpQ60QqfZid3_1ZkMTjqcLLWo', ['SBPDCL_Sections', 'NBPDCL_Sections']],
    'nsc':           ['1lmx8BaG7WW_hHfG5d_zcbgeY6caHCkFVJNIV2_BOaMY', ['Analysis Summary', 'North - NSC <19 kW', 'South - NSC <19 kW', 'North - NSC 19-59 kW', 'South - NSC 19-59 kW', 'North - Load Change', 'South - Load Change']],
    'pmsg_regular':  ['1SGSTSGNwcolOtdoGjtKFv7NxN3MCh-05nM0j3jRtuuI', ['District KPIs', 'Division KPIs', 'Headline KPIs', 'Recent progress', 'Sub-division KPIs']],
    'pmsg_ula':      ['1aSsrArGFTplJsWZL5HAjoiAPgsjSQU3Vyd7F51uXxDM', ['CMDDashboard', 'Inspections']],
    'utility':       ['1hUYR8fhvOWFag3vuLL7to5Iqmybhy-OLeeLvz4qCJFw', []],
    'pss_projects':  ['1K6Mh4wfoJEW5IcSAaDc9YF0xN5WqGKFy_y_JQONvGdU', ['RDSSPSS', 'StatePlanPSS']],
    'electrocution': ['10HK2DmvM8JYcFpnQmO9SrKQwSrh0s4qbqY3xQpIEVm0', []],
    'offgrid':       ['1ZSyBMvGlivLk3sk5WBIjpQqhEhjIwSnkjaUHTVSUXG4', []],
    'cmd_cell':      ['1ykQTXA207fqsz3Tc_9tKF1xWN0FkoP0vnzGt3zua8Ek', []]
  },
  DM_FEEDBACK: ['18aCnpOGBAnBOGtDLhHGS6v4woMWccmTAmvRL7xYVOeg', 'Issues'],
  PENDING_JSON_ID: '1XS03gRoap-XsGAYXL1JazVFRtaRWAKNm',
  NEWS_PREPARED_FOLDER: '1aD89nFkBOVVDVPfFLFrg8UdPIu2H17Hi',
  IN_DIR: 'build/data_in/',               // where the inputs are saved in the repo
  NEWS_DIR: 'docs/news/'                  // where the clipping pictures are saved
};

/* ------------------------------------------------------------------ main */
function updateNow() {
  const files = {}, notes = [], t0 = Date.now();
  // 1. sheets -> JSON (one list of rows per tab)
  Object.keys(CFG.SHEETS).forEach(name => {
    try { files[CFG.IN_DIR + name + '.json'] = JSON.stringify(sheetToJson_(CFG.SHEETS[name][0], CFG.SHEETS[name][1], name)); }
    catch (e) { notes.push(name + ': could not be read (' + e.message + '), yesterday\'s copy is kept'); }
  });
  // 2. DM feedback (displayed values, as the bulletin reads them)
  try {
    const sh = SpreadsheetApp.openById(CFG.DM_FEEDBACK[0]).getSheetByName(CFG.DM_FEEDBACK[1]);
    files[CFG.IN_DIR + 'dm_feedback.json'] = JSON.stringify({ values: sh.getDataRange().getDisplayValues() });
  } catch (e) { notes.push('DM feedback: ' + e.message); }
  // 3. pending issues summary file
  try { files[CFG.IN_DIR + 'summary.json'] = DriveApp.getFileById(CFG.PENDING_JSON_ID).getBlob().getDataAsString('UTF-8'); }
  catch (e) { notes.push('Pending issues: ' + e.message); }
  // 4. news clippings (only when a newer paper has been prepared)
  let news = null;
  try { news = readNews_(); } catch (e) { notes.push('News: ' + e.message); }
  // 5. one commit
  const n = Object.keys(files).length;
  if (!n && !news) throw new Error('Nothing could be read: ' + notes.join('; '));
  const msg = 'Update review inputs ' + Utilities.formatDate(new Date(), CFG.TZ, 'dd MMM yyyy HH:mm') + (notes.length ? ' (' + notes.length + ' kept from before)' : '');
  commit_(files, news, msg);
  if (news) PropertiesService.getScriptProperties().setProperty('NEWS_KEY', news.key);
  Logger.log('Saved to GitHub: %s inputs%s in %s s. %s', n, news ? ' + ' + news.index.clips.length + ' clippings of ' + news.index.date : '', Math.round((Date.now() - t0) / 1000), notes.join(' | '));
  if (notes.length) mailOwner_('All-Agenda Review: some inputs were not read', notes.join('\n'));
}

/* ------------------------------------------------------------------ sheets */
function sheetToJson_(id, tabs, name) {
  const ss = SpreadsheetApp.openById(id), out = {};
  const sheets = tabs.length ? tabs.map(t => { const s = ss.getSheetByName(t); if (!s) throw new Error('tab "' + t + '" not found'); return s; }) : ss.getSheets();
  sheets.forEach(sh => {
    if (sh.getLastRow() < 1 || sh.getLastColumn() < 1) { out[sh.getName()] = []; return; }
    let rows = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getValues().map(r => r.map(cell_));
    while (rows.length && rows[rows.length - 1].every(c => c === null)) rows.pop();
    out[sh.getName()] = rows;
  });
  if (name === 'revenue') {                        // keep the latest day only (all the review reads)
    Object.keys(out).forEach(t => {
      const rows = out[t]; if (!rows.length) return;
      const i = rows[0].indexOf('Period'); if (i < 0) return;
      const key = v => v && v.__dt ? v.__dt.slice(0, 10) : String(v == null ? '' : v);
      let last = ''; rows.slice(1).forEach(r => { if (r[i] != null && key(r[i]) > last) last = key(r[i]); });
      out[t] = [rows[0]].concat(rows.slice(1).filter(r => key(r[i]) === last));
    });
  }
  return out;
}
function cell_(v) {
  if (v === '' || v === null || v === undefined) return null;
  if (Object.prototype.toString.call(v) === '[object Date]') {
    if (isNaN(v.getTime())) return null;
    if (v.getFullYear() < 1900) return Utilities.formatDate(v, CFG.TZ, 'HH:mm:ss');      // a time with no date
    return { __dt: Utilities.formatDate(v, CFG.TZ, "yyyy-MM-dd'T'HH:mm:ss") };
  }
  return v;
}

/* ------------------------------------------------------------------ news */
function readNews_() {
  const folder = DriveApp.getFolderById(CFG.NEWS_PREPARED_FOLDER);
  const it = folder.searchFiles("title contains '.manifest.json' and trashed = false");
  let man = null;
  while (it.hasNext()) { const f = it.next(); if (!man || f.getDateCreated() > man.getDateCreated()) man = f; }
  if (!man) return null;
  const key = man.getName().replace('.manifest.json', '');
  if (PropertiesService.getScriptProperties().getProperty('NEWS_KEY') === key) return null;   // same paper as last time
  const m = JSON.parse(man.getBlob().getDataAsString());
  const paperId = (key.match(/^prep_(.+)_\d+$/) || [])[1] || '';
  let date = '';
  try { date = dateFromName_(DriveApp.getFileById(paperId).getName()); } catch (e) { date = ''; }
  if (!date) date = Utilities.formatDate(man.getDateCreated(), CFG.TZ, 'dd MMM yyyy');
  const pics = {}, clips = [];
  m.parts.forEach(p => {
    const pi = folder.getFilesByName(p); if (!pi.hasNext()) throw new Error('part ' + p + ' missing');
    JSON.parse(pi.next().getBlob().getDataAsString()).clips.forEach(c => {
      const n = clips.length, page = Math.floor(+c.p || 0);
      const f = 'news/clip_' + ('0' + page).slice(-2) + '_' + n + '.jpg';
      pics[CFG.NEWS_DIR + f.slice(5)] = String(c.i).split(',').pop();         // base64 picture
      clips.push({ place: String(c.d || '').trim(), p: page, w: c.w, h: c.h, f: f });
    });
  });
  return { key: key, paper: paperId, pics: pics, index: { date: date, paper: 'BSPHCL News Tracking', paper_id: paperId, clips: clips } };
}
function dateFromName_(s) {
  const t = String(s || ''), MON = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  let m = t.match(/(\d{1,2})[ ._-](\d{1,2})[ ._-](20\d\d)/);
  if (m && +m[2] >= 1 && +m[2] <= 12) return ('0' + m[1]).slice(-2) + ' ' + MON[+m[2] - 1] + ' ' + m[3];
  m = t.match(/(20\d\d)[ ._-](\d{1,2})[ ._-](\d{1,2})/);
  if (m && +m[2] >= 1 && +m[2] <= 12) return ('0' + m[3]).slice(-2) + ' ' + MON[+m[2] - 1] + ' ' + m[1];
  return '';
}

/* ------------------------------------------------------------------ GitHub (one commit) */
function gh_(method, path, body) {
  const token = PropertiesService.getScriptProperties().getProperty('GITHUB_TOKEN');
  if (!token) throw new Error('Script property GITHUB_TOKEN is not set');
  const opt = { method: method, muteHttpExceptions: true, contentType: 'application/json',
                headers: { Authorization: 'Bearer ' + token, Accept: 'application/vnd.github+json', 'X-GitHub-Api-Version': '2022-11-28' } };
  if (body) opt.payload = JSON.stringify(body);
  for (let a = 0; a < 3; a++) {
    const r = UrlFetchApp.fetch('https://api.github.com/repos/' + CFG.REPO + path, opt), code = r.getResponseCode();
    if (code < 300) return JSON.parse(r.getContentText() || '{}');
    if (code !== 409 && code !== 502 && code < 500) throw new Error('GitHub ' + code + ' on ' + path + ': ' + r.getContentText().slice(0, 200));
    Utilities.sleep(1500 * (a + 1));
  }
  throw new Error('GitHub did not answer on ' + path);
}
function blob_(text, isBase64) {
  const content = isBase64 ? text : Utilities.base64Encode(text, Utilities.Charset.UTF_8);
  return gh_('post', '/git/blobs', { content: content, encoding: 'base64' }).sha;
}
function commit_(files, news, message) {
  for (let attempt = 0; attempt < 3; attempt++) {
    const head = gh_('get', '/git/ref/heads/' + CFG.BRANCH).object.sha;
    const baseTree = gh_('get', '/git/commits/' + head).tree.sha;
    const tree = [];
    Object.keys(files).forEach(p => tree.push({ path: p, mode: '100644', type: 'blob', sha: blob_(files[p], false) }));
    if (news) {
      const keep = {};
      Object.keys(news.pics).forEach(p => { keep[p] = 1; tree.push({ path: p, mode: '100644', type: 'blob', sha: blob_(news.pics[p], true) }); });
      tree.push({ path: CFG.NEWS_DIR + 'clips.json', mode: '100644', type: 'blob', sha: blob_(JSON.stringify(news.index), false) });
      // remove yesterday's pictures
      gh_('get', '/git/trees/' + baseTree + '?recursive=1').tree.forEach(e => {
        if (e.type === 'blob' && e.path.indexOf(CFG.NEWS_DIR + 'clip_') === 0 && !keep[e.path]) tree.push({ path: e.path, mode: '100644', type: 'blob', sha: null });
      });
    }
    const newTree = gh_('post', '/git/trees', { base_tree: baseTree, tree: tree }).sha;
    const c = gh_('post', '/git/commits', { message: message, tree: newTree, parents: [head] }).sha;
    try { gh_('patch', '/git/refs/heads/' + CFG.BRANCH, { sha: c }); return c; }
    catch (e) { if (!/ 422 /.test(' ' + e.message + ' ')) throw e; Utilities.sleep(2000); }   // someone pushed meanwhile: redo on top
  }
  throw new Error('Could not save to GitHub after 3 tries');
}

/* ------------------------------------------------------------------ schedule and mail */
function installTriggers() {
  removeTriggers();
  CFG.RUN_AT.forEach(t => ScriptApp.newTrigger('updateNow').timeBased().everyDays(1).atHour(t[0]).nearMinute(t[1]).inTimezone(CFG.TZ).create());
  Logger.log('Updates set for ' + CFG.RUN_AT.map(t => t[0] + ':' + ('0' + t[1]).slice(-2)).join(' and ') + ' IST every day.');
}
function removeTriggers() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'updateNow').forEach(t => ScriptApp.deleteTrigger(t));
}
function mailOwner_(subject, body) {
  try { MailApp.sendEmail(Session.getEffectiveUser().getEmail(), subject, body); } catch (e) { /* mail is optional */ }
}
