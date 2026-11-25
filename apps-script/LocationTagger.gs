/**
 * Location Tagger – Bihar Power Location Master
 * Install in the DistrictWiseReviewSheet-CMDLeveDashboard spreadsheet (Extensions > Apps Script).
 *
 * Menu "Location Master":
 *   1. Add new master rows from GitHub   – appends rows published in the repo (skips rows already present)
 *   2. Tag all dashboard sheets now      – writes an "IDs – <tab>" tab in every configured spreadsheet
 *   3. Tag automatically every morning   – daily trigger at about 10:15 IST (after the dashboards refresh)
 *
 * The tagger never changes a dashboard's own data tab. Each tagged tab gets a companion tab
 * "IDs – <tab name>" with one row per data row: the source row number, the names as written and the Master IDs.
 */

var REPO_RAW = 'https://raw.githubusercontent.com/mdsbpdcl-hub/all-agenda-review-dashboard/main/master_updates/';

/* ---------- which tabs to tag ----------
 * cols: level -> header text in that tab. A fixed DISCOM can be given as discom: 'SBPDCL'.
 * headerRow: row number of the header (1 by default). Levels are resolved top-down:
 * Circle, District, Division, Subdivision, Section, Substation.                                 */
var SOURCES = [
  {name: '1912 Complaint Analytics', id: '1MyjzcCz-IUs1GNEtKQHHohRwtG-FEhpFirb4uAWrguU', tabs: [
    {tab: 'DB_CIRCLE',   cols: {Discom: 'discom', Circle: 'circle'}},
    {tab: 'DB_DIVISION', cols: {Discom: 'discom', Circle: 'circle', Division: 'division'}},
    {tab: 'DB_SECTION',  cols: {Discom: 'discom', Circle: 'circle', Division: 'division', Subdivision: 'subdivision', Section: 'section'}}]},
  {name: 'Bihar Outage Analytics', id: '19cDJxYkrnSSUBB30-IMztkg-9sCtZhxd0aD4iFo29rA', tabs: [
    {tab: 'SBPDCL Division', cols: {Discom: 'Discom', Circle: 'Circle', Division: 'Division'}},
    {tab: 'NBPDCL Division', cols: {Discom: 'Discom', Circle: 'Circle', Division: 'Division'}}]},
  {name: 'Revenue Collection Analytics', id: '1nbkHaVm0eNDoz7xL5vRpQ60QqfZid3_1ZkMTjqcLLWo', tabs: [
    {tab: 'SBPDCL_Sections', discom: 'SBPDCL', cols: {Circle: 'Circle', Division: 'Division', Subdivision: 'Sub Division', Section: 'Section'}},
    {tab: 'NBPDCL_Sections', discom: 'NBPDCL', cols: {Circle: 'Circle', Division: 'Division', Subdivision: 'Sub Division', Section: 'Section'}}]},
  {name: 'NSC and Load Change', id: '1lmx8BaG7WW_hHfG5d_zcbgeY6caHCkFVJNIV2_BOaMY', tabs: [
    {tab: 'North - NSC <19 kW',   cols: {Discom: 'Discom', Circle: 'Circle', Division: 'Division', Subdivision: 'SubDivision', Section: 'Section'}},
    {tab: 'South - NSC <19 kW',   cols: {Discom: 'Discom', Circle: 'Circle', Division: 'Division', Subdivision: 'SubDivision', Section: 'Section'}},
    {tab: 'North - NSC 19-59 kW', cols: {Discom: 'Discom', Circle: 'Circle', Division: 'Division', Subdivision: 'SubDivision', Section: 'Section'}},
    {tab: 'South - NSC 19-59 kW', cols: {Discom: 'Discom', Circle: 'Circle', Division: 'Division', Subdivision: 'SubDivision', Section: 'Section'}},
    {tab: 'North - Load Change',  cols: {Discom: 'Discom', Circle: 'Circle', Division: 'Division'}},
    {tab: 'South - Load Change',  cols: {Discom: 'Discom', Circle: 'Circle', Division: 'Division'}}]},
  {name: 'PMSG-MBY (Regular)', id: '1SGSTSGNwcolOtdoGjtKFv7NxN3MCh-05nM0j3jRtuuI', tabs: [
    {tab: 'District KPIs',     cols: {Discom: 'DISCOM', District: 'District'}},
    {tab: 'Division KPIs',     cols: {Discom: 'DISCOM', District: 'District', Circle: 'Circle', Division: 'Division'}},
    {tab: 'Sub-division KPIs', cols: {Discom: 'DISCOM', District: 'District', Division: 'Division', Subdivision: 'Sub-division'}}]},
  {name: 'PMSG-MBY (ULA)', id: '1aSsrArGFTplJsWZL5HAjoiAPgsjSQU3Vyd7F51uXxDM', tabs: [
    {tab: 'CMDDashboard', cols: {District: 'Name of District'}}]},
  {name: 'Pending Utility Shifting', id: '1hUYR8fhvOWFag3vuLL7to5Iqmybhy-OLeeLvz4qCJFw', tabs: [
    {tab: 'UtilityShifting', discom: 'SBPDCL', cols: {Circle: 'Circle', Division: 'Division'}}]},
  {name: 'Project Monitoring - PSS', id: '1K6Mh4wfoJEW5IcSAaDc9YF0xN5WqGKFy_y_JQONvGdU', tabs: [
    {tab: 'RDSSPSS',      discom: 'SBPDCL', cols: {Circle: 'Circle', District: 'District', Division: 'Division', Substation: 'Name of PSS'}},
    {tab: 'StatePlanPSS', discom: 'SBPDCL', cols: {District: 'District', Substation: 'PSS Name'}}]},
  {name: 'Pending Electrocution Cases', id: '10HK2DmvM8JYcFpnQmO9SrKQwSrh0s4qbqY3xQpIEVm0', tabs: [
    {tab: 'Pending Cases', discom: 'SBPDCL', cols: {Circle: 'Circle', Division: 'Division'}}]},
  {name: 'OffGrid Electrification', id: '1ZSyBMvGlivLk3sk5WBIjpQqhEhjIwSnkjaUHTVSUXG4', tabs: [
    {tab: 'EnergyAccounting', headerRow: 3, cols: {District: 'District'}}]},
  {name: 'PSS Details', id: '1f2yrXhmQqi865OwT6ZvQgPbTzHWEkKWT6idHYfTAhdw', tabs: [
    {tab: 'PSS', discom: 'SBPDCL', cols: {Circle: 'Circle', District: 'District', Block: 'Block', Substation: 'Name of PSS', Section: 'Electric Supply Section'}}]},
  {name: 'PMSG DM Feedback', id: '18aCnpOGBAnBOGtDLhHGS6v4woMWccmTAmvRL7xYVOeg', tabs: [
    {tab: 'Issues', cols: {District: 'District'}}]},
  {name: 'Weather Prediction', id: '1TutxPDHeNdIuC0mGgTDn8Y5fpVmPaOqm2XYYcELiU9E', tabs: [
    {tab: 'Daily Circle',   cols: {Discom: 'Discom', Circle: 'Circle'}},
    {tab: 'Daily Division', cols: {Discom: 'Discom', Circle: 'Circle', Division: 'Division'}}]}
];

var LEVELS = ['Circle', 'District', 'Division', 'Subdivision', 'Section', 'Block', 'Substation'];

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Location Master')
    .addItem('Add new master rows from GitHub', 'importMasterUpdates')
    .addItem('Tag all dashboard sheets now', 'tagAll')
    .addItem('Tag automatically every morning', 'installDailyTrigger')
    .addToUi();
}

/* ======================= spelling rule (same as the master build) ======================= */
var PREFIX = /^(ESSD|ESS|ESD|ESC|SEC|SECTION|SUBDIVISION|SUB DIVISION|DIVISION|CIRCLE|PSS|33\/11 ?KV|33\/11KV|33 ?KV)[\s\-:.]+/;
var SUBS = [[/\(E\)|\bEAST\b|-E\b/g, ' EAST '], [/\(W\)|\bWEST\b|-W\b/g, ' WEST '], [/\(N\)|\bNORTH\b/g, ' NORTH '], [/\(S\)|\bSOUTH\b/g, ' SOUTH '],
  [/\(U\)|\bURBAN\b|-U\b|\bU\b/g, ' URBAN '], [/\(R\)|\bRURAL\b|-R\b|\bR\b/g, ' RURAL '], [/\bNO\.\s*|\bNO\s+(?=\d)/g, ' '],
  [/\b(III|3)\b/g, ' 3 '], [/\b(II|2)\b/g, ' 2 '], [/\b(I|1)\b/g, ' 1 '], [/\bJI\b/g, ' '],
  [/\bESD\b|\bESS\b|\bESC\b|\bDIVISION\b|\bCIRCLE\b|\bSECTION\b|\bSUB ?DIVISION\b/g, ' ']];
var PH = [['SHARIFF', 'SARIF'], ['SHARIF', 'SARIF'], ['ARRAH', 'ARA'], ['AARA', 'ARA'], ['PURNIA', 'PURNEA'], ['CHHAPRA', 'CHAPRA'],
  ['BHABHUA', 'BHABUA'], ['KOSHI', 'KOSI'], ['BAKHTIYARPUR', 'BAKHTIARPUR'], ['DEHRI', 'DEHARI'], ['JAGDEESHPUR', 'JAGDISHPUR'], ['PIRPAINTI', 'PIRPAINTY']];
function normKey(v) {
  var s = String(v === null || v === undefined ? '' : v).toUpperCase().trim().replace(/_/g, ' ');
  for (var i = 0; i < 2; i++) s = s.replace(PREFIX, '').trim();
  s = s.replace(/([\s-]NEW|\sN)$/, '').trim().replace(/&/g, ' AND ');
  SUBS.forEach(function (p) { s = s.replace(p[0], p[1]); });
  s = s.replace(/[^A-Z0-9 ]/g, ' ').split(/\s+/).filter(String).join(' ');
  PH.forEach(function (p) { s = s.split(p[0]).join(p[1]); });
  s = s.replace(/EE/g, 'I').replace(/OO/g, 'U').replace(/Y/g, 'I');
  s = s.replace(/([BCDGJKPST])H/g, '$1').replace(/(.)\1+/g, '$1').replace(/AA/g, 'A');
  return s.replace(/ /g, '');
}
var CIRCLE_ALIAS = {BOJPUR: 'ARA', ROHTAS: 'SASARAM', NALANDA: 'BIHARSARIF', BIHARSARIFNALANDA: 'BIHARSARIF',
  PATNAEAST: 'PESUEAST', PATNAWEST: 'PESUWEST', GAIAJI: 'GAIA'};
var DIST_ALIAS = {EASTCAMPARAN: 'PURBICAMPARAN', WESTCAMPARAN: 'PASCIMCAMPARAN', KAIMUR: 'KAIMURBABUA', BABUA: 'KAIMURBABUA',
  SASARAM: 'ROHTAS', MOTIHARI: 'PURBICAMPARAN', BETIAH: 'PASCIMCAMPARAN', CAPRA: 'SARAN', HAJIPUR: 'VAISALI', ARA: 'BOJPUR'};
function sig(k) { return (k.match(/EAST|WEST|NORT|SOUT|URBAN|RURAL|TOWN|\d+/g) || []).map(function (x) { return x === 'TOWN' ? 'URBAN' : x; }).join(','); }
function similar(a, b) {                       // share of matching characters, like difflib's ratio
  if (!a || !b) return 0;
  var m = 0, used = {};
  for (var i = 0; i < a.length; i++) for (var j = Math.max(0, i - 2); j < Math.min(b.length, i + 3); j++)
    if (!used[j] && a[i] === b[j]) { used[j] = 1; m++; break; }
  return 2 * m / (a.length + b.length);
}

/* ======================= master index ======================= */
function loadMaster_() {
  var ss = SpreadsheetApp.getActive();
  var M = {exact: {}, norm: {}, circleSec: {}, circleSub: {}, ssCircle: {}, ssAny: {}, divCircle: {}};
  var lk = ss.getSheetByName('Lookup').getDataRange().getValues();
  for (var r = 1; r < lk.length; r++) {
    var lv = lk[r][0], d = lk[r][1], p = lk[r][2], nm = lk[r][3], k = lk[r][4], id = lk[r][5];
    if (!id) continue;
    M.exact[k] = id;
    M.norm[lv + '|' + d + '|' + p + '|' + normKey(nm)] = id;
    if (lv === 'Substation') {                    // name-only index, for sheets that give no division
      var sk = d + '|' + normKey(String(nm).replace(/\(.*?\)/g, ''));
      M.ssAny[sk] = (M.ssAny[sk] && M.ssAny[sk] !== id) ? '#' : id;
    }
  }
  var dv = ss.getSheetByName('Division').getDataRange().getValues();
  for (r = 1; r < dv.length; r++) M.divCircle[dv[r][0]] = dv[r][2];
  [['Section', 'circleSec'], ['Subdivision', 'circleSub']].forEach(function (t) {
    var v = ss.getSheetByName(t[0]).getDataRange().getValues();
    for (var i = 1; i < v.length; i++) {
      var key = v[i][1] + '|' + v[i][2] + '|' + normKey(v[i][t[0] === 'Section' ? 8 : 6]);
      if (!M[t[1]][key]) M[t[1]][key] = v[i][0];
    }
  });
  var sv = ss.getSheetByName('Substation').getDataRange().getValues();
  for (r = 1; r < sv.length; r++) {
    var c = sv[r][4]; if (!c) continue;
    M.ssCircle[c + '|' + normKey(String(sv[r][3]).replace(/\(.*?\)/g, ''))] = sv[r][0];
  }
  return M;
}

function find_(M, level, discom, parent, raw) {
  if (raw === '' || raw === null || raw === undefined) return '';
  var up = String(raw).trim().toUpperCase().replace(/\s+/g, ' ');
  var id = M.exact[level + '|' + discom + '|' + parent + '|' + up];
  if (id) return id;
  var k = normKey(raw);
  if (level === 'Circle') k = CIRCLE_ALIAS[k] || k;
  if (level === 'District') k = DIST_ALIAS[k] || k;
  id = M.norm[level + '|' + discom + '|' + parent + '|' + k];
  if (id) return id;
  if (level === 'Circle' || level === 'District') {              // try the alias target's own spelling
    for (var nk in M.norm) if (nk.indexOf(level + '|' + discom + '|' + parent + '|') === 0 && nk.split('|')[3] === k) return M.norm[nk];
  }
  return '';
}

function fuzzyIn_(map, prefix, k, thr) {     // close spelling inside one circle; never across a North/South or I/II difference
  var best = '', bs = 0;
  for (var key in map) {
    if (key.indexOf(prefix) !== 0) continue;
    var kk = key.substring(prefix.length);
    if (sig(kk) !== sig(k)) continue;
    var sc = similar(k, kk);
    if (sc > bs) { bs = sc; best = map[key]; }
  }
  return bs >= thr ? best : '';
}

function resolveRow_(M, cfg, row, hdr) {
  var col = function (lv) { var h = cfg.cols[lv]; if (!h) return null; var i = hdr.indexOf(h); return i < 0 ? null : row[i]; };
  var d = cfg.discom || String(col('Discom') || '').toUpperCase();
  d = /SB|SOUTH/.test(d) ? 'SBPDCL' : (/NB|NORTH/.test(d) ? 'NBPDCL' : '');
  var out = {}, raw;
  raw = col('Circle'); if (raw !== null) {
    out.Circle = d ? find_(M, 'Circle', d, '', raw) : (find_(M, 'Circle', 'SBPDCL', '', raw) || find_(M, 'Circle', 'NBPDCL', '', raw));
    if (!d && out.Circle) d = out.Circle.indexOf('-SB-') > 0 ? 'SBPDCL' : 'NBPDCL';
  }
  raw = col('District'); if (raw !== null) out.District = find_(M, 'District', '', '', raw);
  raw = col('Division'); if (raw !== null) {
    var parts = String(raw).split(/[\n,]+/).map(function (x) { return x.trim(); }).filter(String);
    out.Division = parts.map(function (p) { return find_(M, 'Division', d, '', p); }).filter(String).join(', ');
  }
  var div = (out.Division || '').split(', ')[0];
  var cir = out.Circle || M.divCircle[div] || '';
  raw = col('Subdivision'); if (raw !== null) {
    var nm = String(raw).replace(/\s+[RU]$/, '');
    out.Subdivision = (div && find_(M, 'Subdivision', d, div, nm)) || M.circleSub[d + '|' + cir + '|' + normKey(nm)] ||
      (cir ? fuzzyIn_(M.circleSub, d + '|' + cir + '|', normKey(nm), 0.88) : '');
  }
  raw = col('Section'); if (raw !== null) {
    out.Section = (div && find_(M, 'Section', d, div, raw)) || M.circleSec[d + '|' + cir + '|' + normKey(raw)] ||
      (cir ? fuzzyIn_(M.circleSec, d + '|' + cir + '|', normKey(raw), 0.86) : '');
  }
  raw = col('Block'); if (raw !== null && out.District) out.Block = find_(M, 'Block', '', out.District, raw);
  raw = col('Substation'); if (raw !== null) {
    out.Substation = (div && find_(M, 'Substation', d, div, raw)) || find_(M, 'Substation', d, '', raw) ||
      M.ssCircle[cir + '|' + normKey(String(raw).replace(/\(.*?\)/g, ''))] ||
      (M.ssAny[d + '|' + normKey(String(raw).replace(/\(.*?\)/g, ''))] || '').replace('#', '') || '';
  }
  out.Discom = d;
  return out;
}

/* ======================= tagging ======================= */
function tagAll() {
  var M = loadMaster_(), gaps = [], log = [];
  SOURCES.forEach(function (src) {
    var book;
    try { book = SpreadsheetApp.openById(src.id); } catch (e) { log.push([src.name, '', 'Cannot open: ' + e.message]); return; }
    src.tabs.forEach(function (cfg) {
      var sh = book.getSheetByName(cfg.tab);
      if (!sh) { log.push([src.name, cfg.tab, 'Tab not found']); return; }
      var hr = cfg.headerRow || 1, vals = sh.getDataRange().getValues();
      if (vals.length <= hr) return;
      var hdr = vals[hr - 1].map(function (h) { return String(h).trim(); });
      var levels = LEVELS.filter(function (lv) { return cfg.cols[lv]; });
      var head = ['Source row', 'Discom ID'];
      levels.forEach(function (lv) { head.push(lv + ' (as written)', lv + ' ID'); });
      head.push('Status');
      var rows = [head], found = 0, total = 0;
      for (var r = hr; r < vals.length; r++) {
        var row = vals[r];
        if (row.join('') === '') continue;
        var o = resolveRow_(M, cfg, row, hdr), miss = [], line = [r + 1, o.Discom ? (o.Discom === 'SBPDCL' ? 'DC-SB' : 'DC-NB') : ''];
        levels.forEach(function (lv) {
          var rawv = row[hdr.indexOf(cfg.cols[lv])];
          line.push(rawv === undefined ? '' : rawv, o[lv] || '');
          if (rawv !== '' && rawv !== undefined && !o[lv] && !/total|not stated|not identified/i.test(String(rawv))) {
            miss.push(lv); gaps.push([src.name, cfg.tab, r + 1, lv, rawv, o.Discom || '']);
          }
        });
        line.push(miss.length ? 'Not found: ' + miss.join(', ') : 'All found');
        total++; if (!miss.length) found++;
        rows.push(line);
      }
      var name = ('IDs – ' + cfg.tab).substring(0, 99), out = book.getSheetByName(name) || book.insertSheet(name);
      out.clearContents();
      out.getRange(1, 1, rows.length, head.length).setValues(rows);
      out.getRange(1, 1, 1, head.length).setFontWeight('bold').setBackground('#1F4E79').setFontColor('#FFFFFF');
      out.setFrozenRows(1);
      log.push([src.name, cfg.tab, found + ' of ' + total + ' rows fully tagged']);
    });
  });
  writeTab_('Tagging Gaps', [['Dashboard', 'Tab', 'Row', 'Level', 'Name as written', 'Discom']].concat(gaps));
  writeTab_('Tagging Log', [['Dashboard', 'Tab', 'Result']].concat(log).concat([['', '', 'Run at ' + new Date()]]));
}

function writeTab_(name, rows) {
  var ss = SpreadsheetApp.getActive(), sh = ss.getSheetByName(name) || ss.insertSheet(name);
  sh.clearContents();
  if (rows.length) sh.getRange(1, 1, rows.length, rows[0].length).setValues(rows.map(function (r) {
    while (r.length < rows[0].length) r.push(''); return r; }));
  sh.getRange(1, 1, 1, rows[0].length).setFontWeight('bold').setBackground('#1F4E79').setFontColor('#FFFFFF');
  sh.setFrozenRows(1);
}

function installDailyTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'tagAll') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('tagAll').timeBased().everyDays(1).atHour(10).nearMinute(15).inTimezone('Asia/Kolkata').create();
  SpreadsheetApp.getUi().alert('Tagging will run every day at about 10:15 AM.');
}

/* ======================= master rows from GitHub ======================= */
function importMasterUpdates() {
  var ss = SpreadsheetApp.getActive(), report = [];
  [['lookup_add.csv', 'Lookup', 4], ['substation_add.csv', 'Substation', 0], ['section_add.csv', 'Section', 0]].forEach(function (f) {
    var csv = Utilities.parseCsv(UrlFetchApp.fetch(REPO_RAW + f[0]).getContentText('UTF-8'));
    var sh = ss.getSheetByName(f[1]), have = {};
    sh.getRange(2, f[2] + 1, Math.max(sh.getLastRow() - 1, 1), 1).getValues().forEach(function (v) { have[v[0]] = 1; });
    var add = csv.filter(function (r) { return r.length > 1 && r[f[2]] && !have[r[f[2]]]; });
    if (add.length) {
      var width = Math.max.apply(null, add.map(function (r) { return r.length; }));
      add = add.map(function (r) { r = r.slice(); while (r.length < width) r.push(''); return r.map(function (c) { return /^[=+\-]/.test(c) ? "'" + c : c; }); });
      if (sh.getMaxRows() < sh.getLastRow() + add.length) sh.insertRowsAfter(sh.getMaxRows(), add.length + 20);
      sh.getRange(sh.getLastRow() + 1, 1, add.length, width).setValues(add);
    }
    report.push(f[1] + ': ' + add.length + ' rows added');
  });
  SpreadsheetApp.getUi().alert(report.join('\n'));
}
