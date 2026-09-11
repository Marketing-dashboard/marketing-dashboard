// ============================================================
// Enterprise Dashboard — Web App + Auto-Updater v2
//
// WEB APP MODE (primary):
//   Deploy as a Web App → doGet() reads all sheets live and
//   serves the dashboard HTML with fresh data on every request.
//   No GitHub token required.
//
// GITHUB PUSH MODE (optional cache):
//   Run updateEnterpriseDashboard() to push a static snapshot
//   to GitHub Pages for offline / backup access.
// ============================================================

// ── CONFIGURATION ─────────────────────────────────────────────
var REPO_OWNER = 'Marketing-dashboard';
var REPO_NAME  = 'marketing-dashboard';
var FILE_PATH  = 'enterprise_dashboard.html';

// Raw URL of the HTML template on GitHub (used by doGet)
var TEMPLATE_URL = 'https://raw.githubusercontent.com/' + REPO_OWNER + '/' + REPO_NAME + '/main/' + FILE_PATH;

// Google Sheet IDs
var ENTERPRISE_SS_ID  = '1c5mtbsRiA6axOKMA87KIPiCUvnjAXZGSELW2-jLt-Fk';
var SEP_TRIGGERS_ID   = '1KXaUoUsEIrMHajuMGegVSYCnfeVmCLEfQFX9eujavp0';
var AUG_TRIGGERS_ID   = '1kD6z8CpeaII8Ls545_ZSLv4hqQrwUeWoPWgegtzyka8';

// Sheet / tab names
var SHEET_RAW     = 'Raw';
var SHEET_CPL_VAL = 'Sold_CPL_Val';
var SHEET_TRIG    = 'Triggers';

// Raw tab columns (0-based)
// Month(0) Day(1) Account(2) Campaign(3) CampType(4) Currency(5) Cost(6) Conversions(7) Brand(8) Model(9) Source(10)
var RC = {MO:0, DAY:1, COST:6, LEADS:7, BRAND:8, MODEL:9, SRC:10};

// Sold_CPL_Val tab columns (0-based)
// Brand(0) Model(1) Segment(2) Channel(3) Validation%(4) SoldCPL(5) Month(6)
var CC = {BR:0, MD:1, SG:2, CH:3, VP:4, SC:5, MO:6};

// Triggers tab columns (0-based)
// Date(0) utm_campaign(1) final_source(2) brand(3) model(4) Triggered(5) Triggered_in_List_ID(6)
// Type(7) CampaignType(8) Brand_Mapped(9) Model_Mapped(10)
var TC = {DT:0, SRC:2, BR:3, MD:4, TRIG:5, TRIG_LIST:6, BR_MAP:9, MD_MAP:10};

// Brands that use Triggered + Triggered_in_List_ID (others use only Triggered_in_List_ID)
var SPECIAL_BRANDS = ['JLR','CITROEN','LEXUS'];

// ── TOKEN SETUP (run once, only needed for GitHub push mode) ───
function setGitHubToken(token) {
  PropertiesService.getScriptProperties().setProperty('GITHUB_TOKEN', token);
  Logger.log('GitHub token saved.');
}

// ── STEP 1: PROCESS & CACHE (run once from script editor) ────
// Reads all 3 sheets, processes the data, saves to Drive.
// Takes ~2 minutes. Run this whenever data needs refreshing.
function processAndCacheData() {
  var rows    = buildRows();
  var dataRaw = serializeRows(rows);
  var buildTs = new Date().toISOString();

  var payload = JSON.stringify({dataRaw: dataRaw, buildTs: buildTs, rowCount: rows.length});
  var fname   = 'enterprise_dashboard_cache.json';
  var files   = DriveApp.getFilesByName(fname);
  var fileId;

  if (files.hasNext()) {
    var f = files.next();
    f.setContent(payload);
    fileId = f.getId();
  } else {
    fileId = DriveApp.createFile(fname, payload, 'application/json').getId();
  }
  PropertiesService.getScriptProperties().setProperty('CACHE_FILE_ID', fileId);
  Logger.log('Cache saved: ' + rows.length + ' rows | Build: ' + buildTs + ' | File: ' + fileId);
}

// ── STEP 2: WEB APP (doGet serves from cache in ~2 seconds) ──
// After processAndCacheData() completes, every URL open is instant.
function doGet(e) {
  try {
    var fileId = PropertiesService.getScriptProperties().getProperty('CACHE_FILE_ID');
    if (!fileId) {
      return HtmlService.createHtmlOutput(
        '<div style="font-family:Arial,sans-serif;padding:48px;text-align:center;color:#64748b">' +
        '<h2 style="color:#0f3460">No cached data yet</h2>' +
        '<p>Open the Apps Script editor and run <strong>processAndCacheData()</strong> first.<br>' +
        'It takes ~2 minutes to process all sheets, then this URL will load instantly.</p></div>'
      );
    }

    var cache   = JSON.parse(DriveApp.getFileById(fileId).getBlob().getDataAsString());
    var resp    = UrlFetchApp.fetch(TEMPLATE_URL, {muteHttpExceptions: true});
    if (resp.getResponseCode() !== 200)
      return HtmlService.createHtmlOutput('<pre>Template fetch failed: ' + resp.getResponseCode() + '</pre>');

    var html = resp.getContentText();
    html = replaceBlock(html, 'var DATA_RAW = [', cache.dataRaw);
    html = html.replace(/var BUILD_TS = '[^']*';/, "var BUILD_TS = '" + cache.buildTs + "';");

    return HtmlService.createHtmlOutput(html)
      .setTitle('Enterprise Dashboard')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch(err) {
    return HtmlService.createHtmlOutput('<pre style="color:red;padding:24px">Error: ' + err.message + '</pre>');
  }
}

// ── GITHUB PUSH MODE (optional static snapshot) ───────────────
function updateEnterpriseDashboard() {
  var token = PropertiesService.getScriptProperties().getProperty('GITHUB_TOKEN');
  if (!token) throw new Error('GitHub token not set. Run setGitHubToken("ghp_...") first.');

  var rows    = buildRows();
  var buildTs = new Date().toISOString();
  var dataRaw = serializeRows(rows);

  pushToGitHub(token, dataRaw, buildTs);
  Logger.log('Done. ' + rows.length + ' rows pushed. Build: ' + buildTs);
}

// ── CORE DATA BUILDER ─────────────────────────────────────────
// Returns array of row objects with all fields merged from all sheets.
function buildRows() {
  Logger.log('Reading spreadsheets…');

  var entSS   = SpreadsheetApp.openById(ENTERPRISE_SS_ID);
  var rawData = entSS.getSheetByName(SHEET_RAW).getDataRange().getValues();
  var cplData = entSS.getSheetByName(SHEET_CPL_VAL).getDataRange().getValues();

  var sepSS   = SpreadsheetApp.openById(SEP_TRIGGERS_ID);
  var sepData = sepSS.getSheetByName(SHEET_TRIG).getDataRange().getValues();

  var augSS   = SpreadsheetApp.openById(AUG_TRIGGERS_ID);
  var augData = augSS.getSheetByName(SHEET_TRIG).getDataRange().getValues();

  Logger.log('Raw:'+rawData.length+' CPL:'+cplData.length+' Sep:'+sepData.length+' Aug:'+augData.length);

  // Build Sold_CPL_Val lookup
  // key: normStr(brand)||normStr(model)||channel_norm||month_label  →  {seg, vp, sc}
  var cplMap = {};
  for (var i = 1; i < cplData.length; i++) {
    var c  = cplData[i];
    var br = trim(c[CC.BR]), md = trim(c[CC.MD]), sg = trim(c[CC.SG]);
    var ch = normCh(trim(c[CC.CH]));
    var vp = parseVP(c[CC.VP]), sc = parseSC(c[CC.SC]);
    var mo = normMo(trim(c[CC.MO]));
    if (!br || !md || !ch || !mo) continue;
    var brKey = normStr(br), mdKey = normStr(md);
    cplMap[brKey+'||'+mdKey+'||'+ch+'||'+mo] = {seg:sg, vp:vp, sc:sc};
  }

  // Build trigger maps
  var sepTrigMap = buildTrigMap(sepData);
  var augTrigMap = buildTrigMap(augData);
  Logger.log('CPL map:'+Object.keys(cplMap).length+' Sep trig:'+Object.keys(sepTrigMap).length+' Aug trig:'+Object.keys(augTrigMap).length);

  // Aggregate Raw spends by date + normStr(brand) + normStr(model) + channel
  var rawMap = {};
  for (var j = 1; j < rawData.length; j++) {
    var r  = rawData[j];
    var mo = normMo(trim(r[RC.MO]));
    var sp = parseFloat(r[RC.COST])  || 0;
    var ld = parseFloat(r[RC.LEADS]) || 0;
    var br = trim(r[RC.BRAND]);
    var md = trim(r[RC.MODEL]);
    var ch = normCh(trim(r[RC.SRC]));
    if (!r[RC.DAY] || !br || !mo) continue;
    if (br === '#N/A' || normStr(br) === 'n/a' || normStr(br) === '#n/a') continue;
    if (!sp && !ld) continue;
    var dt = fmtDate(r[RC.DAY]);
    if (!dt) continue;
    var key = dt+'||'+normStr(br)+'||'+normStr(md)+'||'+ch;
    if (!rawMap[key]) rawMap[key] = {dt:dt, mo:mo, br:br, md:md, ch:ch, sp:0, ld:0};
    rawMap[key].sp += sp;
    rawMap[key].ld += ld;
  }
  Logger.log('Raw map:'+Object.keys(rawMap).length);

  // Merge: join spends with triggers + CPL validation (use normalized keys for lookup)
  var rows = [];
  Object.keys(rawMap).forEach(function(key) {
    var r = rawMap[key];
    var brKey = normStr(r.br), mdKey = normStr(r.md);
    var trigMap = r.mo === 'Sep' ? sepTrigMap : (r.mo === 'Aug' ? augTrigMap : {});
    var triggers = (trigMap[r.dt+'||'+brKey+'||'+mdKey+'||'+r.ch] || {}).tr || 0;
    var cplInfo  = cplMap[brKey+'||'+mdKey+'||'+r.ch+'||'+r.mo] || {};
    rows.push({
      dt:r.dt, mo:r.mo, br:r.br, md:r.md,
      sg:cplInfo.seg||'', ch:r.ch,
      sp:round2(r.sp), ld:round2(r.ld), tr:triggers,
      vp:cplInfo.vp!=null?cplInfo.vp:0, sc:cplInfo.sc||0
    });
  });

  rows.sort(function(a,b){
    return a.dt.localeCompare(b.dt)||a.br.localeCompare(b.br)||a.md.localeCompare(b.md)||a.ch.localeCompare(b.ch);
  });
  Logger.log('Final rows:'+rows.length);
  return rows;
}

// Serialize rows into the DATA_RAW block string
function serializeRows(rows) {
  var lines = rows.map(function(r){
    return '{dt:'+jstr(r.dt)+',mo:'+jstr(r.mo)+',br:'+jstr(r.br)+',md:'+jstr(r.md)
          +',sg:'+jstr(r.sg)+',ch:'+jstr(r.ch)+',sp:'+r.sp+',ld:'+r.ld
          +',tr:'+r.tr+',vp:'+r.vp+',sc:'+r.sc+'}';
  });
  return 'var DATA_RAW = [\n' + lines.join(',\n') + '\n];';
}

// ── TRIGGER MAP ────────────────────────────────────────────────
function buildTrigMap(data) {
  var map = {};
  for (var i = 1; i < data.length; i++) {
    var r = data[i];
    var dateVal = r[TC.DT];
    if (!dateVal) continue;
    var dt = fmtDate(dateVal);
    if (!dt) continue;

    // Use Brand_Mapped if present, fall back to brand column; normalize for join
    var brRaw = trim(r[TC.BR_MAP]) || trim(r[TC.BR]);
    var mdRaw = trim(r[TC.MD_MAP]) || trim(r[TC.MD]);
    var brKey = normStr(brRaw);
    var mdKey = normStr(mdRaw);
    var ch = normCh(trim(r[TC.SRC]));

    var triggered     = parseInt(r[TC.TRIG])      || 0;
    var triggeredList = parseInt(r[TC.TRIG_LIST]) || 0;

    // Special brands (JLR/Citroen/Lexus): Triggered + Triggered_in_List_ID
    // Standard brands: Triggered only
    var isSpecial = ['jlr','citroen','lexus'].indexOf(brKey) !== -1;
    var count = isSpecial ? (triggered + triggeredList) : triggered;

    if (!brKey || !ch || count <= 0) continue;

    var key = dt+'||'+brKey+'||'+mdKey+'||'+ch;
    if (!map[key]) map[key] = {tr:0};
    map[key].tr += count;
  }
  return map;
}

// ── SOURCE / CHANNEL NORMALIZATION ────────────────────────────
function normCh(s) {
  if (!s) return '';
  var l = s.toLowerCase().replace(/[-_\s]/g,'');
  if (l.indexOf('adword')!==-1||l.indexOf('google')!==-1||l==='ga'||l.endsWith('ga')||l==='mediasalesga') return 'adwords';
  if (l.indexOf('ctwa')!==-1||l.indexOf('whatsapp')!==-1||l.indexOf('wha')!==-1) return 'ctwa';
  if (l.indexOf('fb')!==-1||l.indexOf('facebook')!==-1||l.indexOf('meta')!==-1) return 'fb';
  return l;
}

// ── MONTH NORMALIZATION ────────────────────────────────────────
function normMo(s) {
  var m = String(s||'').toLowerCase();
  if (m.indexOf('aug')!==-1) return 'Aug';
  if (m.indexOf('sep')!==-1) return 'Sep';
  if (m.indexOf('oct')!==-1) return 'Oct';
  if (m.indexOf('jul')!==-1) return 'Jul';
  if (m.indexOf('jun')!==-1) return 'Jun';
  return s;
}

// ── VALIDATION % PARSE ────────────────────────────────────────
// Cells may be stored as decimal (0.98) or number (98) or string "98%"
function parseVP(v) {
  if (v == null || v === '') return 0;
  if (typeof v === 'number') {
    // Decimal ≤1 means it's a fraction (e.g. 0.98 → 98%)
    return (v > 0 && v <= 1) ? v * 100 : v;
  }
  var s = String(v).trim();
  if (s.indexOf('%') !== -1) return parseFloat(s) || 0;
  var n = parseFloat(s) || 0;
  return (n > 0 && n <= 1) ? n * 100 : n;
}

// ── SOLD CPL PARSE ─────────────────────────────────────────────
// Handle formats like "₹8,835" or 8835 or "8,835"
function parseSC(v) {
  if (v == null || v === '') return 0;
  if (typeof v === 'number') return v;
  return parseFloat(String(v).replace(/[₹,\s]/g,'')) || 0;
}

// ── UTILITIES ──────────────────────────────────────────────────
function trim(v) { return String(v||'').trim(); }
function normStr(s) { return String(s||'').trim().toLowerCase(); }  // for join keys only
function round2(n) { return Math.round(n * 100) / 100; }
function jstr(s) { return JSON.stringify(String(s||'')); }

function fmtDate(val) {
  if (!val) return '';
  var tz = Session.getScriptTimeZone();
  if (val instanceof Date) return Utilities.formatDate(val, tz, 'yyyy-MM-dd');
  var s = String(val).trim();
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.substring(0,10);
  var d = new Date(s);
  if (!isNaN(d)) return Utilities.formatDate(d, tz, 'yyyy-MM-dd');
  return '';
}

// ── GITHUB PUSH ───────────────────────────────────────────────
function pushToGitHub(token, newDataRaw, buildTs) {
  var apiUrl = 'https://api.github.com/repos/' + REPO_OWNER + '/' + REPO_NAME + '/contents/' + FILE_PATH;

  var getResp = UrlFetchApp.fetch(apiUrl, {
    headers: {'Authorization':'token '+token,'Accept':'application/vnd.github.v3+json'},
    muteHttpExceptions: true
  });
  if (getResp.getResponseCode() !== 200)
    throw new Error('GitHub GET failed (' + getResp.getResponseCode() + '): ' + getResp.getContentText());

  var info = JSON.parse(getResp.getContentText());
  var current = Utilities.newBlob(Utilities.base64Decode(info.content.replace(/\n/g,''))).getDataAsString();

  var updated = replaceBlock(current, 'var DATA_RAW = [', newDataRaw);
  updated = updated.replace(/var BUILD_TS = '[^']*';/, "var BUILD_TS = '" + buildTs + "';");

  var putResp = UrlFetchApp.fetch(apiUrl, {
    method: 'PUT',
    headers: {'Authorization':'token '+token,'Content-Type':'application/json','Accept':'application/vnd.github.v3+json'},
    payload: JSON.stringify({
      message: 'Auto-update enterprise data [' + new Date().toISOString().split('T')[0] + ']',
      content: Utilities.base64Encode(Utilities.newBlob(updated).getBytes()),
      sha: info.sha, branch: 'main'
    }),
    muteHttpExceptions: true
  });

  if (putResp.getResponseCode() !== 200 && putResp.getResponseCode() !== 201)
    throw new Error('GitHub PUT failed (' + putResp.getResponseCode() + '): ' + putResp.getContentText());

  Logger.log('Pushed to GitHub: ' + putResp.getResponseCode());
}

function replaceBlock(content, marker, replacement) {
  var start = content.indexOf(marker);
  if (start === -1) { Logger.log('WARN: marker not found: ' + marker); return content; }
  var end = content.indexOf('];', start + marker.length);
  if (end === -1) { Logger.log('WARN: closing ]; not found for: ' + marker); return content; }
  return content.substring(0, start) + replacement + content.substring(end + 2);
}

// ── TRIGGER SETUP ─────────────────────────────────────────────
function createDailyTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t){
    if (t.getHandlerFunction()==='updateEnterpriseDashboard') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('updateEnterpriseDashboard').timeBased().everyDays(1).atHour(10).create();
  Logger.log('Daily trigger created (10 AM).');
}
