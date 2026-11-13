// ============================================================
// NCBD Requirement vs Generation — GitHub Pages auto-updater
//
// Reads the "Req VS Gen" tab of the NCBD Requirement VS Generation
// sheet, turns each block (NCBD, Overall NCD, NCD_DCB-Facebook,
// NCD_DCB-Google-Ads, NBD_DCB-Facebook) into one row per day, and
// rewrites the DATA_RAW / BUILD_TS blocks in ncbd_req_vs_gen.html
// on GitHub.
//
// Setup (once):
//   1. Paste this file into an Apps Script project (Extensions →
//      Apps Script from the sheet, or any project with access).
//   2. Run setGitHubToken('<token with repo contents write>').
//   3. Run updateNcbdDashboard() once to test, then createDailyTrigger().
// ============================================================

// ── CONFIGURATION ─────────────────────────────────────────────
var REPO_OWNER = 'Marketing-dashboard';
var REPO_NAME  = 'marketing-dashboard';
var FILE_PATH  = 'ncbd_req_vs_gen.html';

var NCBD_SS_ID = '1ViRpguJzGsNpKOC50-9SoIrTEx2dASe9ZU7fuh-bLRQ';
var NCBD_TAB   = 'Req VS Gen';

var MONTHS = {Jan:1, Feb:2, Mar:3, Apr:4, May:5, Jun:6, Jul:7, Aug:8, Sep:9, Oct:10, Nov:11, Dec:12};

function setGitHubToken(token) {
  PropertiesService.getScriptProperties().setProperty('GITHUB_TOKEN', token);
  Logger.log('GitHub token saved.');
}

// ── MAIN ──────────────────────────────────────────────────────
function updateNcbdDashboard() {
  var token = PropertiesService.getScriptProperties().getProperty('GITHUB_TOKEN');
  if (!token) throw new Error('Run setGitHubToken() first.');

  var rows = buildRows();
  if (!rows.length) throw new Error('No segment blocks found in "' + NCBD_TAB + '".');

  var dataRaw = 'var DATA_RAW = [\n' + rows.map(function(r) {
    return '{s:' + JSON.stringify(r.s) + ',d:"' + r.d + '",r:' + r.r + ',g:' + r.g + '}';
  }).join(',\n') + '\n];';

  pushToGitHub(token, dataRaw, new Date().toISOString());
  Logger.log('Pushed ' + rows.length + ' rows.');
}

// Reads the sheet as displayed (same as the CSV export) and returns
// [{s: segment, d: 'YYYY-MM-DD', r: requirement, g: generation}]
function buildRows() {
  var ss = SpreadsheetApp.openById(NCBD_SS_ID);
  var sh = ss.getSheetByName(NCBD_TAB) || ss.getSheets()[0];
  var v = sh.getDataRange().getDisplayValues();
  var out = [];
  var title = '';

  for (var i = 0; i < v.length; i++) {
    var row = v[i];
    var first = String(row[0] || '').trim();
    var b = String(row[1] || '').trim();

    if (/^[A-Z][a-z]{2} \d{2}$/.test(b)) {          // header row: "Aug 26", "01-Aug-26", ...
      var seg = first || title;
      var req = null, gen = null;
      for (var k = i + 1; k < v.length && k <= i + 4; k++) {
        var lab = String(v[k][0] || '').trim().toLowerCase();
        if (lab === 'requirement') req = v[k];
        else if (lab === 'generation') gen = v[k];
        else if (!lab) break;
      }
      if (!req || !gen) continue;
      for (var j = 1; j < row.length; j++) {
        var d = parseDay(row[j]);
        if (d) out.push({s: seg, d: d, r: num(req[j]), g: num(gen[j])});
      }
    } else if (first && !/^(requirement|generation|req vs gen %)$/i.test(first)) {
      title = first;                                 // block title row: "NCBD", "Overall NCD"
    }
  }
  return out;
}

// "01-Aug-26" → "2026-08-01"; month-total headers ("Aug 26") return null
function parseDay(h) {
  var m = /^(\d{1,2})-([A-Z][a-z]{2})-(\d{2})$/.exec(String(h || '').trim());
  if (!m || !MONTHS[m[2]]) return null;
  return '20' + m[3] + '-' + ('0' + MONTHS[m[2]]).slice(-2) + '-' + ('0' + m[1]).slice(-2);
}

function num(s) {
  var n = parseFloat(String(s == null ? '' : s).replace(/[,%\s]/g, ''));
  return isNaN(n) ? 0 : Math.round(n * 100) / 100;
}

// ── GITHUB PUSH ───────────────────────────────────────────────
function pushToGitHub(token, newDataRaw, buildTs) {
  var apiUrl = 'https://api.github.com/repos/' + REPO_OWNER + '/' + REPO_NAME + '/contents/' + FILE_PATH;

  var getResp = UrlFetchApp.fetch(apiUrl, {
    headers: {'Authorization': 'token ' + token, 'Accept': 'application/vnd.github.v3+json'},
    muteHttpExceptions: true
  });
  if (getResp.getResponseCode() !== 200)
    throw new Error('GitHub GET failed (' + getResp.getResponseCode() + '): ' + getResp.getContentText());

  var info = JSON.parse(getResp.getContentText());
  var current = Utilities.newBlob(Utilities.base64Decode(info.content.replace(/\n/g, ''))).getDataAsString();

  var updated = replaceBlock(current, 'var DATA_RAW = [', newDataRaw);
  updated = updated.replace(/var BUILD_TS = '[^']*';/, "var BUILD_TS = '" + buildTs + "';");

  var putResp = UrlFetchApp.fetch(apiUrl, {
    method: 'PUT',
    headers: {'Authorization': 'token ' + token, 'Content-Type': 'application/json', 'Accept': 'application/vnd.github.v3+json'},
    payload: JSON.stringify({
      message: 'Auto-update NCBD req vs gen data [' + new Date().toISOString().split('T')[0] + ']',
      content: Utilities.base64Encode(Utilities.newBlob(updated).getBytes()),
      sha: info.sha, branch: 'main'
    }),
    muteHttpExceptions: true
  });

  if (putResp.getResponseCode() !== 200 && putResp.getResponseCode() !== 201)
    throw new Error('GitHub PUT failed (' + putResp.getResponseCode() + '): ' + putResp.getContentText());
}

function replaceBlock(content, marker, replacement) {
  var start = content.indexOf(marker);
  if (start === -1) throw new Error('Marker not found: ' + marker);
  var end = content.indexOf('];', start + marker.length);
  if (end === -1) throw new Error('Closing ]; not found for: ' + marker);
  return content.substring(0, start) + replacement + content.substring(end + 2);
}

// ── TRIGGER SETUP ─────────────────────────────────────────────
function createDailyTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'updateNcbdDashboard') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('updateNcbdDashboard').timeBased().everyDays(1).atHour(11).create();
  Logger.log('Daily trigger created (11 AM).');
}
