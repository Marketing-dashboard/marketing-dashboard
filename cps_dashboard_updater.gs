// ============================================================
// CPS Marketing Performance Dashboard — Data Updater
//
// SETUP (one-time):
//   1. Open the CPS Google Sheet (tabs: Raw, Raw_TVS_Jawa, GA_Raw, FB_Raw)
//   2. Extensions → Apps Script → paste this file → Save
//   3. Run  setGitHubToken('ghp_yourTokenHere')  once
//   4. Run  updateCPSDashboard()  to push fresh data to GitHub
//   5. Optionally run  createDailyTrigger()  for auto-refresh at 9 AM
// ============================================================

var REPO_OWNER = 'Marketing-dashboard';
var REPO_NAME  = 'marketing-dashboard';
var FILE_PATH  = 'dashboard_data.json';

// ── ONE-TIME TOKEN SETUP ──────────────────────────────────────
function setGitHubToken(token) {
  PropertiesService.getScriptProperties().setProperty('GITHUB_TOKEN', token);
  Logger.log('GitHub token saved.');
}

// ── MAIN ENTRY POINT ──────────────────────────────────────────
function updateCPSDashboard() {
  var token = PropertiesService.getScriptProperties().getProperty('GITHUB_TOKEN');
  if (!token) throw new Error('Run setGitHubToken("ghp_...") first.');

  var ss = SpreadsheetApp.getActiveSpreadsheet();

  Logger.log('Reading sheet tabs...');
  var rawData = ss.getSheetByName('Raw').getDataRange().getValues();
  var tvsData = ss.getSheetByName('Raw_TVS_Jawa').getDataRange().getValues();
  var gaData  = ss.getSheetByName('GA_Raw').getDataRange().getValues();
  var fbData  = ss.getSheetByName('FB_Raw').getDataRange().getValues();
  Logger.log('Rows — Raw:' + (rawData.length-1) + ' TVS_Jawa:' + (tvsData.length-1) + ' GA:' + (gaData.length-1) + ' FB:' + (fbData.length-1));

  var triggers = processTriggers(rawData);
  var gaSpends = processGASpends(gaData, tvsData);
  var fbSpends = processFBSpends(fbData);

  var daywise    = buildDaywise(fbSpends, gaSpends, triggers);
  var mtdSummary = buildMTD(fbSpends, gaSpends, triggers);

  // Sort daywise: date → brand → model → FB/GA/Combined
  var srcOrder = {FB: 1, GA: 2, Combined: 3};
  daywise.sort(function(a, b) {
    var d = a.date.localeCompare(b.date);   if (d) return d;
    var b_ = (a.brand||'').localeCompare(b.brand||''); if (b_) return b_;
    var m = (a.model||'').localeCompare(b.model||''); if (m) return m;
    return (srcOrder[a.source]||9) - (srcOrder[b.source]||9);
  });

  var allDates = daywise.map(function(r) { return r.date; }).sort();
  var dateMin  = allDates[0] || '';
  var dateMax  = allDates[allDates.length - 1] || '';

  var output = {
    generated_at: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm'),
    tab_row_counts: {
      raw:      rawData.length - 1,
      tvs_jawa: tvsData.length - 1,
      ga_raw:   gaData.length  - 1,
      fb_raw:   fbData.length  - 1
    },
    summary: {
      date_min:         dateMin,
      date_max:         dateMax,
      daywise_total:    daywise.length,
      daywise_FB:       daywise.filter(function(r){return r.source==='FB';}).length,
      daywise_GA:       daywise.filter(function(r){return r.source==='GA';}).length,
      daywise_Combined: daywise.filter(function(r){return r.source==='Combined';}).length,
      mtd_total:        mtdSummary.length
    },
    daywise:     daywise,
    mtd_summary: mtdSummary
  };

  var jsonStr = JSON.stringify(output);
  pushJsonToGitHub(token, jsonStr, dateMin, dateMax);
  Logger.log('✓ Done: ' + daywise.length + ' daywise rows, ' + mtdSummary.length + ' MTD rows. Range: ' + dateMin + ' → ' + dateMax);
}

// ── DRY RUN: log stats without pushing ───────────────────────
function testRun() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var rawData = ss.getSheetByName('Raw').getDataRange().getValues();
  var tvsData = ss.getSheetByName('Raw_TVS_Jawa').getDataRange().getValues();
  var gaData  = ss.getSheetByName('GA_Raw').getDataRange().getValues();
  var fbData  = ss.getSheetByName('FB_Raw').getDataRange().getValues();

  var triggers = processTriggers(rawData);
  var gaSpends = processGASpends(gaData, tvsData);
  var fbSpends = processFBSpends(fbData);
  var daywise    = buildDaywise(fbSpends, gaSpends, triggers);
  var mtdSummary = buildMTD(fbSpends, gaSpends, triggers);

  var allDates = daywise.map(function(r){return r.date;}).sort();
  Logger.log('=== TEST RUN ===');
  Logger.log('Raw rows: ' + (rawData.length-1) + ', GA rows: ' + (gaData.length-1) + ', TVS rows: ' + (tvsData.length-1) + ', FB rows: ' + (fbData.length-1));
  Logger.log('Date range: ' + allDates[0] + ' → ' + allDates[allDates.length-1]);
  Logger.log('Daywise records: ' + daywise.length + ' (FB:' +
    daywise.filter(function(r){return r.source==='FB';}).length + ' GA:' +
    daywise.filter(function(r){return r.source==='GA';}).length + ' Combined:' +
    daywise.filter(function(r){return r.source==='Combined';}).length + ')');
  Logger.log('MTD records: ' + mtdSummary.length);

  // Log brand-level MTD triggers
  var brandTotals = {};
  mtdSummary.filter(function(r){return r.source==='Combined';}).forEach(function(r) {
    if (!brandTotals[r.brand]) brandTotals[r.brand] = {spends:0, leads:0, triggered:0};
    brandTotals[r.brand].spends    += r.spends;
    brandTotals[r.brand].leads     += r.leads;
    brandTotals[r.brand].triggered += r.triggered_leads;
  });
  Logger.log('=== MTD Brand Totals (Combined) ===');
  Object.keys(brandTotals).sort().forEach(function(b) {
    var t = brandTotals[b];
    Logger.log(b + ': spends=' + Math.round(t.spends) + ' leads=' + Math.round(t.leads) + ' triggered=' + t.triggered);
  });
}

// ── PROCESS RAW TAB (triggered leads) ────────────────────────
// Cols: Date(0)|Model(1)|brand(2)|Medium(3)|LeadType(4)|ModelType(5)|Trigger(6)|Total(7)|Process(8)
function processTriggers(rawData) {
  var fbByModel      = {};   // model → FB+WA trigger total (for MTD)
  var gaByModel      = {};   // model → Google trigger total (for MTD)
  var allByModel     = {};   // model → all-medium trigger total (for MTD)
  var fbByDateModel  = {};   // "date||model" → FB+WA count (for daywise)
  var gaByDateModel  = {};   // "date||model" → Google count (for daywise)
  var allByDateModel = {};   // "date||model" → all-medium count (for daywise)
  var modelToBrand   = {};   // model → normalized brand (Process col)

  for (var i = 1; i < rawData.length; i++) {
    var row = rawData[i];
    if (String(row[6] || '').trim() !== 'Trigger') continue;

    var model   = String(row[1] || '').trim();
    var medium  = String(row[3] || '').trim();
    var brand   = String(row[8] || '').trim(); // Process column = normalized brand
    var count   = parseInt(row[7]) || 0;
    var dateStr = fmtDate(row[0]);

    if (!model || count <= 0 || !dateStr) continue;

    var isFB = (medium === 'Facebook' || medium === 'Whatsapp');
    var isGA = (medium === 'Google');

    if (brand && !modelToBrand[model]) modelToBrand[model] = brand;

    // MTD-level aggregation (by model only)
    allByModel[model] = (allByModel[model] || 0) + count;
    if (isFB) fbByModel[model] = (fbByModel[model] || 0) + count;
    if (isGA) gaByModel[model] = (gaByModel[model] || 0) + count;

    // Daywise aggregation (by date + model)
    var dk = dateStr + '||' + model;
    allByDateModel[dk] = (allByDateModel[dk] || 0) + count;
    if (isFB) fbByDateModel[dk] = (fbByDateModel[dk] || 0) + count;
    if (isGA) gaByDateModel[dk] = (gaByDateModel[dk] || 0) + count;
  }

  return {
    fbByModel: fbByModel,   gaByModel: gaByModel,   allByModel: allByModel,
    fbByDateModel: fbByDateModel, gaByDateModel: gaByDateModel, allByDateModel: allByDateModel,
    modelToBrand: modelToBrand
  };
}

// ── PROCESS GA SPENDS (GA_Raw + Raw_TVS_Jawa) ─────────────────
// GA_Raw:      Month(0)|Day(1)|AdGroup(2)|Campaign(3)|Cost(4)|Conv(5)|Platform(6)|Brand(7)|Model(8)
// Raw_TVS_Jawa: Month(0)|Day(1)|Campaign(2)|Cost(3)|Conv(4)|Platform(5)|Brand(6)|Model(7)
function processGASpends(gaData, tvsData) {
  var map = {};

  function add(dateStr, brand, model, spends, leads) {
    if (!dateStr || !brand || !model) return;
    if (!spends && !leads) return;
    var k = dateStr + '||' + brand + '||' + model;
    if (!map[k]) map[k] = {date: dateStr, brand: brand, model: model, spends: 0, leads: 0};
    map[k].spends += spends;
    map[k].leads  += leads;
  }

  for (var i = 1; i < gaData.length; i++) {
    var r = gaData[i];
    add(fmtDate(r[1]), String(r[7]||'').trim(), String(r[8]||'').trim(), parseFloat(r[4])||0, parseFloat(r[5])||0);
  }
  for (var j = 1; j < tvsData.length; j++) {
    var r = tvsData[j];
    add(fmtDate(r[1]), String(r[6]||'').trim(), String(r[7]||'').trim(), parseFloat(r[3])||0, parseFloat(r[4])||0);
  }

  return map;
}

// ── PROCESS FB SPENDS (FB_Raw) ────────────────────────────────
// Cols: Month(0)|Date(1)|CampaignName(2)|AdSet(3)|AdName(4)|Spends(5)|Leads(6)|Platform(7)|Brand(8)|Model(9)
function processFBSpends(fbData) {
  var map = {};

  for (var i = 1; i < fbData.length; i++) {
    var r       = fbData[i];
    var dateStr = fmtDate(r[1]);
    var brand   = String(r[8] || '').trim();
    var model   = String(r[9] || '').trim();
    var spends  = parseFloat(r[5]) || 0;
    var leads   = parseFloat(r[6]) || 0;
    if (!dateStr || !brand || !model || (!spends && !leads)) continue;
    var k = dateStr + '||' + brand + '||' + model;
    if (!map[k]) map[k] = {date: dateStr, brand: brand, model: model, spends: 0, leads: 0};
    map[k].spends += spends;
    map[k].leads  += leads;
  }

  return map;
}

// ── BUILD DAYWISE RECORDS ─────────────────────────────────────
function buildDaywise(fbSpends, gaSpends, triggers) {
  // Collect all unique date||brand||model keys
  var allKeys = {};
  Object.keys(fbSpends).forEach(function(k) { allKeys[k] = true; });
  Object.keys(gaSpends).forEach(function(k) { allKeys[k] = true; });

  // Add trigger-only date+model keys (if we know the brand)
  Object.keys(triggers.allByDateModel).forEach(function(dk) {
    var parts   = dk.split('||');
    var dateStr = parts[0], model = parts[1];
    var brand   = triggers.modelToBrand[model];
    if (brand) allKeys[dateStr + '||' + brand + '||' + model] = true;
  });

  var rows = [];

  Object.keys(allKeys).forEach(function(key) {
    var parts   = key.split('||');
    var dateStr = parts[0], brand = parts[1], model = parts[2];

    var fb = fbSpends[key] || {spends: 0, leads: 0};
    var ga = gaSpends[key] || {spends: 0, leads: 0};
    var dk = dateStr + '||' + model;

    var fbTrig  = triggers.fbByDateModel[dk]  || 0;
    var gaTrig  = triggers.gaByDateModel[dk]  || 0;
    var allTrig = triggers.allByDateModel[dk] || 0;

    if (fb.spends > 0 || fb.leads > 0 || fbTrig > 0)
      rows.push(makeRow(dateStr, brand, model, 'FB', fb.spends, fb.leads, fbTrig));

    if (ga.spends > 0 || ga.leads > 0 || gaTrig > 0)
      rows.push(makeRow(dateStr, brand, model, 'GA', ga.spends, ga.leads, gaTrig));

    var totS = fb.spends + ga.spends;
    var totL = fb.leads  + ga.leads;
    if (totS > 0 || totL > 0 || allTrig > 0)
      rows.push(makeRow(dateStr, brand, model, 'Combined', totS, totL, allTrig));
  });

  return rows;
}

// ── BUILD MTD SUMMARY ─────────────────────────────────────────
function buildMTD(fbSpends, gaSpends, triggers) {
  // Aggregate all spend dates into brand+model totals
  var fbBM = {}, gaBM = {};

  function addBM(map, brand, model, spends, leads) {
    var k = brand + '||' + model;
    if (!map[k]) map[k] = {brand: brand, model: model, spends: 0, leads: 0};
    map[k].spends += spends;
    map[k].leads  += leads;
  }

  Object.keys(fbSpends).forEach(function(k) {
    var r = fbSpends[k]; addBM(fbBM, r.brand, r.model, r.spends, r.leads);
  });
  Object.keys(gaSpends).forEach(function(k) {
    var r = gaSpends[k]; addBM(gaBM, r.brand, r.model, r.spends, r.leads);
  });

  // All brand+model keys from spend + trigger data
  var allBMKeys = {};
  Object.keys(fbBM).forEach(function(k) { allBMKeys[k] = true; });
  Object.keys(gaBM).forEach(function(k) { allBMKeys[k] = true; });
  Object.keys(triggers.allByModel).forEach(function(model) {
    var brand = triggers.modelToBrand[model];
    if (brand) allBMKeys[brand + '||' + model] = true;
  });

  var rows = [];

  Object.keys(allBMKeys).forEach(function(bmKey) {
    var parts = bmKey.split('||');
    var brand = parts[0], model = parts[1];

    var fb = fbBM[bmKey] || {spends: 0, leads: 0};
    var ga = gaBM[bmKey] || {spends: 0, leads: 0};

    var fbTrig  = triggers.fbByModel[model]  || 0;
    var gaTrig  = triggers.gaByModel[model]  || 0;
    var allTrig = triggers.allByModel[model] || 0;

    if (fb.spends > 0 || fb.leads > 0 || fbTrig > 0)
      rows.push(makeRow(null, brand, model, 'FB', fb.spends, fb.leads, fbTrig));

    if (ga.spends > 0 || ga.leads > 0 || gaTrig > 0)
      rows.push(makeRow(null, brand, model, 'GA', ga.spends, ga.leads, gaTrig));

    var totS = fb.spends + ga.spends;
    var totL = fb.leads  + ga.leads;
    if (totS > 0 || totL > 0 || allTrig > 0)
      rows.push(makeRow(null, brand, model, 'Combined', totS, totL, allTrig));
  });

  return rows;
}

// ── ROW BUILDER ───────────────────────────────────────────────
function makeRow(date, brand, model, source, spends, leads, triggered) {
  var s = Math.round(spends    * 100) / 100;
  var l = Math.round(leads     * 100) / 100;
  var t = Math.round(triggered) || 0;
  var cpl  = (l > 0) ? Math.round(s / l * 100) / 100 : null;
  var tcpl = (t > 0 && s > 0) ? Math.round(s / t * 100) / 100 : null;
  var tpct = (l > 0) ? Math.round(t / l * 10000) / 100 : null;

  var row = {};
  if (date) row.date = date;
  row.brand          = brand;
  row.model          = model;
  row.source         = source;
  row.spends         = s;
  row.leads          = l;
  row.triggered_leads = t;
  row.cpl            = cpl;
  row.tcpl           = tcpl;
  row.trigger_pct    = tpct;
  return row;
}

// ── PUSH JSON TO GITHUB ───────────────────────────────────────
function pushJsonToGitHub(token, jsonStr, dateMin, dateMax) {
  var apiUrl = 'https://api.github.com/repos/' + REPO_OWNER + '/' + REPO_NAME + '/contents/' + FILE_PATH;

  // GET current SHA (needed for update)
  var getResp = UrlFetchApp.fetch(apiUrl, {
    headers: {'Authorization': 'token ' + token, 'Accept': 'application/vnd.github.v3+json'},
    muteHttpExceptions: true
  });

  var sha = '';
  var code = getResp.getResponseCode();
  if (code === 200) {
    sha = JSON.parse(getResp.getContentText()).sha;
    Logger.log('Existing file SHA: ' + sha);
  } else if (code === 404) {
    Logger.log('File not found — will create it.');
  } else {
    throw new Error('GitHub GET failed (' + code + '): ' + getResp.getContentText());
  }

  var commitMsg = 'Update CPS dashboard data ' + dateMin + ' → ' + dateMax;
  var payload   = {
    message: commitMsg,
    content: Utilities.base64Encode(Utilities.newBlob(jsonStr).getBytes()),
    branch:  'main'
  };
  if (sha) payload.sha = sha;

  var putResp = UrlFetchApp.fetch(apiUrl, {
    method:  'PUT',
    headers: {
      'Authorization': 'token ' + token,
      'Content-Type':  'application/json',
      'Accept':        'application/vnd.github.v3+json'
    },
    payload:            JSON.stringify(payload),
    muteHttpExceptions: true
  });

  var putCode = putResp.getResponseCode();
  if (putCode !== 200 && putCode !== 201) {
    throw new Error('GitHub PUT failed (' + putCode + '): ' + putResp.getContentText());
  }
  Logger.log('Pushed to GitHub (' + putCode + '). Commit: ' + commitMsg);
}

// ── DATE FORMATTER ────────────────────────────────────────────
function fmtDate(val) {
  if (!val) return '';
  var tz = Session.getScriptTimeZone();
  if (val instanceof Date) return Utilities.formatDate(val, tz, 'yyyy-MM-dd');
  var s = String(val).trim();
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.substring(0, 10);
  var mdy = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (mdy) {
    var d = new Date(parseInt(mdy[3]), parseInt(mdy[1]) - 1, parseInt(mdy[2]));
    return Utilities.formatDate(d, tz, 'yyyy-MM-dd');
  }
  var d2 = new Date(s);
  if (!isNaN(d2.getTime())) return Utilities.formatDate(d2, tz, 'yyyy-MM-dd');
  return '';
}

// ── OPTIONAL: DAILY TRIGGER ───────────────────────────────────
function createDailyTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'updateCPSDashboard') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('updateCPSDashboard')
    .timeBased().everyDays(1).atHour(9).create();
  Logger.log('Daily trigger set for 9 AM.');
}
