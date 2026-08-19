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

  // Sort daywise: date → brand → model
  daywise.sort(function(a, b) {
    var d = a.date.localeCompare(b.date);          if (d) return d;
    var br = (a.brand||'').localeCompare(b.brand||''); if (br) return br;
    return (a.model||'').localeCompare(b.model||'');
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
      date_min:      dateMin,
      date_max:      dateMax,
      daywise_total: daywise.length,
      mtd_total:     mtdSummary.length
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
  Logger.log('Daywise records: ' + daywise.length + ', MTD records: ' + mtdSummary.length);

  // Log brand-level MTD totals
  var brandTotals = {};
  mtdSummary.forEach(function(r) {
    if (!brandTotals[r.brand]) brandTotals[r.brand] = {paidS:0, paidL:0, paidT:0, waT:0, combT:0};
    brandTotals[r.brand].paidS += r.paid_spends;
    brandTotals[r.brand].paidL += r.paid_leads;
    brandTotals[r.brand].paidT += r.paid_triggered;
    brandTotals[r.brand].waT   += r.wa_triggered;
    brandTotals[r.brand].combT += r.combined_triggered;
  });
  Logger.log('=== MTD Brand Totals ===');
  Object.keys(brandTotals).sort().forEach(function(b) {
    var t = brandTotals[b];
    Logger.log(b + ': spends='+Math.round(t.paidS)+' leads='+Math.round(t.paidL)+
      ' paid_trig='+t.paidT+' wa_trig='+t.waT+' comb_trig='+t.combT);
  });
}

// ── FIND COLUMN INDEX BY HEADER NAME ─────────────────────────
function findCol(headers, candidates) {
  var cl = candidates.map(function(c) { return c.toLowerCase().trim(); });
  for (var i = 0; i < headers.length; i++) {
    var h = String(headers[i] || '').toLowerCase().trim();
    if (cl.indexOf(h) !== -1) return i;
  }
  return -1;
}

// ── DIAGNOSE: log column headers of all sheets (run once to verify) ──
function diagnoseSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ['Raw','Raw_TVS_Jawa','GA_Raw','FB_Raw'].forEach(function(name) {
    var sheet = ss.getSheetByName(name);
    if (!sheet) { Logger.log(name + ': NOT FOUND'); return; }
    var first2 = sheet.getRange(1, 1, 2, sheet.getLastColumn()).getValues();
    Logger.log('=== ' + name + ' headers ===');
    Logger.log(first2[0].join(' | '));
    Logger.log('Row 2: ' + first2[1].join(' | '));
  });
}

// ── PROCESS RAW TAB (triggered leads only) ────────────────────
// Only maps Facebook, Adwords/Google, and Whatsapp mediums.
// Non-MS, Organic, and all other mediums are excluded.
function processTriggers(rawData) {
  var fbByModel = {}, waByModel = {}, gaByModel = {};
  var fbByDateModel = {}, waByDateModel = {}, gaByDateModel = {};
  var modelToBrand = {};

  if (rawData.length === 0) return {
    fbByModel: fbByModel, waByModel: waByModel, gaByModel: gaByModel,
    fbByDateModel: fbByDateModel, waByDateModel: waByDateModel, gaByDateModel: gaByDateModel,
    modelToBrand: modelToBrand
  };

  var h = rawData[0].map(function(v) { return String(v||'').toLowerCase().trim(); });
  var iDate    = findCol(h, ['date','day']);
  var iModel   = findCol(h, ['model']);
  var iMedium  = findCol(h, ['medium']);
  var iTrigCol = findCol(h, ['trigger']);
  var iTotal   = findCol(h, ['total','count','leads']);
  var iProcess = findCol(h, ['process','brand name','brandname']);
  if (iProcess < 0) iProcess = findCol(h, ['brand']);
  Logger.log('Raw idx — date:'+iDate+' model:'+iModel+' medium:'+iMedium+' trigcol:'+iTrigCol+' total:'+iTotal+' process:'+iProcess);

  for (var i = 1; i < rawData.length; i++) {
    var row = rawData[i];
    if (iTrigCol >= 0 && String(row[iTrigCol] || '').trim() !== 'Trigger') continue;

    var model   = iModel   >= 0 ? String(row[iModel]   || '').trim() : '';
    var medium  = iMedium  >= 0 ? String(row[iMedium]  || '').trim() : '';
    var brand   = iProcess >= 0 ? String(row[iProcess] || '').trim() : '';
    var count   = iTotal   >= 0 ? (parseInt(row[iTotal]) || 0) : 0;
    var dateStr = iDate    >= 0 ? fmtDate(row[iDate]) : '';

    if (!model || count <= 0 || !dateStr) continue;

    var isFB = (medium === 'Facebook');
    var isWA = (medium === 'Whatsapp');
    var isGA = (medium === 'Google' || medium === 'Adwords');
    // Skip Non-MS, Organic, and any unrecognised medium
    if (!isFB && !isWA && !isGA) continue;

    if (brand && !modelToBrand[model]) modelToBrand[model] = brand;

    var dk = dateStr + '||' + model;
    if (isFB) {
      fbByModel[model]    = (fbByModel[model]    || 0) + count;
      fbByDateModel[dk]   = (fbByDateModel[dk]   || 0) + count;
    }
    if (isWA) {
      waByModel[model]    = (waByModel[model]    || 0) + count;
      waByDateModel[dk]   = (waByDateModel[dk]   || 0) + count;
    }
    if (isGA) {
      gaByModel[model]    = (gaByModel[model]    || 0) + count;
      gaByDateModel[dk]   = (gaByDateModel[dk]   || 0) + count;
    }
  }

  return {
    fbByModel: fbByModel, waByModel: waByModel, gaByModel: gaByModel,
    fbByDateModel: fbByDateModel, waByDateModel: waByDateModel, gaByDateModel: gaByDateModel,
    modelToBrand: modelToBrand
  };
}

// ── PROCESS GA SPENDS (GA_Raw + Raw_TVS_Jawa) ─────────────────
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

  // ── GA_Raw ──
  if (gaData.length > 0) {
    var gh = gaData[0].map(function(v) { return String(v||'').toLowerCase().trim(); });
    var gDate  = findCol(gh, ['day','date']);
    var gCost  = findCol(gh, ['cost','spend','spends','amount spent','amount_spent']);
    var gConv  = findCol(gh, ['conv','conversions','leads','lead','results','result']);
    var gBrand = findCol(gh, ['brand']);
    var gModel = findCol(gh, ['model']);
    Logger.log('GA_Raw headers: ' + gh.join(' | '));
    Logger.log('GA_Raw idx — date:'+gDate+' cost:'+gCost+' conv:'+gConv+' brand:'+gBrand+' model:'+gModel);
    for (var i = 1; i < gaData.length; i++) {
      var r = gaData[i];
      var d = gDate  >= 0 ? fmtDate(r[gDate])               : '';
      var b = gBrand >= 0 ? String(r[gBrand]||'').trim()    : '';
      var m = gModel >= 0 ? String(r[gModel]||'').trim()    : '';
      var s = gCost  >= 0 ? (parseFloat(r[gCost])  || 0)   : 0;
      var l = gConv  >= 0 ? (parseFloat(r[gConv])  || 0)   : 0;
      add(d, b, m, s, l);
    }
  }

  // ── Raw_TVS_Jawa ──
  if (tvsData.length > 0) {
    var th = tvsData[0].map(function(v) { return String(v||'').toLowerCase().trim(); });
    var tDate  = findCol(th, ['day','date']);
    var tCost  = findCol(th, ['cost','spend','spends','amount spent','amount_spent']);
    var tConv  = findCol(th, ['conv','conversions','leads','lead','results','result']);
    var tBrand = findCol(th, ['brand']);
    var tModel = findCol(th, ['model']);
    Logger.log('TVS_Jawa headers: ' + th.join(' | '));
    Logger.log('TVS_Jawa idx — date:'+tDate+' cost:'+tCost+' conv:'+tConv+' brand:'+tBrand+' model:'+tModel);
    for (var j = 1; j < tvsData.length; j++) {
      var r = tvsData[j];
      var d = tDate  >= 0 ? fmtDate(r[tDate])               : '';
      var b = tBrand >= 0 ? String(r[tBrand]||'').trim()    : '';
      var m = tModel >= 0 ? String(r[tModel]||'').trim()    : '';
      var s = tCost  >= 0 ? (parseFloat(r[tCost])  || 0)   : 0;
      var l = tConv  >= 0 ? (parseFloat(r[tConv])  || 0)   : 0;
      add(d, b, m, s, l);
    }
  }

  return map;
}

// ── PROCESS FB SPENDS (FB_Raw) ────────────────────────────────
function processFBSpends(fbData) {
  var map = {};
  if (fbData.length === 0) return map;

  var fh = fbData[0].map(function(v) { return String(v||'').toLowerCase().trim(); });
  var fDate  = findCol(fh, ['date','day']);
  var fSpend = findCol(fh, ['spend','spends','cost','amount spent','amount_spent']);
  var fLead  = findCol(fh, ['lead','leads','result','results','conv','conversions']);
  var fBrand = findCol(fh, ['brand']);
  var fModel = findCol(fh, ['model']);
  Logger.log('FB_Raw headers: ' + fh.join(' | '));
  Logger.log('FB_Raw idx — date:'+fDate+' spend:'+fSpend+' lead:'+fLead+' brand:'+fBrand+' model:'+fModel);

  for (var i = 1; i < fbData.length; i++) {
    var r      = fbData[i];
    var d      = fDate  >= 0 ? fmtDate(r[fDate])              : '';
    var brand  = fBrand >= 0 ? String(r[fBrand]||'').trim()   : '';
    var model  = fModel >= 0 ? String(r[fModel]||'').trim()   : '';
    var spends = fSpend >= 0 ? (parseFloat(r[fSpend]) || 0)   : 0;
    var leads  = fLead  >= 0 ? (parseFloat(r[fLead])  || 0)   : 0;
    if (!d || !brand || !model || (!spends && !leads)) continue;
    var k = d + '||' + brand + '||' + model;
    if (!map[k]) map[k] = {date: d, brand: brand, model: model, spends: 0, leads: 0};
    map[k].spends += spends;
    map[k].leads  += leads;
  }

  return map;
}

// ── BUILD DAYWISE RECORDS ─────────────────────────────────────
// One row per date+brand+model with paid (FB+GA), WA, and combined fields.
function buildDaywise(fbSpends, gaSpends, triggers) {
  var allKeys = {};
  Object.keys(fbSpends).forEach(function(k) { allKeys[k] = true; });
  Object.keys(gaSpends).forEach(function(k) { allKeys[k] = true; });

  // Add trigger-only date+model keys for all three medium buckets
  function addTrigKeys(tMap) {
    Object.keys(tMap).forEach(function(dk) {
      var parts   = dk.split('||');
      var dateStr = parts[0], model = parts[1];
      var brand   = triggers.modelToBrand[model];
      if (brand) allKeys[dateStr + '||' + brand + '||' + model] = true;
    });
  }
  addTrigKeys(triggers.fbByDateModel);
  addTrigKeys(triggers.waByDateModel);
  addTrigKeys(triggers.gaByDateModel);

  var rows = [];

  Object.keys(allKeys).forEach(function(key) {
    var parts   = key.split('||');
    var dateStr = parts[0], brand = parts[1], model = parts[2];

    var fb = fbSpends[key] || {spends: 0, leads: 0};
    var ga = gaSpends[key] || {spends: 0, leads: 0};
    var dk = dateStr + '||' + model;

    var fbTrig = triggers.fbByDateModel[dk] || 0;
    var waTrig = triggers.waByDateModel[dk] || 0;
    var gaTrig = triggers.gaByDateModel[dk] || 0;

    var paidS    = fb.spends + ga.spends;
    var paidL    = fb.leads  + ga.leads;
    var paidTrig = fbTrig + gaTrig;
    var combTrig = fbTrig + waTrig + gaTrig;

    if (!paidS && !paidL && !combTrig) return;

    rows.push(makeRow(dateStr, brand, model, paidS, paidL, paidTrig, waTrig, combTrig));
  });

  return rows;
}

// ── BUILD MTD SUMMARY ─────────────────────────────────────────
function buildMTD(fbSpends, gaSpends, triggers) {
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
  ['fbByModel','waByModel','gaByModel'].forEach(function(bucket) {
    Object.keys(triggers[bucket]).forEach(function(model) {
      var brand = triggers.modelToBrand[model];
      if (brand) allBMKeys[brand + '||' + model] = true;
    });
  });

  var rows = [];

  Object.keys(allBMKeys).forEach(function(bmKey) {
    var parts = bmKey.split('||');
    var brand = parts[0], model = parts[1];

    var fb = fbBM[bmKey] || {spends: 0, leads: 0};
    var ga = gaBM[bmKey] || {spends: 0, leads: 0};

    var fbTrig = triggers.fbByModel[model] || 0;
    var waTrig = triggers.waByModel[model] || 0;
    var gaTrig = triggers.gaByModel[model] || 0;

    var paidS    = fb.spends + ga.spends;
    var paidL    = fb.leads  + ga.leads;
    var paidTrig = fbTrig + gaTrig;
    var combTrig = fbTrig + waTrig + gaTrig;

    if (!paidS && !paidL && !combTrig) return;

    rows.push(makeRow(null, brand, model, paidS, paidL, paidTrig, waTrig, combTrig));
  });

  return rows;
}

// ── ROW BUILDER ───────────────────────────────────────────────
function makeRow(date, brand, model, paidSpends, paidLeads, paidTriggered, waTriggered, combTriggered) {
  var ps = Math.round(paidSpends    * 100) / 100;
  var pl = Math.round(paidLeads     * 100) / 100;
  var pt = Math.round(paidTriggered) || 0;
  var wt = Math.round(waTriggered)   || 0;
  var ct = Math.round(combTriggered) || 0;

  var paidCpl  = (pl > 0) ? Math.round(ps / pl * 100) / 100 : null;
  var paidTcpl = (pt > 0 && ps > 0) ? Math.round(ps / pt * 100) / 100 : null;
  var paidTpct = (pl > 0) ? Math.round(pt / pl * 10000) / 100 : null;
  var combTcpl = (ct > 0 && ps > 0) ? Math.round(ps / ct * 100) / 100 : null;
  var combTpct = (pl > 0) ? Math.round(ct / pl * 10000) / 100 : null;

  var row = {};
  if (date) row.date = date;
  row.brand                = brand;
  row.model                = model;
  row.paid_spends          = ps;
  row.paid_leads           = pl;
  row.paid_triggered       = pt;
  row.paid_cpl             = paidCpl;
  row.paid_tcpl            = paidTcpl;
  row.paid_trigger_pct     = paidTpct;
  row.wa_triggered         = wt;
  row.combined_spends      = ps;
  row.combined_leads       = pl;
  row.combined_triggered   = ct;
  row.combined_cpl         = paidCpl;
  row.combined_tcpl        = combTcpl;
  row.combined_trigger_pct = combTpct;
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
