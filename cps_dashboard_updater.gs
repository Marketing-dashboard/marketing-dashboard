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

// ── BRAND ALIASES: map variant names → canonical name ─────────
// Add entries here whenever two brand names should be treated as one.
var BRAND_ALIAS = {
  'hero xtreme': 'Hero'
};
function normBrand(brand) {
  if (!brand) return brand;
  return BRAND_ALIAS[brand.toLowerCase().trim()] || brand;
}

// ── EXCLUDED BRANDS: drop these from all output ───────────────
var EXCLUDED_BRANDS = ['others', '#n/a', ''];
function isExcluded(brand) {
  return !brand || EXCLUDED_BRANDS.indexOf(brand.toLowerCase().trim()) !== -1;
}

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
  Logger.log('Rows — Raw:' + (rawData.length-1) + ' TVS_Jawa:' + (tvsData.length-1) + ' GA:' + (gaData.length-1) + ' FB_Raw:' + (fbData.length-1));

  var triggers = processTriggers(rawData);
  var gaSpends = processGASpends(gaData, tvsData);
  var fbwa     = processFBandWASpends(fbData);
  var fbSpends = fbwa.fbMap;
  var waSpends = fbwa.waMap;

  var daywise    = buildDaywise(fbSpends, gaSpends, waSpends, triggers);
  var mtdSummary = buildMTD(fbSpends, gaSpends, waSpends, triggers);

  // Sort daywise: date → brand → model
  daywise.sort(function(a, b) {
    var d = (a.date||'').localeCompare(b.date||''); if (d) return d;
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
  var fbwa     = processFBandWASpends(fbData);
  var fbSpends = fbwa.fbMap;
  var waSpends = fbwa.waMap;
  var daywise    = buildDaywise(fbSpends, gaSpends, waSpends, triggers);
  var mtdSummary = buildMTD(fbSpends, gaSpends, waSpends, triggers);

  var allDates = daywise.map(function(r){return r.date;}).sort();
  Logger.log('=== TEST RUN ===');
  Logger.log('Raw rows: ' + (rawData.length-1) + ', GA rows: ' + (gaData.length-1) + ', TVS rows: ' + (tvsData.length-1) + ', FB_Raw rows: ' + (fbData.length-1));
  Logger.log('Date range: ' + allDates[0] + ' → ' + allDates[allDates.length-1]);
  Logger.log('Daywise records: ' + daywise.length + ', MTD records: ' + mtdSummary.length);

  var brandTotals = {};
  mtdSummary.forEach(function(r) {
    if (!brandTotals[r.brand]) brandTotals[r.brand] = {fS:0,fL:0,fT:0,gS:0,gL:0,gT:0,wS:0,wL:0,wT:0,cT:0};
    var b = brandTotals[r.brand];
    b.fS += r.fb_spends; b.fL += r.fb_leads; b.fT += r.fb_triggered;
    b.gS += r.ga_spends; b.gL += r.ga_leads; b.gT += r.ga_triggered;
    b.wS += r.wa_spends; b.wL += r.wa_leads; b.wT += r.wa_triggered;
    b.cT += r.combined_triggered;
  });
  Logger.log('=== MTD Brand Totals ===');
  Object.keys(brandTotals).sort().forEach(function(b) {
    var t = brandTotals[b];
    Logger.log(b + ': fb_s='+Math.round(t.fS)+' fb_l='+Math.round(t.fL)+' fb_t='+t.fT+
      ' ga_s='+Math.round(t.gS)+' ga_l='+Math.round(t.gL)+' ga_t='+t.gT+
      ' wa_s='+Math.round(t.wS)+' wa_l='+Math.round(t.wL)+' wa_t='+t.wT+' comb_t='+t.cT);
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

// ── DIAGNOSE: log column headers of all sheets ────────────────
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
// Maps Facebook, Google/Adwords, Whatsapp triggers. Excludes Non-MS, Organic, etc.
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
  var iModel   = findCol(h, ['model_map', 'modelmap']);
  if (iModel < 0) iModel = findCol(h, ['model']);
  var iMedium  = findCol(h, ['medium']);
  var iTrigCol = findCol(h, ['trigger']);
  var iTotal   = findCol(h, ['total','count','leads']);
  var iProcess = findCol(h, ['process', 'brand name', 'brandname', 'segment', 'category']);
  var iBrand   = findCol(h, ['brand']);
  Logger.log('Raw idx — date:'+iDate+' model_map:'+iModel+' medium:'+iMedium+' trigcol:'+iTrigCol+' total:'+iTotal+' process:'+iProcess+' brand:'+iBrand);

  for (var i = 1; i < rawData.length; i++) {
    var row = rawData[i];
    if (iTrigCol >= 0 && String(row[iTrigCol] || '').trim() !== 'Trigger') continue;

    var model   = iModel   >= 0 ? String(row[iModel]   || '').trim() : '';
    var medium  = iMedium  >= 0 ? String(row[iMedium]  || '').trim().toLowerCase() : '';
    var processVal = iProcess >= 0 ? String(row[iProcess] || '').trim() : '';
    var brandVal   = iBrand   >= 0 ? String(row[iBrand]   || '').trim() : '';
    // When Process column says "Others", use the Brand column for the real brand name
    var brand = normBrand(processVal.toLowerCase() === 'others' ? brandVal : (processVal || brandVal));
    var count   = iTotal   >= 0 ? (parseInt(row[iTotal]) || 0) : 0;
    var dateStr = iDate    >= 0 ? fmtDate(row[iDate]) : '';

    if (!model || count <= 0 || !dateStr) continue;

    var isFB = (medium === 'facebook');
    var isWA = (medium === 'whatsapp');
    var isGA = (medium === 'google' || medium === 'adwords');
    // Skip Non-MS, Organic, and any unrecognised medium
    if (!isFB && !isWA && !isGA) continue;

    // Normalise model key to lowercase so case differences don't break matching
    var mk = model.toLowerCase();
    if (brand && !modelToBrand[mk]) modelToBrand[mk] = brand;

    var dk = dateStr + '||' + mk;
    if (isFB) {
      fbByModel[mk]       = (fbByModel[mk]       || 0) + count;
      fbByDateModel[dk]   = (fbByDateModel[dk]   || 0) + count;
    }
    if (isWA) {
      waByModel[mk]       = (waByModel[mk]       || 0) + count;
      waByDateModel[dk]   = (waByDateModel[dk]   || 0) + count;
    }
    if (isGA) {
      gaByModel[mk]       = (gaByModel[mk]       || 0) + count;
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
      var d = gDate  >= 0 ? fmtDate(r[gDate])                        : '';
      var b = normBrand(gBrand >= 0 ? String(r[gBrand]||'').trim() : '');
      var m = gModel >= 0 ? String(r[gModel]||'').trim()            : '';
      var s = gCost  >= 0 ? (parseFloat(r[gCost])  || 0)           : 0;
      var l = gConv  >= 0 ? (parseFloat(r[gConv])  || 0)           : 0;
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
      var d = tDate  >= 0 ? fmtDate(r[tDate])                        : '';
      var b = normBrand(tBrand >= 0 ? String(r[tBrand]||'').trim() : '');
      var m = tModel >= 0 ? String(r[tModel]||'').trim()            : '';
      var s = tCost  >= 0 ? (parseFloat(r[tCost])  || 0)           : 0;
      var l = tConv  >= 0 ? (parseFloat(r[tConv])  || 0)           : 0;
      add(d, b, m, s, l);
    }
  }

  return map;
}

// ── PROCESS FB_Raw (splits FB vs WhatsApp by source column) ───
// Rows where source/medium = "whatsapp" → waMap; everything else → fbMap.
function processFBandWASpends(fbData) {
  var fbMap = {}, waMap = {};
  if (fbData.length === 0) return { fbMap: fbMap, waMap: waMap };

  var fh = fbData[0].map(function(v) { return String(v||'').toLowerCase().trim(); });
  var fDate   = findCol(fh, ['date','day']);
  var fSpend  = findCol(fh, ['spend','spends','cost','amount spent','amount_spent']);
  var fLead   = findCol(fh, ['lead','leads','result','results','conv','conversions']);
  var fBrand  = findCol(fh, ['brand']);
  var fModel  = findCol(fh, ['model']);
  var fSource = findCol(fh, ['source','platform','channel','medium']);
  Logger.log('FB_Raw headers: ' + fh.join(' | '));
  Logger.log('FB_Raw idx — date:'+fDate+' spend:'+fSpend+' lead:'+fLead+' brand:'+fBrand+' model:'+fModel+' source:'+fSource);

  for (var i = 1; i < fbData.length; i++) {
    var r      = fbData[i];
    var d      = fDate   >= 0 ? fmtDate(r[fDate])                          : '';
    var brand  = normBrand(fBrand >= 0 ? String(r[fBrand]||'').trim()    : '');
    var model  = fModel  >= 0 ? String(r[fModel]||'').trim()              : '';
    var spends = fSpend  >= 0 ? (parseFloat(r[fSpend]) || 0)   : 0;
    var leads  = fLead   >= 0 ? (parseFloat(r[fLead])  || 0)   : 0;
    var source = fSource >= 0 ? String(r[fSource]||'').trim().toLowerCase() : '';
    if (!d || !brand || !model || (!spends && !leads)) continue;
    var k = d + '||' + brand + '||' + model;
    var isWA = (source === 'whatsapp');
    var map = isWA ? waMap : fbMap;
    if (!map[k]) map[k] = {date: d, brand: brand, model: model, spends: 0, leads: 0};
    map[k].spends += spends;
    map[k].leads  += leads;
  }

  return { fbMap: fbMap, waMap: waMap };
}

// ── BUILD DAYWISE RECORDS ─────────────────────────────────────
// One row per date+brand+model. Triggers are only attached to models
// that have spend/lead data — trigger-only models are excluded.
function buildDaywise(fbSpends, gaSpends, waSpends, triggers) {
  // Keys come from spend maps only — no trigger-only rows
  var allKeys = {};
  Object.keys(fbSpends).forEach(function(k) { allKeys[k] = true; });
  Object.keys(gaSpends).forEach(function(k) { allKeys[k] = true; });
  Object.keys(waSpends).forEach(function(k) { allKeys[k] = true; });

  var rows = [];

  Object.keys(allKeys).forEach(function(key) {
    var parts   = key.split('||');
    var dateStr = parts[0], brand = parts[1], model = parts[2];
    if (isExcluded(brand)) return;
    var fb = fbSpends[key] || {spends: 0, leads: 0};
    var ga = gaSpends[key] || {spends: 0, leads: 0};
    var wa = waSpends[key] || {spends: 0, leads: 0};
    var dk = dateStr + '||' + model.toLowerCase();
    var fbTrig = triggers.fbByDateModel[dk] || 0;
    var waTrig = triggers.waByDateModel[dk] || 0;
    var gaTrig = triggers.gaByDateModel[dk] || 0;
    rows.push(makeRow(dateStr, brand, model,
      fb.spends, fb.leads, fbTrig,
      ga.spends, ga.leads, gaTrig,
      wa.spends, wa.leads, waTrig));
  });

  return rows;
}

// ── BUILD MTD SUMMARY ─────────────────────────────────────────
function buildMTD(fbSpends, gaSpends, waSpends, triggers) {
  var fbBM = {}, gaBM = {}, waBM = {};

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
  Object.keys(waSpends).forEach(function(k) {
    var r = waSpends[k]; addBM(waBM, r.brand, r.model, r.spends, r.leads);
  });

  // Keys come from spend maps only — no trigger-only rows
  var allBMKeys = {};
  Object.keys(fbBM).forEach(function(k) { allBMKeys[k] = true; });
  Object.keys(gaBM).forEach(function(k) { allBMKeys[k] = true; });
  Object.keys(waBM).forEach(function(k) { allBMKeys[k] = true; });

  var rows = [];

  Object.keys(allBMKeys).forEach(function(bmKey) {
    var parts = bmKey.split('||');
    var brand = parts[0], model = parts[1];
    if (isExcluded(brand)) return;
    var fb = fbBM[bmKey] || {spends: 0, leads: 0};
    var ga = gaBM[bmKey] || {spends: 0, leads: 0};
    var wa = waBM[bmKey] || {spends: 0, leads: 0};
    var mk = model.toLowerCase();
    var fbTrig = triggers.fbByModel[mk] || 0;
    var waTrig = triggers.waByModel[mk] || 0;
    var gaTrig = triggers.gaByModel[mk] || 0;
    rows.push(makeRow(null, brand, model,
      fb.spends, fb.leads, fbTrig,
      ga.spends, ga.leads, gaTrig,
      wa.spends, wa.leads, waTrig));
  });

  return rows;
}

// ── ROW BUILDER ───────────────────────────────────────────────
function makeRow(date, brand, model,
    fbSpends, fbLeads, fbTriggered,
    gaSpends, gaLeads, gaTriggered,
    waSpends, waLeads, waTriggered) {
  var fS = Math.round(fbSpends    * 100) / 100;
  var fL = Math.round(fbLeads     * 100) / 100;
  var fT = Math.round(fbTriggered) || 0;
  var gS = Math.round(gaSpends    * 100) / 100;
  var gL = Math.round(gaLeads     * 100) / 100;
  var gT = Math.round(gaTriggered) || 0;
  var wS = Math.round(waSpends    * 100) / 100;
  var wL = Math.round(waLeads     * 100) / 100;
  var wT = Math.round(waTriggered) || 0;

  // Combined includes FB + GA + WA spends and leads
  var cS = fS + gS + wS;
  var cL = fL + gL + wL;
  var cT = fT + gT + wT;

  function kpi(spends, leads, trig) {
    return {
      cpl:  leads > 0 ? Math.round(spends / leads * 100) / 100 : null,
      tcpl: trig  > 0 && spends > 0 ? Math.round(spends / trig * 100) / 100 : null,
      tpct: leads > 0 ? Math.round(trig / leads * 10000) / 100 : null
    };
  }
  var fk = kpi(fS, fL, fT), gk = kpi(gS, gL, gT), wk = kpi(wS, wL, wT), ck = kpi(cS, cL, cT);

  var row = {};
  if (date) row.date = date;
  row.brand = brand;
  row.model = model;
  row.fb_spends       = fS;  row.fb_leads       = fL;  row.fb_triggered       = fT;
  row.fb_cpl          = fk.cpl;  row.fb_tcpl    = fk.tcpl;  row.fb_trigger_pct = fk.tpct;
  row.ga_spends       = gS;  row.ga_leads       = gL;  row.ga_triggered       = gT;
  row.ga_cpl          = gk.cpl;  row.ga_tcpl    = gk.tcpl;  row.ga_trigger_pct = gk.tpct;
  row.wa_spends       = wS;  row.wa_leads       = wL;  row.wa_triggered       = wT;
  row.wa_cpl          = wk.cpl;  row.wa_tcpl    = wk.tcpl;  row.wa_trigger_pct = wk.tpct;
  row.combined_spends = cS;  row.combined_leads = cL;  row.combined_triggered = cT;
  row.combined_cpl    = ck.cpl;  row.combined_tcpl = ck.tcpl; row.combined_trigger_pct = ck.tpct;
  return row;
}

// ── PUSH JSON TO GITHUB ───────────────────────────────────────
function pushJsonToGitHub(token, jsonStr, dateMin, dateMax) {
  var apiUrl = 'https://api.github.com/repos/' + REPO_OWNER + '/' + REPO_NAME + '/contents/' + FILE_PATH;

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
