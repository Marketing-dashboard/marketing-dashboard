// ============================================================
// Used Cars Dashboard Auto-Updater
// Install in Google Sheet: Extensions → Apps Script
// ============================================================

// ── CONFIGURATION ─────────────────────────────────────────────
var REPO_OWNER = 'marketing-dashboard';
var REPO_NAME  = 'marketing-dashboard';
var FILE_PATH  = 'used_cars.html';

// Run this once to store your GitHub token:
//   setGitHubToken('ghp_yourTokenHere')
function setGitHubToken(token) {
  PropertiesService.getScriptProperties().setProperty('ghp_DQOizXKsizU9klaKj14qm52WQ2HSIQ3Ax', token);
  Logger.log('Token saved.');
}

// ── ENTRY POINT (also set as time-driven trigger) ─────────────
function updateDashboard() {
  var token = PropertiesService.getScriptProperties().getProperty('GITHUB_TOKEN');
  if (!token) throw new Error('GitHub token not set. Run setGitHubToken("ghp_...") first.');

  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var refData    = ss.getSheetByName('Reference').getDataRange().getValues();
  var googleData = ss.getSheetByName('Google').getDataRange().getValues();
  var fbData     = ss.getSheetByName('Facebook').getDataRange().getValues();
  var trigData   = ss.getSheetByName('Triggers').getDataRange().getValues();

  // ── 1. Build campaign → type map from Reference tab ──────────
  // Reference cols: [0]=Campaign, [1]=Campaign_type
  var KNOWN_TYPES = {
    'Search':true, 'Catalogue':true, 'Demand Gen':true,
    'Performance Max':true, 'LeadForm':true, 'Website Conversion':true, 'WhatsApp':true
  };
  var typeMap = {};
  for (var i = 1; i < refData.length; i++) {
    var n = String(refData[i][0] || '').trim();
    var t = String(refData[i][1] || '').trim();
    if (n && !typeMap[n]) typeMap[n] = t;
  }
  // Fix any rows where col B was accidentally filled with wrong data
  for (var campKey in typeMap) {
    if (!KNOWN_TYPES[typeMap[campKey]]) {
      var nl = campKey.toLowerCase();
      if (nl.indexOf('leadform') !== -1 || nl.indexOf('lead-form') !== -1) typeMap[campKey] = 'LeadForm';
      else if (nl.indexOf('catalogue') !== -1) typeMap[campKey] = 'Catalogue';
      else if (nl.indexOf('whatsup') !== -1 || nl.indexOf('whatsapp') !== -1) typeMap[campKey] = 'WhatsApp';
      else if (nl.indexOf('pmax') !== -1) typeMap[campKey] = 'Performance Max';
      else if (nl.indexOf('static') !== -1 || nl.indexOf('website') !== -1) typeMap[campKey] = 'Website Conversion';
      else if (nl.indexOf('demand') !== -1) typeMap[campKey] = 'Demand Gen';
      else typeMap[campKey] = 'Search';
    }
  }

  // ── 2. Aggregate Google spend data ───────────────────────────
  // Google cols: [0]=Day, [1]=Campaign, [2]=Campaign_type, [3]=Impressions, [4]=Cost, [5]=Clicks, [6]=Conversions
  var gByC = {};
  var gDailyRows = [];
  for (var i = 1; i < googleData.length; i++) {
    var row  = googleData[i];
    var camp = String(row[1] || '').trim();
    var day  = row[0];
    if (!camp || !day) continue;
    var cost  = parseFloat(row[4]) || 0;
    var impr  = parseFloat(row[3]) || 0;
    var cl    = parseFloat(row[5]) || 0;
    var cv    = parseFloat(row[6]) || 0;
    if (!gByC[camp]) gByC[camp] = {sp:0, imp:0, cl:0, cv:0, spendDays:{}};
    gByC[camp].sp  += cost;
    gByC[camp].imp += impr;
    gByC[camp].cl  += cl;
    gByC[camp].cv  += cv;
    if (cost > 0) gByC[camp].spendDays[fmtDate(day)] = 1;
    var dateStr = fmtDate(day);
    if (dateStr && (cost > 0 || impr > 0)) {
      gDailyRows.push({Date:dateStr, Campaign:camp, Spends:Math.round(cost*100)/100, Impressions:Math.round(impr), Clicks:Math.round(cl), Conversions:Math.round(cv*100)/100});
    }
  }

  // ── 3. Aggregate Facebook spend data ─────────────────────────
  // Facebook cols: [0]=Day, [1]=Campaign, [2]=Conversions, [3]=Cost, [4]=Impressions, [5]=Clicks(all)
  var fbByC = {};
  var fbDailyRows = [];
  for (var i = 1; i < fbData.length; i++) {
    var row  = fbData[i];
    var camp = String(row[1] || '').trim();
    var day  = row[0];
    if (!camp || !day) continue;
    var cost  = parseFloat(row[3]) || 0;
    var impr  = parseFloat(row[4]) || 0;
    var cl    = parseFloat(row[5]) || 0;
    var cv    = parseFloat(row[2]) || 0;
    if (!fbByC[camp]) fbByC[camp] = {sp:0, imp:0, cl:0, cv:0, spendDays:{}};
    fbByC[camp].sp  += cost;
    fbByC[camp].imp += impr;
    fbByC[camp].cl  += cl;
    fbByC[camp].cv  += cv;
    if (cost > 0) fbByC[camp].spendDays[fmtDate(day)] = 1;
    var dateStr = fmtDate(day);
    if (dateStr && (cost > 0 || impr > 0)) {
      fbDailyRows.push({Date:dateStr, Campaign:camp, Spends:Math.round(cost*100)/100, Impressions:Math.round(impr), Clicks:Math.round(cl), Conversions:Math.round(cv*100)/100});
    }
  }

  // ── 4. Aggregate Triggers data ───────────────────────────────
  // Triggers cols: [0]=Date, [1]=utm_campaign, [2]=Listing_Group, [3]=Total_lead, [4]=Unique, [5]=Type
  var trigByC = {};
  for (var i = 1; i < trigData.length; i++) {
    var row      = trigData[i];
    var camp     = String(row[1] || '').trim();
    var listGrp  = String(row[2] || '').trim().toLowerCase();
    var unique   = parseInt(row[4]) || 0;
    var date     = row[0];
    if (!camp) continue;
    if (!trigByC[camp]) trigByC[camp] = {total:0, dealer:0, dealerDays:{}};
    trigByC[camp].total += unique;
    if (listGrp === 'dealer') {
      trigByC[camp].dealer += unique;
      if (date) trigByC[camp].dealerDays[fmtDate(date)] = 1;
    }
  }

  // ── 5. Build DATA_RAW rows ───────────────────────────────────
  var EXP_TYPES = {
    'Demand Gen':true, 'Performance Max':true, 'LeadForm':true,
    'Website Conversion':true, 'WhatsApp':true
  };
  var today = new Date();

  var rows = [];
  var campaigns = Object.keys(typeMap);

  for (var ci = 0; ci < campaigns.length; ci++) {
    var campName = campaigns[ci];
    var campType = typeMap[campName];

    var isGoogle = !!gByC[campName];
    var isFb     = !!fbByC[campName];
    if (!isGoogle && !isFb) continue; // No spend data — skip

    var spend    = isGoogle ? gByC[campName] : fbByC[campName];
    var platform = isGoogle ? 'Google' : 'Meta';

    // Dates
    var spendDayKeys = Object.keys(spend.spendDays).sort();
    var startDate    = spendDayKeys.length ? spendDayKeys[0] : null;
    var lastDate     = spendDayKeys.length ? spendDayKeys[spendDayKeys.length - 1] : null;

    // EndDate: if last spend was > 5 days ago, campaign is ended
    var endDate = null;
    if (lastDate) {
      var diffDays = (today - new Date(lastDate)) / 86400000;
      if (diffDays > 5) endDate = lastDate;
    }

    // Status
    var isExp    = !!EXP_TYPES[campType];
    var isActive = !endDate;
    var status;
    if (isExp  && isActive)  status = 'Experiment · Ongoing';
    else if (isExp  && !isActive) status = 'Experiment · Paused';
    else if (!isExp && isActive)  status = 'Ongoing';
    else                          status = 'Paused';

    // City
    var city = 'Unknown';
    var nameLc = campName.toLowerCase();
    if (nameLc.indexOf('ahm') !== -1 || nameLc.indexOf('ahmedabad') !== -1) city = 'Ahmedabad';
    else if (nameLc.indexOf('chd') !== -1 || nameLc.indexOf('chandigarh') !== -1) city = 'Chandigarh';
    else if (nameLc.indexOf('nashik') !== -1 || nameLc.indexOf('nasik') !== -1) city = 'Nashik';

    // Triggers — exact match, then prefix-based fuzzy match
    var trig = trigByC[campName];
    if (!trig) {
      var prefix = campName.substring(0, Math.min(campName.length, 12));
      var trigKeys = Object.keys(trigByC);
      for (var ti = 0; ti < trigKeys.length; ti++) {
        if (trigKeys[ti].indexOf(prefix) === 0 || campName.indexOf(trigKeys[ti].substring(0, 12)) === 0) {
          trig = trigByC[trigKeys[ti]];
          break;
        }
      }
    }

    var triggeredLeads = trig ? trig.total  : 0;
    var dealerTriggers = trig ? trig.dealer : 0;

    // LeadForm: extrapolate dealer triggers via daily rate × spend days
    if (campType === 'LeadForm' && trig && trig.dealer > 0) {
      var daysWithTrigData = Object.keys(trig.dealerDays).length;
      if (daysWithTrigData > 0) {
        var dailyRate  = trig.dealer / daysWithTrigData;
        var spendDays  = spendDayKeys.length;
        dealerTriggers = Math.round(dailyRate * spendDays);
      }
    }

    // WhatsApp: no trigger data — set 0 (dashboard makes it editable)
    if (campType === 'WhatsApp') {
      triggeredLeads = 0;
      dealerTriggers = 0;
    }

    rows.push({
      Campaign:        campName,
      Spends:          Math.round(spend.sp  * 100) / 100,
      Impressions:     Math.round(spend.imp),
      Clicks:          Math.round(spend.cl),
      Conversions:     Math.round(spend.cv  * 100) / 100,
      Platform:        platform,
      City:            city,
      Type:            campType,
      Triggered_Leads: triggeredLeads,
      Dealer_Triggers: dealerTriggers,
      Status:          status,
      StartDate:       startDate,
      EndDate:         endDate
    });
  }

  // Sort: city → campaign name
  rows.sort(function(a, b) {
    return a.City.localeCompare(b.City) || a.Campaign.localeCompare(b.Campaign);
  });

  // ── 6. Serialize to JS ───────────────────────────────────────
  var lines = rows.map(function(r) {
    return '  {Campaign:'        + jstr(r.Campaign)
         + ', Spends:'           + r.Spends
         + ', Impressions:'      + r.Impressions
         + ', Clicks:'           + r.Clicks
         + ', Conversions:'      + r.Conversions
         + ', Platform:'         + jstr(r.Platform)
         + ', City:'             + jstr(r.City)
         + ', Type:'             + jstr(r.Type)
         + ', Triggered_Leads:'  + r.Triggered_Leads
         + ', Dealer_Triggers:'  + r.Dealer_Triggers
         + ', Status:'           + jstr(r.Status)
         + ', StartDate:'        + (r.StartDate ? jstr(r.StartDate) : 'null')
         + ', EndDate:'          + (r.EndDate   ? jstr(r.EndDate)   : 'null')
         + '}';
  });
  var newDataRaw = 'var DATA_RAW = [\n' + lines.join(',\n') + '\n];';

  // ── 7. Serialize DATA_DAILY ──────────────────────────────────
  var allDaily = gDailyRows.concat(fbDailyRows);
  allDaily.sort(function(a, b) { return a.Date < b.Date ? -1 : a.Date > b.Date ? 1 : 0; });
  var dailyLines = allDaily.map(function(d) {
    return '  {Date:' + jstr(d.Date) + ',Campaign:' + jstr(d.Campaign)
         + ',Spends:' + d.Spends + ',Impressions:' + d.Impressions
         + ',Clicks:' + d.Clicks + ',Conversions:' + d.Conversions + '}';
  });
  var newDataDaily = 'var DATA_DAILY = [\n' + dailyLines.join(',\n') + '\n];';

  // ── 8. Push to GitHub ────────────────────────────────────────
  pushToGitHub(token, newDataRaw, newDataDaily);
  Logger.log('Done. ' + rows.length + ' campaigns, ' + allDaily.length + ' daily rows pushed.');
}

// ── GITHUB HELPER ─────────────────────────────────────────────
function pushToGitHub(token, newDataRaw, newDataDaily) {
  var apiUrl = 'https://api.github.com/repos/' + REPO_OWNER + '/' + REPO_NAME + '/contents/' + FILE_PATH;

  // Fetch current file
  var getResp = UrlFetchApp.fetch(apiUrl, {
    headers: {
      'Authorization': 'token ' + token,
      'Accept': 'application/vnd.github.v3+json'
    },
    muteHttpExceptions: true
  });
  if (getResp.getResponseCode() !== 200) {
    throw new Error('GitHub GET failed: ' + getResp.getContentText());
  }

  var fileInfo = JSON.parse(getResp.getContentText());
  var sha = fileInfo.sha;
  // Decode base64 content (GitHub wraps in newlines)
  var b64 = fileInfo.content.replace(/\n/g, '');
  var currentContent = Utilities.newBlob(Utilities.base64Decode(b64)).getDataAsString();

  // Replace DATA_RAW and DATA_DAILY blocks
  var newContent = currentContent.replace(/var DATA_RAW = \[[\s\S]*?\];/, newDataRaw);
  if (newDataDaily) {
    newContent = newContent.replace(/var DATA_DAILY = \[[\s\S]*?\];/, newDataDaily);
  }

  if (newContent === currentContent) {
    Logger.log('No changes detected — skipping push.');
    return;
  }

  // Push updated file
  var payload = JSON.stringify({
    message: 'Auto-update DATA_RAW from Google Sheet [' + new Date().toISOString().split('T')[0] + ']',
    content: Utilities.base64Encode(Utilities.newBlob(newContent).getBytes()),
    sha:     sha,
    branch:  'main'
  });

  var putResp = UrlFetchApp.fetch(apiUrl, {
    method: 'PUT',
    headers: {
      'Authorization':  'token ' + token,
      'Content-Type':   'application/json',
      'Accept':         'application/vnd.github.v3+json'
    },
    payload: payload,
    muteHttpExceptions: true
  });

  if (putResp.getResponseCode() !== 200 && putResp.getResponseCode() !== 201) {
    throw new Error('GitHub PUT failed (' + putResp.getResponseCode() + '): ' + putResp.getContentText());
  }
  Logger.log('Pushed successfully: ' + putResp.getResponseCode());
}

// ── UTILITIES ─────────────────────────────────────────────────
function fmtDate(val) {
  if (!val) return '';
  if (val instanceof Date) return Utilities.formatDate(val, 'UTC', 'yyyy-MM-dd');
  var s = val.toString().trim();
  // Already ISO format
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.substring(0, 10);
  // Try parsing
  var d = new Date(s);
  if (!isNaN(d)) return Utilities.formatDate(d, 'UTC', 'yyyy-MM-dd');
  return s;
}

function jstr(s) {
  return JSON.stringify(String(s));
}

// ── TRIGGER SETUP (run once to schedule daily updates) ────────
function createDailyTrigger() {
  // Delete existing triggers for updateDashboard
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'updateDashboard') ScriptApp.deleteTrigger(t);
  });
  // Create new daily trigger at 8 AM
  ScriptApp.newTrigger('updateDashboard')
    .timeBased()
    .everyDays(1)
    .atHour(10)
    .create();
  Logger.log('Daily trigger created (8 AM).');
}
