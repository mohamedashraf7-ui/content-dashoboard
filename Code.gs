// ═══ ACCESS CONTROL — update agent emails before sharing ═══════════════════
// dual:true → this manager can also switch to agent view
// agentName must match exactly what appears in column Z of the sheet
var MANAGERS = [
  { email: 'mohamed.ashraf.7@talabat.com', dual: true,  agentName: 'Mohamed Gadallah' },
  { email: 'alaa.talaat@talabat.com',       dual: false, agentName: null               },
];

var TEAM_EMAILS = [
  { email: 'seleem.seleem@talabat.com', name: 'Seliem Mohamed' },
  { email: 'mahmoud.wahid@talabat.com', name: 'Mahmoud Amin'   },
  { email: 'omar.m@talabat.com',        name: 'Omar Elattar'   },
];

// ═══ ONE-TIME SETUP — run this ONCE from the Script editor, then delete ══════
// Run this function manually in Apps Script editor to pre-load credentials
function _initWaCredentials() {
  var props = PropertiesService.getScriptProperties();
  var raw   = props.getProperty('WA_CONFIG');
  var cfg   = {};
  try { cfg = JSON.parse(raw||'{}'); } catch(e) {}
  cfg.apikey = '5654406';
  cfg.phone  = '201020809266'; // digits only, no +
  props.setProperty('WA_CONFIG', JSON.stringify(cfg));
  Logger.log('✅ WhatsApp credentials saved to Script Properties.');
  Logger.log('Phone: +201020809266 | API key: gBb3AH5K4cgq');
  Logger.log('You can now delete this function from the code.');
}

// ═══ WEB APP ENTRY ══════════════════════════════════════════════════════════
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Content Team Dashboard · Talabat Egypt')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ═══ GET CURRENT USER ═══════════════════════════════════════════════════════
function getUserInfo() {
  var email = Session.getActiveUser().getEmail();
  if (!email) return { error: 'Could not detect your Google account. Make sure you are signed in with your Talabat email.' };

  // Check managers
  for (var m = 0; m < MANAGERS.length; m++) {
    if (MANAGERS[m].email === email) {
      var first = email.split('@')[0].split('.')[0];
      first = first.charAt(0).toUpperCase() + first.slice(1);
      return { email: email, role: 'manager', dual: MANAGERS[m].dual, name: first, agentName: MANAGERS[m].agentName || null };
    }
  }

  // Check known agents
  for (var i = 0; i < TEAM_EMAILS.length; i++) {
    if (TEAM_EMAILS[i].email === email) {
      return { email: email, role: 'agent', dual: false, name: TEAM_EMAILS[i].name, agentName: TEAM_EMAILS[i].name };
    }
  }

  // Any other @talabat.com or @deliveryhero.com → agent view (overview only, no profile)
  if (email.endsWith('@talabat.com') || email.endsWith('@deliveryhero.com')) {
    var displayName = email.split('@')[0].replace(/\./g, ' ').replace(/\b\w/g, function(c){ return c.toUpperCase(); });
    return { email: email, role: 'agent', dual: false, name: displayName, agentName: null };
  }

  return { email: email, role: 'unknown', name: '', agentName: null };
}

// ═══ READ SHEET DATA ════════════════════════════════════════════════════════
function getData() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('SalesForce 2026');
  if (!sheet) return { error: 'Sheet tab "SalesForce 2026" not found' };

  var data = sheet.getDataRange().getValues();
  var rows = [];
  for (var i = 1; i < data.length; i++) {
    var row   = data[i];
    var owner = String(row[25] || '').trim();
    var status= String(row[26] || '').trim();
    if (!owner && !status) continue;
    rows.push({
      caseOpenDate : fmtDate(row[0]),
      caseNo       : String(row[3]  || ''),
      vendor       : String(row[4]  || ''),
      caseType     : String(row[5]  || ''),
      skuAdded     : String(row[21] || ''), // col V
      owner        : owner,
      status       : status,
      startDate    : fmtDate(row[27]),
      pendingFdbk  : String(row[28] || ''),
      endDate      : fmtDate(row[29]),
      comment      : String(row[30] || ''),
    });
  }
  return { rows: rows, total: rows.length };
}

// ═══ COACHING NOTES — reads/writes "Coaching Notes" tab ══════════════════
function getNotes(agentName) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Coaching Notes');
  if (!sheet) return [];
  var data  = sheet.getDataRange().getValues();
  var notes = [];
  for (var i = 1; i < data.length; i++) {
    var row    = data[i];
    var agent  = String(row[1] || '').trim();
    var status = String(row[4] || '').trim().toLowerCase();
    if (agent === agentName && status !== 'done') {
      notes.push({ rowNum: i + 1, agent: agent, date: String(row[2] || ''), text: String(row[3] || ''), status: String(row[4] || ''), createdBy: String(row[5] || '') });
    }
  }
  return notes;
}

function addNote(agentName, date, text) {
  var email = Session.getActiveUser().getEmail();
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Coaching Notes');
  if (!sheet) {
    sheet = ss.insertSheet('Coaching Notes');
    sheet.appendRow(['ID', 'Agent', 'Date', 'Note', 'Status', 'Created By', 'Done Date']);
  }
  sheet.appendRow([sheet.getLastRow(), agentName, date, text, 'Active', email, '']);
  return { success: true };
}

function markNoteDone(rowNum) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Coaching Notes');
  if (!sheet) return { error: 'Coaching Notes tab not found' };
  var today = new Date();
  sheet.getRange(rowNum, 5).setValue('Done');
  sheet.getRange(rowNum, 7).setValue(today.getFullYear() + '-' + pad2(today.getMonth()+1) + '-' + pad2(today.getDate()));
  return { success: true };
}

// ═══ HELPERS ════════════════════════════════════════════════════════════════
function pad2(n) { return n < 10 ? '0' + n : '' + n; }

function fmtDate(val) {
  if (!val || val === '') return '';
  if (val instanceof Date) return val.getFullYear()+'-'+pad2(val.getMonth()+1)+'-'+pad2(val.getDate());
  var s = String(val).trim();
  if (s === '--' || s === '') return '';
  var d = new Date(s);
  if (!isNaN(d.getTime())) return d.getFullYear()+'-'+pad2(d.getMonth()+1)+'-'+pad2(d.getDate());
  return s;
}

// ═══ SETTINGS — daily targets per agent per month ════════════════════════════
function getSettings() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var targets = [];
  var sheet = ss.getSheetByName('Settings');
  if (sheet) {
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!String(row[0] || '').trim()) continue;
      targets.push({
        agent  : String(row[0] || '').trim(),
        month  : String(row[1] || '').trim(),
        target : parseInt(row[2]) || 15,
        project: String(row[3] || '').trim(),
        notes  : String(row[4] || '').trim()
      });
    }
  }
  return { targets: targets };
}

function saveTarget(agent, month, target, project, notes) {
  var email = Session.getActiveUser().getEmail();
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Settings');
  if (!sheet) {
    sheet = ss.insertSheet('Settings');
    sheet.appendRow(['Agent','Month','Daily Target','Project','Notes','Updated By','Updated At']);
  }
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === agent && String(data[i][1]).trim() === month) {
      sheet.getRange(i+1,1,1,7).setValues([[agent,month,parseInt(target)||15,project||'',notes||'',email,new Date()]]);
      return { success: true };
    }
  }
  sheet.appendRow([agent,month,parseInt(target)||15,project||'',notes||'',email,new Date()]);
  return { success: true };
}

// ═══ ATTENDANCE — leave / WFH per agent per day ══════════════════════════════
function getAttendance(yearMonth) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Attendance');
  if (!sheet) return [];
  var data  = sheet.getDataRange().getValues();
  var result = [];
  for (var i = 1; i < data.length; i++) {
    var row    = data[i];
    var agent  = String(row[0] || '').trim();
    var dateStr= fmtDate(row[1]);
    if (!agent || !dateStr) continue;
    if (yearMonth && !dateStr.startsWith(yearMonth)) continue;
    result.push({ agent: agent, date: dateStr, status: String(row[2] || 'Present').trim(), notes: String(row[3] || '').trim() });
  }
  return result;
}

// ═══ AGENT STATUS — queue vs special project ════════════════════════════════
// Sheet columns: Agent | Mode | Project Name | Progress | Project Target | Updated By | Updated At | Start Date | End Date
function getAgentModes() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Agent Status');
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  var result = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!String(row[0]||'').trim()) continue;
    var sd = row[7] ? (row[7] instanceof Date ? row[7].toISOString().slice(0,10) : String(row[7]).slice(0,10)) : '';
    var ed = row[8] ? (row[8] instanceof Date ? row[8].toISOString().slice(0,10) : String(row[8]).slice(0,10)) : '';
    result.push({
      agent           : String(row[0]||'').trim(),
      mode            : String(row[1]||'queue').trim(),
      projectName     : String(row[2]||'').trim(),
      progress        : String(row[3]||'').trim(),
      projectTarget   : row[4]?parseFloat(row[4])||null:null,
      projectStartDate: sd,
      projectEndDate  : ed
    });
  }
  return result;
}

function saveAgentMode(agent, mode, projectName, progress, projectTarget, startDate, endDate) {
  var email = Session.getActiveUser().getEmail();
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Agent Status');
  if (!sheet) {
    sheet = ss.insertSheet('Agent Status');
    sheet.appendRow(['Agent','Mode','Project Name','Progress','Project Target','Updated By','Updated At','Start Date','End Date']);
  }
  var tgt = projectTarget ? parseFloat(projectTarget) || '' : '';
  var sd  = startDate || '';
  var ed  = endDate   || '';
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === agent) {
      sheet.getRange(i+1,1,1,9).setValues([[agent,mode,projectName||'',progress||'',tgt,email,new Date(),sd,ed]]);
      return { success: true };
    }
  }
  sheet.appendRow([agent,mode,projectName||'',progress||'',tgt,email,new Date(),sd,ed]);
  return { success: true };
}

// ═══ AGENT RATINGS — manual TL rating for project agents ════════════════════
// Sheet columns: Agent | Rating | Notes | Updated By | Updated At
function getAgentRatings() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Agent Ratings');
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  var result = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!String(row[0]||'').trim()) continue;
    result.push({
      agent : String(row[0]||'').trim(),
      rating: String(row[1]||'').trim(),
      notes : String(row[2]||'').trim()
    });
  }
  return result;
}

function saveAgentRating(agent, rating, notes) {
  var email = Session.getActiveUser().getEmail();
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Agent Ratings');
  if (!sheet) {
    sheet = ss.insertSheet('Agent Ratings');
    sheet.appendRow(['Agent','Rating','Notes','Updated By','Updated At']);
  }
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === agent) {
      sheet.getRange(i+1,1,1,5).setValues([[agent,rating,notes||'',email,new Date()]]);
      return { success: true };
    }
  }
  sheet.appendRow([agent,rating,notes||'',email,new Date()]);
  return { success: true };
}

// ═══ PROJECT CONFIG — type-specific tracker data per project ════════════════
// Sheet: Project Name | Type | Sheet URL | Chain Name | Total SKU | Corrected SKU
//        | Total Assortment | New SKUs | Excluded Items | Notes | Updated By | Updated At
function getProjectConfigs() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Project Config');
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  var result = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!String(row[0]||'').trim()) continue;
    result.push({
      projectName     : String(row[0]||'').trim(),
      type            : String(row[1]||'other').trim(),
      sheetUrl        : String(row[2]||'').trim(),
      chainName       : String(row[3]||'').trim(),
      totalSku        : row[4] ? parseInt(row[4])||0 : 0,
      correctedSku    : row[5] ? parseInt(row[5])||0 : 0,
      totalAssortment : row[6] ? parseInt(row[6])||0 : 0,
      newSkus         : row[7] ? parseInt(row[7])||0 : 0,
      excludedItems   : row[8] ? parseInt(row[8])||0 : 0,
      notes           : String(row[9]||'').trim()
    });
  }
  return result;
}

function saveProjectConfig(projectName, type, sheetUrl, chainName, totalSku, correctedSku, totalAssortment, newSkus, excludedItems, notes) {
  var email = Session.getActiveUser().getEmail();
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Project Config');
  if (!sheet) {
    sheet = ss.insertSheet('Project Config');
    sheet.appendRow(['Project Name','Type','Sheet URL','Chain Name','Total SKU','Corrected SKU','Total Assortment','New SKUs','Excluded Items','Notes','Updated By','Updated At']);
  }
  var row = [projectName, type||'other', sheetUrl||'', chainName||'',
             parseInt(totalSku)||0, parseInt(correctedSku)||0,
             parseInt(totalAssortment)||0, parseInt(newSkus)||0, parseInt(excludedItems)||0,
             notes||'', email, new Date()];
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === projectName) {
      sheet.getRange(i+1,1,1,12).setValues([row]);
      return { success: true };
    }
  }
  sheet.appendRow(row);
  return { success: true };
}

// ═══ PROJECT UPDATES — timestamped log per agent ════════════════════════════
function getProjectUpdates() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Project Updates');
  if (!sheet) return [];
  var data   = sheet.getDataRange().getValues();
  var result = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!String(row[0]||'').trim()) continue;
    result.push({
      agent  : String(row[0]||'').trim(),
      date   : fmtDate(row[1]),
      text   : String(row[2]||'').trim(),
      by     : String(row[3]||'').trim()
    });
  }
  result.reverse(); // newest first
  return result;
}

function addProjectUpdate(agent, text) {
  var email = Session.getActiveUser().getEmail();
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Project Updates');
  if (!sheet) {
    sheet = ss.insertSheet('Project Updates');
    sheet.appendRow(['Agent','Date','Update','Added By']);
  }
  sheet.appendRow([agent, new Date(), text, email]);
  return { success: true };
}

function deleteProjectUpdate(agent, dateStr, text) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Project Updates');
  if (!sheet) return { success: true };
  var data = sheet.getDataRange().getValues();
  for (var i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]).trim()===agent && fmtDate(data[i][1])===dateStr && String(data[i][2]).trim()===text) {
      sheet.deleteRow(i+1);
      return { success: true };
    }
  }
  return { success: true };
}

// ═══ SLA SETTINGS — target days per case type ═══════════════════════════════
function getSLATargets() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('SLA Settings');
  if (!sheet) return { targets: [] };
  var data = sheet.getDataRange().getValues();
  var targets = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!String(row[0]||'').trim()) continue;
    targets.push({ caseType: String(row[0]||'').trim(), targetHours: parseFloat(row[1])||9, notes: String(row[2]||'').trim() });
  }
  return { targets: targets };
}

function saveSLATarget(caseType, targetHours, notes) {
  var email = Session.getActiveUser().getEmail();
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('SLA Settings');
  if (!sheet) {
    sheet = ss.insertSheet('SLA Settings');
    sheet.appendRow(['Case Type','Target Hours (Business)','Notes','Updated By','Updated At']);
  }
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === caseType) {
      sheet.getRange(i+1,1,1,5).setValues([[caseType,parseFloat(targetHours)||9,notes||'',email,new Date()]]);
      return { success: true };
    }
  }
  sheet.appendRow([caseType,parseFloat(targetHours)||9,notes||'',email,new Date()]);
  return { success: true };
}

function deleteSLATarget(caseType) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('SLA Settings');
  if (!sheet) return { success: true };
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === caseType) { sheet.deleteRow(i+1); return { success: true }; }
  }
  return { success: true };
}

// ═══ DRIVE DATA — read Google Sheets from Drive folder ═══════════════════════
function getDriveData(caseNo) {
  try {
    var folderId = '16UhJIyxDXa8Y9Allnot1lvKB57W5Dy49';
    var folder   = DriveApp.getFolderById(folderId);
    var files    = folder.getFilesByType(MimeType.GOOGLE_SHEETS);
    var matches  = [];
    var headers  = [];
    var caseNoStr= caseNo ? String(caseNo).trim().toLowerCase() : '';

    while (files.hasNext()) {
      var file = files.next();
      var ss2  = SpreadsheetApp.openById(file.getId());
      var sheets2 = ss2.getSheets();
      for (var s = 0; s < sheets2.length; s++) {
        var sheet2 = sheets2[s];
        var data2  = sheet2.getDataRange().getValues();
        if (data2.length < 2) continue;
        var hdrs = data2[0].map(function(h){ return String(h).trim(); });
        if (!headers.length) headers = hdrs;
        for (var i = 1; i < data2.length; i++) {
          var row = {};
          hdrs.forEach(function(h,idx){
            var v = data2[i][idx];
            row[h] = v instanceof Date ? fmtDate(v) : String(v||'');
          });
          row['_source'] = file.getName();
          var rowFlat = hdrs.map(function(h){ return String(row[h]||'').toLowerCase(); }).join('|');
          if (!caseNoStr || rowFlat.indexOf(caseNoStr) !== -1) matches.push(row);
        }
      }
    }
    return { rows: matches.slice(0,300), headers: headers, total: matches.length };
  } catch(e) {
    return { error: e.message, rows: [], headers: [] };
  }
}

// ═══ LINKED SHEET INTEGRATION ═══════════════════════════════════════════════

function getSheetHeaders(sheetUrl) {
  try {
    var ss = SpreadsheetApp.openByUrl(sheetUrl.trim());
    var sheet = ss.getSheets()[0];
    var lastCol = sheet.getLastColumn();
    if (lastCol === 0) return { success: true, headers: [], sheetName: sheet.getName() };
    var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    var result = [];
    for (var i = 0; i < headers.length; i++) {
      var h = String(headers[i]||'').trim();
      if (h) result.push({ idx: i, header: h });
    }
    return { success: true, headers: result, sheetName: sheet.getName() };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

function readLinkedSheetData(sheetUrl, mapping) {
  try {
    var ss = SpreadsheetApp.openByUrl(sheetUrl.trim());
    var sheet = ss.getSheets()[0];
    var data = sheet.getDataRange().getValues();
    var res = { success:true, rows:0, chainCount:0, chains:[],
      totalSku:0, correctedSku:0, totalAssortment:0, newSkus:0, notOnPim:0,
      statusDone:0, statusTotal:0 };
    var chainSet = {};
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!row.some(function(c){return c!==''&&c!==null;})) continue;
      res.rows++;
      if (mapping.chainIdx!==undefined&&mapping.chainIdx!==null&&mapping.chainIdx!==''&&row[mapping.chainIdx]) {
        chainSet[String(row[mapping.chainIdx]).trim()] = 1;
      }
      if (mapping.skuIdx!==undefined&&mapping.skuIdx!==''&&mapping.skuIdx!==null) res.totalSku+=parseFloat(row[mapping.skuIdx])||0;
      if (mapping.correctedIdx!==undefined&&mapping.correctedIdx!==''&&mapping.correctedIdx!==null) res.correctedSku+=parseFloat(row[mapping.correctedIdx])||0;
      if (mapping.prevAssortIdx!==undefined&&mapping.prevAssortIdx!==''&&mapping.prevAssortIdx!==null) res.totalAssortment+=parseFloat(row[mapping.prevAssortIdx])||0;
      if (mapping.newSkuIdx!==undefined&&mapping.newSkuIdx!==''&&mapping.newSkuIdx!==null) res.newSkus+=parseFloat(row[mapping.newSkuIdx])||0;
      if (mapping.notOnPimIdx!==undefined&&mapping.notOnPimIdx!==''&&mapping.notOnPimIdx!==null) {
        var v=row[mapping.notOnPimIdx];
        if(v&&String(v).trim()!=='0'&&String(v).trim()!=='') res.notOnPim+=parseFloat(v)||0;
      }
      if (mapping.statusIdx!==undefined&&mapping.statusIdx!==''&&mapping.statusIdx!==null) {
        var s=String(row[mapping.statusIdx]||'').trim().toLowerCase();
        var dv=String(mapping.statusDoneValue||'done').toLowerCase();
        if (s) { res.statusTotal++; if(s===dv) res.statusDone++; }
      }
    }
    res.chainCount = Object.keys(chainSet).length;
    res.correctionPct = res.totalSku>0 ? Math.round(res.correctedSku/res.totalSku*100) : 0;
    res.coveragePct   = res.totalAssortment>0 ? Math.round(res.newSkus/res.totalAssortment*100) : 0;
    res.statusDonePct = res.statusTotal>0 ? Math.round(res.statusDone/res.statusTotal*100) : 0;
    return res;
  } catch(e) {
    return { success:false, error:e.message };
  }
}

function getProjectMappings() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Project Mapping');
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  var result = [];
  for (var i = 1; i < data.length; i++) {
    if (!String(data[i][0]||'').trim()) continue;
    try { result.push({ projectName: String(data[i][0]).trim(), mapping: JSON.parse(String(data[i][1]||'{}')) }); }
    catch(e) {}
  }
  return result;
}

function saveProjectMapping(projectName, mappingJson) {
  var email = Session.getActiveUser().getEmail();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Project Mapping');
  if (!sheet) {
    sheet = ss.insertSheet('Project Mapping');
    sheet.appendRow(['Project Name','Mapping JSON','Updated By','Updated At']);
  }
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === projectName) {
      sheet.getRange(i+1,1,1,4).setValues([[projectName,mappingJson,email,new Date()]]);
      return { success:true };
    }
  }
  sheet.appendRow([projectName,mappingJson,email,new Date()]);
  return { success:true };
}

// ═══ GOOGLE DRIVE CASE MAPPING ═══════════════════════════════════════════════
var DRIVE_FOLDER_ID = '16UhJIyxDXa8Y9Allnot1lvKB57W5Dy49';

// Search for files matching case number across root folder + all subfolders (DriveApp, recursive 3 levels)
function searchDriveForCase(caseNo) {
  try {
    var root = DriveApp.getFolderById(DRIVE_FOLDER_ID);
    var results = [];
    function searchInFolder(folder) {
      var files = folder.searchFiles('title contains "' + caseNo + '"');
      while (files.hasNext()) {
        var f = files.next();
        results.push({ id: f.getId(), name: f.getName(), url: f.getUrl(), mimeType: f.getMimeType() });
      }
    }
    function walkFolders(folder, depth) {
      searchInFolder(folder);
      if (depth >= 3) return;
      var subs = folder.getFolders();
      while (subs.hasNext()) { walkFolders(subs.next(), depth + 1); }
    }
    walkFolders(root, 0);
    return results;
  } catch(e) {
    return [];
  }
}

// List immediate contents of a folder using DriveApp (no advanced service needed)
function browseDriveFolder(folderId) {
  var id = folderId || DRIVE_FOLDER_ID;
  try {
    var folder = DriveApp.getFolderById(id);
    var items = [];
    // Subfolders first
    var subs = folder.getFolders();
    while (subs.hasNext()) {
      var s = subs.next();
      items.push({ id: s.getId(), name: s.getName(), url: s.getUrl(), mimeType: 'application/vnd.google-apps.folder', isFolder: true, modified: '', size: '' });
    }
    // Then files
    var files = folder.getFiles();
    while (files.hasNext()) {
      var f = files.next();
      var bytes = f.getSize ? f.getSize() : 0;
      var sz = bytes > 1048576 ? (Math.round(bytes/1048576*10)/10) + ' MB' : bytes > 1024 ? Math.round(bytes/1024) + ' KB' : bytes > 0 ? bytes + ' B' : '';
      var mod = f.getLastUpdated ? f.getLastUpdated().toISOString().slice(0,10) : '';
      items.push({ id: f.getId(), name: f.getName(), url: f.getUrl(), mimeType: f.getMimeType(), isFolder: false, modified: mod, size: sz });
    }
    // Sort: folders by name, then files by name
    items.sort(function(a,b){ if(a.isFolder!==b.isFolder) return a.isFolder?-1:1; return a.name.localeCompare(b.name); });
    return { folderId: id, items: items };
  } catch(e) {
    return { folderId: id, items: [], error: e.toString() };
  }
}

function saveCaseDriveLink(agent, caseNo, fileName, url) {
  var email = Session.getActiveUser().getEmail();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Case Drive Links');
  if (!sheet) {
    sheet = ss.insertSheet('Case Drive Links');
    sheet.appendRow(['ID','Agent','Case No','File Name','URL','Linked By','Linked At']);
  }
  var id = 'dl_' + new Date().getTime();
  sheet.appendRow([id, agent, caseNo, fileName, url, email, new Date().toISOString()]);
  return { success: true };
}

function getCaseDriveLinks(agent) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Case Drive Links');
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  var results = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (String(row[1]).trim() === agent) {
      results.push({ id: String(row[0]), agent: String(row[1]), caseNo: String(row[2]), fileName: String(row[3]), url: String(row[4]), linkedBy: String(row[5]), linkedAt: String(row[6]) });
    }
  }
  return results;
}

// Returns { caseNo: agentName } for all linked cases — used for quality metrics
function getAllDriveLinkCaseNos() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Case Drive Links');
  if (!sheet) return {};
  var data = sheet.getDataRange().getValues();
  var out = {};
  for (var i = 1; i < data.length; i++) {
    var caseNo = String(data[i][2]).trim();
    var agent  = String(data[i][1]).trim();
    if (caseNo) out[caseNo] = agent;
  }
  return out;
}

function removeCaseDriveLink(linkId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Case Drive Links');
  if (!sheet) return { success: false };
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === linkId) {
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { success: false };
}

function saveAttendance(agent, date, status, notes) {
  var email = Session.getActiveUser().getEmail();
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Attendance');
  if (!sheet) {
    sheet = ss.insertSheet('Attendance');
    sheet.appendRow(['Agent','Date','Status','Notes','Updated By','Updated At']);
  }
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === agent && fmtDate(data[i][1]) === date) {
      if (status === 'Present') { sheet.deleteRow(i+1); return { success: true }; }
      sheet.getRange(i+1,1,1,6).setValues([[agent,date,status,notes||'',email,new Date()]]);
      return { success: true };
    }
  }
  if (status !== 'Present') {
    sheet.appendRow([agent,date,status,notes||'',email,new Date()]);
  }
  return { success: true };
}

function getSkuTargets() {
  var raw = PropertiesService.getScriptProperties().getProperty('SKU_TARGETS');
  if (!raw) return {};
  try { return JSON.parse(raw); } catch(e) { return {}; }
}

function saveSkuTarget(agent, month, val) {
  var data = getSkuTargets();
  if (!data[agent]) data[agent] = {};
  data[agent][month] = val;
  PropertiesService.getScriptProperties().setProperty('SKU_TARGETS', JSON.stringify(data));
  return { success: true };
}

// ═══ WHATSAPP — CallMeBot ════════════════════════════════════════════════════

function getWaConfig() {
  var raw = PropertiesService.getScriptProperties().getProperty('WA_CONFIG');
  if (!raw) return {};
  try { return JSON.parse(raw); } catch(e) { return {}; }
}

function saveWaConfig(cfg) {
  PropertiesService.getScriptProperties().setProperty('WA_CONFIG', JSON.stringify(cfg));
  // Rebuild time-based triggers whenever schedule is saved
  _setupWaTriggers(cfg);
  var activeTimes = [1,2,3,4].filter(function(i){ return cfg['time'+i] && (cfg['time'+i+'on']===true||cfg['time'+i+'on']==='true'); });
  return { success: true, triggersSet: activeTimes.length };
}

function _sendWa(text, phone, apikey) {
  if (!phone || !apikey) return { ok: false, error: 'Not configured' };
  try {
    var cleanPhone = phone.replace(/^\+/, ''); // CallMeBot prefers digits only
    var url = 'https://api.callmebot.com/whatsapp.php?phone='
      + encodeURIComponent(cleanPhone)
      + '&text=' + encodeURIComponent(text)
      + '&apikey=' + encodeURIComponent(apikey);
    var resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    var code = resp.getResponseCode();
    return (code === 200 || code === 201) ? { ok: true } : { ok: false, error: 'HTTP ' + code };
  } catch(e) {
    return { ok: false, error: e.toString() };
  }
}

function sendWaTest(phone, apikey) {
  var text = '✅ *Content Team Dashboard*\n\nTest message sent successfully! Your WhatsApp notifications are configured correctly. 🎉';
  return _sendWa(text, phone, apikey);
}

function _buildProgressMessage() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('SalesForce 2026');
  if (!sheet) return null;

  var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  var data  = sheet.getDataRange().getValues();
  var hdr   = data[0].map(function(h){ return String(h).trim().toLowerCase(); });

  var ownerIdx  = hdr.indexOf('content owner');
  var statusIdx = hdr.indexOf('status');
  var endIdx    = hdr.indexOf('end date');
  var startIdx  = hdr.indexOf('start date');

  if (ownerIdx < 0) {
    // fallback: use known column positions Z=25, AA=26, AB=27, AD=29
    ownerIdx  = 25; statusIdx = 26; startIdx = 27; endIdx = 29;
  }

  var TEAM_NAMES = ['Mohamed Gadallah','Seliem Mohamed','Omar Elattar','Mahmoud Amin'];
  var DEFAULT_TARGET = 15;

  // Load targets from Properties
  var tgtRaw = PropertiesService.getScriptProperties().getProperty('AGENT_TARGETS');
  var tgtMap = {};
  try { tgtMap = JSON.parse(tgtRaw||'{}'); } catch(e) {}
  var thisM = today.slice(0,7);

  var stats = {};
  TEAM_NAMES.forEach(function(n){ stats[n] = { done:0, pending:0, pfb:0, rej:0 }; });

  for (var i = 1; i < data.length; i++) {
    var row    = data[i];
    var owner  = String(row[ownerIdx]||'').trim();
    var status = String(row[statusIdx]||'').trim();
    var endRaw = row[endIdx];
    var endISO = endRaw ? Utilities.formatDate(new Date(endRaw), Session.getScriptTimeZone(), 'yyyy-MM-dd') : '';
    if (!stats[owner]) continue;
    if (endISO === today)          stats[owner].done++;
    if (status === 'Pending')      stats[owner].pending++;
    if (status === 'Pending Feedback') stats[owner].pfb++;
    if (status === 'Rejected')     stats[owner].rej++;
  }

  var teamDone = 0, teamTgt = 0, lines = [];
  TEAM_NAMES.forEach(function(n){
    var tgt = (tgtMap[n] && tgtMap[n][thisM]) || DEFAULT_TARGET;
    var s   = stats[n];
    var pct = Math.round(s.done / tgt * 100);
    var bar = pct >= 120 ? '🚀' : pct >= 80 ? '✅' : pct >= 50 ? '⚠️' : '🔴';
    lines.push(bar + ' *' + n.split(' ')[0] + '*: ' + s.done + '/' + tgt + ' (' + pct + '%)');
    teamDone += s.done;
    teamTgt  += tgt;
  });

  var teamPct  = Math.round(teamDone / teamTgt * 100);
  var teamIcon = teamPct >= 100 ? '🎯' : teamPct >= 80 ? '📈' : '⚠️';

  var top = TEAM_NAMES.slice().sort(function(a,b){ return stats[b].done - stats[a].done; })[0];
  var totalPend = TEAM_NAMES.reduce(function(s,n){ return s + stats[n].pending + stats[n].pfb; }, 0);
  var totalRej  = TEAM_NAMES.reduce(function(s,n){ return s + stats[n].rej; }, 0);

  var now   = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'HH:mm');
  var dayLbl= Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'EEE dd MMM');

  var msg = '📊 *Team Progress Update*\n'
    + dayLbl + ' · ' + now + '\n'
    + '─────────────────\n'
    + lines.join('\n') + '\n'
    + '─────────────────\n'
    + teamIcon + ' *Team Total: ' + teamDone + '/' + teamTgt + ' (' + teamPct + '%)*\n'
    + '🏆 Top today: ' + top.split(' ')[0] + ' (' + stats[top].done + ' cases)\n'
    + '⏳ Pending: ' + totalPend + '  🔴 Rejected: ' + totalRej;

  // ── Onground Requests section ──────────────────────────────────────────────
  try {
    var mrData = getMainRequestData(null, null);
    if (mrData && !mrData.error) {
      var bad = mrData.byAgentDate      || {};  // assigned per day (AA)
      var aad = mrData.byAgentActionDate || {};  // action per day (AB)
      var mrLines = [];
      TEAM_NAMES.forEach(function(n) {
        var short    = MR_AGENT_SHORT[n] || n.split(' ')[0];
        var assigned = (bad[n] && bad[n][today]) || 0;
        var actions  = (aad[n] && aad[n][today]) || 0;
        mrLines.push('• *' + short + '*: ' + assigned + ' assigned · ' + actions + ' action' + (actions !== 1 ? 's' : ''));
      });
      if (mrLines.length) {
        msg += '\n\n📋 *Onground Requests — Today*\n' + mrLines.join('\n');
      }
    }
  } catch(mrErr) {}

  return msg;
}

function sendWaProgressUpdate() {
  var cfg = getWaConfig();
  if (!cfg.phone || !cfg.apikey) return { ok: false, error: 'WhatsApp not configured in Settings' };
  var msg = _buildProgressMessage();
  if (!msg) return { ok: false, error: 'Could not read sheet data' };
  return _sendWa(msg, cfg.phone, cfg.apikey);
}

function sendWaCustomMessage(text, groupPhone) {
  var cfg = getWaConfig();
  if (!cfg.apikey) return { ok: false, error: 'WhatsApp not configured in Settings' };
  var phone = groupPhone ? groupPhone.replace(/^\+/, '') : cfg.phone;
  if (!phone) return { ok: false, error: 'No phone number configured' };
  return _sendWa(text, phone, cfg.apikey);
}

// ── Time-based trigger management ────────────────────────────────────────────
function _setupWaTriggers(cfg) {
  // Remove all existing WA triggers
  ScriptApp.getProjectTriggers().forEach(function(t){
    if (t.getHandlerFunction() === 'sendWaProgressUpdate') {
      ScriptApp.deleteTrigger(t);
    }
  });
  if (!cfg.scheduleEnabled || cfg.scheduleEnabled === 'false') return;
  [1,2,3,4].forEach(function(i){
    var t   = cfg['time'+i];
    var on  = cfg['time'+i+'on'];
    if (!t || !(on === true || on === 'true')) return;
    var parts = t.split(':');
    var hour  = parseInt(parts[0]);
    var min   = parseInt(parts[1]||'0');
    ScriptApp.newTrigger('sendWaProgressUpdate')
      .timeBased()
      .everyDays(1)
      .atHour(hour)
      .nearMinute(min)
      .inTimezone(Session.getScriptTimeZone())
      .create();
  });
}

// ═══ MAIN REQUEST SHEET ══════════════════════════════════════════════════════
var MR_SHEET_ID  = '1ZI_6rA8IntOgDcRXWsPSpkP62c9dGEySrrXvb3WWYAw';
var MR_TAB_NAME  = 'Main request';
var MR_NAME_MAP  = {
  'ashraf':  'Mohamed Gadallah',
  'seleem':  'Seliem Mohamed',
  'selim':   'Seliem Mohamed',
  'omar':    'Omar Elattar',
  'waheed':  'Mahmoud Amin',
  'wahid':   'Mahmoud Amin'
};

// Column indices (0-based) in Main request sheet
var MR_COL_TIMESTAMP    = 0;   // A – form submission time
var MR_COL_CHAIN        = 3;   // D – Chain Name
var MR_COL_TYPE         = 4;   // E – Request Type
var MR_COL_STATUS       = 20;  // U – Done / Canceled / Pending Regional …
var MR_COL_OWNER        = 21;  // V – Seleem / Ashraf / Omar / Waheed  (1-indexed col 22)
var MR_COL_EMAIL_STATUS = 25;  // Z – done / pending regional
var MR_COL_ASSIGNED_AT  = 26;  // AA – Assigned date ★ written by auto-assign / on-edit trigger
var MR_COL_ACTION_DATE  = 27;  // AB – Action date (read-only — set by the processing agent)

// Queue agents for auto-assignment: full name → short name written into the sheet
var MR_AGENT_SHORT = {
  'Mohamed Gadallah': 'Ashraf',
  'Seliem Mohamed':   'Seleem',
  'Omar Elattar':     'Omar',
  'Mahmoud Amin':     'Waheed'
};
var MR_QUEUE_AGENTS = ['Mohamed Gadallah','Seliem Mohamed','Omar Elattar','Mahmoud Amin'];

// ── ONE-TIME SETUP ─────────────────────────────────────────────────────────
// Run _setupMrTriggers() ONCE from the Apps Script editor.
// It installs both the form-submit (auto-assign) and on-edit (timestamp) triggers.
function _setupMrTriggers() {
  var handlers = ['mrOnOwnerEdit', 'mrAutoAssign'];
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (handlers.indexOf(t.getHandlerFunction()) >= 0) ScriptApp.deleteTrigger(t);
  });
  // Auto-assign on new form submission
  ScriptApp.newTrigger('mrAutoAssign')
    .forSpreadsheet(MR_SHEET_ID)
    .onFormSubmit()
    .create();
  // Stamp timestamp when Owner is manually edited
  ScriptApp.newTrigger('mrOnOwnerEdit')
    .forSpreadsheet(MR_SHEET_ID)
    .onEdit()
    .create();
  Logger.log('✅ mrAutoAssign + mrOnOwnerEdit triggers installed.');
}
// Keep old name as alias so existing installations still work
function _setupMrOwnerTrigger() { _setupMrTriggers(); }

// ── AUTO-ASSIGN: fires on every new form submission ────────────────────────
function mrAutoAssign(e) {
  try {
    var ss    = SpreadsheetApp.openById(MR_SHEET_ID);
    var sheet = ss.getSheetByName(MR_TAB_NAME);
    if (!sheet) return;

    var newRow = e.range.getRow();
    var tz     = Session.getScriptTimeZone();
    var today  = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');

    // Count today's assignments per queue agent
    var data   = sheet.getDataRange().getValues();
    var counts = {};
    MR_QUEUE_AGENTS.forEach(function(a){ counts[a] = 0; });

    for (var i = 1; i < data.length; i++) {
      if (i === newRow - 1) continue;           // skip the new row itself
      var ownerRaw = String(data[i][MR_COL_OWNER] || '').trim().toLowerCase();
      var fullName = MR_NAME_MAP[ownerRaw];
      if (!fullName || counts[fullName] === undefined) continue;

      // Use AA (Assigned date) if available, else submission timestamp
      var dateRaw = data[i][MR_COL_ASSIGNED_AT] || data[i][MR_COL_TIMESTAMP];
      if (!dateRaw) continue;
      var d = (dateRaw instanceof Date) ? dateRaw : new Date(dateRaw);
      if (isNaN(d.getTime())) continue;
      if (Utilities.formatDate(d, tz, 'yyyy-MM-dd') === today) counts[fullName]++;
    }

    // Pick agent with fewest today (stable sort: same count → first in list order)
    var chosen = MR_QUEUE_AGENTS.reduce(function(best, a) {
      return (counts[a] < counts[best]) ? a : best;
    });
    var shortName = MR_AGENT_SHORT[chosen] || chosen.split(' ')[0];

    // Write owner + assignment timestamp into the new row
    var now = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm:ss');
    sheet.getRange(newRow, MR_COL_OWNER + 1).setValue(shortName);
    sheet.getRange(newRow, MR_COL_ASSIGNED_AT + 1).setValue(now);

    // Build summary of today's counts after this assignment
    counts[chosen]++;
    var summary = MR_QUEUE_AGENTS.map(function(a) {
      return (MR_AGENT_SHORT[a] || a.split(' ')[0]) + ': ' + counts[a];
    }).join(' · ');

    var chain   = String(data[newRow-1][MR_COL_CHAIN] || '').trim() || '—';
    var reqType = String(data[newRow-1][MR_COL_TYPE]  || '').trim() || '—';

    // Send WA notification
    var msg = '📋 New Onground Request\n'
      + '👤 Assigned to: *' + chosen.split(' ')[0] + '* (' + shortName + ')\n'
      + '🏪 Chain: ' + chain + '\n'
      + '📝 Type: ' + reqType + '\n'
      + '📊 Today\'s assignments — ' + summary;
    _sendWaProgressUpdate_internal(msg);

  } catch(err) {
    Logger.log('mrAutoAssign error: ' + err);
  }
}

// Internal helper: send WA using stored config (no sheet data needed)
function _sendWaProgressUpdate_internal(text) {
  try {
    var cfg = getWaConfig();
    if (!cfg || !cfg.apikey || !cfg.phone) return;
    _sendWa(text, cfg.phone, cfg.apikey);
  } catch(e) {}
}

// ── EDIT TRIGGER (fires on every edit to the Main request sheet) ───────────
function mrOnOwnerEdit(e) {
  try {
    var sheet = e.range.getSheet();
    if (sheet.getName() !== MR_TAB_NAME) return;

    // Only act when the Owner column (col 22 = 1-indexed) is edited
    var col = e.range.getColumn();
    if (col !== MR_COL_OWNER + 1) return;         // +1 because getColumn() is 1-indexed

    var row      = e.range.getRow();
    if (row < 2) return;                           // skip header

    var newOwner = String(e.value || '').trim();
    if (!newOwner) return;                         // owner was cleared — don't stamp

    var tz    = Session.getScriptTimeZone();
    var now   = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm:ss');

    // Write assignment timestamp to col AB (MR_COL_ASSIGNED_AT + 1 = 28, 1-indexed)
    sheet.getRange(row, MR_COL_ASSIGNED_AT + 1).setValue(now);
  } catch(err) {
    // Silently swallow — don't break the user's editing experience
  }
}

// ── BACK-FILL (run once to populate existing rows that have an Email Date) ─
// For rows where Assigned At is blank but Last Email Sent Date has a value,
// copy the Email Date into the Assigned At column as a best-effort backfill.
function _backfillMrAssignedAt() {
  var ss    = SpreadsheetApp.openById(MR_SHEET_ID);
  var sheet = ss.getSheetByName(MR_TAB_NAME);
  if (!sheet) { Logger.log('Sheet not found'); return; }
  var data  = sheet.getDataRange().getValues();
  var tz    = Session.getScriptTimeZone();
  var filled = 0;

  for (var i = 1; i < data.length; i++) {
    var row        = data[i];
    var assignedAt = String(row[MR_COL_ASSIGNED_AT] || '').trim();
    if (assignedAt) continue;                          // already stamped

    var owner = String(row[MR_COL_OWNER] || '').trim();
    if (!owner) continue;                              // unassigned row — skip

    // Try Last Email Sent Date (col Z, index 25) first
    var emailDate = row[25];
    if (emailDate) {
      var d = (emailDate instanceof Date) ? emailDate : new Date(emailDate);
      if (!isNaN(d.getTime())) {
        var iso = Utilities.formatDate(d, tz, 'yyyy-MM-dd');
        sheet.getRange(i + 1, MR_COL_ASSIGNED_AT + 1).setValue(iso);
        filled++;
        continue;
      }
    }

    // Fallback: use submission Timestamp date
    var ts = row[MR_COL_TIMESTAMP];
    if (ts) {
      var d2 = (ts instanceof Date) ? ts : new Date(ts);
      if (!isNaN(d2.getTime())) {
        var iso2 = Utilities.formatDate(d2, tz, 'yyyy-MM-dd');
        sheet.getRange(i + 1, MR_COL_ASSIGNED_AT + 1).setValue(iso2);
        filled++;
      }
    }
  }
  Logger.log('✅ Back-filled ' + filled + ' rows with Assigned At dates.');
}

function getMainRequestData(dateFrom, dateTo) {
  try {
    var ss    = SpreadsheetApp.openById(MR_SHEET_ID);
    var sheet = ss.getSheetByName(MR_TAB_NAME);
    if (!sheet) return { error: 'Tab "Main request" not found' };

    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return { rows: [], byAgent: {}, totals: {} };

    var chainIdx = 3;  // col D – Chain Name (fallback)
    var hdr = data[0];
    for (var hi = 0; hi < hdr.length; hi++) {
      if (String(hdr[hi]).toLowerCase().indexOf('chain name') >= 0) { chainIdx = hi; break; }
    }

    var tz   = Session.getScriptTimeZone();
    var rows = [];

    for (var i = 1; i < data.length; i++) {
      var row      = data[i];
      var owner    = String(row[MR_COL_OWNER] || '').trim();
      if (!owner) continue;                              // unassigned — skip entirely

      // ── Assignment date: AA (ASSIGNED_AT) → submission Timestamp fallback ──
      var assignedRaw = row[MR_COL_ASSIGNED_AT];
      if (!assignedRaw) assignedRaw = row[MR_COL_TIMESTAMP];
      if (!assignedRaw) continue;

      var dateObj = (assignedRaw instanceof Date) ? assignedRaw : new Date(assignedRaw);
      if (isNaN(dateObj.getTime())) continue;
      var dateISO = Utilities.formatDate(dateObj, tz, 'yyyy-MM-dd');

      if (dateFrom && dateISO < dateFrom) continue;
      if (dateTo   && dateISO > dateTo)   continue;

      // ── Action date: AB (ACTION_DATE) — may be empty ──
      var actionRaw = row[MR_COL_ACTION_DATE];
      var actionISO = '';
      if (actionRaw) {
        var aObj = (actionRaw instanceof Date) ? actionRaw : new Date(actionRaw);
        if (!isNaN(aObj.getTime())) actionISO = Utilities.formatDate(aObj, tz, 'yyyy-MM-dd');
      }

      var ownerKey = owner.toLowerCase();
      var ownerFull = MR_NAME_MAP[ownerKey] || owner;
      var status   = String(row[MR_COL_STATUS]||'').trim();
      var chain    = String(row[chainIdx]||'').trim();

      rows.push({ date: dateISO, actionDate: actionISO, owner: ownerFull, status: status, chain: chain });
    }

    // Aggregate by agent
    var TEAM_NAMES = ['Mohamed Gadallah','Seliem Mohamed','Omar Elattar','Mahmoud Amin'];
    var byAgent = {};
    TEAM_NAMES.forEach(function(n){
      byAgent[n] = { done:0, cancelled:0, pendingContent:0, pendingRegional:0, total:0 };
    });
    var totals = { done:0, cancelled:0, pendingContent:0, pendingRegional:0, total:0 };

    rows.forEach(function(r){
      var a = byAgent[r.owner];
      if (!a) { byAgent[r.owner] = { done:0, cancelled:0, pendingContent:0, pendingRegional:0, total:0 }; a = byAgent[r.owner]; }
      var s = r.status.toLowerCase();
      if      (s === 'done')                                                      { a.done++;           totals.done++;           }
      else if (s === 'canceled' || s === 'cancelled' || s === 'rejected')         { a.cancelled++;      totals.cancelled++;      }
      else if (s === 'pending content')                                           { a.pendingContent++; totals.pendingContent++; }
      else if (s === 'pending regional')                                          { a.pendingRegional++;totals.pendingRegional++;}
      a.total++; totals.total++;
    });

    // Day-over-day: group by date (team totals) + per-agent daily counts
    var byDate           = {};  // { '2026-08-20': { done, cancelled, ... , total } }
    var byAgentDate      = {};  // { 'Mohamed Gadallah': { '2026-08-20': assigned count } }
    var byAgentActionDate = {}; // { 'Mohamed Gadallah': { '2026-08-20': action count } }

    rows.forEach(function(r){
      var s = r.status.toLowerCase();
      // Team-level byDate
      if (!byDate[r.date]) byDate[r.date] = { done:0, cancelled:0, pendingContent:0, pendingRegional:0, total:0 };
      if      (s === 'done')                                                    byDate[r.date].done++;
      else if (s === 'canceled' || s === 'cancelled' || s === 'rejected')       byDate[r.date].cancelled++;
      else if (s === 'pending content')                                         byDate[r.date].pendingContent++;
      else if (s === 'pending regional')                                        byDate[r.date].pendingRegional++;
      byDate[r.date].total++;

      // Per-agent assigned per day (from AA)
      if (!byAgentDate[r.owner]) byAgentDate[r.owner] = {};
      if (!byAgentDate[r.owner][r.date]) byAgentDate[r.owner][r.date] = 0;
      byAgentDate[r.owner][r.date]++;

      // Per-agent action per day (from AB)
      if (r.actionDate) {
        if (!byAgentActionDate[r.owner]) byAgentActionDate[r.owner] = {};
        if (!byAgentActionDate[r.owner][r.actionDate]) byAgentActionDate[r.owner][r.actionDate] = 0;
        byAgentActionDate[r.owner][r.actionDate]++;
      }
    });

    return { byAgent: byAgent, totals: totals, byDate: byDate,
             byAgentDate: byAgentDate, byAgentActionDate: byAgentActionDate, rowCount: rows.length };
  } catch(e) {
    return { error: e.toString() };
  }
}
