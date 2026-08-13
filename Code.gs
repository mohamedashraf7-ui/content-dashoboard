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
