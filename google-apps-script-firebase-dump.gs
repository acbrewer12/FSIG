// FSIG — Firebase to Google Sheets 6-hour dump
// Paste this into Google Apps Script (script.google.com)
// Set a time-based trigger: syncFirebaseToSheets → every 6 hours
// No web app deployment needed.

var FIREBASE_URL = 'https://fsig-5f7e9-default-rtdb.firebaseio.com';
var SPREADSHEET_ID = ''; // <-- Paste your new Google Spreadsheet ID here

function syncFirebaseToSheets() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  // ── Jobs ──────────────────────────────────────────────────────────────────
  try {
    var jobsRes = UrlFetchApp.fetch(FIREBASE_URL + '/fsig/jobs.json');
    var jobs = JSON.parse(jobsRes.getContentText()) || {};
    var jobSheet = getOrCreate(ss, 'Jobs');
    var jobHeaders = ['Job #','Customer','Phone','Address','Status','Start Date','End Date',
                      'Crew Lead','Job Type','Total Charge','Stain Gal','Bleach Gal','Notes','Updated'];
    setHeaders(jobSheet, jobHeaders);
    var jobRows = Object.values(jobs).map(function(j) {
      var s = j.state || {};
      var t = s.totals || {};
      return [
        s.jobNumber||'', s.customerName||'', s.customerPhone||'',
        [s.addr,s.city,s.state,s.zip].filter(Boolean).join(', '),
        s.jobStatus||'', s.jobStartDate||'', s.jobEndDate||'',
        s.crewLead||'', j.jobType||'',
        t.totalCharge||'', t.stainGal||'', t.bleachGal||'',
        s.changeOrder||'', new Date(j.updatedAt||j.id||0).toLocaleDateString()
      ];
    });
    writeRows(jobSheet, jobRows);
  } catch(e) { Logger.log('Jobs sync error: ' + e); }

  // ── Pipeline ──────────────────────────────────────────────────────────────
  try {
    var pipeRes = UrlFetchApp.fetch(FIREBASE_URL + '/fsig/pipeline.json');
    var pipeline = JSON.parse(pipeRes.getContentText()) || {};
    var pipeSheet = getOrCreate(ss, 'Pipeline');
    var pipeHeaders = ['Name','Phone','Address','Type','Status','Bid Amount','Start Date','Notes','Created'];
    setHeaders(pipeSheet, pipeHeaders);
    var pipeRows = Object.values(pipeline).map(function(j) {
      return [
        j.name||'', j.phone||'', j.address||'',
        j.type||'', j.status||'', j.bidAmount||'',
        j.startDate||'', j.notes||'',
        j.createdAt ? new Date(j.createdAt).toLocaleDateString() : ''
      ];
    });
    writeRows(pipeSheet, pipeRows);
  } catch(e) { Logger.log('Pipeline sync error: ' + e); }

  // ── Inventory ─────────────────────────────────────────────────────────────
  try {
    var invRes = UrlFetchApp.fetch(FIREBASE_URL + '/fsig/inventory.json');
    var inventory = JSON.parse(invRes.getContentText()) || {};
    var invSheet = getOrCreate(ss, 'Inventory');
    var invHeaders = ['Name','Category','Unit','Qty','Paid','Cost/Unit','Reorder At','Notes','Updated'];
    setHeaders(invSheet, invHeaders);
    var invRows = Object.values(inventory).map(function(i) {
      var cpu = (parseFloat(i.qty)>0 && parseFloat(i.paid)>0) ? (parseFloat(i.paid)/parseFloat(i.qty)).toFixed(2) : '0';
      return [i.name||'',i.category||'',i.unit||'',i.qty||0,i.paid||0,cpu,i.reorderAt||0,i.notes||'',i.updatedAt||''];
    });
    writeRows(invSheet, invRows);
  } catch(e) { Logger.log('Inventory sync error: ' + e); }

  Logger.log('Firebase → Sheets sync complete: ' + new Date().toISOString());
}

function getOrCreate(ss, name) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  return sheet;
}

function setHeaders(sheet, headers) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
    sheet.getRange(1,1,1,headers.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
}

function writeRows(sheet, rows) {
  // Clear existing data rows (keep header)
  if (sheet.getLastRow() > 1) sheet.getRange(2, 1, sheet.getLastRow()-1, sheet.getLastColumn()).clearContent();
  if (rows.length) sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
}
