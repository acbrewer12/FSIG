// Holliday Welding & Fence — Google Sheets Integration
// Paste this entire file into Google Apps Script (script.google.com)
// Then: Deploy → New deployment → Web App → Execute as "Me" → Anyone can access

var SPREADSHEET_ID = ''; // <-- Paste your Google Spreadsheet ID here

var SHEET_LEADS    = 'Website Leads';
var SHEET_JOBS     = 'Jobs';
var SHEET_PIPELINE = 'Pipeline';

// ── GET: return website leads for FSIG app polling ─────────────────────────
function doGet(e) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = getOrCreateSheet(ss, SHEET_LEADS);
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return jsonOut({ rows: [] });
  var rows = sheet.getDataRange().getValues();
  var headers = rows[0];
  var data = rows.slice(1).map(function(row) {
    var obj = {};
    headers.forEach(function(h, i) { obj[h] = row[i]; });
    return obj;
  });
  return jsonOut({ rows: data });
}

// ── POST: write to the appropriate sheet by type ────────────────────────────
function doPost(e) {
  var payload;
  try {
    payload = JSON.parse(e.postData.contents);
  } catch(err) {
    return jsonOut({ ok: false, error: 'bad json' });
  }

  var type = (payload.type || 'contact').toLowerCase();
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  if (type === 'contact' || type === 'lead') {
    writeRow(ss, SHEET_LEADS,
      ['timestamp','firstName','lastName','phone','email','address','services','details','source'],
      payload);
  } else if (type === 'job') {
    writeRow(ss, SHEET_JOBS,
      ['timestamp','jobNumber','customerName','phone','address','jobType','status',
       'startDate','endDate','linFt','totalCharge','materialCost','grossProfit',
       'stainGal','bleachGal','crewLead','stainProduct','notes'],
      payload);
  } else if (type === 'pipeline') {
    writeRow(ss, SHEET_PIPELINE,
      ['timestamp','name','phone','address','type','status','bidAmount','startDate','notes'],
      payload);
  }

  return jsonOut({ ok: true, type: type });
}

// ── helpers ─────────────────────────────────────────────────────────────────
function getOrCreateSheet(ss, name) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  return sheet;
}

function writeRow(ss, sheetName, headers, data) {
  var sheet = getOrCreateSheet(ss, sheetName);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  data.timestamp = data.timestamp || new Date().toISOString();
  var row = headers.map(function(h) { return data[h] !== undefined ? String(data[h]) : ''; });
  sheet.appendRow(row);
}

function jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
