// Holliday Welding & Fence — Google Sheets Integration
var SPREADSHEET_ID = ''; // <-- Paste your Google Spreadsheet ID here

var SHEET_LEADS     = 'Website Leads';
var SHEET_JOBS      = 'Jobs';
var SHEET_PIPELINE  = 'Pipeline';
var SHEET_INVENTORY = 'Inventory';
var SHEET_SETTINGS  = 'Settings';

// ── GET: route by ?type= parameter ─────────────────────────────────────────
function doGet(e) {
  var type = (e && e.parameter && e.parameter.type) || 'leads';
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  if (type === 'jobs')      return getSheetRows(ss, SHEET_JOBS, 'rows');
  if (type === 'inventory') return getSheetRows(ss, SHEET_INVENTORY, 'items');
  if (type === 'settings')  return getSettings(ss);
  return getSheetRows(ss, SHEET_LEADS, 'rows');
}

function getSheetRows(ss, sheetName, key) {
  var sheet = getOrCreateSheet(ss, sheetName);
  if (sheet.getLastRow() <= 1) return jsonOut(key === 'items' ? { items: [] } : { rows: [] });
  var rows = sheet.getDataRange().getValues();
  var headers = rows[0];
  var data = rows.slice(1).map(function(row) {
    var obj = {};
    headers.forEach(function(h, i) { obj[h] = row[i]; });
    return obj;
  });
  var out = {};
  out[key] = data;
  return jsonOut(out);
}

function getSettings(ss) {
  var sheet = getOrCreateSheet(ss, SHEET_SETTINGS);
  if (sheet.getLastRow() <= 1) return jsonOut({ settings: {} });
  var rows = sheet.getDataRange().getValues();
  var settings = {};
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][0]) settings[rows[i][0]] = rows[i][1];
  }
  return jsonOut({ settings: settings });
}

// ── POST: write to the appropriate sheet by type ────────────────────────────
function doPost(e) {
  var payload;
  try { payload = JSON.parse(e.postData.contents); } catch(err) { return jsonOut({ ok: false, error: 'bad json' }); }

  var type = (payload.type || 'contact').toLowerCase();
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  if (type === 'contact' || type === 'lead') {
    writeRow(ss, SHEET_LEADS,
      ['timestamp','firstName','lastName','phone','email','address','services','details','source'],
      payload);
  } else if (type === 'job') {
    upsertJob(ss, payload);
  } else if (type === 'pipeline') {
    writeRow(ss, SHEET_PIPELINE,
      ['timestamp','name','phone','address','type','status','bidAmount','startDate','notes'],
      payload);
  } else if (type === 'inventory') {
    upsertInventoryItem(ss, payload);
  } else if (type === 'settings') {
    upsertSetting(ss, payload);
  }

  return jsonOut({ ok: true, type: type });
}

// ── Job upsert: find row by jobNumber and update, or append ────────────────
function upsertJob(ss, data) {
  var headers = ['timestamp','jobNumber','customerName','phone','address','jobType','status',
                 'startDate','endDate','linFt','totalCharge','materialCost','grossProfit',
                 'stainGal','bleachGal','crewLead','stainProduct','notes','stateJson'];
  var sheet = getOrCreateSheet(ss, SHEET_JOBS);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  data.timestamp = data.timestamp || new Date().toISOString();
  if (data.jobNumber && sheet.getLastRow() > 1) {
    var allData = sheet.getDataRange().getValues();
    var jnCol = allData[0].indexOf('jobNumber');
    for (var i = 1; i < allData.length; i++) {
      if (String(allData[i][jnCol]) === String(data.jobNumber)) {
        var updatedRow = headers.map(function(h) {
          return data[h] !== undefined ? String(data[h]) : String(allData[i][headers.indexOf(h)] || '');
        });
        sheet.getRange(i + 1, 1, 1, headers.length).setValues([updatedRow]);
        return;
      }
    }
  }
  var row = headers.map(function(h) { return data[h] !== undefined ? String(data[h]) : ''; });
  sheet.appendRow(row);
}

// ── Settings upsert: find row by key and update, or append ─────────────────
function upsertSetting(ss, data) {
  var sheet = getOrCreateSheet(ss, SHEET_SETTINGS);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['key', 'value', 'updatedAt']);
    sheet.getRange(1, 1, 1, 3).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  var updatedAt = new Date().toISOString();
  if (data.key && sheet.getLastRow() > 1) {
    var rows = sheet.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      if (rows[i][0] === data.key) {
        sheet.getRange(i + 1, 1, 1, 3).setValues([[data.key, data.value, updatedAt]]);
        return;
      }
    }
  }
  sheet.appendRow([data.key, data.value, updatedAt]);
}

// ── Inventory upsert: find row by id and update, or append ─────────────────
function upsertInventoryItem(ss, data) {
  var headers = ['id','name','category','unit','qty','paid','costPerUnit','reorderAt','notes','updatedAt'];
  var sheet = getOrCreateSheet(ss, SHEET_INVENTORY);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
    sheet.getRange(1,1,1,headers.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  data.updatedAt = new Date().toISOString();
  if (data.qty !== undefined && data.paid !== undefined) {
    var q = parseFloat(data.qty) || 0;
    var p = parseFloat(data.paid) || 0;
    data.costPerUnit = (q > 0 && p > 0) ? String((p / q).toFixed(2)) : '0';
  }
  if (data.id && sheet.getLastRow() > 1) {
    var allData = sheet.getDataRange().getValues();
    var idCol = allData[0].indexOf('id');
    for (var i = 1; i < allData.length; i++) {
      if (String(allData[i][idCol]) === String(data.id)) {
        var updatedRow = headers.map(function(h) {
          return data[h] !== undefined ? String(data[h]) : String(allData[i][headers.indexOf(h)] || '');
        });
        sheet.getRange(i + 1, 1, 1, headers.length).setValues([updatedRow]);
        return;
      }
    }
  }
  var row = headers.map(function(h) { return data[h] !== undefined ? String(data[h]) : ''; });
  sheet.appendRow(row);
}

// ── helpers ─────────────────────────────────────────────────────────────────
function getOrCreateSheet(ss, name) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
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
