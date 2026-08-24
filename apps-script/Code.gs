// ===== FrameLedger backend (Google Apps Script) =====
// Bind this script to a Google Sheet: create a new Sheet, then
// Extensions > Apps Script, paste this file in as Code.gs, then
// Deploy > New deployment > Type: Web app
//   Execute as: Me
//   Who has access: Anyone
// Copy the resulting /exec URL and send it back so the frontend can be wired up.

var FOLDER_ID = '1fACSjyzIjhUUEh-8cS7sDwThv34hPliq'; // Drive folder for photos

var COLLECTION_SHEETS = {
  transactions: 'Transactions',
  inventory: 'Inventory',
  invoices: 'Invoices',
  receipts: 'Receipts'
};

var SCHEMAS = {
  Transactions: ['id', 'type', 'vendor', 'amount', 'date', 'category', 'note', 'linkedItemId'],
  Inventory: ['id', 'name', 'sn', 'productCode', 'condition', 'supplier', 'purchaseCost', 'dateIn', 'status', 'salePrice', 'saleDate', 'customer', 'saleSlip', 'evidencePhoto'],
  Invoices: ['id', 'invNo', 'itemName', 'sn', 'productCode', 'customer', 'price', 'date', 'cost', 'profit'],
  Receipts: ['id', 'recNo', 'desc', 'amount', 'shipping', 'total', 'date']
};

function getSheet_(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  return sh;
}

function ensureHeaders_(sheet, headers) {
  var range = sheet.getRange(1, 1, 1, headers.length);
  var existing = range.getValues()[0];
  if (existing.join('') === '') {
    range.setValues([headers]);
  }
}

function readCollection_(sheetName) {
  var headers = SCHEMAS[sheetName];
  var sheet = getSheet_(sheetName);
  ensureHeaders_(sheet, headers);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var values = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
  var out = [];
  for (var r = 0; r < values.length; r++) {
    if (values[r][0] === '') continue;
    var obj = {};
    for (var c = 0; c < headers.length; c++) obj[headers[c]] = values[r][c];
    out.push(obj);
  }
  return out;
}

function writeCollection_(sheetName, items) {
  var headers = SCHEMAS[sheetName];
  var sheet = getSheet_(sheetName);
  ensureHeaders_(sheet, headers);
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, headers.length).clearContent();
  if (!items || !items.length) return;
  var rows = items.map(function (it) {
    return headers.map(function (h) {
      var v = it[h];
      return (v === undefined || v === null) ? '' : v;
    });
  });
  sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
}

function readSettings_() {
  var sheet = getSheet_('Settings');
  ensureHeaders_(sheet, ['key', 'value']);
  var lastRow = sheet.getLastRow();
  var flat = {};
  if (lastRow >= 2) {
    var values = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
    for (var i = 0; i < values.length; i++) {
      if (values[i][0]) flat[values[i][0]] = values[i][1];
    }
  }
  var profile = {};
  try { profile = JSON.parse(flat.profile || '{}'); } catch (err) {}
  return {
    taxRate: Number(flat.taxRate || 5),
    costMode: flat.costMode || 'detailed',
    profile: profile
  };
}

function writeSettings_(settingsObj) {
  var sheet = getSheet_('Settings');
  ensureHeaders_(sheet, ['key', 'value']);
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, 2).clearContent();
  var rows = [
    ['taxRate', settingsObj.taxRate || 5],
    ['costMode', settingsObj.costMode || 'detailed'],
    ['profile', JSON.stringify(settingsObj.profile || {})]
  ];
  sheet.getRange(2, 1, rows.length, 2).setValues(rows);
}

function jsonResponse_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  try {
    var action = (e.parameter && e.parameter.action) || 'loadAll';
    if (action === 'loadAll') {
      return jsonResponse_({
        ok: true,
        transactions: readCollection_('Transactions'),
        inventory: readCollection_('Inventory'),
        invoices: readCollection_('Invoices'),
        receipts: readCollection_('Receipts'),
        settings: readSettings_()
      });
    }
    return jsonResponse_({ ok: false, error: 'unknown action: ' + action });
  } catch (err) {
    return jsonResponse_({ ok: false, error: String(err) });
  }
}

// Diagnostic helper: select this function in the dropdown next to "Run" in the
// Apps Script editor, click Run, then View > Logs (or Execution log) to see
// exactly what loadAll returns, without needing the deployed URL or a browser.
function testLoadAll() {
  var result = doGet({ parameter: { action: 'loadAll' } }).getContent();
  Logger.log(result);
}

function doPost(e) {
  try {
    // Same wire format as the older working Camera Flip Backend: action in
    // the query string, a single 'payload' form field carrying the JSON.
    // This avoids relying on a raw JSON request body, which is more prone
    // to being mangled/blocked between a static GitHub Pages site and a
    // script.google.com Web App than a plain form POST is.
    var action = e.parameter.action;
    var payload = {};
    if (e.parameter.payload) {
      payload = JSON.parse(e.parameter.payload);
    }

    if (action === 'save') {
      var key = payload.key;
      if (key === 'settings') {
        writeSettings_(payload.data || {});
      } else if (COLLECTION_SHEETS[key]) {
        writeCollection_(COLLECTION_SHEETS[key], payload.data || []);
      } else {
        return jsonResponse_({ ok: false, error: 'unknown key: ' + key });
      }
      return jsonResponse_({ ok: true });
    }

    if (action === 'uploadImage') {
      var folder = DriveApp.getFolderById(FOLDER_ID);
      var bytes = Utilities.base64Decode(payload.base64);
      var blob = Utilities.newBlob(bytes, payload.mimeType || 'image/jpeg', payload.filename || ('img-' + Date.now() + '.jpg'));
      var file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      var url = 'https://drive.google.com/uc?export=view&id=' + file.getId();
      return jsonResponse_({ ok: true, url: url, fileId: file.getId() });
    }

    return jsonResponse_({ ok: false, error: 'unknown action: ' + action });
  } catch (err) {
    return jsonResponse_({ ok: false, error: String(err) });
  }
}
