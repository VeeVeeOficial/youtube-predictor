// ===== FrameLedger backend (Google Apps Script) =====
// Bind this script to a Google Sheet: create a new Sheet, then
// Extensions > Apps Script, paste this file in as Code.gs, then
// Deploy > New deployment > Type: Web app
//   Execute as: Me
//   Who has access: Anyone
// Copy the resulting /exec URL and send it back so the frontend can be wired up.

var SPREADSHEET_ID = '1-7LvGF-YykwFllEjyj1HscK2cxQ-Gr6VCcQoEsknscg'; // FrameLedger Data sheet

// Slips get routed to different Drive folders depending on what they
// document, chosen by the frontend via the 'folderKey' upload param.
var FOLDER_MAP = {
  general: '1fACSjyzIjhUUEh-8cS7sDwThv34hPliq',  // camera/lens purchase evidence (default)
  income: '19DVWk9e3aDApM0g2qg2lPeWHRHjv7XKe',   // payment slips from customers (sales, invoices)
  shipping: '170mJWu_8hZD7j7fJ5kPgnnksnkPfanQ4',  // shipping cost slips (ค่าจัดส่ง)
  parts: '1QHJM-vfiWCLj-boR9Ldm7FYHUuUmQVBM'      // parts/software purchase slips (อะไหล่)
};

var COLLECTION_SHEETS = {
  transactions: 'Transactions',
  inventory: 'Inventory',
  invoices: 'Invoices',
  receipts: 'Receipts'
};

var SCHEMAS = {
  Transactions: ['id', 'type', 'vendor', 'amount', 'date', 'category', 'note', 'linkedItemId'],
  Inventory: ['id', 'name', 'sn', 'productCode', 'condition', 'supplier', 'purchaseCost', 'dateIn', 'status', 'salePrice', 'saleDate', 'customer', 'saleSlip', 'evidencePhoto'],
  Invoices: ['id', 'invNo', 'itemName', 'sn', 'productCode', 'customer', 'price', 'shipping', 'date', 'cost', 'profit', 'incomeTxId', 'shippingTxId'],
  Receipts: ['id', 'recNo', 'desc', 'amount', 'shipping', 'total', 'date']
};

function getSheet_(name) {
  // openById (matching the older, proven-working Camera Flip Backend) instead
  // of getActiveSpreadsheet() — the latter is meant for scripts running
  // interactively inside the Sheets UI and can behave unreliably when the
  // script is only ever invoked over HTTP as a Web App, which is our case.
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  return sh;
}

function ensureHeaders_(sheet, headers) {
  // Self-healing: also fixes up a sheet that already has rows but is
  // missing newly-added columns (e.g. after a schema change), not just
  // a brand-new empty sheet — otherwise new columns never get a header
  // label on sheets that already had data before the schema changed.
  var lastCol = sheet.getLastColumn();
  var existing = lastCol > 0 ? sheet.getRange(1, 1, 1, lastCol).getValues()[0] : [];
  var matches = existing.length >= headers.length;
  if (matches) {
    for (var i = 0; i < headers.length; i++) {
      if (existing[i] !== headers[i]) { matches = false; break; }
    }
  }
  if (!matches) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
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
    profile: profile,
    claudeApiKey: flat.claudeApiKey || '',
    claudeModel: flat.claudeModel || 'claude-sonnet-5'
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
    ['profile', JSON.stringify(settingsObj.profile || {})],
    ['claudeApiKey', settingsObj.claudeApiKey || ''],
    ['claudeModel', settingsObj.claudeModel || 'claude-sonnet-5']
  ];
  sheet.getRange(2, 1, rows.length, 2).setValues(rows);
}

function jsonResponse_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// GET responses from an Apps Script Web App are served through a redirect to
// script.googleusercontent.com for caching, and that redirected response
// doesn't carry CORS headers a cross-origin fetch() can read — confirmed by
// opening the /exec URL directly (works, shows correct JSON) vs. the site's
// own fetch() call (always fails). A <script> tag isn't subject to CORS at
// all, so when a 'callback' param is present, wrap the JSON as a JS function
// call the frontend can load via a script tag instead of fetch().
function respond_(obj, callbackName) {
  if (callbackName) {
    return ContentService
      .createTextOutput(callbackName + '(' + JSON.stringify(obj) + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return jsonResponse_(obj);
}

function doGet(e) {
  var callbackName = e && e.parameter && e.parameter.callback;
  try {
    var action = (e.parameter && e.parameter.action) || 'loadAll';
    if (action === 'loadAll') {
      return respond_({
        ok: true,
        transactions: readCollection_('Transactions'),
        inventory: readCollection_('Inventory'),
        invoices: readCollection_('Invoices'),
        receipts: readCollection_('Receipts'),
        settings: readSettings_()
      }, callbackName);
    }
    return respond_({ ok: false, error: 'unknown action: ' + action }, callbackName);
  } catch (err) {
    return respond_({ ok: false, error: String(err) }, callbackName);
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
      var folderId = FOLDER_MAP[payload.folderKey] || FOLDER_MAP.general;
      var folder = DriveApp.getFolderById(folderId);
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
