/**
 * Cash Tray Manager — Apps Script backend.
 * See CONTRACT.md for the full API contract.
 *
 * Deploy as a Web App (Execute as: Me). The token is a shared secret checked on every
 * request. For stronger auth, remove the token check and rely on the web-app
 * access setting + Session.getEffectiveUser().
 */

var EXPECTED_TOKEN = 'MY_SECRET_2025';   // must match SECRET_TOKEN in index.html
var CONFIG_SHEET_ID = '1p3RJHBlNoX2agEIYkFScXZhOMh8NrZjQnmT9fK8KcHs'; // default sheet
var TAB_NAME = 'Data';

var COLUMNS = [
  'id', 'date', 'bankIndentNo', 'bankIndentVal', 'atmIds',
  'totalNotesAll', 'totalValueAll', 'totalLoadedVal', 'trays', 'summary', 'updatedAt'
];

function doGet(e) {
  var token = e.parameter.token;
  if (token !== EXPECTED_TOKEN) {
    return jsonResponse({ status: 'error', message: 'Unauthorized' }, 401);
  }
  try {
    switch (e.parameter.action) {
      case 'getAll': return jsonResponse({ status: 'ok', data: getAll(e.parameter.sheetId) });
      case 'delete': return deleteRow({ id: e.parameter.id, sheetId: e.parameter.sheetId });
      default:       return jsonResponse({ status: 'error', message: 'Unknown action: ' + e.parameter.action }, 400);
    }
  } catch (err) {
    Logger.log('doGet error: %s', err);
    return jsonResponse({ status: 'error', message: 'Server error: ' + err.message }, 500);
  }
}

function doPost(e) {
  var body = {};
  try { body = JSON.parse(e.postData.contents); } catch (err) {
    return jsonResponse({ status: 'error', message: 'Invalid JSON body' }, 400);
  }
  if (body._token !== EXPECTED_TOKEN) {
    return jsonResponse({ status: 'error', message: 'Unauthorized' }, 401);
  }
  try {
    switch (body.action) {
      case 'getAll':  return jsonResponse({ status: 'ok', data: getAll(body.sheetId) });
      case 'save':    return saveRow(body, false);
      case 'update':  return saveRow(body, true);
      case 'delete':  return deleteRow(body);
      default:        return jsonResponse({ status: 'error', message: 'Unknown action: ' + body.action }, 400);
    }
  } catch (err) {
    Logger.log('Handler error: %s', err);
    return jsonResponse({ status: 'error', message: 'Server error: ' + err.message }, 500);
  }
}

function getSheet(id) {
  var ss = (id && id !== CONFIG_SHEET_ID)
    ? SpreadsheetApp.openById(id)
    : SpreadsheetApp.openById(CONFIG_SHEET_ID);
  var tab = ss.getSheetByName(TAB_NAME);
  if (!tab) {
    tab = ss.insertSheet(TAB_NAME);
    tab.appendRow(COLUMNS);
    tab.setFrozenRows(1);
  }
  return tab;
}

function getAll(id) {
  var sheet = getSheet(id);
  var values = sheet.getDataRange().getValues();
  var out = [];
  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    var obj = {};
    for (var c = 0; c < COLUMNS.length; c++) obj[COLUMNS[c]] = row[c];
    if (!obj.id) continue;
    obj.atmIds = parseJson(obj.atmIds, []);
    obj.trays = parseJson(obj.trays, {});
    obj.summary = parseJson(obj.summary, {});
    out.push(obj);
  }
  return out;
}

function saveRow(body, isUpdate) {
  var sheet = getSheet(body.sheetId);
  var id = String(body.id || '');
  if (!id) return jsonResponse({ status: 'error', message: 'Missing id' }, 400);
  if (!body.date || !body.bankIndentNo) {
    return jsonResponse({ status: 'error', message: 'date and bankIndentNo are required' }, 400);
  }
  if (!Array.isArray(body.atmIds) || body.atmIds.length === 0) {
    return jsonResponse({ status: 'error', message: 'At least one ATM is required' }, 400);
  }
  var row = buildRow(body);
  var data = sheet.getDataRange().getValues();
  var rowIdx = -1;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === id) { rowIdx = i + 1; break; }
  }
  if (isUpdate && rowIdx === -1) rowIdx = data.length + 1; // upsert if not found
  if (rowIdx === -1) {
    sheet.appendRow(row);
  } else {
    sheet.getRange(rowIdx, 1, 1, row.length).setValues([row]);
  }
  return jsonResponse({ status: 'ok' });
}

function deleteRow(body) {
  var sheet = getSheet(body.sheetId);
  var id = String(body.id || '');
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === id) {
      sheet.deleteRow(i + 1);
      return jsonResponse({ status: 'ok' });
    }
  }
  return jsonResponse({ status: 'error', message: 'Row not found' }, 404);
}

function buildRow(body) {
  return [
    String(body.id),
    body.date,
    String(body.bankIndentNo),
    Number(body.bankIndentVal) || 0,
    JSON.stringify(body.atmIds || []),
    Number(body.totalNotesAll) || 0,
    Number(body.totalValueAll) || 0,
    Number(body.totalLoadedVal) || 0,
    JSON.stringify(body.trays || {}),
    JSON.stringify(body.summary || {}),
    new Date()
  ];
}

function parseJson(str, fallback) {
  if (!str) return fallback;
  if (typeof str === 'object') return str;
  try { return JSON.parse(str); } catch (e) { return fallback; }
}

function jsonResponse(obj, code) {
  var payload = JSON.stringify(obj);
  return ContentService.createTextOutput(payload).setMimeType(ContentService.MimeType.JSON);
}
