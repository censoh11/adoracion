const SHEET_NAME = 'Historico';
const HEADERS = [
  'eventId',
  'eventCode',
  'eventName',
  'eventType',
  'eventDate',
  'startTime',
  'manager',
  'collaboratorCode',
  'collaboratorName',
  'eventRole',
  'collaboratorRoles',
  'listedStatus',
  'checked',
  'onTime',
  'punctualityStatus',
  'averageScore',
  'arrival',
  'comment',
  'evaluationComment',
  'manualArrival',
  'arrivalChangeLog',
  'checkedAt',
  'evaluatedAt',
  'scores',
  'syncedAt',
  'lateTolerance'
];

function doPost(e) {
  const payload = JSON.parse((e.parameter && e.parameter.payload) || '{"rows":[]}');
  const rows = Array.isArray(payload.rows) ? payload.rows : [];
  const sheet = getHistorySheet_();
  const now = new Date().toISOString();
  const values = rows.map(row => HEADERS.map(header => header === 'syncedAt' ? now : (row[header] || '')));
  const existing = getExistingRowMap_(sheet);
  const appendValues = [];
  values.forEach(valueRow => {
    const rowObject = rowToObject_(valueRow);
    const key = historyKey_(rowObject);
    if (existing[key]) {
      sheet.getRange(existing[key], 1, 1, HEADERS.length).setValues([valueRow]);
    } else {
      appendValues.push(valueRow);
    }
  });
  if (appendValues.length) sheet.getRange(sheet.getLastRow() + 1, 1, appendValues.length, HEADERS.length).setValues(appendValues);
  return ContentService
    .createTextOutput(JSON.stringify({ ok: true, received: values.length, inserted: appendValues.length }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  const callback = (e.parameter && e.parameter.callback) || 'callback';
  const rows = readRows_();
  const output = `${callback}(${JSON.stringify({ ok: true, rows })});`;
  return ContentService
    .createTextOutput(output)
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

function getHistorySheet_() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = spreadsheet.insertSheet(SHEET_NAME);
  const firstRow = sheet.getRange(1, 1, 1, HEADERS.length).getValues()[0];
  const needsHeaders = HEADERS.some((header, index) => firstRow[index] !== header);
  if (needsHeaders) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function readRows_() {
  const sheet = getHistorySheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const values = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();
  return values.map(rowToObject_);
}

function getExistingRowMap_(sheet) {
  const rows = readRows_();
  const map = {};
  rows.forEach((row, index) => {
    map[historyKey_(row)] = index + 2;
  });
  return map;
}

function rowToObject_(row) {
  const item = {};
  HEADERS.forEach((header, index) => {
    const value = row[index];
    item[header] = value instanceof Date ? value.toISOString() : value;
  });
  return item;
}

function historyKey_(row) {
  return `${row.eventId || ''}|${row.collaboratorCode || row.collaboratorName || ''}`;
}
