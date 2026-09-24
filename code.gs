const SHEET_NAME = 'Data SIPATERA';
const DATA_HEADER = ['id', 'nomor', 'pemohon', 'unit', 'aset', 'keperluan', 'tglMulai', 'tglSelesai', 'dokumen', 'status', 'catatan', 'jadwalDitetapkan', 'buktiBayar', 'riwayat'];

function getDataSheet_() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = spreadsheet.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) sheet.appendRow(DATA_HEADER);
  return sheet;
}

function doGet() {
  return jsonResponse_({ok: true, data: readData_()});
}

function doPost(event) {
  try {
    const payload = JSON.parse(event.postData.contents || '{}');
    if (payload.action !== 'sync' || !Array.isArray(payload.data)) {
      return jsonResponse_({ok: false, error: 'Payload sync tidak valid.'});
    }
    writeData_(payload.data);
    return jsonResponse_({ok: true, count: payload.data.length});
  } catch (error) {
    return jsonResponse_({ok: false, error: String(error)});
  }
}

function readData_() {
  const sheet = getDataSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  return sheet.getRange(2, 1, lastRow - 1, DATA_HEADER.length).getValues().map(row => ({
    id: row[0], nomor: row[1], pemohon: row[2], unit: row[3], aset: row[4],
    keperluan: row[5], tglMulai: row[6], tglSelesai: row[7],
    dokumen: parseJson_(row[8], []), status: row[9], catatan: row[10],
    jadwalDitetapkan: row[11], buktiBayar: row[12], riwayat: parseJson_(row[13], [])
  }));
}

function writeData_(data) {
  const sheet = getDataSheet_();
  const rows = data.map(item => [
    item.id || '', item.nomor || '', item.pemohon || '', item.unit || '', item.aset || '',
    item.keperluan || '', item.tglMulai || '', item.tglSelesai || '',
    JSON.stringify(item.dokumen || []), item.status || '', item.catatan || '',
    item.jadwalDitetapkan || '', item.buktiBayar || '', JSON.stringify(item.riwayat || [])
  ]);
  const rowsToClear = Math.max(sheet.getLastRow() - 1, 1);
  sheet.getRange(2, 1, rowsToClear, DATA_HEADER.length).clearContent();
  if (rows.length) sheet.getRange(2, 1, rows.length, DATA_HEADER.length).setValues(rows);
  sheet.setFrozenRows(1);
}

function parseJson_(value, fallback) {
  try { return value ? JSON.parse(value) : fallback; } catch (error) { return fallback; }
}

function jsonResponse_(body) {
  return ContentService.createTextOutput(JSON.stringify(body))
    .setMimeType(ContentService.MimeType.JSON);
}