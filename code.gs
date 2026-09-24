/**
 * SIPATERA — Apps Script backend
 * -------------------------------
 * Menghubungkan aplikasi SIPATERA (HTML) dengan Google Sheet + Google Drive
 * sehingga data permohonan dapat dilihat/diedit langsung dari Sheet, dan
 * berkas yang diupload pemohon (KTP, Surat Tugas, dll) otomatis tersimpan
 * ke folder Drive lalu tautannya dicatat di kolom "dokumen" pada Sheet.
 *
 * CARA PASANG:
 * 1. Buka https://script.google.com -> New Project.
 * 2. Hapus isi Code.gs bawaan, tempel seluruh isi file ini.
 * 3. Jalankan fungsi `setup` sekali (menu Run > setup) untuk membuat
 *    Google Sheet baru bernama "SIPATERA-DATA" dengan header kolom.
 *    (Boleh juga pakai Spreadsheet yang sudah ada — lihat catatan di setup()).
 * 4. Deploy -> New deployment -> Type: Web app.
 *      - Execute as: Me
 *      - Who has access: Anyone
 * 5. Salin URL Web App yang muncul (diakhiri /exec), tempel di tab
 *    "Monitoring & Laporan" -> "Hubungkan ke Google Sheet" pada SIPATERA.
 */

const SHEET_NAME = 'Data';
const DRIVE_FOLDER_NAME = 'SIPATERA-Berkas';
const HEADERS = [
  'id','nomor','pemohon','unit','aset','keperluan','tglMulai','tglSelesai',
  'dokumen','status','catatan','jadwalDitetapkan','buktiBayar','riwayat','updatedAt'
];

function setup() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    ss = SpreadsheetApp.create('SIPATERA-DATA');
  }
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  sheet.clear();
  sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  sheet.setFrozenRows(1);

  // Siapkan folder Drive untuk menyimpan berkas yang diupload pemohon
  const folder = getFolder_();

  Logger.log('Sheet siap: ' + ss.getUrl());
  Logger.log('Folder berkas siap: ' + folder.getUrl());
}

/** Folder Drive tempat berkas upload disimpan (dibuat sekali, lalu diingat lewat Script Properties). */
function getFolder_() {
  const props = PropertiesService.getScriptProperties();
  const folderId = props.getProperty('DRIVE_FOLDER_ID');
  if (folderId) {
    try { return DriveApp.getFolderById(folderId); } catch (e) { /* folder dihapus, buat baru */ }
  }
  const folder = DriveApp.createFolder(DRIVE_FOLDER_NAME);
  props.setProperty('DRIVE_FOLDER_ID', folder.getId());
  return folder;
}

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.openById(
    PropertiesService.getScriptProperties().getProperty('SHEET_ID')
  );
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  }
  return sheet;
}

function rowToItem_(row) {
  const item = {};
  HEADERS.forEach((h, i) => {
    let v = row[i];
    if (h === 'dokumen' || h === 'riwayat') {
      try { v = v ? JSON.parse(v) : []; } catch (e) { v = []; }
    }
    item[h] = v === undefined ? '' : v;
  });
  return item;
}

function itemToRow_(item) {
  return HEADERS.map(h => {
    const v = item[h];
    if (h === 'dokumen' || h === 'riwayat') return JSON.stringify(v || []);
    if (h === 'updatedAt') return new Date().toISOString();
    return v === undefined ? '' : v;
  });
}

/** GET -> mengembalikan seluruh data sebagai JSON: {data:[...]} */
function doGet(e) {
  const sheet = getSheet_();
  const values = sheet.getDataRange().getValues();
  const rows = values.slice(1).filter(r => r[0]); // skip header, skip blank rows
  const data = rows.map(rowToItem_);
  return ContentService
    .createTextOutput(JSON.stringify({ data: data }))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * POST -> menerima {action:'sync', data:[...]} dari SIPATERA dan
 * menimpa seluruh isi Sheet dengan data terbaru (upsert sederhana
 * berbasis "id" lewat penulisan ulang penuh — cocok untuk skala kecil
 * satu instansi).
 */
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);

    if (body.action === 'sync' && Array.isArray(body.data)) {
      const sheet = getSheet_();
      sheet.clearContents();
      sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
      if (body.data.length > 0) {
        const rows = body.data.map(itemToRow_);
        sheet.getRange(2, 1, rows.length, HEADERS.length).setValues(rows);
      }
      return ContentService
        .createTextOutput(JSON.stringify({ ok: true, count: body.data.length }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // Berkas yang diupload pemohon (KTP, Surat Tugas, dll) disimpan ke Google Drive,
    // lalu link file-nya dikembalikan agar disimpan di kolom "dokumen" pada Sheet.
    if (body.action === 'uploadFile') {
      const folder = getFolder_();
      const bytes = Utilities.base64Decode(body.base64);
      const safeName = [body.nomor, body.docName, body.fileName].filter(Boolean).join('_');
      const blob = Utilities.newBlob(bytes, body.mimeType || 'application/octet-stream', safeName);
      const file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      return ContentService
        .createTextOutput(JSON.stringify({
          ok: true, url: file.getUrl(), fileId: file.getId(), fileName: file.getName()
        }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: 'Aksi tidak dikenali' }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}