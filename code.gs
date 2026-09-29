const SHEET_NAME = 'Data';
const DRIVE_FOLDER_NAME = 'SIPATERA-DOKUMEN';

const HEADERS = [
  'id',
  'nomor',
  'tanggal',
  'pemohon',
  'instansi',
  'alamat',
  'telepon',
  'email',
  'alat',
  'jumlah',
  'keperluan',
  'lokasi',
  'mulai',
  'selesai',
  'status',
  'catatan',
  'createdAt',
  'updatedAt'
];

/* =========================================================
   SETUP
   Jalankan fungsi setup() SATU KALI dari Apps Script
   ========================================================= */
function setup() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();

  if (!ss) {
    ss = SpreadsheetApp.create('SIPATERA-DATA');
  }

  // Simpan ID spreadsheet agar Web App selalu menggunakan
  // spreadsheet yang sama.
  PropertiesService
    .getScriptProperties()
    .setProperty('SHEET_ID', ss.getId());

  let sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }

  // Buat header jika sheet masih kosong
  if (sheet.getLastRow() === 0) {
    sheet
      .getRange(1, 1, 1, HEADERS.length)
      .setValues([HEADERS]);

    sheet.setFrozenRows(1);
  }

  // Buat folder Google Drive
  let folders = DriveApp.getFoldersByName(DRIVE_FOLDER_NAME);

  if (!folders.hasNext()) {
    DriveApp.createFolder(DRIVE_FOLDER_NAME);
  }

  return {
    spreadsheetId: ss.getId(),
    spreadsheetUrl: ss.getUrl(),
    sheetName: SHEET_NAME
  };
}


/* =========================================================
   AMBIL SHEET SIPATERA
   ========================================================= */
function getSheet_() {
  const props = PropertiesService.getScriptProperties();
  const sheetId = props.getProperty('SHEET_ID');

  let ss = null;

  if (sheetId) {
    ss = SpreadsheetApp.openById(sheetId);
  } else {
    ss = SpreadsheetApp.getActiveSpreadsheet();
  }

  if (!ss) {
    throw new Error(
      'Spreadsheet SIPATERA belum dikonfigurasi. Jalankan setup() satu kali.'
    );
  }

  let sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);

    sheet
      .getRange(1, 1, 1, HEADERS.length)
      .setValues([HEADERS]);

    sheet.setFrozenRows(1);
  }

  return sheet;
}


/* =========================================================
   GET
   Dipakai aplikasi untuk mengambil data dari Google Sheet
   ========================================================= */
function doGet(e) {
  try {

    const action = e && e.parameter
      ? e.parameter.action
      : '';

    if (action === 'getData') {
      const sheet = getSheet_();
      const values = sheet.getDataRange().getValues();

      if (values.length <= 1) {
        return json_({
          ok: true,
          data: []
        });
      }

      const headers = values[0];

      const data = values
        .slice(1)
        .map(function(row) {
          const obj = {};

          headers.forEach(function(header, index) {
            obj[header] = row[index];
          });

          return obj;
        });

      return json_({
        ok: true,
        data: data
      });
    }

    return json_({
      ok: true,
      message: 'SIPATERA Web App aktif.'
    });

  } catch (err) {

    return json_({
      ok: false,
      error: err.message
    });

  }
}


/* =========================================================
   POST
   Menerima data dari aplikasi SIPATERA
   ========================================================= */
function doPost(e) {

  try {

    if (!e || !e.postData || !e.postData.contents) {
      throw new Error('Data POST tidak ditemukan.');
    }

    const body = JSON.parse(e.postData.contents);

    /* -----------------------------------------------------
       SIMPAN SATU DATA
       Dipakai saat user klik KIRIM pada aplikasi
       ----------------------------------------------------- */
    if (body.action === 'saveItem' && body.item) {

      const sheet = getSheet_();
      const item = body.item;

      const row = itemToRow_(item);

      const idCol = HEADERS.indexOf('id') + 1;
      const lastRow = sheet.getLastRow();

      const lock = LockService.getScriptLock();

      lock.waitLock(20000);

      try {

        let targetRow = -1;

        // Cari ID yang sama
        if (lastRow >= 2) {

          const ids = sheet
            .getRange(
              2,
              idCol,
              lastRow - 1,
              1
            )
            .getValues()
            .flat();

          const idx = ids.findIndex(function(v) {
            return String(v) === String(item.id);
          });

          if (idx >= 0) {
            targetRow = idx + 2;
          }
        }

        // Jika ID sudah ada → UPDATE
        if (targetRow >= 2) {

          sheet
            .getRange(
              targetRow,
              1,
              1,
              HEADERS.length
            )
            .setValues([row]);

        }

        // Jika belum ada → INSERT
        else {

          sheet
            .getRange(
              sheet.getLastRow() + 1,
              1,
              1,
              HEADERS.length
            )
            .setValues([row]);

        }

      } finally {

        lock.releaseLock();

      }

      return json_({
        ok: true,
        id: item.id
      });
    }


    /* -----------------------------------------------------
       SYNC SEMUA DATA
       Digunakan sebagai backup/manual recovery
       ----------------------------------------------------- */
    if (body.action === 'sync') {

      const sheet = getSheet_();
      const data = body.data || [];

      const lock = LockService.getScriptLock();

      lock.waitLock(20000);

      try {

        data.forEach(function(item) {

          const row = itemToRow_(item);

          const idCol = HEADERS.indexOf('id') + 1;
          const lastRow = sheet.getLastRow();

          let targetRow = -1;

          if (lastRow >= 2) {

            const ids = sheet
              .getRange(
                2,
                idCol,
                lastRow - 1,
                1
              )
              .getValues()
              .flat();

            const idx = ids.findIndex(function(v) {
              return String(v) === String(item.id);
            });

            if (idx >= 0) {
              targetRow = idx + 2;
            }
          }

          if (targetRow >= 2) {

            sheet
              .getRange(
                targetRow,
                1,
                1,
                HEADERS.length
              )
              .setValues([row]);

          } else {

            sheet
              .getRange(
                sheet.getLastRow() + 1,
                1,
                1,
                HEADERS.length
              )
              .setValues([row]);

          }

        });

      } finally {

        lock.releaseLock();

      }

      return json_({
        ok: true,
        count: data.length
      });
    }


    /* -----------------------------------------------------
       UPLOAD FILE KE GOOGLE DRIVE
       ----------------------------------------------------- */
    if (body.action === 'uploadFile') {

      if (!body.base64) {
        throw new Error('Data file/base64 kosong.');
      }

      const folder = getDriveFolder_();

      const bytes = Utilities.base64Decode(body.base64);

      const blob = Utilities.newBlob(
        bytes,
        body.mimeType || 'application/octet-stream',
        body.fileName || 'dokumen'
      );

      const file = folder.createFile(blob);

      // Simpan metadata jika tersedia
      if (body.nomor) {
        file.setDescription(
          'SIPATERA - Nomor: ' + body.nomor
        );
      }

      return json_({
        ok: true,
        fileId: file.getId(),
        fileName: file.getName(),
        url: file.getUrl()
      });
    }


    throw new Error(
      'Action tidak dikenali: ' + body.action
    );


  } catch (err) {

    return json_({
      ok: false,
      error: err.message,
      stack: err.stack
    });

  }
}


/* =========================================================
   KONVERSI DATA ITEM → BARIS GOOGLE SHEET
   ========================================================= */
function itemToRow_(item) {

  return HEADERS.map(function(header) {

    let value = item[header];

    if (
      value === undefined ||
      value === null
    ) {
      value = '';
    }

    // Object / Array → JSON
    if (
      typeof value === 'object'
    ) {
      try {
        value = JSON.stringify(value);
      } catch (err) {
        value = String(value);
      }
    }

    return value;
  });

}


/* =========================================================
   GOOGLE DRIVE FOLDER
   ========================================================= */
function getDriveFolder_() {

  const folders =
    DriveApp.getFoldersByName(DRIVE_FOLDER_NAME);

  if (folders.hasNext()) {
    return folders.next();
  }

  return DriveApp.createFolder(
    DRIVE_FOLDER_NAME
  );
}


/* =========================================================
   RESPONSE JSON
   ========================================================= */
function json_(data) {

  return ContentService
    .createTextOutput(
      JSON.stringify(data)
    )
    .setMimeType(
      ContentService.MimeType.JSON
    );

}