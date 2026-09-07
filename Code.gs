// === TIMESHEET MULTIMEDIA UNAI — BACKEND (Google Apps Script) ===
// Cara pakai: buka Google Sheet > Extensions > Apps Script > tempel kode ini
// menggantikan isi Code.gs, lalu Deploy > New deployment > Web app.

const SHEET_NAME = "Timesheet";

function doPost(e) {
  try {
    const sheet = getSheet();
    const data = JSON.parse(e.postData.contents);

    sheet.appendRow([
      new Date(),                 // Timestamp
      data.nama || "",            // Nama
      data.divisi || "",          // Divisi
      data.tanggal || "",         // Tanggal
      data.jenisPekerjaan || "",  // Jenis Pekerjaan
      data.deskripsi || "",       // Deskripsi
      data.jamMulai || "",        // Jam Mulai
      data.jamSelesai || "",      // Jam Selesai
      data.totalJam || "",        // Total Jam
      data.project || ""          // Project
    ]);

    return jsonResponse({ status: "success" });
  } catch (err) {
    return jsonResponse({ status: "error", message: err.message });
  }
}

function doGet(e) {
  try {
    const sheet = getSheet();
    const values = sheet.getDataRange().getValues();
    const headers = values.shift();

    const data = values
      .filter(row => row.some(cell => cell !== ""))
      .map(row => {
        const obj = {};
        headers.forEach((h, i) => {
          if (h === "Tanggal" && row[i] instanceof Date) {
            obj[h] = Utilities.formatDate(row[i], Session.getScriptTimeZone(), "yyyy-MM-dd");
          } else if (row[i] instanceof Date) {
            obj[h] = row[i].toISOString();
          } else {
            obj[h] = row[i];
          }
        });
        return obj;
      });

    return jsonResponse(data);
  } catch (err) {
    return jsonResponse({ status: "error", message: err.message });
  }
}

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow([
      "Timestamp", "Nama", "Divisi", "Tanggal", "Jenis Pekerjaan",
      "Deskripsi", "Jam Mulai", "Jam Selesai", "Total Jam", "Project"
    ]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
