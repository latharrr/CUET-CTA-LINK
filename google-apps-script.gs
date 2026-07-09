// Google Apps Script for Picapool tracking
// 1. Create a new Google Sheet
// 2. Extensions > Apps Script
// 3. Paste this code and click Deploy > New deployment > Web app
// 4. Set "Who has access" to "Anyone"
// 5. Copy the web app URL and paste it into index.html as GAS_WEB_APP_URL

const SHEET_NAME = "Tracking";

function doGet(e) {
  return jsonResponse({ ok: true });
}

function doPost(e) {
  try {
    const sheet = getSheet();
    const raw = e.postData ? e.postData.contents : "{}";
    const data = JSON.parse(raw || "{}");

    const headers = ["Timestamp", "IP", "Page", "Event", "Target", "Detail", "User Agent", "Referrer"];
    if (sheet.getRange(1, 1).getValue() !== "Timestamp") {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    }

    const row = [
      new Date().toISOString(),
      data.ip || "",
      data.page || "",
      data.event || "",
      data.target || "",
      data.detail || "",
      data.ua || "",
      data.ref || ""
    ];

    sheet.appendRow(row);
    return jsonResponse({ ok: true, row: row });
  } catch (err) {
    return jsonResponse({ ok: false, error: err.toString() });
  }
}

function doOptions(e) {
  return jsonResponse({ ok: true });
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }
  return sheet;
}
