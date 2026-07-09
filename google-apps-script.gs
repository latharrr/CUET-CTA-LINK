// Google Apps Script for Picapool tracking
// 1. Create a new Google Sheet
// 2. Extensions > Apps Script
// 3. Paste this code and click Deploy > New deployment > Web app
// 4. Set "Who has access" to "Anyone"
// 5. Copy the web app URL and paste it into index.html as GAS_WEB_APP_URL

const SHEET_NAME = "Tracking";
const LEAD_SHEET_NAME = "Lead";

function doGet(e) {
  return jsonResponse({ ok: true });
}

function doPost(e) {
  try {
    const sheet = getSheet(SHEET_NAME);
    const raw = e.postData ? e.postData.contents : "{}";
    const data = JSON.parse(raw || "{}");

    const headers = ["Timestamp", "IP", "Page", "Event", "Target", "Detail", "User Agent", "Referrer"];
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(headers);
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

    if (data.event === "form" && data.target === "phone_submit") {
      appendLead(data);
    }

    return jsonResponse({ ok: true, row: row });
  } catch (err) {
    return jsonResponse({ ok: false, error: err.toString() });
  }
}

function appendLead(data) {
  const leadSheet = getSheet(LEAD_SHEET_NAME);
  const leadHeaders = ["Timestamp", "Phone", "Campus", "IP", "Page", "User Agent", "Referrer"];
  if (leadSheet.getLastRow() === 0) {
    leadSheet.appendRow(leadHeaders);
  }

  leadSheet.appendRow([
    new Date().toISOString(),
    data.detail || "",
    data.campus || "",
    data.ip || "",
    data.page || "",
    data.ua || "",
    data.ref || ""
  ]);
}

function doOptions(e) {
  return jsonResponse({ ok: true });
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON)
    .setHeaders({
      "Access-Control-Allow-Origin": "*",
      "Access-Control-Allow-Methods": "GET, POST, OPTIONS",
      "Access-Control-Allow-Headers": "Content-Type"
    });
}

function getSheet(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  return sheet;
}
