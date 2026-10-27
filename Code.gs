const SHEET_NAME = "Registrations";

// Set SPREADSHEET_ID under Apps Script > Project Settings > Script properties.
function doGet() {
  return ContentService.createTextOutput("NEXPLORA '26 registration endpoint is running.");
}
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const p = (e && e.parameter) ? e.parameter : {};
    if (p.website) return output_({ok:false, message:"Invalid submission"});
    const fields = ["schoolName","schoolAddress","coordinatorName","designation","email","mobile"];
    const values = fields.map(k => String(p[k] || "").trim());
    if (values.some(v => !v)) return output_({ok:false, message:"All fields are required"});
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(values[4])) return output_({ok:false, message:"Invalid email"});
    const id = PropertiesService.getScriptProperties().getProperty("SPREADSHEET_ID");
    if (!id) throw new Error("Missing SPREADSHEET_ID script property");
    const ss = SpreadsheetApp.openById(id);
    let sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(["Timestamp","School Name","School Address","Coordinator Name","Coordinator Designation","Email Address","Mobile Number"]);
      sheet.setFrozenRows(1);
      sheet.getRange(1,1,1,7).setFontWeight("bold");
    }
    const safe = values.map(v => /^[=+\-@]/.test(v) ? "'" + v : v);
    sheet.appendRow([new Date()].concat(safe));
    return output_({ok:true});
  } catch (err) {
    console.error(err);
    return output_({ok:false, message:"Unable to save registration"});
  } finally {
    lock.releaseLock();
  }
}
function output_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
