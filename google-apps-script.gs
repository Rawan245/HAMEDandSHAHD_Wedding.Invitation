// Google Apps Script — paste into Extensions > Apps Script in your Google Sheet
// Row 1 of the sheet should be: Time | Name | Message
function clean(v, n) {
  v = String(v || '').slice(0, n);
  return /^[=+\-@]/.test(v) ? "'" + v : v; // stop spreadsheet formulas
}
function doGet() {
  const s = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  const rows = s.getLastRow() > 1 ? s.getRange(2, 1, s.getLastRow() - 1, 3).getValues() : [];
  const out = rows.map(r => ({ n: r[1], m: r[2] }));
  return ContentService.createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}
function doPost(e) {
  const d = JSON.parse(e.postData.contents);
  if (!d.n || !d.m) return ContentService.createTextOutput('skip');
  SpreadsheetApp.getActiveSpreadsheet().getSheets()[0]
    .appendRow([new Date(), clean(d.n, 40), clean(d.m, 400)]);
  return ContentService.createTextOutput('ok');
}
