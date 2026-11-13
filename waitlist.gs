// Saves joyna.ai sign-ups to the "joyna.ai waitlist" Google Sheet.
// Paste into Extensions > Apps Script in that sheet, then Deploy > New deployment > Web app
// (Execute as: Me, Who has access: Anyone). The page POSTs { email, page, company }.

function doPost(e) {
  const p = (e && e.parameter) || {};
  const email = String(p.email || '').trim().toLowerCase();
  // company is a hidden field people never see; bots fill it in.
  if (p.company) return ok();
  // Starting with = + - @ would make the cell a formula, so those are refused too.
  if (email.length > 254 || !/^[^\s@=+\-][^\s@]*@[^\s@]+\.[^\s@]+$/.test(email)) return ok();

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
    const last = sheet.getLastRow();
    const known = last > 1 ? sheet.getRange(2, 2, last - 1, 1).getValues().flat() : [];
    if (!known.includes(email)) {
      sheet.appendRow([new Date().toISOString(), email, String(p.page || '').slice(0, 100)]);
    }
  } finally {
    lock.releaseLock();
  }
  return ok();
}

function ok() {
  return ContentService.createTextOutput('ok');
}
