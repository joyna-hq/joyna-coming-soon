// Saves joyna.ai and joynasaturdays.com sign-ups to the "joyna.ai waitlist" Google Sheet.
// Paste into Extensions > Apps Script in that sheet, then Deploy > Manage deployments > Edit > New version
// (Execute as: Me, Who has access: Anyone), so the web address stays the same.
//
// Sign-ups POST { email, page, company }: a new row of date, email, page.
// Host sign-ups POST { action: 'host', email, idea, key }: fills that person's "Wants to host" and
// "What they'd host" columns. key must match the HOST_KEY script property (Project Settings >
// Script properties); it is not in this file because this repo is public.

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
    const index = known.indexOf(email);

    if (p.action === 'host') {
      const key = PropertiesService.getScriptProperties().getProperty('HOST_KEY');
      if (!key || p.key !== key) return ok();
      // A leading ' keeps text that starts like a formula as plain text.
      const idea = String(p.idea || '').slice(0, 280).replace(/^([=+\-@])/, "'$1");
      let row = index + 2;
      if (index === -1) {
        sheet.appendRow([new Date().toISOString(), email, String(p.page || '').slice(0, 100)]);
        row = sheet.getLastRow();
      }
      sheet.getRange(row, 4, 1, 2).setValues([['Yes', idea]]);
    } else if (index === -1) {
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
