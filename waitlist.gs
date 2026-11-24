// Saves joyna.ai and joynasaturdays.com sign-ups to the "joyna.ai waitlist" Google Sheet.
// Paste into Extensions > Apps Script in that sheet, then Deploy > Manage deployments > Edit > New version
// (Execute as: Me, Who has access: Anyone), so the web address stays the same.
//
// Columns: A signed up, B email, C page, D wants to host, E what they'd host, F friends invited, G tier, H city.
// Sign-ups POST { email, page, company }: a new row of date, email, page.
// joynasaturdays.com also POSTs, with key:
//   { action: 'friends', email, friends, tier }: sets F and G on that person's row (adding the row if missing)
//   { action: 'host', email, idea }: sets D and E on that person's row
//   { action: 'city', email, city }: sets H on that person's row
// key must match the HOST_KEY script property (Project Settings > Script properties); it is not in this
// file because this repo is public. A 'friends' request with a wrong key still counts as a plain sign-up.

function doPost(e) {
  const p = (e && e.parameter) || {};
  const email = String(p.email || '').trim().toLowerCase();
  // company is a hidden field people never see; bots fill it in.
  if (p.company) return ok();
  // Starting with = + - @ would make the cell a formula, so those are refused too.
  if (email.length > 254 || !/^[^\s@=+\-][^\s@]*@[^\s@]+\.[^\s@]+$/.test(email)) return ok();

  const key = PropertiesService.getScriptProperties().getProperty('HOST_KEY');
  const trusted = Boolean(key) && p.key === key;

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
    const last = sheet.getLastRow();
    const known = last > 1 ? sheet.getRange(2, 2, last - 1, 1).getValues().flat() : [];
    const index = known.indexOf(email);
    // Their row, added (date, email, page) when they aren't on the sheet yet.
    const row = () => {
      if (index !== -1) return index + 2;
      sheet.appendRow([new Date().toISOString(), email, String(p.page || '').slice(0, 100)]);
      return sheet.getLastRow();
    };

    if (p.action === 'host') {
      if (!trusted) return ok();
      // A leading ' keeps text that starts like a formula as plain text.
      const idea = String(p.idea || '').slice(0, 280).replace(/^([=+\-@])/, "'$1");
      sheet.getRange(row(), 4, 1, 2).setValues([['Yes', idea]]);
    } else if (p.action === 'city' && trusted) {
      const city = String(p.city || '').slice(0, 60).replace(/^([=+\-@])/, "'$1");
      sheet.getRange(row(), 8).setValue(city);
    } else if (p.action === 'friends' && trusted) {
      const friends = Math.max(0, Math.min(100000, parseInt(p.friends, 10) || 0));
      const tier = String(p.tier || '').slice(0, 40).replace(/^([=+\-@])/, "'$1");
      sheet.getRange(row(), 6, 1, 2).setValues([[friends, tier]]);
    } else if (index === -1) {
      row();
    }
  } finally {
    lock.releaseLock();
  }
  return ok();
}

function ok() {
  return ContentService.createTextOutput('ok');
}
