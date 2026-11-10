// Výsledky třídy pro hru „90 dní“.
// Vlož do Google tabulky: Rozšíření → Apps Script, a nasaď jako webovou aplikaci (přístup: Kdokoli).
// Tabulka ukládá jen: čas, kód konce a firmu. Žádná osobní data.

const HOURS = 12; // statistiky počítají jen hry z posledních 12 hodin (= jedna prezentace)

function doGet(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  const p = (e && e.parameter) || {};

  if (p.action === 'add' && /^[a-zA-Z]{2,20}$/.test(p.end || '')) {
    const lock = LockService.getScriptLock();
    lock.waitLock(5000);
    try {
      sheet.appendRow([new Date(), p.end, String(p.firm || '').slice(0, 20)]);
    } finally {
      lock.releaseLock();
    }
  }

  const since = Date.now() - HOURS * 3600 * 1000;
  const rows = sheet.getDataRange().getValues();
  const counts = {};
  let total = 0;
  rows.forEach(function (r) {
    if (r[0] instanceof Date && r[0].getTime() >= since && r[1]) {
      counts[r[1]] = (counts[r[1]] || 0) + 1;
      total++;
    }
  });

  return ContentService
    .createTextOutput(JSON.stringify({ counts: counts, total: total }))
    .setMimeType(ContentService.MimeType.JSON);
}
