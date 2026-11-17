/**
 * Sincronizare check-in -> Google Sheet.
 *
 * Aplicatia de check-in trimite situatia la fiecare 2 minute. Scriptul o scrie
 * intr-o foaie (tab) cu numele evenimentului; daca foaia nu exista, o creeaza.
 * Foaia e rescrisa complet la fiecare sincronizare, deci nu edita manual in ea
 * (copiaza datele in alta foaie daca vrei sa lucrezi pe ele).
 *
 * INSTALARE (o singura data, in Google Sheet-ul tinta):
 *  1. Extensions -> Apps Script. Sterge codul existent, lipeste acest fisier, Save.
 *  2. Project Settings (rotita) -> Script properties -> Add script property:
 *       TOKEN = un sir lung, aleator (acelasi il pui in Render la GSHEET_TOKEN)
 *  3. Deploy -> New deployment -> tip "Web app":
 *       Execute as: Me      Who has access: Anyone
 *     -> Deploy -> Authorize access (contul tau Google) -> copiezi "Web app URL".
 *  4. In Render, la serviciul evenimentului: GSHEET_WEBAPP_URL = URL-ul copiat.
 *
 * "Anyone" inseamna ca oricine are URL-ul poate apela scriptul, dar fara TOKEN
 * corect nu scrie nimic. Daca modifici codul: Deploy -> Manage deployments ->
 * Edit -> Version: New version (URL-ul ramane acelasi).
 */

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var token = PropertiesService.getScriptProperties().getProperty('TOKEN');
    if (!token || data.token !== token) {
      return out_({ok: false, error: 'token invalid: verifica TOKEN (Script properties) si GSHEET_TOKEN (Render)'});
    }
    var lock = LockService.getScriptLock();
    if (!lock.tryLock(30000)) return out_({ok: false, error: 'sheet ocupat, se reincearca'});
    try {
      var ss = data.spreadsheetId ? SpreadsheetApp.openById(data.spreadsheetId) : SpreadsheetApp.getActiveSpreadsheet();
      if (!ss) return out_({ok: false, error: 'spreadsheet negasit (lipseste googleSheet in event.json?)'});
      var name = String(data.sheetName || 'Check-in').replace(/[\[\]\*\?\/\\:]/g, '-').slice(0, 99);
      var sh = ss.getSheetByName(name);
      var created = false;
      if (!sh) {
        sh = ss.insertSheet(name);
        created = true;
      }
      write_(sh, data, created);
      return out_({ok: true, sheet: name, created: created, rows: (data.rows || []).length});
    } finally {
      lock.releaseLock();
    }
  } catch (err) {
    return out_({ok: false, error: String((err && err.message) || err)});
  }
}

// Deschis in browser, URL-ul arata daca scriptul e publicat corect.
function doGet() {
  return out_({ok: true, info: 'Scriptul de sincronizare check-in este publicat.'});
}

function write_(sh, d, created) {
  var header = d.header || [];
  var cols = Math.max(header.length, 1);
  var rows = d.rows || [];
  var values = [pad_([d.title || 'Check-in'], cols), pad_([d.summary || ''], cols), pad_([], cols), pad_(header, cols)];
  rows.forEach(function (r) { values.push(pad_(r, cols)); });
  values.push(pad_([], cols));
  values.push(pad_(d.totals || [], cols));

  var need = values.length;
  if (sh.getMaxRows() < need) sh.insertRowsAfter(sh.getMaxRows(), need - sh.getMaxRows());
  if (sh.getMaxColumns() < cols) sh.insertColumnsAfter(sh.getMaxColumns(), cols - sh.getMaxColumns());

  sh.clear();                                              // continut + formatare
  sh.getRange(1, 3, need, 1).setNumberFormat('@');         // "Ora sosirii" ramane text (19:15)
  sh.getRange(1, 1, need, cols).setValues(values);

  sh.getRange(1, 1).setFontSize(13).setFontWeight('bold');
  sh.getRange(2, 1).setFontStyle('italic').setFontColor('#666666');
  sh.getRange(4, 1, 1, cols).setFontWeight('bold').setBackground('#262626').setFontColor('#ffffff');
  if (rows.length) {
    var bg = rows.map(function (r) { return pad_([], cols).map(function () { return r[1] === 'A venit' ? '#e6f4ea' : null; }); });
    sh.getRange(5, 1, rows.length, cols).setBackgrounds(bg);
  }
  sh.getRange(need, 1, 1, cols).setFontWeight('bold').setBackground('#ede9e0');
  sh.setFrozenRows(4);

  if (created) {
    var widths = [50, 90, 90, 220, 260, 130, 130];
    for (var i = 0; i < Math.min(cols, widths.length); i++) sh.setColumnWidth(i + 1, widths[i]);
  }
}

function pad_(r, n) {
  var o = [];
  for (var i = 0; i < n; i++) {
    var v = (r && i < r.length && r[i] !== null && r[i] !== undefined) ? r[i] : '';
    if (typeof v === 'string' && /^[=+\-@]/.test(v)) v = "'" + v;   // un nume nu devine formula
    o.push(v);
  }
  return o;
}

function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
