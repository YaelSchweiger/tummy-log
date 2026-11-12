// Tummy Log: connects the app to this Google Sheet.
// Paste this whole file into Extensions > Apps Script, then Deploy > New deployment > Web app
// (Execute as: Me, Who has access: Anyone).

const ENTRY_COLS = ["id", "device", "who", "kind", "day", "part", "dairy", "dairyOther",
                    "wheat", "wheatOther", "type", "note", "at", "look"];
const ISSUE_COLS = ["id", "who", "day", "text", "at"];

function sheet_(name, cols) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(cols);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, cols.length).setFontWeight("bold");
  } else if (sh.getLastColumn() < cols.length) {
    // A column was added later (like "look"): write the full header row again
    sh.getRange(1, 1, 1, cols.length).setValues([cols]).setFontWeight("bold");
  }
  return sh;
}

function rows_(sh, cols) {
  const values = sh.getDataRange().getValues().slice(1);
  return values.map((r) => {
    const o = {};
    cols.forEach((c, i) => { o[c] = r[i] === "" ? null : r[i]; });
    return o;
  });
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// Text from people can start with = + - @, which a sheet would run as a formula
function safe_(v) {
  if (v === null || v === undefined) return "";
  const s = String(v).slice(0, 2000);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function doGet(e) {
  // ?issues=1 lists the problem reports (read by the evening check)
  if (e && e.parameter && e.parameter.issues) {
    return out_({ ok: true, issues: rows_(sheet_("Issues", ISSUE_COLS), ISSUE_COLS) });
  }
  const entries = rows_(sheet_("Entries", ENTRY_COLS), ENTRY_COLS);
  return out_({ ok: true, entries: entries.map((e) => { delete e.device; return e; }) });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const req = JSON.parse(e.postData.contents);
    const sh = sheet_("Entries", ENTRY_COLS);

    if (req.action === "put") {
      const entry = req.entry || {};
      if (!entry.id || !req.device) return out_({ ok: false, error: "missing" });
      const ids = sh.getRange(1, 1, sh.getLastRow(), 1).getValues().map((r) => r[0]);
      const rowIndex = ids.indexOf(entry.id);
      // Only the phone that made an entry may change it
      if (rowIndex > 0 && sh.getRange(rowIndex + 1, 2).getValue() !== req.device) {
        return out_({ ok: false, error: "not_yours" });
      }
      const row = ENTRY_COLS.map((c) => (c === "device" ? req.device : c === "at" ? Number(entry.at) || Date.now() : safe_(entry[c])));
      if (rowIndex > 0) sh.getRange(rowIndex + 1, 1, 1, row.length).setValues([row]);
      else sh.appendRow(row);
      return out_({ ok: true });
    }

    if (req.action === "delete") {
      const ids = sh.getRange(1, 1, sh.getLastRow(), 1).getValues().map((r) => r[0]);
      const rowIndex = ids.indexOf(req.id);
      if (rowIndex <= 0) return out_({ ok: true });
      if (sh.getRange(rowIndex + 1, 2).getValue() !== req.device) return out_({ ok: false, error: "not_yours" });
      sh.deleteRow(rowIndex + 1);
      return out_({ ok: true });
    }

    if (req.action === "issue") {
      const ish = sheet_("Issues", ISSUE_COLS);
      ish.appendRow([Utilities.getUuid(), safe_(req.who), safe_(req.day), safe_(req.text), Date.now()]);
      return out_({ ok: true });
    }

    return out_({ ok: false, error: "unknown_action" });
  } finally {
    lock.releaseLock();
  }
}
