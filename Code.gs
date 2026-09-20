/**
 * Yogalates Form — Google Apps Script backend
 * -------------------------------------------------------------
 * Handles two things:
 *   doGet()  -> returns { registered, max, full } so form.html can show live
 *               capacity and switch to waitlist mode when the class is full.
 *   doPost() -> receives a submission from form.html, uploads the payment
 *               proof image to a Drive folder, and appends a row to your
 *               Spreadsheet.
 *
 * TESTING (run from the Apps Script editor, not the web app):
 *   - doTest()        -> checks config + Spreadsheet + Drive + sheet access.
 *   - doInsertDummy() -> inserts dummy members so you can test the flow.
 *
 * SETUP (do this once):
 *   1. Create a Google Spreadsheet, copy its ID from the URL:
 *        https://docs.google.com/spreadsheets/d/<THIS_IS_THE_ID>/edit
 *      and paste it into SPREADSHEET_ID below.
 *   2. Create a Drive folder for the uploaded payment proofs, open it,
 *      copy the folder ID from the URL:
 *        https://drive.google.com/drive/folders/<THIS_IS_THE_ID>
 *      and paste it into DRIVE_FOLDER_ID below.
 *   3. Deploy:  Deploy > New deployment > Web app
 *        - Execute as:  Me
 *        - Who has access:  Anyone
 *      Copy the /exec URL and paste it into form.html (the ENDPOINT variable).
 *   4. Re-deploy (Manage deployments > Edit > New version) whenever you
 *      change this script.
 * -------------------------------------------------------------
 */

// ====== CONFIG ======
var SPREADSHEET_ID = 'PASTE_YOUR_SPREADSHEET_ID_HERE';
var SHEET_NAME     = 'Pendaftaran';
var DRIVE_FOLDER_ID = 'PASTE_YOUR_DRIVE_FOLDER_ID_HERE';

// Settings live in a second sheet named "Config" (column A = key, B = value):
//   Date | Theme | Max Registered | Location
// "Max Registered" is read from there. DEFAULT_MAX_REGISTERED is only used as
// a fallback if that sheet or row is missing.
var CONFIG_SHEET_NAME = 'Config';
var DEFAULT_MAX_REGISTERED = 30;

var HEADERS = ['Timestamp', 'Email', 'Seatalks', 'Sewa Mat', 'Bukti Pembayaran', 'Status'];

// ====== GET: report live registration capacity + config ======
// Polled by form.html every 2s. Returns how many are registered, the cap,
// whether the class is full, and all values from the "Config" sheet.
function doGet() {
  var cfg = getConfig();
  var max = getMaxRegistered(cfg);
  var registered = countRegistered();
  return json({
    registered: registered,
    max: max,
    full: registered >= max,
    config: cfg
  });
}

// ====== POST: store a submission ======
function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000); // avoid two submissions writing at once

    var p = (e && e.parameter) ? e.parameter : {};

    var email    = (p.email || '').toString().trim();
    var seatalks = (p.seatalks || '').toString().trim();
    var sewaMat  = (p.sewaMat || '').toString().trim();

    if (!email || !seatalks || !sewaMat) {
      return json({ result: 'error', message: 'Data wajib belum lengkap.' });
    }

    // ---- decide Registered vs Waitlist based on current capacity ----
    var registered = countRegistered();
    var isWaitlist = registered >= getMaxRegistered();
    var status = isWaitlist ? 'Waitlist' : 'Registered';

    // payment proof is only required for actual registrants (not waitlist)
    if (!isWaitlist && !p.fileData) {
      return json({ result: 'error', message: 'Bukti pembayaran wajib diunggah.' });
    }

    // ---- upload payment-proof image to Drive (if provided) ----
    var fileUrl = '';
    if (p.fileData) {
      var folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
      var mime = p.mimeType || 'image/jpeg';
      var name = sanitize(p.fileName || ('bukti_' + email));
      var stamped = seatalks.replace(/^@/, '') + '_' + name;
      var bytes = Utilities.base64Decode(p.fileData);
      var blob = Utilities.newBlob(bytes, mime, stamped);
      var file = folder.createFile(blob);
      // make it viewable by anyone with the link (comment out if you want it private)
      try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (err) {}
      fileUrl = file.getUrl();
    }

    // ---- append row to the sheet ----
    var sheet = getSheet();
    sheet.appendRow([new Date(), email, seatalks, sewaMat, fileUrl, status]);

    return json({ result: 'success', status: status });
  } catch (err) {
    return json({ result: 'error', message: err.message || String(err) });
  } finally {
    try { lock.releaseLock(); } catch (e2) {}
  }
}

// ====== helpers ======
// Count rows whose Status is "Registered" (waitlist rows don't take a slot).
function countRegistered() {
  var sheet = getSheet();
  var last = sheet.getLastRow();
  if (last < 2) return 0; // only header (or empty)
  var statusCol = HEADERS.indexOf('Status') + 1;
  var values = sheet.getRange(2, statusCol, last - 1, 1).getValues();
  var n = 0;
  for (var i = 0; i < values.length; i++) {
    if (String(values[i][0]).trim().toLowerCase() === 'registered') n++;
  }
  return n;
}

// Read the "Config" sheet (col A = key, col B = value) into a plain object,
// e.g. { "Date": "22 September 2026", "Theme": "test tema", "Max Registered": 10, "Location": 12 }
function getConfig() {
  var cfg = {};
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(CONFIG_SHEET_NAME);
  if (!sheet) return cfg;
  var last = sheet.getLastRow();
  if (last < 1) return cfg;
  var values = sheet.getRange(1, 1, last, 2).getValues();
  for (var i = 0; i < values.length; i++) {
    var key = String(values[i][0]).trim();
    if (!key) continue;
    var val = values[i][1];
    // dates -> "YYYY-MM-DD" so they serialize cleanly to JSON
    if (Object.prototype.toString.call(val) === '[object Date]') {
      val = Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    }
    cfg[key] = val;
  }
  return cfg;
}

// Max registered pulled from the Config sheet ("Max Registered"), with fallback.
function getMaxRegistered(cfg) {
  cfg = cfg || getConfig();
  var n = parseInt(cfg['Max Registered'], 10);
  return (!isNaN(n) && n > 0) ? n : DEFAULT_MAX_REGISTERED;
}

function getSheet() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }
  // ensure header row exists
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function sanitize(name) {
  return String(name).replace(/[\\/:*?"<>|]+/g, '_').slice(0, 120);
}

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// TEST HELPERS — run these from the Apps Script editor (Run menu).
// They are NOT called by doGet/doPost. Open View > Logs to see output.
// ============================================================

/**
 * doTest() — smoke-test the whole setup end to end.
 * Verifies config is filled in, the Spreadsheet opens, the sheet + headers
 * exist, the Drive folder is reachable and writable, and counting works.
 */
function doTest() {
  var results = [];
  function check(label, fn) {
    try {
      var detail = fn();
      results.push({ ok: true, label: label, detail: detail || 'OK' });
      Logger.log('✅ ' + label + (detail ? '  →  ' + detail : ''));
    } catch (err) {
      results.push({ ok: false, label: label, detail: err.message || String(err) });
      Logger.log('❌ ' + label + '  →  ' + (err.message || err));
    }
  }

  check('Config: SPREADSHEET_ID set', function () {
    if (!SPREADSHEET_ID || SPREADSHEET_ID.indexOf('PASTE_') === 0) throw new Error('SPREADSHEET_ID is not set');
    return SPREADSHEET_ID;
  });

  check('Config: DRIVE_FOLDER_ID set', function () {
    if (!DRIVE_FOLDER_ID || DRIVE_FOLDER_ID.indexOf('PASTE_') === 0) throw new Error('DRIVE_FOLDER_ID is not set');
    return DRIVE_FOLDER_ID;
  });

  check('Spreadsheet opens', function () {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    return ss.getName();
  });

  check('Sheet + headers ready', function () {
    var sheet = getSheet();
    return '"' + sheet.getName() + '" with columns: ' + HEADERS.join(', ');
  });

  check('Drive folder reachable + writable', function () {
    var folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
    // write a tiny probe file then remove it, to prove we can create files
    var probe = folder.createFile(Utilities.newBlob('yogalates test', 'text/plain', '__test_probe.txt'));
    probe.setTrashed(true);
    return 'folder "' + folder.getName() + '" (write OK)';
  });

  check('Config sheet read', function () {
    var cfg = getConfig();
    var keys = Object.keys(cfg);
    if (!keys.length) throw new Error('"' + CONFIG_SHEET_NAME + '" sheet is empty or missing');
    return keys.map(function (k) { return k + '=' + cfg[k]; }).join(', ');
  });

  check('Count registered works', function () {
    return countRegistered() + ' / ' + getMaxRegistered() + ' registered';
  });

  var passed = results.filter(function (r) { return r.ok; }).length;
  var summary = passed + '/' + results.length + ' checks passed';
  Logger.log('———————————————\n' + summary +
    (passed === results.length ? '  🎉 all good!' : '  ⚠️ fix the ❌ items above'));
  return { summary: summary, passed: passed, total: results.length, results: results };
}

/**
 * doInsertDummy() — insert dummy members to test the flow / waitlist logic.
 * Each dummy gets a placeholder "bukti" link (no real file upload) and its
 * status is computed the same way as a real submission, so once you cross
 * MAX_REGISTERED the extras correctly become "Waitlist".
 * Change HOW_MANY to insert more/less per run.
 */
function doInsertDummy() {
  var HOW_MANY = 5;
  var sheet = getSheet();
  var max = getMaxRegistered();
  var inserted = [];

  for (var i = 0; i < HOW_MANY; i++) {
    var registered = countRegistered();
    var status = registered >= max ? 'Waitlist' : 'Registered';
    var n = sheet.getLastRow(); // rough unique-ish suffix
    var email = 'dummy' + n + '@test.com';
    var seatalks = '@dummy' + n;
    var sewaMat = (n % 2 === 0) ? 'Ya' : 'Tidak';
    var fileUrl = (status === 'Registered') ? 'https://example.com/dummy-bukti-' + n + '.jpg' : '';

    sheet.appendRow([new Date(), email, seatalks, sewaMat, fileUrl, status]);
    inserted.push(email + ' → ' + status);
    Logger.log('➕ ' + email + '  (' + status + ')');
  }

  Logger.log('———————————————\nInserted ' + inserted.length + ' dummy member(s). ' +
    'Now ' + countRegistered() + '/' + max + ' registered.');
  return { inserted: inserted, registered: countRegistered(), max: max };
}
