/**
 * RHYC registration service (Google Apps Script, bound to a Google Sheet).
 * - Anyone can send a registration.
 * - Only someone who knows ADMIN_CODE can read, confirm or delete registrations.
 */

// Change this to your own secret code (use at least 8 characters). Keep it private.
const ADMIN_CODE = 'CHANGE-THIS-CODE';

const SHEET_NAME = 'Registrations';
const HEAD = ['id', 'createdAt', 'name', 'age', 'cls', 'phone', 'phoneKey', 'gname', 'gphone', 'question', 'heard', 'status'];
const MAX_ROWS = 5000;

function doGet() {
  return ContentService.createTextOutput('RHYC registration service is running.');
}

function doPost(e) {
  let out;
  try {
    const req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    out = route_(req);
  } catch (err) {
    out = { ok: false, error: 'Bad request.' };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function route_(r) {
  if (r.action === 'register') return register_(r);
  const gate = checkCode_(r.code);
  if (gate) return gate;
  if (r.action === 'list') return { ok: true, rows: readAll_() };
  if (r.action === 'update') return update_(r);
  if (r.action === 'delete') return remove_(r);
  return { ok: false, error: 'Unknown request.' };
}

/* ---------- code check ---------- */
function checkCode_(code) {
  const cache = CacheService.getScriptCache();
  const fails = Number(cache.get('fails') || 0);
  if (fails >= 10) return { ok: false, error: 'locked' };
  if (typeof code === 'string' && code.length > 0 && code === ADMIN_CODE) {
    cache.remove('fails');
    return null;
  }
  cache.put('fails', String(fails + 1), 300);
  Utilities.sleep(1500); // slows down guessing
  return { ok: false, error: 'code' };
}

/* ---------- sheet ---------- */
function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.getRange(1, 1, 1, HEAD.length).setValues([HEAD]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.getRange(1, 1, sh.getMaxRows(), HEAD.length).setNumberFormat('@'); // keep everything as plain text
  }
  return sh;
}

function readAll_() {
  const sh = sheet_();
  const last = sh.getLastRow();
  if (last < 2) return [];
  const vals = sh.getRange(2, 1, last - 1, HEAD.length).getValues();
  const rows = vals.map(function (v) {
    const o = {};
    HEAD.forEach(function (h, i) { o[h] = v[i] === null || v[i] === undefined ? '' : String(v[i]); });
    o.age = Number(o.age);
    return o;
  });
  rows.sort(function (a, b) { return a.createdAt < b.createdAt ? 1 : -1; });
  return rows;
}

function findRow_(sh, id) {
  const last = sh.getLastRow();
  if (last < 2) return -1;
  const ids = sh.getRange(2, 1, last - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return i + 2;
  return -1;
}

/* ---------- actions ---------- */
function clean_(v, max) {
  return String(v === null || v === undefined ? '' : v).replace(/[\u0000-\u001f\u007f]/g, ' ').replace(/\s+/g, ' ').trim().slice(0, max);
}
function phoneOk_(p) { return /^[0-9+\-\s()]{7,20}$/.test(p) && p.replace(/\D/g, '').length >= 7; }

function register_(r) {
  if (r.website) return { ok: true }; // honeypot: bots fill this hidden field
  const name = clean_(r.name, 80);
  const age = Number(r.age);
  const phone = clean_(r.phone, 20);
  const gname = clean_(r.gname, 80);
  const gphone = clean_(r.gphone, 20);
  if (name.length < 2 || !Number.isInteger(age) || !phoneOk_(phone)) return { ok: false, error: 'Please fill in the name, age and a valid phone number.' };
  if (age < 5 || age > 80) return { ok: false, error: 'Please enter a valid age.' };
  if (age < 18 && (gname.length < 2 || !phoneOk_(gphone))) return { ok: false, error: 'Please add the parent or guardian name and phone number.' };
  if (r.consent !== true) return { ok: false, error: 'Please tick the confirmation box to continue.' };

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const sh = sheet_();
    if (sh.getLastRow() - 1 >= MAX_ROWS) return { ok: false, error: 'Registration is full. Please contact the organizers.' };
    const key = phone.replace(/\D/g, '');
    const existing = readAll_();
    for (let i = 0; i < existing.length; i++) {
      if (existing[i].phoneKey === key && existing[i].name.toLowerCase() === name.toLowerCase()) {
        return { ok: false, error: 'This name and phone number are already registered.' };
      }
    }
    const id = Utilities.getUuid().slice(0, 8) + Date.now().toString(36);
    const row = [id, new Date().toISOString(), name, String(age), clean_(r.cls, 100), phone, key,
      age < 18 ? gname : '', age < 18 ? gphone : '', clean_(r.question, 1000), clean_(r.heard, 200), 'Pending'];
    sh.getRange(sh.getLastRow() + 1, 1, 1, HEAD.length).setNumberFormat('@').setValues([row]);
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

function update_(r) {
  const status = r.status === 'Confirmed' ? 'Confirmed' : 'Pending';
  const sh = sheet_();
  const row = findRow_(sh, r.id);
  if (row < 0) return { ok: false, error: 'Not found.' };
  sh.getRange(row, HEAD.indexOf('status') + 1).setValue(status);
  return { ok: true };
}

function remove_(r) {
  const sh = sheet_();
  const row = findRow_(sh, r.id);
  if (row < 0) return { ok: false, error: 'Not found.' };
  sh.deleteRow(row);
  return { ok: true };
}
