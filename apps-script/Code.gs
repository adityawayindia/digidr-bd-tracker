/**
 * DigiDr BD Tracker — Google Sheets backend (Apps Script web app).
 *
 * The Sheet is the single source of truth. Tabs (created by setup()):
 *   Doctors   one row per doctor, updated in place
 *   Calls     one row per logged call
 *   Plan      notes for a given day's plan
 *   Pitch     key/value content for the "About DigiDr" screen
 *   Activity  append-only audit log of every change (shown in "Sheet Log")
 *
 * See README.md in this folder for deployment steps.
 */

// Optional shared key. If set, the app must be configured with the same key.
const TEAM_KEY = '';
const TZ = 'Asia/Kolkata';

const SCHEMA = {
  Doctors:  ['id', 'name', 'email', 'mobile', 'dept', 'status', 'followup', 'remarks',
             'addedBy', 'addedByEmail', 'addedOn', 'updatedBy', 'updatedAt'],
  Calls:    ['id', 'date', 'time', 'docId', 'docName', 'dept', 'outcome', 'notes', 'next',
             'by', 'byEmail', 'byPhone'],
  Plan:     ['id', 'forDate', 'text', 'by', 'createdAt'],
  Pitch:    ['key', 'value', 'updatedBy', 'updatedAt'],
  Activity: ['timestamp', 'by', 'action', 'details']
};
const PITCH_KEYS = ['what', 'why', 'steps', 'time', 'script'];
const AUDIT_FIELDS = ['updatedBy', 'updatedAt'];

// ---------- one-time setup (run from the Apps Script editor) ----------
function setup() {
  Object.keys(SCHEMA).forEach(getSheet_);
  const n = importFormResponses_();
  Logger.log('Setup done. Imported ' + n + ' row(s) from Google Form responses.');
}

// ---------- web app entry points ----------
function doGet() {
  return json_({ ok: true, result: 'DigiDr BD Tracker API is running' });
}

function doPost(e) {
  let req;
  try {
    req = JSON.parse(e.postData.contents);
  } catch (err) {
    return json_({ ok: false, error: 'Bad request' });
  }
  if (TEAM_KEY && req.key !== TEAM_KEY) return json_({ ok: false, error: 'Invalid team key' });

  const isRead = req.action === 'getAll';
  const lock = LockService.getScriptLock();
  let locked = false;
  if (!isRead) {
    // Writes run one at a time so simultaneous edits can't interleave. If the sheet
    // is busy for too long the app is told to retry, not to discard the change.
    try {
      lock.waitLock(25000);
      locked = true;
    } catch (err) {
      return json_({ ok: false, retry: true, error: 'Sheet is busy, will retry' });
    }
  }
  try {
    return json_({ ok: true, result: handle_(req.action, req.data || {}, req.by || '') });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  } finally {
    if (locked) lock.releaseLock();
  }
}

function handle_(action, data, by) {
  switch (action) {
    case 'getAll':       return getAll_();
    case 'saveDoctor':   return saveDoctor_(data, by);
    case 'deleteDoctor': return deleteDoctor_(data, by);
    case 'logCall':      return logCall_(data, by);
    case 'addPlan':      return addPlan_(data, by);
    case 'deletePlan':   return deletePlan_(data, by);
    case 'savePitch':    return savePitch_(data, by);
    default: throw new Error('Unknown action: ' + action);
  }
}

// ---------- actions ----------
function getAll_() {
  const pitch = { meta: { by: '', at: '' } };
  readRows_('Pitch').forEach(r => {
    pitch[r.key] = r.value;
    if (r.updatedAt > pitch.meta.at) pitch.meta = { by: r.updatedBy, at: r.updatedAt };
  });
  return {
    doctors: readRows_('Doctors'),
    calls: readRows_('Calls'),
    plan: readRows_('Plan'),
    pitch: pitch,
    activity: readRows_('Activity').slice(-300).reverse(),
    sheetUrl: SpreadsheetApp.getActive().getUrl()
  };
}

function saveDoctor_(d, by) {
  if (!d.id) throw new Error('Doctor id missing');
  const sh = getSheet_('Doctors');
  // An edit (a patch without "addedBy") for a doctor someone deleted meanwhile
  // must not bring back a half-empty row. Only a full "add" payload creates a row.
  const exists = findRow_(sh, d.id);
  if (!exists && !d.addedBy) return null;
  if (!exists) {
    // Checked against the sheet itself, so a doctor a teammate just added is caught too.
    const digits = s => String(s || '').replace(/\D/g, '');
    const mobile = digits(d.mobile);
    const email = String(d.email || '').trim().toLowerCase();
    const dup = readRows_('Doctors').find(r =>
      (mobile && digits(r.mobile) === mobile) ||
      (email && String(r.email || '').trim().toLowerCase() === email));
    if (dup) throw new Error(dup.name + ' already exists in the sheet (same phone/email), so it was not added again');
  }
  d.updatedBy = by;
  d.updatedAt = nowStr_();
  const res = upsert_('Doctors', d);
  if (res.created) {
    log_(by, 'Doctor added', res.row.name + ' (' + (res.row.dept || '—') + ') · ' + (res.row.status || 'New'));
  } else if (res.changes.length) {
    log_(by, 'Doctor updated', res.row.name + ': ' + res.changes.join('; '));
  }
  return res.row;
}

function deleteDoctor_(d, by) {
  const sh = getSheet_('Doctors');
  const idx = findRow_(sh, d.id);
  if (!idx) return false;
  const name = sh.getRange(idx, SCHEMA.Doctors.indexOf('name') + 1).getValue();
  sh.deleteRow(idx);
  log_(by, 'Doctor deleted', String(name));
  return true;
}

function logCall_(data, by) {
  const c = data.call;
  if (!c || !c.id) throw new Error('Call id missing');
  const res = upsert_('Calls', c);
  const doc = data.doctor;
  if (doc && doc.id && findRow_(getSheet_('Doctors'), doc.id)) {
    doc.updatedBy = by;
    doc.updatedAt = nowStr_();
    upsert_('Doctors', doc);
  }
  if (res.created) {
    log_(by, 'Call logged', c.docName + ': ' + c.outcome + (c.next ? ' · next action ' + c.next : ''));
  }
  return true;
}

function addPlan_(p, by) {
  if (!p.id) throw new Error('Plan id missing');
  p.createdAt = p.createdAt || nowStr_();
  const res = upsert_('Plan', p);
  if (res.created) log_(by, 'Plan note added', 'For ' + p.forDate + ': ' + p.text);
  return true;
}

function deletePlan_(p, by) {
  const sh = getSheet_('Plan');
  const idx = findRow_(sh, p.id);
  if (!idx) return false;
  const text = sh.getRange(idx, SCHEMA.Plan.indexOf('text') + 1).getValue();
  sh.deleteRow(idx);
  log_(by, 'Plan note removed', String(text));
  return true;
}

function savePitch_(p, by) {
  const now = nowStr_();
  PITCH_KEYS.forEach(k => {
    let v = p[k];
    if (Array.isArray(v)) v = v.join('\n');
    upsert_('Pitch', { key: k, value: v || '', updatedBy: by, updatedAt: now });
  });
  log_(by, 'Pitch updated', 'About DigiDr content edited');
  return true;
}

// ---------- sheet helpers ----------
function getSheet_(name) {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    const headers = SCHEMA[name];
    sh = ss.insertSheet(name);
    // Plain-text format so dates/phones are stored exactly as sent.
    sh.getRange(1, 1, sh.getMaxRows(), headers.length).setNumberFormat('@');
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function readRows_(name) {
  const sh = getSheet_(name);
  const headers = SCHEMA[name];
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, headers.length).getValues()
    .filter(r => r.some(v => v !== ''))
    .map(r => {
      const o = {};
      headers.forEach((h, i) => { o[h] = cell_(r[i]); });
      return o;
    });
}

function findRow_(sh, id) {
  if (!id) return 0;
  const last = sh.getLastRow();
  if (last < 2) return 0;
  const ids = sh.getRange(2, 1, last - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) return i + 2;
  }
  return 0;
}

// Insert or merge-update a row keyed by the first column. Only fields present
// in `obj` are changed, so two people editing different fields don't clobber each other.
function upsert_(name, obj) {
  const sh = getSheet_(name);
  const headers = SCHEMA[name];
  const idx = findRow_(sh, obj[headers[0]]);
  const current = {};
  if (idx) {
    const vals = sh.getRange(idx, 1, 1, headers.length).getValues()[0];
    headers.forEach((h, i) => { current[h] = cell_(vals[i]); });
  }
  const changes = [];
  const row = {};
  headers.forEach(h => {
    const has = Object.prototype.hasOwnProperty.call(obj, h) && obj[h] != null;
    row[h] = has ? String(obj[h]) : (current[h] || '');
    if (idx && has && row[h] !== (current[h] || '') && AUDIT_FIELDS.indexOf(h) < 0) {
      changes.push(h + ': ' + short_(current[h]) + ' → ' + short_(row[h]));
    }
  });
  const values = [headers.map(h => safe_(row[h]))];
  const target = idx || sh.getLastRow() + 1;
  sh.getRange(target, 1, 1, headers.length).setNumberFormat('@').setValues(values);
  return { created: !idx, row: row, changes: changes };
}

function log_(by, action, details) {
  const sh = getSheet_('Activity');
  sh.getRange(sh.getLastRow() + 1, 1, 1, 4).setNumberFormat('@')
    .setValues([[nowStr_(), safe_(by || '—'), action, safe_(details || '')]]);
}

// ---------- import of the old Google Form responses ----------
function importFormResponses_() {
  const ss = SpreadsheetApp.getActive();
  let imported = 0;
  ss.getSheets().forEach(sh => {
    if (SCHEMA[sh.getName()]) return;
    const values = sh.getDataRange().getValues();
    if (values.length < 2) return;
    const headers = values[0].map(String);
    if (headers.indexOf('Doctor Name') < 0) return;
    const get = (r, h) => headers.indexOf(h) >= 0 ? cell_(r[headers.indexOf(h)]) : '';

    values.slice(1).forEach((r, i) => {
      const name = get(r, 'Doctor Name');
      if (!name) return;
      const ts = r[headers.indexOf('Timestamp')];
      const day = isDate_(ts) ? Utilities.formatDate(ts, TZ, 'yyyy-MM-dd') : '';
      const id = 'form-' + sh.getSheetId() + '-' + (i + 2);

      if (get(r, 'Call Outcome')) {
        if (findRow_(getSheet_('Calls'), id)) return;
        const doc = readRows_('Doctors').filter(d => d.name.toLowerCase() === name.toLowerCase())[0];
        upsert_('Calls', {
          id: id, date: day, time: isDate_(ts) ? ts.toISOString() : '',
          docId: doc ? doc.id : '', docName: name, dept: get(r, 'Speciality'),
          outcome: get(r, 'Call Outcome'), notes: get(r, 'Call Notes'), next: get(r, 'Next Action Date'),
          by: get(r, 'BD Name'), byEmail: get(r, 'BD Email'), byPhone: get(r, 'BD Phone')
        });
      } else {
        if (findRow_(getSheet_('Doctors'), id)) return;
        upsert_('Doctors', {
          id: id, name: name, email: get(r, 'Doctor Email'), mobile: get(r, 'Doctor Mobile'),
          dept: get(r, 'Speciality'), status: get(r, 'Status') || 'New',
          followup: get(r, 'Follow-up Date'), remarks: get(r, 'Remarks'),
          addedBy: get(r, 'BD Name'), addedByEmail: get(r, 'BD Email'), addedOn: day
        });
      }
      imported++;
    });
  });
  return imported;
}

// ---------- small utils ----------
function cell_(v) {
  if (isDate_(v)) return Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
  const s = v == null ? '' : String(v);
  return s.charAt(0) === "'" ? s.slice(1) : s;
}

// Prevent user text from being interpreted as a formula.
function safe_(s) {
  s = s == null ? '' : String(s);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function short_(s) {
  s = s || '—';
  return s.length > 60 ? s.slice(0, 57) + '…' : s;
}

function nowStr_() {
  return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss');
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function isDate_(v) {
  return Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v);
}
