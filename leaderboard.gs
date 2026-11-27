/**
 * Турецкий экран — leaderboard backend (Google Apps Script bound to a Google Sheet).
 * Deploy: Deploy > New deployment > Web app, Execute as: Me, Who has access: Anyone.
 *
 * POST (text/plain JSON) {id, name, total, week, weekKey, acc, lvl, rounds}  -> upserts one row per device id
 * GET                                                                        -> {ok, week, rows:[{name,total,week,acc,lvl,me}]}
 */
const SHEET = 'Рейтинг';
const HEAD = ['Ad', 'Toplam puan', 'Haftalık puan', 'Hafta', 'Doğruluk %', 'En çok seviye', 'Tur sayısı', 'Son güncelleme', 'id'];
const LEVELS = ['A1', 'A2', 'B1', 'B2', 'C1'];

function sheet_() {
  const ss = SpreadsheetApp.getActive();
  let s = ss.getSheetByName(SHEET);
  if (!s) {
    s = ss.insertSheet(SHEET);
    s.appendRow(HEAD);
    s.setFrozenRows(1);
    s.getRange(1, 1, 1, HEAD.length).setFontWeight('bold');
  }
  return s;
}

// ISO week key like "2026-W48", same algorithm as the site
function isoWeek_(date) {
  const d = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
  const day = d.getUTCDay() || 7;
  d.setUTCDate(d.getUTCDate() + 4 - day);
  const y0 = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
  return d.getUTCFullYear() + '-W' + String(Math.ceil(((d - y0) / 864e5 + 1) / 7)).padStart(2, '0');
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

const num_ = (v, max) => Math.max(0, Math.min(max, Math.round(Number(v) || 0)));

// Owner's C1 videos: served only from this Drive folder, as base64 JSON so the site can play them in its own player
const C1_FOLDER = '1gllPIMwQGDfAAzeVVTMdU9L3eynDbWcX';
const MAX_VIDEO_BYTES = 20 * 1024 * 1024;

function video_(rawId) {
  const id = String(rawId || '').replace(/[^A-Za-z0-9_-]/g, '');
  let f;
  try { f = DriveApp.getFileById(id); } catch (err) { return json_({ ok: false, error: 'not found' }); }
  let inFolder = false;
  const parents = f.getParents();
  while (parents.hasNext()) { if (parents.next().getId() === C1_FOLDER) { inFolder = true; break; } }
  if (!inFolder || f.isTrashed()) return json_({ ok: false, error: 'forbidden' });
  if (f.getSize() > MAX_VIDEO_BYTES) return json_({ ok: false, error: 'too big' });
  const blob = f.getBlob();
  return json_({ ok: true, mime: blob.getContentType() || 'video/mp4', b64: Utilities.base64Encode(blob.getBytes()) });
}

function doGet(e) {
  if (e && e.parameter && e.parameter.video) return video_(e.parameter.video);
  const rows = sheet_().getDataRange().getValues().slice(1);
  const wk = isoWeek_(new Date());
  const me = String((e && e.parameter && e.parameter.me) || '');
  const out = rows
    .filter(r => String(r[0]).trim() !== '')
    .map(r => ({
      name: String(r[0]).replace(/^'/, ''),
      total: Number(r[1]) || 0,
      week: String(r[3]) === wk ? (Number(r[2]) || 0) : 0,
      acc: Number(r[4]) || 0,
      lvl: String(r[5] || ''),
      me: me !== '' && String(r[8]) === me
    }));
  return json_({ ok: true, week: wk, rows: out });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const d = JSON.parse(e.postData.contents);
    const id = String(d.id || '').replace(/[^A-Za-z0-9]/g, '').slice(0, 32);
    let name = String(d.name || '').replace(/[\u0000-\u001f<>]/g, '').trim().slice(0, 30);
    if (id.length < 8 || !name) return json_({ ok: false, error: 'bad input' });
    if (/^[=+\-@]/.test(name)) name = "'" + name; // never let a name become a spreadsheet formula

    const wk = isoWeek_(new Date());
    const row = [
      name,
      num_(d.total, 1e7),
      String(d.weekKey) === wk ? num_(d.week, 1e6) : 0,
      wk,
      num_(d.acc, 100),
      LEVELS.indexOf(String(d.lvl)) >= 0 ? String(d.lvl) : '',
      num_(d.rounds, 1e6),
      new Date(),
      id
    ];
    const s = sheet_();
    const ids = s.getRange(2, 9, Math.max(1, s.getLastRow() - 1), 1).getValues().map(r => String(r[0]));
    const i = ids.indexOf(id);
    if (i >= 0) s.getRange(i + 2, 1, 1, row.length).setValues([row]);
    else s.appendRow(row);
    return json_({ ok: true });
  } finally {
    lock.releaseLock();
  }
}

// Run once from the editor to create the "Рейтинг" sheet and grant access
function setup() { sheet_(); }

// Run once from the editor to grant Drive read access (needed for C1 videos)
function authDrive() { DriveApp.getFolderById(C1_FOLDER).getName(); }
