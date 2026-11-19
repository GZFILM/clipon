/**
 * קליפון — הסקריפט שמקבל את החומרים מדף הלקוח ושומר אותם ב-Google Drive.
 *
 * מבנה התיקיות שנוצר ב-Drive:
 *   קליפון/
 *     זהר_סנדר_יום_הולדת_60/
 *       01_ינקות/001.jpg, 002.jpg ...
 *       02_גן/...
 *       NN_השלמות/
 *       AI_סרטון/
 *       מוזיקה_והערות.txt   (כל משלוח נוסף מתווסף לסוף הקובץ)
 *
 * התקנה: ראו README.md במאגר.
 */

const ROOT_NAME = 'קליפון';            // התיקייה הראשית ב-Drive
const NOTES_FILE = 'מוזיקה_והערות.txt';
const NOTIFY_BY_EMAIL = true;           // מייל אליך בכל פעם שלקוח שולח
const TZ = 'Asia/Jerusalem';

/* ---------- כניסה ---------- */

function doGet() {
  return out({ ok: true, msg: 'קליפון פועל' });
}

function doPost(e) {
  try {
    const b = JSON.parse(e.postData.contents);
    switch (b.action) {
      case 'begin':  return out(begin(b));
      case 'file':   return out(saveFile(b));
      case 'start':  return out(startUpload(b));
      case 'chunk':  return out(putChunk(b));
      case 'status': return out(uploadStatus(b));
      case 'notes':  return out(addNotes(b));
    }
    return out({ ok: false, error: 'unknown action' });
  } catch (err) {
    return out({ ok: false, error: String((err && err.message) || err) });
  }
}

function out(o) {
  return ContentService.createTextOutput(JSON.stringify(o))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ---------- תיקיות ---------- */

function clean(s) {
  return String(s || '').replace(/[\/\\]/g, '-').trim().slice(0, 120) || 'ללא_שם';
}

function rootFolder() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('ROOT');
  if (id) {
    try { const f = DriveApp.getFolderById(id); if (!f.isTrashed()) return f; } catch (e) {}
  }
  const it = DriveApp.getRootFolder().getFoldersByName(ROOT_NAME);
  const f = it.hasNext() ? it.next() : DriveApp.createFolder(ROOT_NAME);
  props.setProperty('ROOT', f.getId());
  return f;
}

function child(parent, name) {
  const it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}

function countFiles(folder) {
  let n = 0; const it = folder.getFiles();
  while (it.hasNext()) { it.next(); n++; }
  return n;
}

// Only accept folders that sit inside the Clipon root, `depth` levels down,
// so nobody can use the page to write anywhere else in your Drive.
function assertInRoot(id, depth) {
  let f = DriveApp.getFolderById(id);
  for (let i = 0; i < depth; i++) {
    const p = f.getParents();
    if (!p.hasNext()) throw new Error('bad folder');
    f = p.next();
  }
  if (f.getId() !== rootFolder().getId()) throw new Error('bad folder');
}

/* ---------- פעולות ---------- */

// Creates the client folder and all chapter folders; returns their ids and how
// many files each already holds (so new files continue the numbering).
function begin(b) {
  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    const client = child(rootFolder(), clean(b.client));
    const folders = [], counts = [];
    (b.folders || []).slice(0, 80).forEach(function (name) {
      const f = child(client, clean(name));
      folders.push(f.getId()); counts.push(countFiles(f));
    });
    return { ok: true, clientId: client.getId(), folders: folders, counts: counts };
  } finally { lock.releaseLock(); }
}

// A small file in one request.
function saveFile(b) {
  assertInRoot(b.folderId, 2);
  const blob = Utilities.newBlob(Utilities.base64Decode(b.data), b.mime || 'application/octet-stream', clean(b.name));
  const f = DriveApp.getFolderById(b.folderId).createFile(blob);
  return { ok: true, id: f.getId() };
}

// A big file (video): open a Drive resumable upload, then receive it in chunks.
function startUpload(b) {
  assertInRoot(b.folderId, 2);
  const res = UrlFetchApp.fetch('https://www.googleapis.com/upload/drive/v3/files?uploadType=resumable', {
    method: 'post',
    contentType: 'application/json; charset=UTF-8',
    headers: {
      Authorization: 'Bearer ' + ScriptApp.getOAuthToken(),
      'X-Upload-Content-Type': b.mime || 'application/octet-stream',
      'X-Upload-Content-Length': String(b.size)
    },
    payload: JSON.stringify({ name: clean(b.name), parents: [b.folderId] }),
    muteHttpExceptions: true
  });
  if (res.getResponseCode() !== 200) throw new Error('start ' + res.getResponseCode() + ' ' + res.getContentText().slice(0, 200));
  return { ok: true, session: lower(res.getAllHeaders())['location'] };
}

function putChunk(b) {
  checkSession(b.session);
  const bytes = Utilities.base64Decode(b.data);
  const res = UrlFetchApp.fetch(b.session, {
    method: 'put',
    contentType: 'application/octet-stream',
    payload: bytes,
    headers: { 'Content-Range': 'bytes ' + b.start + '-' + (b.start + bytes.length - 1) + '/' + b.size },
    muteHttpExceptions: true,
    followRedirects: false
  });
  return progress(res);
}

// Asks Drive how much of a big file already arrived (used after a dropped connection).
function uploadStatus(b) {
  checkSession(b.session);
  const res = UrlFetchApp.fetch(b.session, {
    method: 'put',
    headers: { 'Content-Range': 'bytes */' + b.size },
    muteHttpExceptions: true,
    followRedirects: false
  });
  return progress(res);
}

function progress(res) {
  const code = res.getResponseCode();
  if (code === 200 || code === 201) return { ok: true, done: true };
  if (code === 308) {
    const r = lower(res.getAllHeaders())['range'];
    const m = r && String(r).match(/bytes=0-(\d+)/);
    return { ok: true, done: false, next: m ? Number(m[1]) + 1 : 0 };
  }
  throw new Error('chunk ' + code + ' ' + res.getContentText().slice(0, 200));
}

function checkSession(s) {
  if (!/^https:\/\/www\.googleapis\.com\/upload\/drive\/v3\/files\?/.test(String(s))) throw new Error('bad session');
}

function lower(h) {
  const o = {}; Object.keys(h).forEach(function (k) { o[k.toLowerCase()] = h[k]; }); return o;
}

// Appends this send's title, songs and notes to the client's text file.
function addNotes(b) {
  assertInRoot(b.clientId, 1);
  const client = DriveApp.getFolderById(b.clientId);
  const text = String(b.text || '').slice(0, 100000);
  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    const it = client.getFilesByName(NOTES_FILE);
    if (it.hasNext()) {
      const f = it.next();
      f.setContent(f.getBlob().getDataAsString('UTF-8') + '\n\n' + text);
    } else {
      client.createFile(NOTES_FILE, text, MimeType.PLAIN_TEXT);
    }
  } finally { lock.releaseLock(); }

  if (NOTIFY_BY_EMAIL) {
    try {
      MailApp.sendEmail(Session.getEffectiveUser().getEmail(),
        'קליפון: הגיעו חומרים — ' + client.getName(),
        text + '\n\nהתיקייה ב-Drive: ' + client.getUrl() + '\n' +
        Utilities.formatDate(new Date(), TZ, 'dd/MM/yyyy HH:mm'));
    } catch (e) {}
  }
  return { ok: true };
}

// Run this once from the editor (▶) to approve Drive + email access before deploying.
function setup() {
  Logger.log('תיקיית קליפון: ' + rootFolder().getUrl());
}
