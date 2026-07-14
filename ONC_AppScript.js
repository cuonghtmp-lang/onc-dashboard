// ══════════════════════════════════════════════════════════════════
// ONC BACKUP — Google Apps Script
// Lưu file JSON vào Google Drive (thư mục "ONC Backup")
// Không dùng Google Sheets
// ══════════════════════════════════════════════════════════════════

const FOLDER_NAME = 'ONC Backup';
const SECRET_KEY  = 'onc2026';
const LATEST_FILE = 'ONC_latest.json';
const MAX_BACKUPS = 30; // giữ tối đa 30 bản backup

// ── POST: nhận dữ liệu từ app, lưu vào Drive ─────────────────────
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    if (body.key !== SECRET_KEY)
      return json({ok: false, error: 'Unauthorized'});

    const db   = body.data;
    const ts   = new Date().toISOString().replace('T',' ').slice(0,19);
    const folder = getOrCreateFolder(FOLDER_NAME);

    // File "latest" — ghi đè để lấy bản mới nhất
    setOrCreateFile(folder, LATEST_FILE, JSON.stringify(db));

    // File backup có timestamp
    const fname = `ONC_${ts.replace(/[: ]/g,'-')}.json`;
    folder.createFile(fname, JSON.stringify(db, null, 1), MimeType.PLAIN_TEXT);

    cleanOldFiles(folder, MAX_BACKUPS);

    return json({ok: true, timestamp: ts, file: fname});
  } catch(err) {
    return json({ok: false, error: err.toString()});
  }
}

// ── GET: health check hoặc trả dữ liệu mới nhất ──────────────────
function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || '';

  if (action === 'load') {
    if ((e.parameter.key||'') !== SECRET_KEY)
      return json({ok: false, error: 'Unauthorized'});

    const folder = getOrCreateFolder(FOLDER_NAME);
    const files  = folder.getFilesByName(LATEST_FILE);
    if (!files.hasNext())
      return json({ok: false, error: 'No backup yet'});

    const content = files.next().getBlob().getDataAsString();
    return json({ok: true, data: JSON.parse(content)});
  }

  return json({ok: true, service: 'ONC Drive Backup', version: '3.0'});
}

// ── Tiện ích ──────────────────────────────────────────────────────
function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function getOrCreateFolder(name) {
  const it = DriveApp.getFoldersByName(name);
  return it.hasNext() ? it.next() : DriveApp.createFolder(name);
}

function setOrCreateFile(folder, name, content) {
  const it = folder.getFilesByName(name);
  if (it.hasNext()) { it.next().setContent(content); return; }
  folder.createFile(name, content, MimeType.PLAIN_TEXT);
}

function cleanOldFiles(folder, keepMax) {
  const files = [], it = folder.getFiles();
  while (it.hasNext()) {
    const f = it.next();
    if (f.getName() !== LATEST_FILE) files.push(f);
  }
  files.sort((a,b) => b.getDateCreated()-a.getDateCreated());
  files.slice(keepMax).forEach(f => f.setTrashed(true));
}
