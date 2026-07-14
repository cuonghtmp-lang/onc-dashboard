// ══════════════════════════════════════════════════════════════════
// ONC SYNC — Google Apps Script v2.0
// Hỗ trợ đồng bộ 2 chiều: save (POST) + load (GET)
// ══════════════════════════════════════════════════════════════════

const FOLDER_NAME   = 'ONC Backup';
const SHEET_NAME    = 'ONC Dashboard Backup';
const SECRET_KEY    = 'onc2026'; // Đổi khóa bảo vệ endpoint nếu muốn
const LATEST_FILE   = 'ONC_latest.json'; // File luôn ghi đè = dữ liệu mới nhất

// ── Hàm nhận POST từ app ONC (save) ─────────────────────────────
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);

    // Xác thực khóa bí mật
    if (body.key !== SECRET_KEY) {
      return ContentService.createTextOutput(
        JSON.stringify({ok: false, error: 'Unauthorized'})
      ).setMimeType(ContentService.MimeType.JSON);
    }

    const db = body.data; // Toàn bộ DB từ app
    const ts = new Date().toISOString().replace('T',' ').slice(0,19);
    const folder = getOrCreateFolder(FOLDER_NAME);

    // 1. Lưu file latest (ghi đè) — dùng để đồng bộ thiết bị khác
    const latestFiles = folder.getFilesByName(LATEST_FILE);
    if (latestFiles.hasNext()) {
      latestFiles.next().setContent(JSON.stringify(db));
    } else {
      folder.createFile(LATEST_FILE, JSON.stringify(db), MimeType.PLAIN_TEXT);
    }

    // 2. Lưu file backup có timestamp
    const filename = `ONC_backup_${ts.replace(/[: ]/g,'-')}.json`;
    folder.createFile(filename, JSON.stringify(db, null, 2), MimeType.PLAIN_TEXT);

    // Xóa file backup cũ hơn 30 ngày (không xóa latest)
    cleanOldFiles(folder, 30);

    // 3. Cập nhật Google Sheet tóm tắt
    updateSummarySheet(db, ts);

    return ContentService.createTextOutput(
      JSON.stringify({ok: true, timestamp: ts, filename: filename})
    ).setMimeType(ContentService.MimeType.JSON);

  } catch(err) {
    return ContentService.createTextOutput(
      JSON.stringify({ok: false, error: err.toString()})
    ).setMimeType(ContentService.MimeType.JSON);
  }
}

// ── Cập nhật Google Sheet tóm tắt ────────────────────────────────
function updateSummarySheet(db, ts) {
  let ss;
  const files = DriveApp.getFilesByName(SHEET_NAME);
  if (files.hasNext()) {
    ss = SpreadsheetApp.open(files.next());
  } else {
    ss = SpreadsheetApp.create(SHEET_NAME);
  }
  
  // Sheet 1: Lịch sử backup
  let shLog = ss.getSheetByName('Lịch sử');
  if (!shLog) shLog = ss.insertSheet('Lịch sử');
  if (shLog.getLastRow() === 0) {
    shLog.appendRow(['Thời gian', 'Đơn nhập', 'Dòng hàng', 'Sản phẩm', 'Thu chi', 'Bán hàng']);
  }
  const orders  = (db.orders  || []).length;
  const lines   = (db.lines   || []).length;
  const prods   = (db.products|| []).length;
  const cash    = (db.cash    || []).length;
  const sales   = (db.sales   || []).length;
  shLog.appendRow([ts, orders, lines, prods, cash, sales]);
  
  // Sheet 2: Tồn kho (sản phẩm)
  let shProd = ss.getSheetByName('Sản phẩm');
  if (!shProd) shProd = ss.insertSheet('Sản phẩm');
  shProd.clearContents();
  shProd.appendRow(['Mã SP', 'Tên VI', 'Tên xuất HĐ', 'ĐVT', 'Tiền tệ', 'VAT bán %', 'Tồn TT']);
  (db.products || []).forEach(p => {
    shProd.appendRow([p.code, p.groupVI||'', p.invoiceName||'', p.unit||'', p.currency||'', (p.vatSale||0)*100, p.minStock||0]);
  });
  
  // Sheet 3: Thu chi tiền
  let shCash = ss.getSheetByName('Thu chi');
  if (!shCash) shCash = ss.insertSheet('Thu chi');
  shCash.clearContents();
  shCash.appendRow(['Ngày', 'Loại', 'Mã đối tác', 'Thu vào', 'Chi ra', 'Nội dung', 'Đơn nhập']);
  (db.cash || []).forEach(c => {
    shCash.appendRow([c.date||'', c.objType||'', c.objCode||'', c.cashIn||0, c.cashOut||0, c.content||'', c.order||'']);
  });
  
  // Sheet 4: Đơn nhập
  let shOrd = ss.getSheetByName('Đơn nhập');
  if (!shOrd) shOrd = ss.insertSheet('Đơn nhập');
  shOrd.clearContents();
  shOrd.appendRow(['Mã đơn', 'NCC', 'Ngày PI', 'Ngày nhập kho', 'Freight', 'Ghi chú']);
  (db.orders || []).forEach(o => {
    shOrd.appendRow([o.code||'', o.supplier||'', o.datePI||'', o.dateWH||'', o.freight||0, o.note||'']);
  });
  
  // Sheet 5: Bán hàng
  let shSales = ss.getSheetByName('Bán hàng');
  if (!shSales) shSales = ss.insertSheet('Bán hàng');
  shSales.clearContents();
  shSales.appendRow(['Ngày', 'Sản phẩm', 'Khách hàng', 'Số lượng', 'Đơn giá', 'Ghi chú']);
  (db.sales || []).forEach(s => {
    shSales.appendRow([s.date||'', s.product||'', s.customer||'', s.qty||0, s.price||0, s.note||'']);
  });
}

// ── Tiện ích ──────────────────────────────────────────────────────
function getOrCreateFolder(name) {
  const folders = DriveApp.getFoldersByName(name);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(name);
}

function cleanOldFiles(folder, days) {
  const cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - days);
  const files = folder.getFiles();
  while (files.hasNext()) {
    const f = files.next();
    if (f.getDateCreated() < cutoff) f.setTrashed(true);
  }
}

// ── Backup theo lịch (trigger tự động) ───────────────────────────
// Hàm này chạy theo lịch nếu anh muốn nhắc từ script
// (không dùng nếu chỉ backup từ app)
function scheduledReminder() {
  const email = Session.getActiveUser().getEmail();
  GmailApp.sendEmail(email, 
    '[ONC] Nhắc backup dữ liệu', 
    'Đã ' + 7 + ' ngày kể từ lần backup cuối. Vào app ONC → Sao lưu để backup ngay.'
  );
}

// ── GET: health check hoặc load dữ liệu mới nhất ────────────────
function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || '';

  if (action === 'load') {
    // Xác thực
    const key = e.parameter.key || '';
    if (key !== SECRET_KEY) {
      return ContentService.createTextOutput(
        JSON.stringify({ok: false, error: 'Unauthorized'})
      ).setMimeType(ContentService.MimeType.JSON);
    }
    // Trả về dữ liệu mới nhất
    const folder = getOrCreateFolder(FOLDER_NAME);
    const files = folder.getFilesByName(LATEST_FILE);
    if (!files.hasNext()) {
      return ContentService.createTextOutput(
        JSON.stringify({ok: false, error: 'No data yet'})
      ).setMimeType(ContentService.MimeType.JSON);
    }
    const content = files.next().getBlob().getDataAsString();
    return ContentService.createTextOutput(
      JSON.stringify({ok: true, data: JSON.parse(content)})
    ).setMimeType(ContentService.MimeType.JSON);
  }

  // Health check
  return ContentService.createTextOutput(
    JSON.stringify({ok: true, service: 'ONC Sync', version: '2.0'})
  ).setMimeType(ContentService.MimeType.JSON);
}
