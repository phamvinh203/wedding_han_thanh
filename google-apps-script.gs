/**
 * Nhận xác nhận tham dự từ thiệp cưới và ghi vào Google Sheet.
 * Dán toàn bộ file này vào: Google Sheet > Tiện ích mở rộng > Apps Script,
 * rồi Triển khai dưới dạng "Ứng dụng web" (Thực thi: Tôi, Quyền truy cập: Bất kỳ ai).
 * Sửa code xong: Triển khai > Quản lý các bản triển khai > bút chì > Phiên bản mới.
 */
const SHEET_ID = "1UdAnADnD12gxJ7ShYpMgDgHH4GC0WJ5Qwln617YwJs4";
const HEADERS = ["Thời gian", "Tên khách", "Tham dự", "Khách của", "Số người", "Lời chúc"];

function doPost(e) {
  return withSheet(sheet => {
    const p = (e && e.parameter) || {};
    sheet.appendRow([
      new Date(),
      clean(p.name, 100),
      clean(p.attend, 20),
      clean(p.side, 20),
      clean(p.guests, 30),
      clean(p.wish, 1000),
    ]);
  });
}

// Mở URL /exec trên trình duyệt: tạo tiêu đề nếu thiếu và trả về {"ok":true}
function doGet() {
  return withSheet(() => {});
}

function withSheet(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000); // nhiều khách gửi cùng lúc không ghi đè nhau
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheets()[0];
    ensureHeaders(sheet);
    fn(sheet);
    return json({ ok: true });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

// Hàng 1 luôn là tiêu đề cột: trống thì ghi vào, đang có dữ liệu thì chèn hàng mới lên trên
function ensureHeaders(sheet) {
  const current = sheet.getRange(1, 1, 1, HEADERS.length).getValues()[0].map(v => String(v).trim());
  if (current.join("|") === HEADERS.join("|")) return;
  if (current.join("") !== "") sheet.insertRowBefore(1);
  sheet.getRange(1, 1, 1, HEADERS.length)
    .setValues([HEADERS])
    .setFontWeight("bold")
    .setBackground("#F4E6D4");
  sheet.setFrozenRows(1);
}

// Cắt độ dài và chặn chèn công thức (ô bắt đầu bằng = + - @ bị Sheets hiểu là công thức)
function clean(v, max) {
  const s = String(v || "").trim().slice(0, max);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
