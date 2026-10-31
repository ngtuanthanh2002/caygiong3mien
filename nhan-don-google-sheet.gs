/**
 * SCRIPT LƯU ĐƠN HÀNG VÀO GOOGLE SHEET
 * ----------------------------------------------------------
 * Sheet đích:
 * https://docs.google.com/spreadsheets/d/1QGYlJHLa-f5O_VCSpZ5y-x3UlUTP1oCIFus5sNEzQfU/edit
 *
 * CÁCH DÙNG:
 * 1. Mở Google Sheet ở link trên.
 * 2. Vào: Tiện ích mở rộng (Extensions) > Apps Script.
 * 3. Xoá code mẫu, dán toàn bộ nội dung file này vào.
 * 4. Bấm "Lưu" (biểu tượng đĩa mềm) rồi đặt tên dự án, ví dụ: "Nhan don cay giong".
 * 5. Bấm "Triển khai" (Deploy) > "Triển khai mới" (New deployment).
 *      - Loại: Ứng dụng web (Web app)
 *      - Thực thi với tư cách (Execute as): Tôi (Me)
 *      - Ai có quyền truy cập (Who has access): Bất kỳ ai (Anyone)
 * 6. Bấm "Triển khai", cấp quyền cho script (Allow).
 * 7. Copy "URL ứng dụng web" (Web app URL) rồi dán vào biến ORDER_URL
 *    trong file index.html (dòng: var ORDER_URL='...').
 *
 * MẸO:
 * - Nếu đổi code sau này: Triển khai > Quản lý triển khai > sửa > "Phiên bản mới".
 * - Có thể chạy hàm testThem() trong Apps Script để kiểm tra ghi dữ liệu.
 */

// ID của Google Sheet (lấy từ URL giữa /d/ và /edit)
var SHEET_ID = '1QGYlJHLa-f5O_VCSpZ5y-x3UlUTP1oCIFus5sNEzQfU';

// Tên tab (sheet con) sẽ ghi đơn. Nếu chưa có sẽ tự tạo.
var SHEET_NAME = 'DonHang';

// Tiêu đề các cột
var HEADERS = ['Thời gian', 'Họ và tên', 'Số điện thoại', 'Địa chỉ', 'Gói cây', 'Phân kích rễ', 'Tổng tiền (₫)'];

/**
 * Trả về sheet đích; tự tạo tab và dòng tiêu đề nếu chưa có.
 */
function laySheet() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
  }
  if (sh.getLastRow() === 0) {
    sh.appendRow(HEADERS);
    sh.getRange(1, 1, 1, HEADERS.length)
      .setFontWeight('bold')
      .setBackground('#ff6b7a')
      .setFontColor('#ffffff');
    sh.setFrozenRows(1);
    sh.setColumnWidth(1, 150); // Thời gian
    sh.setColumnWidth(2, 180); // Họ tên
    sh.setColumnWidth(3, 130); // SĐT
    sh.setColumnWidth(4, 300); // Địa chỉ
    sh.setColumnWidth(5, 180); // Gói
    sh.setColumnWidth(6, 150); // Phân
    sh.setColumnWidth(7, 130); // Tổng
  }
  return sh;
}

/**
 * Nhận POST từ website và ghi 1 dòng xuống sheet.
 * Body JSON: { ten, sdt, diachi, goi, phan, tong, luc }
 *
 * Luôn THÊM đơn mới vào dòng trống kế tiếp sau đơn cũ, không ghi đè.
 * Dùng LockService để nhiều đơn cùng lúc vẫn ghi tuần tự, không đè nhau.
 */
function doPost(e) {
  var ok = true;
  var msg = 'OK';
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000); // chờ tối đa 30 giây nếu đang có đơn khác ghi

    var d = JSON.parse(e.postData.contents);

    // Chống trường hợp thiếu dữ liệu
    var ten    = (d.ten    || '').toString().trim();
    var sdt    = (d.sdt    || '').toString().trim();
    var diachi = (d.diachi || '').toString().trim();
    var goi    = (d.goi    || '').toString().trim();
    var phan   = (d.phan   || '').toString().trim();
    var tong   = Number(d.tong) || 0;
    var luc    = d.luc || new Date().toLocaleString('vi-VN');

    if (!ten || !sdt || !diachi) {
      throw new Error('Thiếu thông tin bắt buộc');
    }

    var sh = laySheet();
    // Tính rõ dòng kế tiếp dựa trên dòng cuối hiện có → chắc chắn ghi sau đơn cũ.
    var dongMoi = sh.getLastRow() + 1;
    sh.getRange(dongMoi, 1, 1, 7).setValues([
      [luc, ten, "'" + sdt, diachi, goi, phan, tong]
      // Dấu nháy đơn trước số điện thoại để Google Sheet không cắt số 0 đầu.
    ]);
    SpreadsheetApp.flush(); // ép Google Sheet lưu ngay, tránh mất đơn khi deploy restart
  } catch (err) {
    ok = false;
    msg = err && err.message ? err.message : 'Lỗi không xác định';
  } finally {
    try { lock.releaseLock(); } catch (ignore) {}
  }
  return ContentService
    .createTextOutput(JSON.stringify({ ok: ok, msg: msg }))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Mở link Web App bằng trình duyệt sẽ trả về dòng này -
 * dùng để kiểm tra deploy đã sống chưa.
 */
function doGet() {
  return ContentService
    .createTextOutput(JSON.stringify({ ok: true, msg: 'Web App dang chay' }))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Chạy thử trong Apps Script để chắc chắn ghi được xuống sheet.
 * Nhấn nút ▶ (Run) sau khi chọn hàm testThem.
 */
function testThem() {
  var sh = laySheet();
  sh.appendRow([
    new Date().toLocaleString('vi-VN'),
    'Nguyễn Văn Test',
    "'0912345678",
    '123 Đường Test, Quận 1, TP.HCM',
    '6 cây',
    'Không mua kèm',
    150000
  ]);
}
