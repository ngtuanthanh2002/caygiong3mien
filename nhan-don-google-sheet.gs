// Dán vào Google Sheet: Tiện ích mở rộng > Apps Script
// Triển khai > Triển khai mới > Ứng dụng web > Thực thi với tư cách: Tôi > Ai có quyền truy cập: Bất kỳ ai
// Copy link Web App dán vào biến ORDER_URL trong index.html
function doPost(e) {
  var ok = true;
  try {
    var d = JSON.parse(e.postData.contents);
    SpreadsheetApp.getActiveSheet().appendRow([d.luc, d.ten, d.sdt, d.diachi, d.goi, d.phan, d.tong]);
  } catch (err) {
    ok = false;
  }
  return ContentService.createTextOutput(JSON.stringify({ ok: ok }))
    .setMimeType(ContentService.MimeType.JSON);
}
