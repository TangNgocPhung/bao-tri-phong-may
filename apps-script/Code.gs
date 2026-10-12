/**
 * Backend ghi lỗi máy tính — Google Apps Script gắn với Google Sheet.
 *
 * Cách cài:
 *  1. Mở Google Sheet → Tiện ích mở rộng → Apps Script, dán toàn bộ file này vào.
 *  2. Chạy hàm taoSheetMau() một lần (cấp quyền khi được hỏi) để tạo 2 trang tính.
 *  3. Điền tài khoản giáo viên vào trang "GiaoVien" (taiKhoan | matKhau | hoTen).
 *  4. Triển khai → Ứng dụng web: "Thực thi bằng: Tôi", "Ai có quyền truy cập: Bất kỳ ai".
 *  5. Dán URL /exec vào data/cau-hinh.js (biến API_URL).
 *
 * Mỗi lần sửa code phải Triển khai → Quản lý bản triển khai → Phiên bản mới.
 *
 * Mật khẩu trong sheet có thể để chữ thường, hoặc dạng "sha256:<mã băm>" (tạo bằng hàm bamMatKhau).
 * Nên chỉ chia sẻ Sheet cho người quản lý vì trang "GiaoVien" chứa mật khẩu.
 */

var SHEET_GV = "GiaoVien";
var SHEET_LOI = "LoiMay";
var TT_CHUA = "Chưa xử lý";
var TT_XONG = "Đã xong";
var TOKEN_TTL = 6 * 60 * 60; // giây (tối đa của CacheService)

// Số máy tối đa của từng phòng — phải khớp data/phong-may.js
var PHONG = { PM1: 48, PM2: 50, PM3: 50, PM4: 60, PM5: 48 };

// Giáo viên được chọn ở ô "Giáo viên ghi nhận" — phải khớp ghi-loi-may.html
var GIAO_VIEN = [
  "Bùi Thị Thanh Nhàn",
  "Đỗ Như Hảo",
  "Nguyễn Thị Kim Ngân",
  "Võ Thị Châu Tiền",
  "Cao Hoàng Thiên Ý",
  "Nguyễn Thị Kim Phụng",
  "Tăng Ngọc Phụng",
  "Nguyễn Thanh Hiền"
];

function taoSheetMau() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var gv = ss.getSheetByName(SHEET_GV) || ss.insertSheet(SHEET_GV);
  if (gv.getLastRow() === 0) {
    gv.appendRow(["taiKhoan", "matKhau", "hoTen", "vaiTro"]);
    gv.appendRow(["gv01", "doi-mat-khau", "Nguyễn Văn A", "admin"]);
  }
  var loi = ss.getSheetByName(SHEET_LOI) || ss.insertSheet(SHEET_LOI);
  if (loi.getLastRow() === 0) {
    loi.appendRow(["thoiGianGhi", "phong", "may", "loi", "ngay", "giaoVien", "taiKhoan", "id", "trangThai", "nguoiXuLy", "hinhAnh"]);
    loi.getRange("E:E").setNumberFormat("@"); // giữ ngày dạng yyyy-mm-dd
  }
  chuanHoaSheet_(loi);
}

// Chạy một lần sau khi cập nhật để tạo cột ảnh và sửa tên cũ.
function capNhatAnhMayHu() {
  taoSheetMau();
  thuMucAnh_();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  [ [SHEET_GV, 3], [SHEET_LOI, 6] ].forEach(function (entry) {
    var sheet = ss.getSheetByName(entry[0]);
    if (sheet.getLastRow() < 2) return;
    sheet.getRange(2, entry[1], sheet.getLastRow() - 1, 1)
      .createTextFinder("Nguyễn Thị Kim Phung").matchEntireCell(true)
      .replaceAllWith("Nguyễn Thị Kim Phụng");
  });
}

function thuMucAnh_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty("THU_MUC_ANH_MAY_HU");
  if (id) return DriveApp.getFolderById(id);
  var folder = DriveApp.createFolder("Ảnh máy tính bị hư - " + SpreadsheetApp.getActiveSpreadsheet().getName());
  props.setProperty("THU_MUC_ANH_MAY_HU", folder.getId());
  return folder;
}

function luuAnh_(image, phong, may) {
  if (!image) return null;
  if (image.mimeType !== "image/jpeg" || typeof image.data !== "string" ||
      !image.data || image.data.length > Math.ceil(4 * 1024 * 1024 / 3) * 4 ||
      !/^[A-Za-z0-9+/]+={0,2}$/.test(image.data)) {
    throw new Error("Ảnh phải ở định dạng JPEG, tối đa 4 MB.");
  }
  var bytes = Utilities.base64Decode(image.data);
  if (bytes.length > 4 * 1024 * 1024 || bytes.length < 3 ||
      (bytes[0] & 255) !== 255 || (bytes[1] & 255) !== 216 || (bytes[2] & 255) !== 255) {
    throw new Error("Dữ liệu ảnh không hợp lệ.");
  }
  var name = phong + "-may-" + may + "-" + Utilities.getUuid() + ".jpg";
  return thuMucAnh_().createFile(Utilities.newBlob(bytes, "image/jpeg", name));
}

// Dùng trong trình soạn thảo: bamMatKhau("abc123") rồi xem nhật ký để lấy chuỗi dán vào sheet.
function bamMatKhau(matKhau) {
  Logger.log(bam_(matKhau));
}

function doPost(e) {
  var out;
  try {
    var req = JSON.parse(e.postData.contents);
    switch (req.action) {
      case "login": out = dangNhap_(req); break;
      case "logout": CacheService.getScriptCache().remove("t_" + req.token); out = { ok: true }; break;
      case "submit": out = ghiLoi_(req); break;
      case "list": out = layDanhSach_(req); break;
      case "resolve": out = doiTrangThai_(req); break;
      case "delete": out = xoaLoi_(req); break;
      default: out = { ok: false, error: "Yêu cầu không hợp lệ." };
    }
  } catch (err) {
    out = { ok: false, error: "Lỗi máy chủ: " + err.message };
  }
  return ContentService.createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet() {
  return ContentService.createTextOutput("API ghi lỗi máy đang chạy.");
}

function bam_(text) {
  var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, text, Utilities.Charset.UTF_8);
  return "sha256:" + bytes.map(function (b) {
    return ("0" + (b & 0xff).toString(16)).slice(-2);
  }).join("");
}

function dangNhap_(req) {
  var taiKhoan = String(req.taiKhoan || "").trim().toLowerCase();
  var matKhau = String(req.matKhau || "");
  var rows = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_GV).getDataRange().getValues();
  for (var i = 1; i < rows.length; i += 1) {
    var tk = String(rows[i][0]).trim().toLowerCase();
    var saved = String(rows[i][1]);
    var khop = saved.indexOf("sha256:") === 0 ? saved === bam_(matKhau) : saved === matKhau;
    if (tk && tk === taiKhoan && khop) {
      var token = Utilities.getUuid();
      var user = {
        taiKhoan: tk,
        hoTen: String(rows[i][2] || tk).replace(/^Nguyễn Thị Kim Phung$/, "Nguyễn Thị Kim Phụng"),
        admin: String(rows[i][3] || "").trim().toLowerCase() === "admin"
      };
      CacheService.getScriptCache().put("t_" + token, JSON.stringify(user), TOKEN_TTL);
      return { ok: true, token: token, hoTen: user.hoTen };
    }
  }
  Utilities.sleep(800); // làm chậm việc dò mật khẩu
  return { ok: false, error: "Sai tài khoản hoặc mật khẩu." };
}

function nguoiDung_(token) {
  var raw = token && CacheService.getScriptCache().get("t_" + token);
  return raw ? JSON.parse(raw) : null;
}

// Chặn ô bắt đầu bằng = + - @ bị Google Sheet hiểu thành công thức.
function antoan_(text) {
  text = String(text);
  return /^[=+\-@]/.test(text) ? "'" + text : text;
}

function ghiLoi_(req) {
  var user = nguoiDung_(req.token);
  if (!user) return { ok: false, error: "Phiên đăng nhập đã hết hạn, hãy đăng nhập lại.", expired: true };

  var phong = String(req.phong || "");
  var may = parseInt(req.may, 10);
  var loi = String(req.loi || "").trim();
  var ngay = String(req.ngay || "");

  if (!PHONG.hasOwnProperty(phong)) return { ok: false, error: "Phòng máy không hợp lệ." };
  if (!(may >= 1 && may <= PHONG[phong])) return { ok: false, error: "Số máy không hợp lệ cho phòng này." };
  if (!loi || loi.length > 500) return { ok: false, error: "Hãy nhập lỗi (tối đa 500 ký tự)." };
  if (!/^\d{4}-\d{2}-\d{2}$/.test(ngay)) return { ok: false, error: "Ngày không hợp lệ." };
  var giaoVien = String(req.giaoVien || "");
  if (GIAO_VIEN.indexOf(giaoVien) < 0) return { ok: false, error: "Hãy chọn giáo viên ghi nhận." };

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_LOI);
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  var files = [];
  try {
    var images = req.hinhAnhs === undefined ? (req.hinhAnh ? [req.hinhAnh] : []) : req.hinhAnhs;
    if (!Array.isArray(images) || images.length > 3 || images.some(function (image) { return !image; })) {
      throw new Error("Chỉ gửi tối đa 3 ảnh cho mỗi lỗi máy.");
    }
    chuanHoaSheet_(sheet);
    images.forEach(function (image) { files.push(luuAnh_(image, phong, may)); });
    var urls = files.map(function (file) { return file.getUrl(); });
    var savedImages = urls.length > 1 ? JSON.stringify(urls) : (urls[0] || "");
    sheet.appendRow([new Date(), phong, may, antoan_(loi), ngay, giaoVien, user.taiKhoan, Utilities.getUuid(), TT_CHUA, "", savedImages]);
  } catch (err) {
    files.forEach(function (file) { try { file.setTrashed(true); } catch (cleanupError) { /* giữ lỗi gốc */ } });
    throw err;
  } finally {
    lock.releaseLock();
  }
  return { ok: true };
}

// Đảm bảo có tiêu đề cột id/trangThai và mọi dòng cũ đều có id, trạng thái.
function chuanHoaSheet_(sheet) {
  var last = sheet.getLastRow();
  if (last < 1) return;
  sheet.getRange("H1:K1").setValues([["id", "trangThai", "nguoiXuLy", "hinhAnh"]]);
  if (last < 2) return;
  var range = sheet.getRange(2, 8, last - 1, 2);
  var values = range.getValues();
  var changed = false;
  values.forEach(function (row) {
    if (!row[0]) { row[0] = Utilities.getUuid(); changed = true; }
    if (!row[1]) { row[1] = TT_CHUA; changed = true; }
  });
  if (changed) range.setValues(values);
}

function timDong_(sheet, id) {
  if (!id) return null;
  var cell = sheet.getRange("H:H").createTextFinder(String(id)).matchEntireCell(true).findNext();
  return cell && cell.getRow() > 1 ? cell.getRow() : null;
}

function docAnh_(value) {
  var text = String(value || "");
  if (!text) return [];
  try {
    var urls = JSON.parse(text);
    if (Array.isArray(urls)) return urls.filter(function (url) { return typeof url === "string"; });
  } catch (err) { /* đường dẫn một ảnh từ dữ liệu cũ */ }
  return [text];
}

function layDanhSach_(req) {
  var user = nguoiDung_(req.token);
  if (!user) return { ok: false, error: "Phiên đăng nhập đã hết hạn, hãy đăng nhập lại.", expired: true };
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_LOI);
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try { chuanHoaSheet_(sheet); } finally { lock.releaseLock(); }
  var rows = sheet.getDataRange().getValues();
  var items = [];
  for (var i = rows.length - 1; i >= 1 && items.length < 500; i -= 1) {
    var ngay = rows[i][4];
    // Sheet có thể tự đổi "2026-10-11" thành ngày thật; chuẩn hóa lại về yyyy-MM-dd.
    if (ngay instanceof Date) ngay = Utilities.formatDate(ngay, Session.getScriptTimeZone(), "yyyy-MM-dd");
    items.push({
      id: rows[i][7],
      phong: rows[i][1],
      may: rows[i][2],
      loi: rows[i][3],
      ngay: String(ngay),
      giaoVien: rows[i][5],
      hinhAnh: docAnh_(rows[i][10])[0] || "",
      hinhAnhs: docAnh_(rows[i][10]),
      xong: rows[i][8] === TT_XONG,
      nguoiXong: rows[i][8] === TT_XONG ? String(rows[i][9] || "") : "",
      coQuyenXoa: user.admin || rows[i][6] === user.taiKhoan
    });
  }
  return { ok: true, items: items, supportsImages: true, maxImages: 3 };
}

function doiTrangThai_(req) {
  var user = nguoiDung_(req.token);
  if (!user) return { ok: false, error: "Phiên đăng nhập đã hết hạn, hãy đăng nhập lại.", expired: true };
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_LOI);
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var row = timDong_(sheet, req.id);
    if (!row) return { ok: false, error: "Không tìm thấy dòng lỗi này (có thể đã bị xóa)." };
    sheet.getRange(row, 9, 1, 2).setValues([req.xong ? [TT_XONG, user.hoTen] : [TT_CHUA, ""]]);
  } finally {
    lock.releaseLock();
  }
  return { ok: true };
}

function xoaLoi_(req) {
  var user = nguoiDung_(req.token);
  if (!user) return { ok: false, error: "Phiên đăng nhập đã hết hạn, hãy đăng nhập lại.", expired: true };
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_LOI);
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var row = timDong_(sheet, req.id);
    if (!row) return { ok: false, error: "Không tìm thấy dòng lỗi này (có thể đã bị xóa)." };
    if (!user.admin && sheet.getRange(row, 7).getValue() !== user.taiKhoan) {
      return { ok: false, error: "Bạn chỉ được xóa lỗi do chính tài khoản của mình ghi." };
    }
    sheet.deleteRow(row);
  } finally {
    lock.releaseLock();
  }
  return { ok: true };
}
