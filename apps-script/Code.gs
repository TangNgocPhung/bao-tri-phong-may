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
      case "ws_load": out = taiPhienCham_(req); break;
      case "ws_save": out = luuPhienCham_(req); break;
      case "ws_image_put": out = luuAnhCham_(req); break;
      case "ws_image_get": out = docAnhCham_(req); break;
      case "ws_publish": out = congBoDiem_(req); break;
      case "ws_unpublish": out = goCongBo_(req); break;
      case "ws_published": out = dsCongBo_(req); break;
      case "tra_cuu": out = traCuuDiem_(req); break;
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

// ===== Lưu phiên chấm trắc nghiệm theo từng tài khoản (Google Drive) =====
// Mỗi giáo viên có một thư mục riêng: ws.json (đáp án, danh sách lớp, kết quả) và ảnh bài làm *.jpg.

function thuMucCham_(user) {
  var props = PropertiesService.getScriptProperties();
  var rootId = props.getProperty("THU_MUC_CHAM_BAI");
  var root;
  try { root = rootId ? DriveApp.getFolderById(rootId) : null; } catch (err) { root = null; }
  if (!root) {
    root = DriveApp.createFolder("Dữ liệu chấm trắc nghiệm - " + SpreadsheetApp.getActiveSpreadsheet().getName());
    props.setProperty("THU_MUC_CHAM_BAI", root.getId());
  }
  var it = root.getFoldersByName(user.taiKhoan);
  return it.hasNext() ? it.next() : root.createFolder(user.taiKhoan);
}

function idAnhHopLe_(id) {
  return typeof id === "string" && /^[A-Za-z0-9_-]{1,64}$/.test(id);
}

function timFile_(folder, name) {
  var it = folder.getFilesByName(name);
  return it.hasNext() ? it.next() : null;
}

function taiPhienCham_(req) {
  var user = nguoiDung_(req.token);
  if (!user) return { ok: false, error: "Phiên đăng nhập đã hết hạn, hãy đăng nhập lại.", expired: true };
  var file = timFile_(thuMucCham_(user), "ws.json");
  return { ok: true, data: file ? file.getBlob().getDataAsString("UTF-8") : null };
}

function luuPhienCham_(req) {
  var user = nguoiDung_(req.token);
  if (!user) return { ok: false, error: "Phiên đăng nhập đã hết hạn, hãy đăng nhập lại.", expired: true };
  if (typeof req.data !== "string" || req.data.length > 20 * 1024 * 1024) return { ok: false, error: "Dữ liệu phiên chấm không hợp lệ." };
  var data = JSON.parse(req.data);
  var keep = {};
  (data.imageIds || []).forEach(function (id) { if (idAnhHopLe_(id)) keep[id + ".jpg"] = true; });
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var folder = thuMucCham_(user);
    var file = timFile_(folder, "ws.json");
    if (file) file.setContent(req.data); else folder.createFile("ws.json", req.data, "application/json");
    // Xóa ảnh không còn thuộc phiên (bài đã bị xóa).
    var files = folder.getFiles();
    while (files.hasNext()) {
      var f = files.next();
      var name = f.getName();
      if (/\.jpg$/.test(name) && !keep[name]) f.setTrashed(true);
    }
  } finally {
    lock.releaseLock();
  }
  return { ok: true };
}

function luuAnhCham_(req) {
  var user = nguoiDung_(req.token);
  if (!user) return { ok: false, error: "Phiên đăng nhập đã hết hạn, hãy đăng nhập lại.", expired: true };
  if (!idAnhHopLe_(req.id)) return { ok: false, error: "Mã ảnh không hợp lệ." };
  if (typeof req.data !== "string" || !req.data || req.data.length > Math.ceil(6 * 1024 * 1024 / 3) * 4 ||
      !/^[A-Za-z0-9+/]+={0,2}$/.test(req.data)) {
    return { ok: false, error: "Ảnh bài làm không hợp lệ hoặc quá lớn." };
  }
  var bytes = Utilities.base64Decode(req.data);
  if (bytes.length < 3 || (bytes[0] & 255) !== 255 || (bytes[1] & 255) !== 216 || (bytes[2] & 255) !== 255) {
    return { ok: false, error: "Ảnh phải là JPEG." };
  }
  var folder = thuMucCham_(user);
  var name = req.id + ".jpg";
  var old = timFile_(folder, name);
  if (old) old.setTrashed(true);
  folder.createFile(Utilities.newBlob(bytes, "image/jpeg", name));
  return { ok: true };
}

function docAnhCham_(req) {
  var user = nguoiDung_(req.token);
  if (!user) return { ok: false, error: "Phiên đăng nhập đã hết hạn, hãy đăng nhập lại.", expired: true };
  if (!idAnhHopLe_(req.id)) return { ok: false, error: "Mã ảnh không hợp lệ." };
  var file = timFile_(thuMucCham_(user), req.id + ".jpg");
  if (!file) return { ok: false, error: "Không tìm thấy ảnh." };
  return { ok: true, data: Utilities.base64Encode(file.getBlob().getBytes()) };
}

// ===== Công bố điểm cho học sinh tự tra cứu (không cần đăng nhập, chỉ xem) =====
// Giáo viên bấm "Công bố" mới có dữ liệu; mỗi giáo viên một file congbo-<tài khoản>.json ở thư mục gốc.
// Chỉ lưu tên, lớp, SBD, mã đề, điểm và đáp án từng câu — không lưu ảnh bài làm.

function goc_() {
  thuMucCham_({ taiKhoan: "_" }); // đảm bảo thư mục gốc đã tạo
  return DriveApp.getFolderById(PropertiesService.getScriptProperties().getProperty("THU_MUC_CHAM_BAI"));
}

function tenCongBo_(user) {
  return "congbo-" + user.taiKhoan.replace(/[^a-z0-9_.-]/g, "_") + ".json";
}

function docCongBo_(file) {
  try { return file ? JSON.parse(file.getBlob().getDataAsString("UTF-8")) : { exams: {} }; }
  catch (err) { return { exams: {} }; }
}

function ghiCongBo_(user, mutate) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var root = goc_();
    var name = tenCongBo_(user);
    var file = timFile_(root, name);
    var data = docCongBo_(file);
    if (!data.exams) data.exams = {};
    mutate(data);
    var text = JSON.stringify(data);
    if (file) file.setContent(text); else root.createFile(name, text, "application/json");
  } finally {
    lock.releaseLock();
  }
}

function congBoDiem_(req) {
  var user = nguoiDung_(req.token);
  if (!user) return { ok: false, error: "Phiên đăng nhập đã hết hạn, hãy đăng nhập lại.", expired: true };
  var exam = String(req.exam || "").trim().slice(0, 120);
  if (!exam || !Array.isArray(req.records) || req.records.length > 2000) return { ok: false, error: "Dữ liệu công bố không hợp lệ." };
  var records = req.records.map(function (r) {
    return {
      name: String(r.name || "").slice(0, 80),
      className: String(r.className || "").slice(0, 30), electiveClass: String(r.electiveClass || "").slice(0, 30),
      sbd: String(r.sbd || "").slice(0, 20),
      code: String(r.code || "").slice(0, 20),
      score: Number(r.score),
      correct: Number(r.correct) || 0,
      total: Number(r.total) || 0,
      detail: (Array.isArray(r.detail) ? r.detail : []).slice(0, 250).map(function (d) {
        return [String(d.q).slice(0, 8), String(d.selected || "").slice(0, 4), String(d.answer || "").slice(0, 4), d.correct ? 1 : 0];
      })
    };
  }).filter(function (r) { return r.name && isFinite(r.score); });
  ghiCongBo_(user, function (data) {
    data.exams[exam] = { at: new Date().toISOString(), teacher: user.hoTen, records: records };
  });
  return { ok: true, count: records.length };
}

function goCongBo_(req) {
  var user = nguoiDung_(req.token);
  if (!user) return { ok: false, error: "Phiên đăng nhập đã hết hạn, hãy đăng nhập lại.", expired: true };
  var exam = String(req.exam || "").trim();
  ghiCongBo_(user, function (data) { delete data.exams[exam]; });
  return { ok: true };
}

function dsCongBo_(req) {
  var user = nguoiDung_(req.token);
  if (!user) return { ok: false, error: "Phiên đăng nhập đã hết hạn, hãy đăng nhập lại.", expired: true };
  var data = docCongBo_(timFile_(goc_(), tenCongBo_(user)));
  return { ok: true, exams: Object.keys(data.exams || {}).map(function (k) {
    return { exam: k, at: data.exams[k].at, count: data.exams[k].records.length };
  }) };
}

function chuanTen_(text) {
  return String(text || "").normalize("NFD").replace(/[̀-ͯ]/g, "").replace(/đ/g, "d").replace(/Đ/g, "D")
    .toLowerCase().replace(/[^a-z0-9 ]/g, " ").replace(/\s+/g, " ").trim();
}

function traCuuDiem_(req) {
  var q = chuanTen_(String(req.ten || "").slice(0, 80));
  var tokens = q ? q.split(" ") : [];
  if (q.length < 5 || tokens.length < 2) return { ok: false, error: "Hãy nhập họ và tên đầy đủ (ít nhất 2 từ) và lớp." };
  var lop = chuanTen_(req.lop).replace(/ /g, "");
  if (!lop) return { ok: false, error: "Hãy nhập lớp của em." };
  var found = [];
  var files = goc_().getFiles();
  while (files.hasNext() && found.length < 15) {
    var file = files.next();
    if (file.getName().indexOf("congbo-") !== 0) continue;
    var data = docCongBo_(file);
    Object.keys(data.exams || {}).forEach(function (exam) {
      data.exams[exam].records.forEach(function (r) {
        var n = " " + chuanTen_(r.name) + " ";
        if (found.length < 15 && (chuanTen_(r.className).replace(/ /g, "") === lop || chuanTen_(r.electiveClass).replace(/ /g, "") === lop) && tokens.every(function (t) { return n.indexOf(" " + t) >= 0; })) {
          found.push({ exam: exam, at: data.exams[exam].at, teacher: data.exams[exam].teacher, name: r.name,
            className: r.className, electiveClass: r.electiveClass || "", sbd: r.sbd, code: r.code, score: r.score, correct: r.correct, total: r.total, detail: r.detail });
        }
      });
    });
  }
  return { ok: true, results: found };
}
