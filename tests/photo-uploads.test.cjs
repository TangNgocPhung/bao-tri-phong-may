const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");
const test = require("node:test");

const root = path.resolve(__dirname, "..");
const backend = fs.readFileSync(path.join(root, "apps-script/Code.gs"), "utf8");
const html = fs.readFileSync(path.join(root, "ghi-loi-may.html"), "utf8");
const script = [...html.matchAll(/<script(?:\s[^>]*)?>([\s\S]*?)<\/script>/g)]
  .map(match => match[1]).find(text => text.includes("function selectPhoto"));

function server() {
  const rows = [], files = [];
  const sheet = { appendRow: row => rows.push(row) };
  const context = vm.createContext({
    SpreadsheetApp: { getActiveSpreadsheet: () => ({ getSheetByName: () => sheet }) },
    LockService: { getScriptLock: () => ({ waitLock() {}, releaseLock() {} }) },
    Utilities: {
      getUuid: () => "test-id",
      base64Decode: text => Array.from(Buffer.from(text, "base64")),
      newBlob: bytes => bytes
    }
  });
  vm.runInContext(backend, context);
  context.nguoiDung_ = () => ({ taiKhoan: "teacher" });
  context.chuanHoaSheet_ = () => {};
  context.thuMucAnh_ = () => ({ createFile: () => {
    const file = { trashed: false, getUrl: () => "https://drive.google.com/file/d/" + files.length,
      setTrashed(value) { this.trashed = value; } };
    files.push(file);
    return file;
  } });
  return { context, rows, files };
}

const image = { mimeType: "image/jpeg", data: Buffer.from([255, 216, 255, 224]).toString("base64") };
const request = { token: "test", phong: "PM1", may: 1, loi: "Broken keyboard", ngay: "2026-10-12",
  giaoVien: "Nguy\u1ec5n Th\u1ecb Kim Ph\u1ee5ng" };

test("syntax is valid for the page and Apps Script", () => {
  new vm.Script(script);
  new vm.Script(backend);
});

test("zero, legacy single, and three photos save with readable links", () => {
  for (const attachment of [{}, { hinhAnh: image }, { hinhAnhs: [image, image, image] }]) {
    const { context, rows, files } = server();
    assert.equal(context.ghiLoi_({ ...request, ...attachment }).ok, true);
    assert.equal(rows.length, 1);
    const urls = context.docAnh_(rows[0][10]);
    assert.equal(urls.length, files.length);
    assert.ok(urls.every(url => url.startsWith("https://drive.google.com/")));
  }
});

test("four photos and malformed arrays are rejected before creating files", () => {
  for (const photos of [[image, image, image, image], "invalid", [null]]) {
    const { context, rows, files } = server();
    assert.throws(() => context.ghiLoi_({ ...request, hinhAnhs: photos }));
    assert.equal(rows.length, 0);
    assert.equal(files.length, 0);
  }
});

test("a failed second photo rolls back the first and does not save a report", () => {
  const { context, rows, files } = server();
  assert.throws(() => context.ghiLoi_({ ...request, hinhAnhs: [image, { ...image, data: "AAAA" }] }));
  assert.equal(rows.length, 0);
  assert.equal(files.length, 1);
  assert.equal(files[0].trashed, true);
});

test("old photo URLs, empty cells and JSON arrays remain readable", () => {
  const { context } = server();
  assert.equal(context.docAnh_("").length, 0);
  assert.equal(context.docAnh_("https://drive.google.com/one")[0], "https://drive.google.com/one");
  assert.equal(context.docAnh_(JSON.stringify(["a", "b", "c"])).length, 3);
});

test("phone selection appends photos, rejects a fourth, removes individually and clears", () => {
  function element() {
    return { children: [], value: "", disabled: false, textContent: "", events: {},
      appendChild(child) { this.children.push(child); },
      addEventListener(name, run) { this.events[name] = run; },
      set innerHTML(value) { this.children = []; } };
  }
  const elements = {}, revoked = [];
  const context = vm.createContext({
    window: { PHONG_MAY: { rooms: [] } },
    document: { getElementById: id => elements[id] ||= element(), createElement: element },
    URL: { createObjectURL: file => "blob:" + file.name, revokeObjectURL: url => revoked.push(url) }
  });
  const cut = script.indexOf('  $("login-form").addEventListener');
  vm.runInContext(script.slice(0, cut) + "globalThis.testing = {selectPhoto, clearPhoto};})();", context);
  const pick = names => context.testing.selectPhoto({ target: {
    files: names.map(name => ({ name, type: "image/jpeg", size: 100 })), value: "selection"
  } });
  pick(["one", "two"]);
  pick(["three"]);
  assert.equal(elements["photo-preview"].children.length, 3);
  pick(["four"]);
  assert.equal(elements["photo-preview"].children.length, 3);
  elements["photo-preview"].children[1].children[1].events.click();
  assert.equal(elements["photo-preview"].children.length, 2);
  assert.ok(revoked.includes("blob:two"));
  pick(["four"]);
  assert.equal(elements["photo-preview"].children.length, 3);
  context.testing.clearPhoto();
  assert.equal(elements["photo-preview"].children.length, 0);
  assert.equal(elements["photo-preview"].hidden, true);
});
