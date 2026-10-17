/**
 * デコラボ POS サーバー (Google Apps Script)
 *
 * Googleスプレッドシートの「拡張機能 > Apps Script」に、このファイルの中身をすべて貼り付けて
 * ウェブアプリとしてデプロイします。手順は README.md の「サーバーの設置」を見てください。
 *
 * 商品・割引・アカウント・売上はすべてこのスプレッドシートに保存されます。
 * ログインのID・パスワードの確認もここで行うので、ページのURLを知っていてもログインしないと何も読めません。
 */

const DOC_SHEET = "設定";
const SALE_SHEET = "売上";
const ACCOUNT_SHEET = "アカウント";
const SALE_HEADER = ["id", "日時", "番号", "明細", "小計", "割引", "合計", "お預かり", "おつり", "担当", "取消", "data"];
const SESSION_HOURS = 24;

function doGet() {
  return ContentService.createTextOutput("デコラボ POS のサーバーです。このURLをClaudeに伝えてください。");
}

function doPost(e) {
  let req;
  try { req = JSON.parse(e.postData.contents); } catch (x) { return out({ ok: false, error: "bad_request" }); }
  try {
    return out(handle(req || {}));
  } catch (x) {
    return out({ ok: false, error: x && x.code ? x.code : "server", message: String(x && x.message || x) });
  }
}

function handle(req) {
  const op = req.op;
  if (op === "status") return { ok: true, hasAccounts: accounts().length > 0 };
  if (op === "setup" || op === "login") return locked(() => op === "setup" ? setup(req) : login(req));

  const me = session(req.token);
  if (op === "sync") return sync(req.since);
  if (op === "logout") { props().deleteProperty("tok_" + req.token); return { ok: true }; }
  return locked(() => {
    if (op === "putSale") putSale(req.sale);
    else if (op === "putSales") (req.sales || []).forEach(putSale);
    else {
      if (me.role !== "admin") fail("forbidden");
      if (op === "setDoc") setDoc(req.name, req.data);
      else if (op === "deleteSales") deleteSales(req.ids || []);
      else if (op === "saveAccount") saveAccount(me, req);
      else if (op === "deleteAccount") deleteAccount(me, req.id);
      else fail("bad_op");
    }
    return { ok: true, rev: bump() };
  });
}

function out(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
function fail(code) { const e = new Error(code); e.code = code; throw e; }
function locked(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}

/* ---------- revision counter: lets clients skip downloads when nothing changed ---------- */
function props() { return PropertiesService.getScriptProperties(); }
function rev() { return Number(props().getProperty("rev") || 0); }
function bump() { const r = rev() + 1; props().setProperty("rev", String(r)); return r; }

/* ---------- sheets ---------- */
function sheet(name, header) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(header);
    sh.setFrozenRows(1);
  }
  return sh;
}
function rows(sh) {
  const n = sh.getLastRow();
  return n < 2 ? [] : sh.getRange(2, 1, n - 1, sh.getLastColumn()).getValues();
}
function findRow(sh, id) {
  const hit = sh.getRange("A:A").createTextFinder(String(id)).matchEntireCell(true).findNext();
  return hit && hit.getRow() > 1 ? hit.getRow() : 0;
}
function upsert(sh, row) {
  const at = findRow(sh, row[0]);
  if (at) sh.getRange(at, 1, 1, row.length).setValues([row]);
  else sh.appendRow(row);
}

/* ---------- accounts & sessions ---------- */
function accountSheet() { return sheet(ACCOUNT_SHEET, ["id", "権限", "salt", "hash"]); }
function accounts() { return rows(accountSheet()).filter(v => v[0] !== "").map(v => ({ id: String(v[0]), role: v[1] === "admin" ? "admin" : "staff", salt: v[2], hash: v[3] })); }
function hashPw(salt, pw) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + ":" + pw, Utilities.Charset.UTF_8);
  return bytes.map(b => ((b + 256) % 256).toString(16).padStart(2, "0")).join("");
}
function writeAccount(id, pw, role) {
  const salt = Utilities.getUuid();
  upsert(accountSheet(), [id, role, salt, hashPw(salt, pw)]);
}
function checkCredentials(id, pw) {
  if (!id || typeof pw !== "string" || pw.length < 4) fail("bad_input");
  if (String(id).length > 32) fail("bad_input");
}
function newSession(a) {
  const token = Utilities.getUuid();
  props().setProperty("tok_" + token, JSON.stringify({ id: a.id, exp: Date.now() + SESSION_HOURS * 3600e3 }));
  return { ok: true, token: token, id: a.id, role: a.role };
}
function session(token) {
  if (!token) fail("auth");
  const raw = props().getProperty("tok_" + token);
  const s = raw && JSON.parse(raw);
  if (!s || s.exp < Date.now()) { if (raw) props().deleteProperty("tok_" + token); fail("auth"); }
  const a = accounts().find(x => x.id === s.id);
  if (!a) fail("auth"); // the account was deleted
  return a;
}
function setup(req) {
  if (accounts().length) fail("already_setup");
  checkCredentials(req.id, req.pw);
  writeAccount(String(req.id), req.pw, "admin");
  bump();
  return newSession({ id: String(req.id), role: "admin" });
}
function login(req) {
  const a = accounts().find(x => x.id === String(req.id || ""));
  if (!a || a.hash !== hashPw(a.salt, String(req.pw || ""))) fail("bad_login");
  return newSession(a);
}
function saveAccount(me, req) {
  checkCredentials(req.id, req.pw);
  if (req.id === me.id && req.role !== "admin") fail("self_role");
  writeAccount(String(req.id), req.pw, req.role === "admin" ? "admin" : "staff");
}
function deleteAccount(me, id) {
  if (id === me.id) fail("self_delete");
  const sh = accountSheet(), at = findRow(sh, id);
  if (at) sh.deleteRow(at);
}

/* ---------- products, discounts, sales ---------- */
function sync(since) {
  const r = rev();
  if (since != null && Number(since) === r) return { ok: true, rev: r, same: true };
  const docs = {};
  rows(sheet(DOC_SHEET, ["name", "data", "更新"])).forEach(v => { try { docs[v[0]] = JSON.parse(v[1]); } catch (x) {} });
  docs.accounts = { items: accounts().map(a => ({ id: a.id, role: a.role })) };
  const dataCol = SALE_HEADER.indexOf("data");
  const sales = [];
  rows(sheet(SALE_SHEET, SALE_HEADER)).forEach(v => { try { sales.push(JSON.parse(v[dataCol])); } catch (x) {} });
  return { ok: true, rev: r, docs: docs, sales: sales };
}
function setDoc(name, data) {
  if (name !== "products" && name !== "discounts") fail("bad_doc");
  upsert(sheet(DOC_SHEET, ["name", "data", "更新"]), [name, JSON.stringify(data), new Date()]);
}
function putSale(s) {
  if (!s || !s.id) fail("bad_sale");
  const detail = (s.items || []).map(i => i.name + "×" + i.qty).join(" / ");
  upsert(sheet(SALE_SHEET, SALE_HEADER), [s.id, new Date(s.ts), s.no, detail, s.subtotal != null ? s.subtotal : s.total,
    s.discount ? s.discount.amount : 0, s.total, s.received, s.change, s.staff, s.voided ? "取消" : "", JSON.stringify(s)]);
}
function deleteSales(ids) {
  const sh = sheet(SALE_SHEET, SALE_HEADER);
  const drop = new Set(ids.map(String));
  // Delete bottom-up so row numbers stay valid.
  const vals = rows(sh);
  for (let i = vals.length - 1; i >= 0; i--) if (drop.has(String(vals[i][0]))) sh.deleteRow(i + 2);
}
