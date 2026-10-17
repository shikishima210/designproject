/**
 * デコラボ POS 共有用 Google Apps Script
 *
 * Googleスプレッドシートの「拡張機能 > Apps Script」に、このファイルの中身をすべて貼り付けて使います。
 * 手順は README.md の「全端末で共有する」を見てください。
 */

// ↓ 好きな合言葉に変えてください(英数字で12文字以上がおすすめ)。POSの接続画面で同じものを入力します。
const KEY = "ここを合言葉に変える";

const DOC_SHEET = "設定";
const SALE_SHEET = "売上";
const SALE_HEADER = ["id", "日時", "番号", "明細", "小計", "割引", "合計", "お預かり", "おつり", "担当", "取消", "data"];

function doGet() {
  return ContentService.createTextOutput("デコラボ POS の接続先です。このURLをPOSの接続画面に入力してください。");
}

function doPost(e) {
  let req;
  try { req = JSON.parse(e.postData.contents); } catch (x) { return out({ ok: false, error: "bad_request" }); }
  if (KEY === "ここを合言葉に変える" || !req || req.key !== KEY) return out({ ok: false, error: "bad_key" });
  try {
    if (req.op === "sync") return out(sync(req.since));
    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      if (req.op === "setDoc") setDoc(req.name, req.data);
      else if (req.op === "putSale") putSale(req.sale);
      else if (req.op === "putSales") (req.sales || []).forEach(putSale);
      else if (req.op === "deleteSale") deleteSale(req.id);
      else return out({ ok: false, error: "bad_op" });
      return out({ ok: true, rev: bump() });
    } finally { lock.releaseLock(); }
  } catch (x) {
    return out({ ok: false, error: "server", message: String(x) });
  }
}

function out(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function props() { return PropertiesService.getScriptProperties(); }
function rev() { return Number(props().getProperty("rev") || 0); }
function bump() { const r = rev() + 1; props().setProperty("rev", String(r)); return r; }

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
function docSheet() { return sheet(DOC_SHEET, ["name", "data", "更新"]); }
function saleSheet() { return sheet(SALE_SHEET, SALE_HEADER); }

function rows(sh) {
  const n = sh.getLastRow();
  return n < 2 ? [] : sh.getRange(2, 1, n - 1, sh.getLastColumn()).getValues();
}
function findRow(sh, id) {
  const hit = sh.getRange("A:A").createTextFinder(String(id)).matchEntireCell(true).findNext();
  return hit && hit.getRow() > 1 ? hit.getRow() : 0;
}

function sync(since) {
  const r = rev();
  if (since != null && Number(since) === r) return { ok: true, rev: r, same: true };
  const docs = {};
  rows(docSheet()).forEach(v => { try { docs[v[0]] = JSON.parse(v[1]); } catch (x) {} });
  const dataCol = SALE_HEADER.indexOf("data");
  const sales = [];
  rows(saleSheet()).forEach(v => { try { sales.push(JSON.parse(v[dataCol])); } catch (x) {} });
  return { ok: true, rev: r, docs: docs, sales: sales };
}

function setDoc(name, data) {
  if (["products", "discounts", "accounts"].indexOf(name) < 0) throw new Error("bad doc " + name);
  const sh = docSheet();
  const row = [name, JSON.stringify(data), new Date()];
  const at = findRow(sh, name);
  if (at) sh.getRange(at, 1, 1, row.length).setValues([row]);
  else sh.appendRow(row);
}

function putSale(s) {
  if (!s || !s.id) throw new Error("bad sale");
  const sh = saleSheet();
  const detail = (s.items || []).map(i => i.name + "×" + i.qty).join(" / ");
  const row = [s.id, new Date(s.ts), s.no, detail, s.subtotal != null ? s.subtotal : s.total,
    s.discount ? s.discount.amount : 0, s.total, s.received, s.change, s.staff, s.voided ? "取消" : "", JSON.stringify(s)];
  const at = findRow(sh, s.id);
  if (at) sh.getRange(at, 1, 1, row.length).setValues([row]);
  else sh.appendRow(row);
}

function deleteSale(id) {
  const sh = saleSheet();
  const at = findRow(sh, id);
  if (at) sh.deleteRow(at);
}
