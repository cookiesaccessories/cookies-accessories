// Google Apps Script: يسجل طلبات الموقع في Google Sheet ويسمح للأدمن بقراءتها وتغيير حالتها.
// 1) افتح Google Sheet جديد، ثم Extensions > Apps Script، والصق الكود ده.
// 2) غيّر ADMIN_KEY لكلمة سر طويلة (هتكتبها في صفحة الأدمن).
// 3) Deploy > New deployment > Web app. Execute as: Me. Who has access: Anyone. انسخ الرابط في config.js (ORDERS_URL).
var ADMIN_KEY = "CHANGE_ME_TO_A_LONG_SECRET";
var STATUSES = ["pending", "done", "cancelled"];

function sheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getSheetByName("orders");
  if (!s) {
    s = ss.insertSheet("orders");
    s.appendRow(["code", "createdAt", "name", "phone", "payment", "items", "total", "status"]);
  }
  return s;
}
function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
function doPost(e) {
  try {
    var d = JSON.parse(e.postData.contents);
    if (!d.code || !d.name || !d.phone || !d.items || !d.items.length) return out_({ ok: false });
    var s = sheet_();
    var n = s.getLastRow() - 1;
    if (n > 0 && s.getRange(2, 1, n, 1).getValues().some(function (r) { return r[0] === d.code; })) return out_({ ok: true, dup: true });
    var total = 0;
    d.items.forEach(function (i) { total += Number(i.q) * Number(i.p); });
    s.appendRow([String(d.code).slice(0, 20), new Date().toISOString(), String(d.name).slice(0, 80), "'" + String(d.phone).slice(0, 25),
      d.pay === "insta" ? "insta" : "cod", JSON.stringify(d.items).slice(0, 4000), total, "pending"]);
    return out_({ ok: true });
  } catch (err) { return out_({ ok: false }); }
}
function doGet(e) {
  var p = e.parameter || {};
  if (p.key !== ADMIN_KEY) return out_({ ok: false, error: "key" });
  var s = sheet_(), n = s.getLastRow() - 1;
  var rows = n > 0 ? s.getRange(2, 1, n, 8).getValues() : [];
  if (p.action === "status") {
    if (STATUSES.indexOf(p.status) < 0) return out_({ ok: false });
    for (var i = 0; i < rows.length; i++) {
      if (rows[i][0] === p.code) { s.getRange(i + 2, 8).setValue(p.status); return out_({ ok: true }); }
    }
    return out_({ ok: false });
  }
  var orders = rows.map(function (r) {
    var items = []; try { items = JSON.parse(r[5]); } catch (x) {}
    return { code: r[0], createdAt: r[1], name: r[2], phone: String(r[3]), payment: r[4], items: items, total: Number(r[6]), status: r[7] };
  }).reverse();
  return out_({ ok: true, orders: orders });
}
