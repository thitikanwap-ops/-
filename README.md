// Google Apps Script — ผูกกับ Google Sheet (Extensions > Apps Script)
const ADMIN_KEY = "เปลี่ยนรหัสนี้";            // รหัสสำหรับเปิดหน้า pack.html
const PRICE = { qty_tomyum: 10, qty_moosub: 8 }; // ต้องตรงกับหน้าเว็บ
const SHEET = "Orders";

function doPost(e) {
  const d = JSON.parse(e.postData.contents);
  const a = Math.max(0, parseInt(d.qty_tomyum) || 0);
  const b = Math.max(0, parseInt(d.qty_moosub) || 0);
  const phone = String(d.phone || "").trim();
  if (!d.name || !d.address || !/^0\d{9}$/.test(phone) || a + b < 1) return out({ ok: false });
  const total = a * PRICE.qty_tomyum + b * PRICE.qty_moosub;
  // ลำดับต้องตรงกับหัวคอลัมน์: timestamp,name,phone,address,qty_tomyum,qty_moosub,total
  SpreadsheetApp.getActive().getSheetByName(SHEET)
    .appendRow([new Date(), d.name, "'" + phone, d.address, a, b, total]);
  return out({ ok: true });
}

function doGet(e) {
  if (!e.parameter.key || e.parameter.key !== ADMIN_KEY) return out({ ok: false });
  const v = SpreadsheetApp.getActive().getSheetByName(SHEET).getDataRange().getValues().slice(1);
  const orders = v.filter(r => r[1]).map((r, i) => ({
    id: "r" + (i + 2),
    name: r[1], phone: String(r[2]).replace(/^(?=\d{9}$)/, "0"), address: r[3],
    qty_tomyum: Number(r[4]) || 0, qty_moosub: Number(r[5]) || 0, total: Number(r[6]) || 0
  }));
  return out({ ok: true, orders });
}

function out(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
