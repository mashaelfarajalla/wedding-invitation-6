/**
 * ربط قسم المباركات بـ Google Sheets
 *
 * طريقة التركيب:
 * 1) افتحي Google Sheet جديد (مثلاً باسم "مباركات الزفاف").
 * 2) من القائمة: Extensions (الإضافات) ← Apps Script.
 * 3) امسحي الكود الموجود والصقي هذا الملف كامل، ثم احفظي.
 * 4) Deploy (نشر) ← New deployment (نشر جديد) ← النوع: Web app
 *      - Execute as (التنفيذ بصفة): Me (أنا)
 *      - Who has access (من يمكنه الوصول): Anyone (أي شخص)
 *    اضغطي Deploy ووافقي على الصلاحيات.
 * 5) انسخي رابط Web app (ينتهي بـ /exec) وحطيه في index.html داخل:
 *      congratsSheetUrl: "https://script.google.com/macros/s/.../exec",
 *
 * لحذف مباركة غير لائقة: احذفي صفها من الشيت وتختفي من الموقع مباشرة.
 * إذا عدّلتي الكود لاحقًا: Deploy ← Manage deployments ← Edit ← Version: New version.
 */

const SHEET_NAME = "المباركات";
const MAX_SHOWN = 200;

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(["التاريخ", "الاسم", "المباركة"]);
    sh.setFrozenRows(1);
    sh.setRightToLeft(true);
  }
  return sh;
}

function json_(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// قراءة المباركات (الأحدث أولًا)
function doGet() {
  const sh = getSheet_();
  const last = sh.getLastRow();
  if (last < 2) return json_([]);
  const start = Math.max(2, last - MAX_SHOWN + 1);
  const rows = sh.getRange(start, 2, last - start + 1, 2).getValues();
  const list = rows
    .filter((r) => r[0] && r[1])
    .map((r) => ({ name: String(r[0]), message: String(r[1]) }))
    .reverse();
  return json_(list);
}

// إضافة مباركة جديدة
function doPost(e) {
  const p = (e && e.parameter) || {};
  // نقطع النص ونمنع بدء الخلية بـ = + - @ حتى ما تتحول لمعادلة
  const clean = (s, n) => String(s || "").trim().slice(0, n).replace(/^[=+\-@]/, "'$&");
  const name = clean(p.name, 40);
  const message = clean(p.message, 300);
  if (!name || !message) return json_({ ok: false });

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    getSheet_().appendRow([new Date(), name, message]);
  } finally {
    lock.releaseLock();
  }
  return json_({ ok: true });
}
