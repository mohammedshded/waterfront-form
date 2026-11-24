/**
 * استقبال طلبات استمارة الووترفرونت داخل Google Sheet
 *
 * الخطوات:
 * 1) سوّي Google Sheet جديد  ←  Extensions  ←  Apps Script
 * 2) امسح الموجود والصق هذا الكود كله، واحفظ.
 * 3) Deploy  ←  New deployment  ←  النوع: Web app
 *      Execute as: Me
 *      Who has access: Anyone
 * 4) انسخ رابط الـ Web app وحطه بـ CONFIG.endpoint داخل index.html
 */

const SHEET_NAME = 'الطلبات';
const HEADERS = ['التاريخ', 'الاسم', 'الهاتف', 'المساحة م²', 'الاستلام', 'الدفع', 'ملاحظات', 'موافقة التواصل', 'المصدر', 'الحملة'];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const d = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
    if (sh.getLastRow() === 0) {
      sh.appendRow(HEADERS);
      sh.setFrozenRows(1);
      sh.setRightToLeft(true);
    }
    const clean = v => String(v == null ? '' : v).replace(/^[=+\-@]/, "'$&").slice(0, 1000);
    sh.appendRow([
      new Date(),
      clean(d.name),
      "'" + String(d.phone || '').slice(0, 20),
      clean(d.area),
      clean(d.handover_label),
      clean(d.payment_label),
      clean(d.notes),
      d.consent ? 'نعم' : 'لا',
      clean(d.source),
      clean(d.campaign)
    ]);
    return ContentService.createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}
