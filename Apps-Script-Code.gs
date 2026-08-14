const SPREADSHEET_ID = '1ChqPWKBO29R7eZRlcTzbwLQelmLB1ZJOjv3KPJ3krbY';
const SHEET_NAME = 'Đăng ký';

function doGet() {
  return ContentService
    .createTextOutput(JSON.stringify({ ok: true, service: 'Trip Biên Hòa registration' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  const output = ContentService.createTextOutput().setMimeType(ContentService.MimeType.JSON);
  const lock = LockService.getScriptLock();

  try {
    if (!e || !e.postData || !e.postData.contents) {
      throw new Error('Thiếu dữ liệu đăng ký.');
    }

    const data = JSON.parse(e.postData.contents);
    if (!data.fullName || !data.phone || !data.plan || !data.motorbike || !data.consent) {
      throw new Error('Vui lòng điền đủ các trường bắt buộc.');
    }

    lock.waitLock(10000);
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
    if (!sheet) throw new Error('Không tìm thấy tab dữ liệu.');

    sheet.appendRow([
      new Date(),
      safeCell_(data.fullName),
      safeCell_(data.phone),
      safeCell_(data.department),
      safeCell_(data.plan),
      safeCell_(data.motorbike),
      safeCell_(data.pickup),
      safeCell_(data.diet),
      safeCell_(data.emergency),
      safeCell_(data.note),
      'Đã đăng ký'
    ]);
    SpreadsheetApp.flush();

    return output.setContent(JSON.stringify({ ok: true }));
  } catch (error) {
    return output.setContent(JSON.stringify({ ok: false, error: String(error.message || error) }));
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

function safeCell_(value) {
  const text = String(value || '').trim().slice(0, 500);
  return /^[=+\-@]/.test(text) ? "'" + text : text;
}
