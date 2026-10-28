/**
 * مبادرة مدارس دار الفرسان الأهلية للدفاع عن الوطن وتعزيز الوعي الفكري
 * Google Apps Script — استقبال «كلمة دافع عن وطنك» وعرض المقبول منها فقط.
 *
 * الاستخدام: أنشئ Google Sheet، ثم Extensions ▸ Apps Script، والصق هذا الملف،
 * ثم شغّل الدالة setup() مرة واحدة، ثم Deploy ▸ New deployment ▸ Web app
 * (Execute as: Me — Who has access: Anyone). التفاصيل في README.md.
 */

var SHEET_NAME = 'المشاركات';
var HEADERS = ['التاريخ', 'الرقم', 'اسم الطالب', 'الصف', 'الكلمة', 'الحالة', 'ملاحظة المراجع'];
var COL = { DATE: 1, ID: 2, NAME: 3, GRADE: 4, WORD: 5, STATUS: 6, NOTE: 7 };
var STATUS = { NEW: 'جديدة', OK: 'مقبولة', NO: 'مرفوضة' };
var GRADES = ['الصف الثالث الابتدائي', 'الصف الرابع الابتدائي', 'الصف الخامس الابتدائي', 'الصف السادس الابتدائي'];
var MAX_PER_MINUTE = 20;   // سقف الإرسال العام في الدقيقة (حماية من الإغراق)
var LIST_LIMIT = 40;       // أقصى عدد كلمات مقبولة تُعرض في الشريط

/* ---------- إعداد أولي (يُشغَّل مرة واحدة) ---------- */
function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  sh.setRightToLeft(true);
  sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold').setBackground('#0B3D28').setFontColor('#FFFFFF');
  sh.setFrozenRows(1);
  // نص عادي في كل الأعمدة: يمنع تحويل ما يكتبه الطالب إلى صيغة (حماية من حقن الصيغ)
  sh.getRange(2, 1, sh.getMaxRows() - 1, HEADERS.length).setNumberFormat('@');
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList([STATUS.NEW, STATUS.OK, STATUS.NO], true).setAllowInvalid(false).build();
  sh.getRange(2, COL.STATUS, sh.getMaxRows() - 1, 1).setDataValidation(rule);
  sh.setColumnWidth(COL.NAME, 160); sh.setColumnWidth(COL.GRADE, 170); sh.setColumnWidth(COL.WORD, 420);
  var cf = [
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(STATUS.OK).setBackground('#D8F3E3').setRanges([sh.getRange(2, COL.STATUS, sh.getMaxRows() - 1, 1)]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(STATUS.NO).setBackground('#FADBD8').setRanges([sh.getRange(2, COL.STATUS, sh.getMaxRows() - 1, 1)]).build()
  ];
  sh.setConditionalFormatRules(cf);
  ss.toast('تم الإعداد. انشر السكربت كتطبيق ويب ثم ضع الرابط في الصفحة.', 'المبادرة', 8);
}

/* ---------- قائمة للمراجع داخل الشيت ---------- */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('المبادرة')
    .addItem('قبول الصفوف المحددة', 'acceptSelected')
    .addItem('رفض الصفوف المحددة', 'rejectSelected')
    .addToUi();
}
function acceptSelected() { setSelectedStatus_(STATUS.OK); }
function rejectSelected() { setSelectedStatus_(STATUS.NO); }
function setSelectedStatus_(status) {
  var sh = SpreadsheetApp.getActiveSheet();
  if (sh.getName() !== SHEET_NAME) { SpreadsheetApp.getUi().alert('افتح ورقة «' + SHEET_NAME + '» أولًا.'); return; }
  var r = sh.getActiveRange(); if (!r) return;
  var first = Math.max(r.getRow(), 2), last = r.getLastRow();
  for (var row = first; row <= last; row++) sh.getRange(row, COL.STATUS).setValue(status);
  CacheService.getScriptCache().remove('approved');
}

/* ---------- استقبال كلمة طالب ---------- */
function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var data = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (data.website) return json_({ ok: true });                 // حقل مخفي للروبوتات: نتجاهل بصمت
    if (!rateOk_()) return json_({ ok: false, error: 'busy' });
    var name = clean_(data.name), grade = clean_(data.grade), word = clean_(data.word);
    var err = validate_(name, grade, word);
    if (err) return json_({ ok: false, error: err });
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    if (!sh) return json_({ ok: false, error: 'not_setup' });
    var row = sh.getLastRow() + 1;
    var id = Utilities.getUuid().slice(0, 8);
    var range = sh.getRange(row, 1, 1, HEADERS.length);
    range.setNumberFormat('@');
    range.setValues([[new Date().toISOString(), id, name, grade, word, STATUS.NEW, '']]);
    return json_({ ok: true, id: id });
  } catch (err2) {
    return json_({ ok: false, error: 'server' });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

/* ---------- قراءة الكلمات المقبولة فقط ---------- */
function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) || 'list';
  if (action !== 'list') return json_({ ok: false, error: 'unknown_action' });
  var cache = CacheService.getScriptCache();
  var hit = cache.get('approved');
  if (hit) return ContentService.createTextOutput(hit).setMimeType(ContentService.MimeType.JSON);
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  var items = [];
  if (sh && sh.getLastRow() > 1) {
    var vals = sh.getRange(2, 1, sh.getLastRow() - 1, HEADERS.length).getValues();
    for (var i = vals.length - 1; i >= 0 && items.length < LIST_LIMIT; i--) {
      if (String(vals[i][COL.STATUS - 1]).trim() === STATUS.OK) {
        // نُرجع الاسم والصف والكلمة فقط: لا تاريخ ولا رقم ولا ملاحظات
        items.push({ name: String(vals[i][COL.NAME - 1]), grade: String(vals[i][COL.GRADE - 1]), word: String(vals[i][COL.WORD - 1]) });
      }
    }
  }
  var out = JSON.stringify({ ok: true, items: items });
  cache.put('approved', out, 30);
  return ContentService.createTextOutput(out).setMimeType(ContentService.MimeType.JSON);
}

/* ---------- أدوات ---------- */
function json_(obj) { return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON); }
function clean_(v) { return String(v == null ? '' : v).replace(/[\u0000-\u001F\u007F]/g, ' ').replace(/\s+/g, ' ').trim(); }
function rateOk_() {
  var cache = CacheService.getScriptCache();
  var key = 'rate_' + Math.floor(Date.now() / 60000);
  var n = Number(cache.get(key) || 0);
  if (n >= MAX_PER_MINUTE) return false;
  cache.put(key, String(n + 1), 120);
  return true;
}
// نفس قواعد الصفحة، وتُعاد هنا لأن الخادم لا يثق بما يأتيه من المتصفح
function validate_(name, grade, word) {
  var words = name.split(' ').filter(Boolean);
  if (name.length < 2 || name.length > 30 || words.length > 3) return 'name';
  if (/[0-9٠-٩@:\/]/.test(name)) return 'name';
  if (/^[=+\-@]/.test(name) || /^[=+\-@]/.test(word)) return 'name';
  if (GRADES.indexOf(grade) === -1) return 'grade';
  if (word.length < 10 || word.length > 160) return 'word';
  if (/(https?:|www\.|@|\.com|\.sa\b|\.net|\.org)/i.test(word)) return 'word';
  if (/[0-9٠-٩]{7,}/.test(word.replace(/[\s\-]/g, ''))) return 'word';
  return '';
}
