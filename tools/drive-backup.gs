/**
 * HRMS Center — สำรองข้อมูลลง Google Drive อัตโนมัติทุกคืน (Google Apps Script)
 * ---------------------------------------------------------------------------
 * วิธีติดตั้ง (ครั้งเดียว · ใช้บัญชี Google ของบริษัทที่จะเก็บไฟล์):
 *   1) เปิด https://script.google.com → "โครงการใหม่"
 *   2) ลบโค้ดเดิมทั้งหมด → วางโค้ดนี้ทั้งไฟล์ → กดบันทึก (รูปแผ่นดิสก์)
 *   3) เลือกฟังก์ชัน "setup" ที่แถบด้านบน → กด "เรียกใช้" → อนุญาตสิทธิ์
 *      (ถ้าขึ้น "Google ไม่ได้ยืนยันแอปนี้" → กด "ขั้นสูง" → "ไปที่ ... (ไม่ปลอดภัย)" — เป็นสคริปต์ของบริษัทเอง)
 *   4) เสร็จ — ระบบจะสำรองทันทีครั้งแรก แล้วทำเองทุกคืนประมาณ 01:00 น.
 *      ไฟล์อยู่ในโฟลเดอร์ "HRMS Backup" ใน Google Drive · ดูสถานะได้ในหน้าเว็บ ตั้งค่าระบบ → สำรองข้อมูล
 *
 * สิ่งที่สำรอง:
 *   • HRMS Backup/Excel รายเดือน/HRMS_2569-10_ตุลาคม.xlsx  — ไฟล์ละ 1 เดือน (เดือนปัจจุบันอัปเดตทุกคืน
 *     ไฟล์เดิมย้ายไปถังขยะ 30 วัน · หลังสิ้นเดือน 7 วันไฟล์เดือนนั้นจะไม่ถูกแก้อีก)
 *   • HRMS Backup/รูปและเอกสาร/...  — รูปใบรับรองแพทย์ รูปหน้างาน PDF รายงาน เอกสารพนักงาน
 *     ส่งครั้งเดียวต่อไฟล์ ไม่ลบ ไม่ทับ (ระบบจะไม่ลบไฟล์ในเว็บจนกว่าจะสำรองแล้ว)
 *
 * ⚠️ รหัสด้านล่างใช้ดึงข้อมูล HR ทั้งหมด — อย่าแชร์สคริปต์นี้ · ถ้าหลุด ให้กด "สร้างรหัสใหม่" ในหน้าเว็บ แล้ววางสคริปต์ใหม่
 */
var BACKUP_KEY = "__BACKUP_KEY__";
var ENDPOINT   = "__ENDPOINT__";
var ROOT_NAME  = "HRMS Backup";
var NIGHT_HOUR = 1;                       // 01:00 น. (ก่อนระบบล้างไฟล์เก่าตอน 02:00)
var TIME_BUDGET_MS = 5 * 60 * 1000;       // Apps Script ให้รันได้ 6 นาที/ครั้ง → เผื่อ 1 นาที
var BATCH = 25;

/** ติดตั้ง: ตั้งเวลาทุกคืน + สำรองครั้งแรกทันที */
function setup() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    var fn = t.getHandlerFunction();
    if (fn === "nightlyBackup" || fn === "continueBackup") ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger("nightlyBackup").timeBased().everyDays(1).atHour(NIGHT_HOUR).inTimezone("Asia/Bangkok").create();
  var root = rootFolder_();
  Logger.log("ตั้งเวลาสำรองทุกคืน %s:00 น. แล้ว · โฟลเดอร์: %s", NIGHT_HOUR, root.getUrl());
  nightlyBackup();
}

/** ทุกคืน: Excel รายเดือน + รูป/เอกสารใหม่ */
function nightlyBackup() { run_(true); }
/** รอบต่อ (มีไฟล์ค้างจากรอบก่อน เพราะเกินเวลา 6 นาที) */
function continueBackup() { run_(false); }

function run_(withExcel) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) { Logger.log("มีรอบอื่นกำลังทำงานอยู่"); return; }
  var t0 = Date.now(), errors = [], excelDone = [], uploaded = 0, more = false;
  clearContinue_();
  try {
    var root = rootFolder_();
    var plan = call_({ action: "plan" });
    if (withExcel) {
      var xf = child_(root, plan.excel_folder || "Excel รายเดือน");
      for (var i = 0; i < plan.months.length; i++) {
        if (Date.now() - t0 > TIME_BUDGET_MS * 0.6) { errors.push("Excel บางเดือนยังไม่ได้ทำ (เวลาไม่พอ) — จะทำคืนถัดไป"); break; }
        var m = plan.months[i];
        try {
          var x = call_({ action: "excel", month: m.month });
          var blob = Utilities.newBlob(Utilities.base64Decode(x.base64), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", x.name);
          var old = xf.getFilesByName(x.name), olds = [];
          while (old.hasNext()) olds.push(old.next());
          var f = xf.createFile(blob);
          olds.forEach(function (o) { o.setTrashed(true); });          // ไฟล์เดิมของเดือนเดียวกัน → ถังขยะ (กู้ได้ 30 วัน)
          call_({ action: "excel_saved", month: m.month, name: x.name, drive_file_id: f.getId() });
          excelDone.push(m.month);
        } catch (e) { errors.push("Excel " + m.month + ": " + msg_(e)); }
      }
    }
    // รูปและเอกสาร
    var cache = {}, tried = {}, failed = 0;
    while (Date.now() - t0 < TIME_BUDGET_MS) {
      var r = call_({ action: "files", limit: Math.min(100, BATCH + failed) });   // ไฟล์ที่พลาดยังค้างหัวคิว → ขอเพิ่มให้พ้น
      var items = (r.items || []).filter(function (it) { return !tried[it.bucket + "|" + it.path]; });
      if (!items.length) { more = false; break; }
      var acks = [];
      for (var k = 0; k < items.length; k++) {
        if (Date.now() - t0 > TIME_BUDGET_MS) { more = true; break; }
        var it = items[k]; tried[it.bucket + "|" + it.path] = true;
        try {
          if (!it.url) throw new Error("สร้างลิงก์ไม่ได้");
          var res = UrlFetchApp.fetch(it.url, { muteHttpExceptions: true });
          if (res.getResponseCode() !== 200) throw new Error("HTTP " + res.getResponseCode());
          var folder = path_(root, it.folder, cache);
          var file = folder.createFile(res.getBlob().setName(it.name));
          acks.push({ bucket: it.bucket, path: it.path, size: it.size, drive_file_id: file.getId(), drive_path: it.folder.join("/") + "/" + it.name });
          uploaded++;
        } catch (e) { failed++; errors.push(it.name + ": " + msg_(e)); }
        if (acks.length >= 10) { call_({ action: "ack", items: acks }); acks = []; }
      }
      if (acks.length) call_({ action: "ack", items: acks });
      more = more || !!r.more;
      if (!r.more && k >= items.length) break;
    }
    var cont = more || Date.now() - t0 >= TIME_BUDGET_MS;
    if (cont) scheduleContinue_();
    call_({ action: "done", excel: excelDone, files_uploaded: uploaded, remaining: cont ? 1 : 0, errors: errors.slice(0, 20), folder_url: root.getUrl(), ms: Date.now() - t0 });
    Logger.log("เสร็จ: Excel %s เดือน · รูป/เอกสาร %s ไฟล์ · %s", excelDone.length, uploaded, cont ? "ยังมีไฟล์ค้าง — ทำต่ออัตโนมัติใน 1 นาที" : "ครบแล้ว");
    if (errors.length) Logger.log("พลาด: " + errors.join(" | "));
  } catch (e) {
    try { call_({ action: "fail", error: msg_(e) }); } catch (_) {}
    throw e;
  } finally {
    lock.releaseLock();
  }
}

// ---------- ตัวช่วย ----------
function call_(payload) {
  payload.key = BACKUP_KEY;
  var res = UrlFetchApp.fetch(ENDPOINT, { method: "post", contentType: "application/json", payload: JSON.stringify(payload), muteHttpExceptions: true });
  var code = res.getResponseCode(), body = {};
  try { body = JSON.parse(res.getContentText()); } catch (e) {}
  if (code !== 200) throw new Error((body && body.error) || ("HTTP " + code));
  return body;
}
function rootFolder_() {
  var props = PropertiesService.getScriptProperties(), id = props.getProperty("ROOT_ID");
  if (id) { try { var f = DriveApp.getFolderById(id); if (!f.isTrashed()) return f; } catch (e) {} }
  var it = DriveApp.getRootFolder().getFoldersByName(ROOT_NAME);
  var root = it.hasNext() ? it.next() : DriveApp.getRootFolder().createFolder(ROOT_NAME);
  props.setProperty("ROOT_ID", root.getId());
  return root;
}
function child_(parent, name) {
  var it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}
function path_(root, parts, cache) {
  var key = "", cur = root;
  for (var i = 0; i < parts.length; i++) {
    key += "/" + parts[i];
    if (!cache[key]) cache[key] = child_(cur, parts[i]);
    cur = cache[key];
  }
  return cur;
}
function scheduleContinue_() {
  clearContinue_();
  ScriptApp.newTrigger("continueBackup").timeBased().after(60 * 1000).create();
}
function clearContinue_() {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === "continueBackup") ScriptApp.deleteTrigger(t); });
}
function msg_(e) { return (e && e.message) ? e.message : String(e); }
