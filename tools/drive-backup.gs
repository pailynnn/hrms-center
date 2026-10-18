/**
 * HRMS Center — สำรองข้อมูลลง Google Drive + Google Sheet อัตโนมัติทุกคืน (Google Apps Script) · รุ่น 2
 * ---------------------------------------------------------------------------
 * วิธีติดตั้ง / อัปเดต (ใช้บัญชี Google ของบริษัทที่เป็นเจ้าของ Google Sheet และโฟลเดอร์ HRMS Backup):
 *   1) เปิด https://script.google.com → โครงการเดิม (HRMS Backup) หรือ "โครงการใหม่"
 *   2) ลบโค้ดเดิมทั้งหมด → วางโค้ดนี้ทั้งไฟล์ → กดบันทึก (รูปแผ่นดิสก์)
 *   3) เลือกฟังก์ชัน "setup" ที่แถบด้านบน → กด "เรียกใช้" → อนุญาตสิทธิ์ (รุ่นนี้ขอสิทธิ์ Google Sheet / Google Docs เพิ่ม)
 *      (ถ้าขึ้น "Google ไม่ได้ยืนยันแอปนี้" → กด "ขั้นสูง" → "ไปที่ ... (ไม่ปลอดภัย)" — เป็นสคริปต์ของบริษัทเอง)
 *   4) เสร็จ — สำรองทันที แล้วทำเองทุกคืนประมาณ 01:00 น. · ดูสถานะในหน้าเว็บ ตั้งค่าระบบ → สำรองข้อมูล
 *
 * สิ่งที่สำรอง:
 *   • Google Sheet รายปี "HRMS Backup2026" (ไฟล์ที่บริษัทกำหนดไว้ด้านล่าง) — 1 แท็บ = 1 เดือน ("10 ตุลาคม")
 *     + แท็บ "สรุป" · เดือนปัจจุบันเขียนใหม่ทุกคืน · หลังสิ้นเดือน 7 วันแท็บเดือนนั้นจะไม่ถูกแก้อีก
 *     ขึ้นปีใหม่ → สร้างไฟล์ "HRMS Backup2027" ในโฟลเดอร์ HRMS Backup ให้เอง
 *   • โฟลเดอร์ "HRMS Backup" (โฟลเดอร์เดียว) แบ่งโฟลเดอร์ย่อย:
 *       📊 ลิงก์ Google Sheet ของแต่ละปี (ทางลัด)
 *       ใบรับรองแพทย์ / 2026 / 10 ตุลาคม /
 *       รูปหน้างาน / 2026 / 10 ตุลาคม / 2026-10-14 DR-20261014-001 /
 *       รายงานประจำวัน PDF / 2026 / 10 ตุลาคม /   (สร้าง PDF ทุกฉบับ + ฉบับที่เคยกดแชร์)
 *       เอกสารพนักงาน / T002 ชื่อ /
 *     ส่งครั้งเดียวต่อไฟล์ ไม่ลบ ไม่ทับ (ระบบจะไม่ลบไฟล์ในเว็บจนกว่าจะสำรองแล้ว)
 *
 * ⚠️ รหัสด้านล่างใช้ดึงข้อมูล HR ทั้งหมด — อย่าแชร์สคริปต์นี้ · ถ้าหลุด ให้กด "สร้างรหัสใหม่" ในหน้าเว็บ แล้ววางสคริปต์ใหม่
 */
var BACKUP_KEY = "__BACKUP_KEY__";
var ENDPOINT   = "__ENDPOINT__";
var ROOT_NAME  = "HRMS Backup";
var SHEET_IDS  = { "2026": "1cp99UeSfO_YPJwUyC0nELvgRUCPUHnU5trkzax1mTso" };   // ปีที่ไม่มีในรายการ → สร้างไฟล์ใหม่ให้เอง
var SHEET_PREFIX = "HRMS Backup";        // ชื่อไฟล์ปีใหม่ = "HRMS Backup2027"
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

/** ทุกคืน: Google Sheet + รูป/เอกสาร + PDF รายงาน */
function nightlyBackup() { run_(true); }
/** รอบต่อ (งานค้างจากรอบก่อนเพราะเกินเวลา 6 นาที) */
function continueBackup() { run_(false); }

function run_(nightly) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) { Logger.log("มีรอบอื่นกำลังทำงานอยู่"); return; }
  var t0 = Date.now(), errors = [], sheetDone = [], uploaded = 0, pdfs = 0, moved = 0, left = false;
  var props = PropertiesService.getScriptProperties();
  var timeLeft = function () { return TIME_BUDGET_MS - (Date.now() - t0); };
  clearContinue_();
  try {
    var root = rootFolder_(), cache = {};
    var plan = call_({ action: "plan" });

    // 1) Google Sheet รายปี (แท็บรายเดือน)
    if (nightly || props.getProperty("SHEET_PENDING")) {
      var pendingSheet = false;
      for (var i = 0; i < (plan.sheet_months || []).length; i++) {
        if (timeLeft() < TIME_BUDGET_MS * 0.5) { pendingSheet = true; break; }
        var m = plan.sheet_months[i];
        try {
          var ss = writeMonth_(m, root);
          call_({ action: "sheet_saved", month: m.month, sheet_url: ss.getUrl() });
          sheetDone.push(m.month);
        } catch (e) { errors.push("Google Sheet " + m.month + ": " + msg_(e)); }
      }
      if (pendingSheet) { props.setProperty("SHEET_PENDING", "1"); left = true; }
      else {
        props.deleteProperty("SHEET_PENDING");
        if (sheetDone.length && !props.getProperty("OLD_EXCEL_TRASHED")) {   // ไฟล์ Excel รุ่นแรกซ้ำกับ Google Sheet แล้ว → ถังขยะ (กู้ได้ 30 วัน)
          var ox = root.getFoldersByName("Excel รายเดือน");
          while (ox.hasNext()) ox.next().setTrashed(true);
          props.setProperty("OLD_EXCEL_TRASHED", "1");
        }
      }
    }

    // 2) ย้ายไฟล์ที่ส่งไปแล้วตามโครงโฟลเดอร์รุ่นแรก ("รูปและเอกสาร/...") เข้าโครงใหม่
    if (plan.relocate > 0) {
      var r1 = relocate_(root, cache, timeLeft, errors);
      moved = r1.moved; left = left || r1.more;
    }

    // 3) รูป/เอกสารใหม่
    var tried = {}, failed = 0, more = false;
    while (timeLeft() > 30 * 1000) {
      var r = call_({ action: "files", limit: Math.min(100, BATCH + failed) });   // ไฟล์ที่พลาดยังค้างหัวคิว → ขอเพิ่มให้พ้น
      var items = (r.items || []).filter(function (it) { return !tried[it.bucket + "|" + it.path]; });
      if (!items.length) { more = false; break; }
      var acks = [], k;
      for (k = 0; k < items.length; k++) {
        if (timeLeft() < 20 * 1000) { more = true; break; }
        var it = items[k]; tried[it.bucket + "|" + it.path] = true;
        try {
          if (!it.url) throw new Error("สร้างลิงก์ไม่ได้");
          var res = UrlFetchApp.fetch(it.url, { muteHttpExceptions: true });
          if (res.getResponseCode() !== 200) throw new Error("HTTP " + res.getResponseCode());
          var file = path_(root, it.folder, cache).createFile(res.getBlob().setName(it.name));
          acks.push({ bucket: it.bucket, path: it.path, size: it.size, drive_file_id: file.getId(), drive_path: it.folder.join("/") + "/" + it.name });
          uploaded++;
        } catch (e) { failed++; errors.push(it.name + ": " + msg_(e)); }
        if (acks.length >= 10) { call_({ action: "ack", items: acks }); acks = []; }
      }
      if (acks.length) call_({ action: "ack", items: acks });
      more = more || !!r.more;
      if (!r.more && k >= items.length) break;
    }
    if (timeLeft() <= 30 * 1000) more = true;
    left = left || more;

    // 4) PDF รายงานประจำวัน (สร้างจาก Google Docs)
    if (!more && plan.reports_pending > 0) {
      var r2 = reports_(root, cache, timeLeft, errors);
      pdfs = r2.done; left = left || r2.more;
    }

    // 5) ทางลัดไปยัง Google Sheet ในโฟลเดอร์ HRMS Backup
    try { linkSheets_(root, plan.sheets || {}); } catch (e) { errors.push("ทางลัด Google Sheet: " + msg_(e)); }

    if (left) scheduleContinue_();
    call_({ action: "done", excel: sheetDone, files_uploaded: uploaded, pdfs: pdfs, moved: moved, remaining: left ? 1 : 0,
            errors: errors.slice(0, 20), folder_url: root.getUrl(), ms: Date.now() - t0 });
    Logger.log("เสร็จ: Google Sheet %s เดือน · รูป/เอกสาร %s ไฟล์ · PDF รายงาน %s ฉบับ · ย้ายไฟล์เดิม %s · %s",
      sheetDone.length, uploaded, pdfs, moved, left ? "ยังมีงานค้าง — ทำต่ออัตโนมัติใน 1 นาที" : "ครบแล้ว");
    if (errors.length) Logger.log("พลาด: " + errors.join(" | "));
  } catch (e) {
    try { call_({ action: "fail", error: msg_(e) }); } catch (_) {}
    throw e;
  } finally {
    lock.releaseLock();
  }
}

// ---------- Google Sheet ----------
function yearSheet_(year, root) {
  var props = PropertiesService.getScriptProperties();
  var id = SHEET_IDS[String(year)] || props.getProperty("SHEET_" + year);
  if (id) {
    try { return SpreadsheetApp.openById(id); }
    catch (e) { throw new Error("เปิด Google Sheet ปี " + year + " ไม่ได้ (ต้องเป็นเจ้าของหรือมีสิทธิ์แก้ไข): " + msg_(e)); }
  }
  var found = root.getFilesByName(SHEET_PREFIX + year);
  var ss = found.hasNext() ? SpreadsheetApp.openById(found.next().getId()) : SpreadsheetApp.create(SHEET_PREFIX + year);
  var f = DriveApp.getFileById(ss.getId());
  if (!found.hasNext()) { try { f.moveTo(root); } catch (e) {} }
  props.setProperty("SHEET_" + year, ss.getId());
  return ss;
}

function writeMonth_(m, root) {
  var d = call_({ action: "month_data", month: m.month });
  var ss = yearSheet_(d.year, root);
  var sh = ss.getSheetByName(d.tab);
  if (!sh) sh = ss.insertSheet(d.tab, tabIndex_(ss, d.tab));
  else { sh.clear(); sh.clearConditionalFormatRules(); }

  var grid = [], titles = [], heads = [], texts = [], kinds = [], width = 2, total = 0, lastKind = "";
  grid.push([d.title]); grid.push(["อัปเดตเมื่อ " + d.generated_at + " (เวลาไทย) · " + (m.final ? "ปิดเดือนแล้ว — แท็บนี้จะไม่ถูกแก้อีก" : "อัปเดตทุกคืนจนถึง 7 วันหลังสิ้นเดือน")]);
  if (d.company) grid.push([d.company]);
  grid.push([]);
  d.sections.forEach(function (s) {
    if (s.kind !== lastKind) { kinds.push(grid.length + 1); grid.push(["▼ " + s.kind]); lastKind = s.kind; }
    titles.push(grid.length + 1);
    grid.push([s.title + "  (" + (s.count == null ? "ยังไม่มีตารางนี้" : s.count + " แถว") + ")", s.desc || ""]);
    if (s.header.length && s.rows.length) {
      heads.push({ row: grid.length + 1, cols: s.header.length });
      grid.push(s.header.slice());
      var first = grid.length + 1;
      s.rows.forEach(function (row) {
        grid.push(row.map(function (v) {
          if (v && typeof v === "object" && v.link) return '=HYPERLINK("' + String(v.link).replace(/"/g, '""') + '","' + String(v.text || "เปิด").replace(/"/g, '""') + '")';
          return v == null ? "" : v;
        }));
      });
      s.types.forEach(function (t, ci) { if (t === "text") texts.push(a1_(first, ci + 1, first + s.rows.length - 1, ci + 1)); });
      width = Math.max(width, s.header.length);
      if (s.kind.indexOf("ข้อมูลของเดือน") === 0) total += s.rows.length;
    }
    grid.push([]);
  });
  grid = grid.map(function (r) { while (r.length < width) r.push(""); return r; });
  if (sh.getMaxRows() < grid.length) sh.insertRowsAfter(sh.getMaxRows(), grid.length - sh.getMaxRows());
  if (sh.getMaxColumns() < width) sh.insertColumnsAfter(sh.getMaxColumns(), width - sh.getMaxColumns());
  for (var i = 0; i < texts.length; i += 400) sh.getRangeList(texts.slice(i, i + 400)).setNumberFormat("@");   // ข้อความล้วน: เลข 0 นำหน้า / เลขบัญชีไม่เพี้ยน
  sh.getRange(1, 1, grid.length, width).setValues(grid);
  sh.getRange(1, 1).setFontSize(15).setFontWeight("bold");
  if (kinds.length) sh.getRangeList(kinds.map(function (r) { return a1_(r, 1, r, width); })).setBackground("#F7941D").setFontColor("#FFFFFF").setFontWeight("bold").setFontSize(12);
  if (titles.length) sh.getRangeList(titles.map(function (r) { return a1_(r, 1, r, width); })).setBackground("#FDEBD3").setFontWeight("bold");
  if (heads.length) sh.getRangeList(heads.map(function (h) { return a1_(h.row, 1, h.row, h.cols); })).setBackground("#20304A").setFontColor("#FFFFFF").setFontWeight("bold");
  sh.setColumnWidths(1, width, 130);
  sh.setColumnWidth(1, 190);
  updateSummary_(ss, sh, d, m, total, root);
  return ss;
}

function tabIndex_(ss, tab) {
  var n = 1;   // หลังแท็บ "สรุป"
  ss.getSheets().forEach(function (s) { var nm = s.getName(); if (/^\d\d /.test(nm) && nm < tab) n++; });
  return Math.min(n, ss.getSheets().length);
}

function updateSummary_(ss, monthSheet, d, m, total, root) {
  var sum = ss.getSheetByName("สรุป");
  if (!sum) {   // แท็บเริ่มต้นที่ว่างอยู่ (ชีต1 / Sheet1) → ใช้เป็น "สรุป"
    var blank = ss.getSheets().filter(function (s) { return /^(ชีต|Sheet)\d+$/.test(s.getName()) && s.getLastRow() === 0; })[0];
    sum = blank ? blank.setName("สรุป") : ss.insertSheet("สรุป", 0);
  }
  ss.setActiveSheet(sum); ss.moveActiveSheet(1);
  var head = [["HRMS Backup " + d.year + " — สำรองข้อมูลระบบ HR (" + (d.year + 543) + ")", "", "", "", ""],
              ["1 แท็บ = 1 เดือน · ข้อมูลของเดือน + ข้อมูลหลัก ณ วันที่สำรอง · เวลาเป็นเวลาไทย", "", "", "", ""],
              ['=HYPERLINK("' + root.getUrl() + '","📂 เปิดโฟลเดอร์รูป / PDF / เอกสาร ใน Google Drive")', "", "", "", ""],
              ["", "", "", "", ""],
              ["เดือน", "อัปเดตล่าสุด", "สถานะ", "จำนวนรายการของเดือน", "เปิดแท็บ"]];
  sum.getRange(1, 1, head.length, 5).setValues(head);
  sum.getRange(1, 1).setFontSize(15).setFontWeight("bold");
  sum.getRange(5, 1, 1, 5).setBackground("#20304A").setFontColor("#FFFFFF").setFontWeight("bold");
  var last = Math.max(sum.getLastRow(), 5), rowIdx = -1;
  var names = last > 5 ? sum.getRange(6, 1, last - 5, 1).getValues() : [];
  for (var i = 0; i < names.length; i++) if (String(names[i][0]) === d.tab) { rowIdx = 6 + i; break; }
  if (rowIdx < 0) rowIdx = last + 1;
  sum.getRange(rowIdx, 1, 1, 5).setValues([[d.tab, d.generated_at, m.final ? "ปิดเดือนแล้ว" : "อัปเดตทุกคืน", total,
    '=HYPERLINK("#gid=' + monthSheet.getSheetId() + '","เปิด")']]);
  sum.getRange(rowIdx, 1, 1, 1).setNumberFormat("@");
  var n = Math.max(sum.getLastRow(), 6) - 5;
  if (n > 1) sum.getRange(6, 1, n, 5).sort(1);
  sum.setColumnWidths(1, 5, 150); sum.setColumnWidth(1, 160);
}

function linkSheets_(root, sheets) {
  var props = PropertiesService.getScriptProperties(), ids = {};
  Object.keys(SHEET_IDS).forEach(function (y) { ids[y] = SHEET_IDS[y]; });
  props.getKeys().forEach(function (k) { if (/^SHEET_\d{4}$/.test(k)) ids[k.slice(6)] = props.getProperty(k); });
  Object.keys(ids).forEach(function (y) {
    var id = ids[y];
    if (props.getProperty("SC_" + id)) return;
    var f; try { f = DriveApp.getFileById(id); } catch (e) { return; }
    var inRoot = false, ps = f.getParents(); while (ps.hasNext()) if (ps.next().getId() === root.getId()) inRoot = true;
    if (!inRoot) {
      var nm = "📊 " + f.getName() + " (Google Sheet สำรองข้อมูล)";
      if (!root.getFilesByName(nm).hasNext()) root.createShortcut(id).setName(nm);
    }
    props.setProperty("SC_" + id, "1");
  });
}

// ---------- ย้ายไฟล์รุ่นแรกเข้าโครงใหม่ ----------
function relocate_(root, cache, timeLeft, errors) {
  var moved = 0, more = false, tried = {}, missing = [];
  while (timeLeft() > 60 * 1000) {
    var r = call_({ action: "relocate", limit: 50 });
    var items = (r.items || []).filter(function (it) { return !tried[it.bucket + "|" + it.path]; });
    if (!items.length) break;
    var acks = [];
    for (var k = 0; k < items.length; k++) {
      if (timeLeft() < 45 * 1000) { more = true; break; }
      var it = items[k]; tried[it.bucket + "|" + it.path] = true;
      try {
        var f = DriveApp.getFileById(it.drive_file_id);
        if (f.isTrashed()) throw new Error("อยู่ในถังขยะ");
        f.moveTo(path_(root, it.folder, cache));
        if (f.getName() !== it.name) f.setName(it.name);
        acks.push({ bucket: it.bucket, path: it.path, size: it.size, drive_file_id: it.drive_file_id, drive_path: it.folder.join("/") + "/" + it.name });
        moved++;
      } catch (e) { missing.push({ bucket: it.bucket, path: it.path }); }
    }
    if (acks.length) call_({ action: "ack", items: acks });
    if (!r.more) break;
  }
  if (missing.length) { call_({ action: "forget", items: missing }); errors.push("ไฟล์เดิมใน Drive หาย " + missing.length + " ไฟล์ — ส่งใหม่จากระบบแทน"); }
  if (!more) { var old = root.getFoldersByName("รูปและเอกสาร"); while (old.hasNext()) trashIfEmpty_(old.next()); }
  return { moved: moved, more: more };
}
function trashIfEmpty_(folder) {
  var subs = folder.getFolders(), empty = !folder.getFiles().hasNext();
  while (subs.hasNext()) if (!trashIfEmpty_(subs.next())) empty = false;
  if (empty) folder.setTrashed(true);
  return empty;
}

// ---------- PDF รายงานประจำวัน ----------
function reports_(root, cache, timeLeft, errors) {
  var done = 0, more = false, tried = {}, failed = 0;
  while (timeLeft() > 45 * 1000) {
    var r = call_({ action: "reports", limit: Math.min(20, 3 + failed) });
    var items = (r.items || []).filter(function (it) { return !tried[it.id]; });
    if (!items.length) break;
    for (var k = 0; k < items.length; k++) {
      if (timeLeft() < 40 * 1000) { more = true; break; }
      var it = items[k]; tried[it.id] = true;
      try {
        var f = makeReportPdf_(it, path_(root, it.folder, cache));
        call_({ action: "ack", items: [{ bucket: "report-pdf", path: "daily_report/" + it.id, drive_file_id: f.getId(), drive_path: it.folder.join("/") + "/" + it.name }] });
        done++;
      } catch (e) { failed++; errors.push(it.report_no + " PDF: " + msg_(e)); }
    }
    if (more) break;
    if (r.remaining <= items.length) break;
  }
  if (timeLeft() <= 45 * 1000) more = true;
  return { done: done, more: more };
}

function makeReportPdf_(r, folder) {
  var doc = DocumentApp.create("HRMS tmp " + r.report_no);
  var tmp = DriveApp.getFileById(doc.getId());
  try {
    var body = doc.getBody();
    body.setMarginTop(30).setMarginBottom(30).setMarginLeft(36).setMarginRight(36);
    var NAVY = "#20304A", ORANGE = "#F7941D", GREY = "#5b6475", INK = "#222222";
    // Paragraph ไม่มี setBold/setFontSize → จัดผ่าน editAsText()
    var st = function (p, o) { var tx = p.editAsText(); if (tx.getText().length) { tx.setBold(!!o.b); if (o.s) tx.setFontSize(o.s); tx.setForegroundColor(o.c || INK); } if (o.before) p.setSpacingBefore(o.before); return p; };
    var para = function (text, o) { return st(body.appendParagraph(String(text == null ? "" : text)), o); };
    st(body.getParagraphs()[0].setText(r.company || "HRMS CENTER"), { b: true, s: 14, c: NAVY });
    if (r.company_th) para(r.company_th, { s: 10, c: GREY });
    para("รายงานประจำวันหน้างาน · DAILY SITE REPORT", { b: true, s: 15, c: ORANGE, before: 6 });
    para(r.report_no + " · " + r.date_th + " · สถานะ: " + r.status, { s: 10, c: GREY });
    var info = body.appendTable(r.info);
    info.setBorderColor("#d8dde8");
    for (var i = 0; i < info.getNumRows(); i++) {
      var c0 = info.getCell(i, 0), c1 = info.getCell(i, 1);
      c0.setBackgroundColor("#f3f5fa").setWidth(120);
      c0.editAsText().setBold(true).setFontSize(10).setForegroundColor(NAVY);
      if (c1.getText().length) c1.editAsText().setBold(false).setFontSize(10).setForegroundColor(INK);
    }
    r.blocks.forEach(function (b) {
      para(b[0], { b: true, s: 11, c: NAVY, before: 10 });
      para(String(b[1] || "-"), { s: 10 });
    });
    var ph = (r.photos || []).filter(function (p) { return p.url; });
    var purged = (r.photos || []).length - ph.length;
    para("รูปถ่ายหน้างาน (" + ph.length + " รูป)" + (purged ? " · ลบตามนโยบายแล้ว " + purged + " รูป" : ""), { b: true, s: 11, c: NAVY, before: 10 });
    if (ph.length) {
      var tbl = body.appendTable();
      tbl.setBorderColor("#ffffff");
      for (var j = 0; j < ph.length; j += 2) {
        var row = tbl.appendTableRow();
        for (var q = j; q < j + 2; q++) {
          var cell = row.appendTableCell("");
          if (q >= ph.length) continue;
          try {
            var res = UrlFetchApp.fetch(ph[q].url, { muteHttpExceptions: true });
            if (res.getResponseCode() !== 200) throw new Error("HTTP " + res.getResponseCode());
            var img = cell.appendImage(res.getBlob());
            var w = img.getWidth(), h = img.getHeight(), W = 245;
            if (w > 0) img.setWidth(W).setHeight(Math.round(h * W / w));
          } catch (e) { st(cell.appendParagraph("(โหลดรูปไม่ได้)"), { s: 9, c: GREY }); }
          st(cell.appendParagraph("รูปที่ " + (q + 1) + (ph[q].label ? " · " + ph[q].label : "")), { s: 9, c: GREY });
        }
      }
    } else para("ไม่มีรูป", { s: 10 });
    para("ลงชื่อ  " + (r.signed.name || "-") + (r.signed.position ? "  ·  " + r.signed.position : "") + (r.signed.at ? "  ·  " + r.signed.at : ""), { s: 10, before: 14 });
    para("สร้างจากระบบ HRMS เพื่อสำรองข้อมูลอัตโนมัติ · ส่งรายงานเมื่อ " + (r.created_at || "-"), { s: 8, c: "#8b94a7" });
    body.editAsText().setFontFamily("Sarabun");
    doc.saveAndClose();
    var pdf = tmp.getAs("application/pdf").setName(r.name);
    return folder.createFile(pdf);
  } finally {
    try { tmp.setTrashed(true); } catch (e) {}
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
function a1_(r1, c1, r2, c2) { return colName_(c1) + r1 + ":" + colName_(c2) + r2; }
function colName_(n) { var s = ""; while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; }
function scheduleContinue_() {
  clearContinue_();
  ScriptApp.newTrigger("continueBackup").timeBased().after(60 * 1000).create();
}
function clearContinue_() {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === "continueBackup") ScriptApp.deleteTrigger(t); });
}
function msg_(e) { return (e && e.message) ? e.message : String(e); }
