/*************************************************************************
 * ระบบรายงานผลการเรียน ม.1-3 • หลักสูตรแกนกลางการศึกษาขั้นพื้นฐาน พุทธศักราช 2551
 * พร้อมเทียบเคียงเป็นผลการเรียนตามหลักสูตรแกนกลางฯ พุทธศักราช 2551
 * --------------------------------------------------------------------
 * Backend: Google Apps Script (REST API) + Google Sheets (ฐานข้อมูล)
 * Frontend: GitHub Pages เรียกผ่าน fetch() (ไม่ใช้ google.script.run)
 * พัฒนาโดย: ครูรุ่งนิรันดร์
 * --------------------------------------------------------------------
 * หมายเหตุ: ฟีเจอร์แปลง PDF→รูป และนำเข้าเช็คชื่อจาก Excel ใช้ Drive REST API
 *   ผ่าน UrlFetch (ไม่ต้องเปิด Advanced Service) — แต่ครั้งแรกต้องอนุญาตสิทธิ์
 *   Drive + External requests ตอน Deploy เวอร์ชันใหม่
 * --------------------------------------------------------------------
 * Part 1/4 : Backend ทั้งหมด
 *************************************************************************/

/** ====================== ค่าคงที่ / โครงสร้างชีต ====================== */

/* ============================================================
   เมนู NSR (แสดงบนแถบเมนูของ Google Sheets)
   ============================================================ */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('NSR')
    .addSubMenu(ui.createMenu('🚀 เผยแพร่/สำรอง')
      .addItem('🚀 เผยแพร่ Edge เดี๋ยวนี้', 'menuEdgePublishNow')
      .addItem('🔁 เปิดเผยแพร่อัตโนมัติ (ทุก 5 นาที)', 'menuEdgeAutoOn')
      .addItem('⏸️ ปิดเผยแพร่อัตโนมัติ', 'menuEdgeAutoOff')
      .addItem('🔑 ตั้งค่า GitHub Token', 'menuEdgeSetToken')
      .addSeparator()
      .addItem('💾 สำรองข้อมูลลง Drive เดี๋ยวนี้', 'menuBackupNow')
      .addItem('📊 สถานะเผยแพร่/สำรอง', 'menuEdgeStatus'))
    .addItem('📥 นำเข้าการเช็คชื่อจาก Excel', 'menuImportAttendance')
    .addItem('🔧 ติดตั้ง/ตรวจสอบชีต', 'menuSetupSheets')
    .addSeparator()
    .addSubMenu(ui.createMenu('🗑️ ล้างข้อมูล (ระวัง)')
      .addItem('ล้างรายชื่อนักเรียน', 'menuClearStudents')
      .addItem('ล้างคะแนน/เกรด', 'menuClearScores')
      .addItem('ลบตารางสอนทั้งหมด', 'menuClearSchedule')
      .addSeparator()
      .addItem('🩹 แก้สถานะ "สาย" ปลอม (จากบั๊กเช็ครายคาบ)', 'menuFixPhantomLate')
      .addItem('ลบนักเรียนที่ย้ายออก "ถาวร" พร้อมข้อมูล', 'menuPurgeMoved')
      .addItem('🧨 ล้างทั้งระบบ', 'menuClearAll'))
    .addToUi();
}

function menuPurgeMoved() {
  var ui = SpreadsheetApp.getUi();
  if (ui.alert('ลบถาวร', 'จะลบนักเรียนที่ "ย้ายออก" ทั้งหมด พร้อมคะแนน/เกรด/เช็คชื่อ/ผลประเมิน ของพวกเขาอย่างถาวร\nกู้คืนไม่ได้! ดำเนินการต่อ?', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  var r = purgeMovedStudents();
  ui.alert('NSR', 'ลบถาวรแล้ว ' + r.purged + ' คน', ui.ButtonSet.OK);
}

function menuFixPhantomLate() {
  var ui = SpreadsheetApp.getUi();
  if (ui.alert('แก้สายปลอม', 'ระบบจะตรวจหากลุ่มเช็คชื่อหน้าเสาธงที่เป็น "สาย" ทั้งห้อง ซึ่งถูกสร้างโดยบั๊กเช็ครายคาบ แล้วแปลงเป็น "มา" ให้\nดำเนินการต่อ?', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  var r = fixPhantomLate();
  ui.alert('NSR', r.fixed ? ('แก้แล้ว ' + r.fixed + ' รายการ ✅\n(' + r.groups.join(', ') + ')') : 'ไม่พบรายการสายปลอม ✅', ui.ButtonSet.OK);
}

function menuClearSchedule() {
  var ui = SpreadsheetApp.getUi();
  if (ui.alert('ลบตารางสอน', 'จะลบตารางสอนทั้งหมด (รายวิชาที่สร้างไว้ยังอยู่) ดำเนินการต่อ?',
    ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  var s = ss().getSheetByName('Schedule');
  if (s && s.getLastRow() > 1) s.getRange(2, 1, s.getLastRow() - 1, s.getLastColumn()).clearContent();
  SpreadsheetApp.flush();
  ui.alert('NSR', 'ลบตารางสอนทั้งหมดแล้ว ✅', ui.ButtonSet.OK);
}

function menuSetupSheets() {
  PropertiesService.getScriptProperties().setProperty('setupVersion', SETUP_VERSION);
  // ลบชีต conflict ที่ Google Sheets สร้างจากการซิงก์ชนกัน (เช่น Supervisions_conflict...)
  ss().getSheets().forEach(function (sh) {
    if (/_conflict/i.test(sh.getName())) { try { ss().deleteSheet(sh); } catch (e) {} }
  });

  setupSheets();
  SpreadsheetApp.getUi().alert('NSR', 'ตรวจสอบ/สร้างชีตครบแล้ว ✅', SpreadsheetApp.getUi().ButtonSet.OK);
}

// ล้างเฉพาะรายชื่อนักเรียน (และข้อมูลที่ผูกกับนักเรียน) เพื่ออัปรายชื่อใหม่
function menuClearStudents() {
  var ui = SpreadsheetApp.getUi();
  var res = ui.alert('ล้างรายชื่อนักเรียน',
    'จะลบ "รายชื่อนักเรียนเดิมทั้งหมด" พร้อมคะแนน/เกรด/ความสามารถ/การเช็คชื่อ ที่ผูกกับนักเรียน\n\nเพื่อให้พร้อมอัปรายชื่อใหม่ — ดำเนินการต่อหรือไม่?',
    ui.ButtonSet.YES_NO);
  if (res !== ui.Button.YES) return;
  ['Students', 'Scores', 'Grades', 'AbilityDetail', 'Attendance', 'Locks'].forEach(function (name) {
    var s = ss().getSheetByName(name);
    if (s && s.getLastRow() > 1) s.getRange(2, 1, s.getLastRow() - 1, s.getLastColumn()).clearContent();
  });
  SpreadsheetApp.flush();
  ui.alert('NSR', 'ล้างรายชื่อนักเรียนและข้อมูลที่เกี่ยวข้องแล้ว ✅\nนำเข้ารายชื่อใหม่ได้เลย', ui.ButtonSet.OK);
}

function menuClearScores() {
  var ui = SpreadsheetApp.getUi();
  if (ui.alert('ล้างผลการเรียน', 'จะลบคะแนน/เกรด/ความสามารถทั้งหมด (รายชื่อนักเรียนยังอยู่) ดำเนินการต่อ?',
    ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  ['Scores', 'Grades', 'AbilityDetail', 'LearningUnits'].forEach(function (name) {
    var s = ss().getSheetByName(name);
    if (s && s.getLastRow() > 1) s.getRange(2, 1, s.getLastRow() - 1, s.getLastColumn()).clearContent();
  });
  SpreadsheetApp.flush();
  ui.alert('NSR', 'ล้างผลการเรียนแล้ว ✅', ui.ButtonSet.OK);
}

function menuClearAll() {
  var ui = SpreadsheetApp.getUi();
  if (ui.alert('⚠️ ล้างข้อมูลทั้งระบบ',
    'จะลบทุกข้อมูล (นักเรียน/รายวิชา/คะแนน/เกรด/เช็คชื่อ ฯลฯ) ยกเว้นค่าตั้งค่าโรงเรียนและตารางแปลงเกรด\n\nดำเนินการต่อหรือไม่?',
    ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  Object.keys(SHEETS).forEach(function (name) {
    if (name === 'Settings' || name === 'GradeMapping' || name === 'Teachers') return;
    var s = ss().getSheetByName(name);
    if (s && s.getLastRow() > 1) s.getRange(2, 1, s.getLastRow() - 1, s.getLastColumn()).clearContent();
  });
  SpreadsheetApp.flush();
  ui.alert('NSR', 'ล้างข้อมูลทั้งระบบแล้ว ✅', ui.ButtonSet.OK);
}

function menuBackupNow() {
  var r = runBackupToDrive();
  SpreadsheetApp.getUi().alert('NSR', 'สำรองข้อมูลลง Drive แล้ว ✅\nไฟล์: ' + r.name, SpreadsheetApp.getUi().ButtonSet.OK);
}

function menuEnableAutoBackup() {
  var ui = SpreadsheetApp.getUi();
  setAutoBackup({ enable: true, hour: 22 });
  ui.alert('NSR', 'เปิดสำรองข้อมูลอัตโนมัติทุกวัน (ประมาณ 22:00) แล้ว ✅\nไฟล์สำรองอยู่ในโฟลเดอร์ Drive:\n' + AUTO_BACKUP_FOLDER, ui.ButtonSet.OK);
}

// นำเข้าการเช็คชื่อจากไฟล์ Excel (รูปแบบระบบเช็คชื่อเดิม: ชีต Students + Attendance)
function menuImportAttendance() {
  var ui = SpreadsheetApp.getUi();
  var resp = ui.prompt('นำเข้าการเช็คชื่อจาก Excel',
    '1) อัปโหลดไฟล์ .xlsx เข้า Google Drive ก่อน\n2) วางลิงก์ (URL) หรือ File ID ของไฟล์ที่นี่\n\n(ไฟล์ต้องมีชีต "Students" และ "Attendance" รูปแบบ uuid|ชั้น|เลขที่|คำนำหน้า|ชื่อ|นามสกุล และ วันที่|uuid|สถานะ)',
    ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  var idOrUrl = resp.getResponseText().trim();
  if (!idOrUrl) { ui.alert('ยังไม่ได้ใส่ลิงก์/ID'); return; }
  try {
    var r = importAttendanceFromFile_(idOrUrl);
    ui.alert('NSR — นำเข้าสำเร็จ ✅',
      'นำเข้า: ' + r.imported + ' รายการ\nข้ามซ้ำ: ' + r.skippedDup + '\nจับคู่นักเรียนไม่ได้: ' + r.unmatched + ' รายการ' +
      (r.unmatchedClasses.length ? '\n(ชั้นที่ไม่มีในระบบ: ' + r.unmatchedClasses.join(', ') + ')' : ''),
      ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('เกิดข้อผิดพลาด', String(e.message || e), ui.ButtonSet.OK);
  }
}

// คัดลอกไฟล์ Drive พร้อมแปลงชนิด (ผ่าน REST API + OAuth token) — ไม่ต้องเปิด Advanced Service
function driveCopyConvert_(fileId, name, targetMime) {
  var res = UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/files/' + fileId + '/copy?supportsAllDrives=true', {
    method: 'post', contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    payload: JSON.stringify({ name: name, mimeType: targetMime }),
    muteHttpExceptions: true
  });
  var data = JSON.parse(res.getContentText() || '{}');
  if (!data.id) throw new Error('แปลงไฟล์ไม่สำเร็จ: ' + ((data.error && data.error.message) || res.getContentText()));
  return data.id;
}

function importAttendanceFromFile_(idOrUrl) {
  var m = String(idOrUrl).match(/[-\w]{25,}/);
  if (!m) throw new Error('ลิงก์/ID ไม่ถูกต้อง');
  var fileId = m[0];

  // แปลง xlsx → Google Sheet ชั่วคราว (ผ่าน Drive REST API — ไม่ต้องเปิด Advanced Service)
  var tmpId = driveCopyConvert_(fileId, '__att_import_tmp', 'application/vnd.google-apps.spreadsheet');
  try {
    var tmp = SpreadsheetApp.openById(tmpId);
    var sStu = tmp.getSheetByName('Students');
    var sAtt = tmp.getSheetByName('Attendance');
    if (!sStu || !sAtt) throw new Error('ไฟล์ต้องมีชีต "Students" และ "Attendance"');

    // uuid → ข้อมูลนักเรียนในไฟล์
    var fsv = sStu.getDataRange().getValues();
    var uuidInfo = {};
    fsv.forEach(function (r) {
      if (!r[0]) return;
      uuidInfo[String(r[0])] = { cls: String(r[1] || '').trim(), num: Number(r[2]) || 0,
        first: String(r[4] || '').trim(), last: String(r[5] || '').trim() };
    });

    // ระบบ: ดัชนีหลายแบบ — ชื่อ-นามสกุล (หลัก), ชั้น+เลขที่ (สำรอง), ชั้น+ชื่อ (แก้ชื่อซ้ำ)
    var nrm = function (x) { return String(x || '').replace(/\s+/g, '').trim(); };
    var sysByName = {}, sysByNum = {}, sysByClassName = {};
    getStudents().forEach(function (s) {
      var nk = nrm(s.firstName) + '|' + nrm(s.lastName);
      (sysByName[nk] = sysByName[nk] || []).push(s.ID);     // อาจมีชื่อซ้ำหลายคน
      sysByNum[s.classLevel + '|' + (Number(s.number) || 0)] = s.ID;
      sysByClassName[s.classLevel + '|' + nk] = s.ID;
    });

    // อ่าน attendance ในไฟล์
    var av = sAtt.getDataRange().getValues();
    var STMAP = { present: 'มา', late: 'สาย', leave: 'ลา', absent: 'ขาด' };

    // กันเขียนชนกับการเช็คชื่อสด (บทเรียนข้อมูลหายจากการเขียนพร้อมกัน)
    var lock = LockService.getScriptLock(); lock.waitLock(30000);
    try {
    // ชุดคีย์ที่มีอยู่แล้วในระบบ (กันซ้ำ): date|studentID|period0
    var sys = sheet('Attendance');
    var existing = {};
    var ev = sys.getDataRange().getValues();
    for (var i = 1; i < ev.length; i++) {
      if (String(ev[i][3]) === '0') existing[ymd(ev[i][1]) + '|' + ev[i][4]] = true;
    }

    var now = Utilities.formatDate(new Date(), 'GMT+7', 'yyyy-MM-dd HH:mm:ss');
    var out = [], imported = 0, skippedDup = 0, unmatched = 0, unmatchedCls = {};
    av.forEach(function (r) {
      if (!r[0] || !r[1]) return;
      var info = uuidInfo[String(r[1])];
      if (!info) { unmatched++; return; }
      var nk = nrm(info.first) + '|' + nrm(info.last);
      var sid = '';
      var byName = sysByName[nk] || [];
      if (byName.length === 1) sid = byName[0];                                  // ชื่อไม่ซ้ำ → จับด้วยชื่อ
      else if (byName.length > 1) sid = sysByClassName[info.cls + '|' + nk] || ''; // ชื่อซ้ำ → ใช้ชั้นช่วย
      if (!sid) sid = sysByClassName[info.cls + '|' + nk] || sysByNum[info.cls + '|' + info.num] || ''; // สำรอง
      if (!sid) { unmatched++; unmatchedCls[info.cls] = true; return; }
      var date = ymd(r[0]);
      var key = date + '|' + sid;
      if (existing[key]) { skippedDup++; return; }
      existing[key] = true;
      var status = STMAP[String(r[2] || '').toLowerCase()] || 'มา';
      out.push([genId('ATT'), "'" + date, info.cls, '0', sid, status, 'นำเข้า', now, curSem_()]);
      imported++;
    });

    if (out.length) sys.getRange(sys.getLastRow() + 1, 1, out.length, out[0].length).setValues(out);
    SpreadsheetApp.flush();
    } finally { lock.releaseLock(); }

    return { imported: imported, skippedDup: skippedDup, unmatched: unmatched,
      unmatchedClasses: Object.keys(unmatchedCls).sort() };
  } finally {
    try { DriveApp.getFileById(tmpId).setTrashed(true); } catch (e) {}
  }
}


var SHEETS = {
  Settings:     ['Key', 'Value'],
  Students:     ['ID', 'studentCode', 'number', 'fullName', 'classLevel', 'dateAdded', 'prefix', 'firstName', 'lastName', 'status', 'movedDate', 'movedNumber'],
  Subjects:     ['ID', 'code', 'name', 'learningArea', 'hours', 'type', 'classLevel', 'sortOrder'],
  SubjectDocs:  ['ID', 'subjectID', 'fileName', 'fileId', 'type', 'updatedAt'],
  LearningUnits:['ID', 'subjectID', 'unitName', 'maxScore', 'sortOrder', 'semester'],
  Scores:       ['ID', 'studentID', 'subjectID', 'unitID', 'score', 'semester'],
  // 'level68' = คอลัมน์เก่า คงไว้เพื่อให้ตรงกับชีต Grades เดิม (ระบบเขียนค่าว่างเสมอ ไม่ใช้งานแล้ว)
  Grades:       ['ID', 'studentID', 'subjectID', 'mode', 'percent', 'level68', 'grade51', 'letter51', 'note', 'semester'],
  AbilityDetail:['ID', 'studentID', 'group', 'item', 'value', 'semester'],
  GradeMapping: ['minPercent', 'maxPercent', 'grade51', 'letter51'],
  Teachers:     ['ID', 'name', 'pin', 'role', 'homeroomClass'],
  Schedule:     ['ID', 'day', 'period', 'classLevel', 'subject', 'teacher'],
  Attendance:   ['ID', 'date', 'classLevel', 'period', 'studentID', 'status', 'checkedBy', 'updatedAt', 'semester'],
  Calendar:     ['ID', 'date', 'type', 'title', 'note'],
  Submissions:  ['ID', 'teacher', 'subject', 'classLevel', 'docType', 'fileName', 'fileId', 'fileUrl', 'status', 'updatedAt', 'semester'],
  SubmissionTypes: ['ID', 'name', 'semester', 'startDate', 'endDate', 'active'],
  Remediation:  ['ID', 'studentID', 'subjectID', 'semester', 'origMark', 'newGrade', 'status', 'note', 'resolvedBy', 'resolvedDate', 'updatedAt'],
  Supervisions: ['ID', 'date', 'classLevel', 'teacher', 'subject', 'topic', 'strengths', 'suggestions', 'rating', 'by', 'updatedAt', 'semester'],
  FormTemplates:['ID', 'name', 'fileName', 'fileId', 'fileUrl', 'updatedAt'],
  TeacherAssign:['ID', 'subjectID', 'room', 'teacher', 'updatedAt'],
  ExamSchedule: ['ID', 'examType', 'date', 'startTime', 'endTime', 'classLevel', 'subject', 'room', 'proctor1', 'proctor2', 'semester', 'updatedAt'],
  SDQ:          ['ID', 'studentID', 'semester', 'answers', 'emotional', 'conduct', 'hyper', 'peer', 'prosocial', 'total', 'result', 'assessedBy', 'updatedAt'],
  HomeVisits:   ['ID', 'studentID', 'date', 'visitors', 'found', 'living', 'needs', 'photoId', 'photoUrl', 'semester', 'by', 'updatedAt'],
  StudentProfiles:['studentID', 'nickname', 'photoId', 'photoUrl', 'citizenId', 'birthDate', 'religion', 'nationality', 'ethnicity', 'bloodGroup', 'weight', 'height', 'address', 'fatherName', 'motherName', 'guardianName', 'guardianRelation', 'guardianPhone', 'updatedAt', 'updatedBy', 'guardianJob', 'fatherJob', 'motherJob', 'disadvantaged'],
  Locks:        ['classLevel', 'status', 'updatedAt', 'by', 'note']
};

// ภาคเรียนปัจจุบัน (default ถ้า frontend ยังไม่ส่ง semester มา)
function curSem_(p) {
  if (p && p.semester) return String(p.semester);
  var cs = getSettings().currentSemester;
  return String(cs || '1');
}

// ลบแถวที่ตรงสองเงื่อนไข (ใช้ header-name หาคอลัมน์ — ปลอดภัยจากคอลัมน์เลื่อน)
function clearRowsWhere2(name, f1, v1, f2, v2) {
  var s = sheet(name);
  var data = s.getDataRange().getValues();
  if (data.length < 2) return;
  var c1 = data[0].indexOf(f1), c2 = data[0].indexOf(f2);
  if (c1 < 0) return;
  var keep = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i].join('') === '') continue;
    var match = String(data[i][c1]) === String(v1) && (c2 < 0 || String(data[i][c2]) === String(v2));
    if (!match) keep.push(data[i]);
  }
  var cols = data[0].length;
  if (s.getLastRow() > 1) s.getRange(2, 1, s.getLastRow() - 1, cols).clearContent();
  if (keep.length) s.getRange(2, 1, keep.length, cols).setValues(keep);
}

// ค่าตั้งต้นข้อมูลสถานศึกษา (แก้ไขได้ในเมนูตั้งค่า)
var DEFAULT_SETTINGS = {
  schoolName:      'โรงเรียนของฉัน',
  slogan:          'ระบบรายงานผลการเรียน หลักสูตรแกนกลางฯ พุทธศักราช 2551',
  academicYear:    '2569',
  area:            '',                 // สำนักงานเขตพื้นที่
  directorName:    '',                 // ชื่อผู้อำนวยการ
  directorSignURL: '',                 // URL ลายเซ็น ผอ.
  deputyName:      '',                 // ชื่อรองผู้อำนวยการ (ฝ่ายวิชาการ)
  deputySignURL:   '',                 // URL ลายเซ็นรอง ผอ.
  logoURL:         '',                 // URL ตราสัญลักษณ์
  remark:          '',
  hoursPerYear:    '200',              // มาตรฐานเวลาเรียน ชม./ปี (ประถม)
  currentSemester: '1',                // ภาคเรียนปัจจุบัน (1/2)
  sem1Start:       '',                 // ช่วงวันที่ภาคเรียนที่ 1 (YYYY-MM-DD) — ใช้กรองเวลาเรียน/มส.
  sem1End:         '',
  sem2Start:       '',
  sem2End:         ''
};

// ตารางแปลงเกรดเริ่มต้น (อิงตารางที่ ๕ ของแนวปฏิบัติฯ) — แก้ไขได้ในเมนูตั้งค่า
var DEFAULT_MAPPING = [
  { minPercent: 80, maxPercent: 100, grade51: 4,   letter51: 'ดีเยี่ยม' },
  { minPercent: 75, maxPercent: 79,  grade51: 3.5, letter51: 'ดีมาก'   },
  { minPercent: 70, maxPercent: 74,  grade51: 3,   letter51: 'ดี'      },
  { minPercent: 65, maxPercent: 69,  grade51: 2.5, letter51: 'ค่อนข้างดี' },
  { minPercent: 60, maxPercent: 64,  grade51: 2,   letter51: 'ปานกลาง'  },
  { minPercent: 55, maxPercent: 59,  grade51: 1.5, letter51: 'พอใช้'    },
  { minPercent: 50, maxPercent: 54,  grade51: 1,   letter51: 'ผ่านเกณฑ์ขั้นต่ำ' },
  { minPercent: 0,  maxPercent: 49,  grade51: 0,   letter51: 'ต่ำกว่าเกณฑ์' }
];

// โฟลเดอร์เก็บรูปบน Drive
var DRIVE_FOLDER_NAME = 'KPS_GradeReport_Images';


/** ====================== ROUTER (REST) ====================== */

function doGet(e)  { return handleRequest(e, 'GET');  }
function doPost(e) { return handleRequest(e, 'POST'); }

function handleRequest(e, method) {
  var action  = '';
  var payload = {};
  try {
    // อ่าน action + payload จาก POST (text/plain) หรือ GET (?action=)
    if (method === 'POST' && e && e.postData && e.postData.contents) {
      var body = JSON.parse(e.postData.contents);
      action  = body.action || '';
      payload = body.payload || {};
    } else if (e && e.parameter) {
      action  = e.parameter.action || '';
      if (e.parameter.payload) payload = JSON.parse(e.parameter.payload);
    }

    ensureSetup_(); // ติดตั้ง/ซ่อมโครงสร้างเฉพาะเมื่อจำเป็น

    var result = dispatch(action, payload);
    return jsonOut({ ok: true, data: result });

  } catch (err) {
    return jsonOut({ ok: false, error: String(err && err.message ? err.message : err) });
  }
}

function dispatch(action, p) {
  switch (action) {
    // โหลดข้อมูลเริ่มต้น
    case 'getBootstrap':      return getBootstrapFull(p);

    // ตั้งค่าสถานศึกษา
    case 'getSettings':       return getSettings();
    case 'saveSettings':      return saveSettings(p);

    // ตารางแปลงเกรด
    case 'getGradeMapping':   return getGradeMapping();
    case 'saveGradeMapping':  return saveGradeMapping(p.mapping);

    // นักเรียน
    case 'getStudents':       return getStudents();
    case 'addStudent':        assertAdmin_(p); return addStudent(p);
    case 'updateStudent':     assertAdmin_(p); return updateStudent(p.id, p);
    case 'deleteStudent':     assertAdmin_(p); return deleteStudent(p.id);
    case 'bulkImportStudents':assertAdmin_(p); return bulkImportStudents(p.rows);

    // รายวิชา
    case 'getSubjects':       return getSubjects();
    case 'addSubject':        return addSubject(p);
    case 'updateSubject':     return updateSubject(p.id, p);
    case 'deleteSubject':     return deleteSubject(p.id);

    // หน่วย/ผลลัพธ์การเรียนรู้
    case 'getUnits':          return getUnits(p.subjectID, p.semester);
    case 'saveUnits':         assertScorePerm_(p, p.subjectID, ''); return saveUnits(p.subjectID, p.units, p.semester);

    // คะแนน (โหมด A)
    case 'getScores':         return getScores(p.subjectID, p.semester);
    case 'saveScores':        assertScorePerm_(p, p.subjectID, ''); return saveScores(p.subjectID, p.scores, p.semester);

    // เกรด/ผลการเรียน
    case 'getGrades':         return getGrades(p.semester);
    case 'saveGrade':         assertScorePerm_(p, p.subjectID, roomOfStudent_(p.studentID)); return saveGrade(p);
    case 'saveGradesBatch':   assertGradesBatchPerm_(p); return saveGradesBatch(p.grades, p.semester);
    case 'recomputeGrades':   assertScorePerm_(p, p.subjectID, ''); return recomputeGrades(p.subjectID, p.semester);

    // ความสามารถ / คุณลักษณะ / กิจกรรม / อ่านคิดวิเคราะห์เขียน
    case 'getAbilityDetailClass': return getAbilityDetailClass(p.classLevel, p.semester);
    case 'saveAbilityDetailBatch': return saveAbilityDetailBatch(p);

    // รายงาน
    case 'getReportBatch':    return getReportBatch(p.classLevel, p.studentID);
    case 'getPP5Cover':       return getPP5Cover(p.classLevel);
    case 'getPP5Book':        return getPP5Book(p);
    case 'getCalendar':       return getCalendar();
    case 'getLocks':          return getLocks();
    case 'getRiskStudents':   return getRiskStudents(p);
    case 'getFlagCheckedDates': return getFlagCheckedDates(p);
    case 'getPeriodCheckedDates': return getPeriodCheckedDates(p);
    case 'getSubjectAttendance': return getSubjectAttendance(p);
    case 'getSubjectTimeCheck': return getSubjectTimeCheck(p);
    case 'getMyTodo':           return getMyTodo(p);
    case 'fixPhantomLate':      return fixPhantomLate();
    case 'getDirectorOverview': return getDirectorOverview(p);
    case 'getExecAttendance':   return getExecAttendance(p);
    case 'getExecScoreStatus':  return getExecScoreStatus(p);
    case 'getRemediation':      return getRemediation(p);
    case 'resolveRemediation':  return resolveRemediation(p);
    case 'getSupervisions':     return getSupervisions(p);
    case 'saveSupervision':     return saveSupervision(p);
    case 'deleteSupervision':   return deleteSupervision(p);
    case 'getClassProfiles':    return getClassProfiles(p);
    case 'saveStudentProfile':  return saveStudentProfile(p);
    case 'uploadStudentPhoto':  return uploadStudentPhoto(p);
    case 'importDMC':           return importDMC(p);
    case 'moveOutStudent':      return moveOutStudent(p);
    case 'restoreStudent':      return restoreStudent(p);
    case 'getMovedStudents':    return getMovedStudents();
    case 'getSAR':              return getSAR(p);
    case 'getSDQClass':         return getSDQClass(p);
    case 'saveSDQ':             return saveSDQ(p);
    case 'getHomeVisitClass':   return getHomeVisitClass(p);
    case 'getHomeVisits':       return getHomeVisits(p);
    case 'saveHomeVisit':       return saveHomeVisit(p);
    case 'deleteHomeVisit':     return deleteHomeVisit(p);
    case 'getExams':            return getExams(p);
    case 'saveExam':            return saveExam(p);
    case 'deleteExam':          return deleteExam(p);
    case 'saveScheduleSlot':    return saveScheduleSlot(p);
    case 'deleteScheduleSlot':  return deleteScheduleSlot(p);
    case 'getTeacherAssign':    return getTeacherAssign();
    case 'saveTeacherAssign':   return saveTeacherAssign(p);
    case 'syncTeacherAssign':   return syncTeacherAssign();
    case 'getMyTeach':          return getMyTeach(p);
    case 'subjectHealthCheck':  return subjectHealthCheck();
    case 'getFormTemplates':    return getFormTemplates();
    case 'uploadFormTemplate':  return uploadFormTemplate(p);
    case 'deleteFormTemplate':  return deleteFormTemplate(p);
    case 'getTeachingTasks':    return getTeachingTasks(p);
    case 'uploadTeachingDoc':   return uploadTeachingDoc(p);
    case 'deleteTeachingDoc':   return deleteTeachingDoc(p);
    case 'getTeachingOverview': return getTeachingOverview(p);
    case 'getSubmissionTypes':  return getSubmissionTypes(p);
    case 'getAllSubmissionTypes': return getAllSubmissionTypes();
    case 'saveSubmissionType':  return saveSubmissionType(p);
    case 'deleteSubmissionType': return deleteSubmissionType(p);
    case 'setLock':           return setLock(p);
    case 'addCalendarEvent':  return addCalendarEvent(p);
    case 'deleteCalendarEvent': return deleteCalendarEvent(p.id);
    case 'getStats':          return getStats();
    case 'getGradeStats':     return getGradeStats(p);
    case 'getClassGradeSummary': return getClassGradeSummary(p.classLevel, p.semester);
    case 'exportSchoolMIS':   return exportSchoolMIS(p.classLevel);
    case 'exportSchoolMISWide': return exportSchoolMISWide(p.classLevel);

    // จัดการข้อมูล
    case 'getEdgeStatus':     assertAdmin_(p); return getEdgeStatus();
    case 'saveEdgeConfig':    assertAdmin_(p); return saveEdgeConfig(p);
    case 'setEdgeAuto':       assertAdmin_(p); return setEdgeAuto(p);
    case 'publishEdgeNow':    assertAdmin_(p); return publishEdgeNow(p);
    case 'backupAllData':     assertAdmin_(p); return backupAllData();
    case 'runBackupToDrive':  assertAdmin_(p); return runBackupToDrive();
    case 'setAutoBackup':     assertAdmin_(p); return setAutoBackup(p);
    case 'getBackupStatus':   assertAdmin_(p); return getBackupStatus();
    case 'restoreBackup':     assertAdmin_(p); return restoreBackup(p);
    case 'importBackupData':  assertAdmin_(p); return importBackupData(p.json);
    case 'clearAllData':      assertAdmin_(p); return clearAllData();

    // อัปโหลดรูป
    case 'uploadImage':       return uploadImageToDrive(p.base64, p.filename);
    case 'getSubjectDocs':    return getSubjectDocs(p.subjectID);
    case 'getDocBase64':      return getDocBase64(p.fileId);
    case 'uploadSubjectDoc':  return uploadSubjectDoc(p);
    case 'deleteSubjectDoc':  return deleteSubjectDoc(p.id);

    case 'ping':              return { pong: true, time: new Date().toString() };

    // ---- Part 5: ล็อกอิน / ครู / ตารางสอน / เช็คชื่อ ----
    case 'login':             return login(p.name, p.pin);
    case 'changeMyPassword':  return changeMyPassword(p);
    case 'getTeachers':       return getTeachers();
    case 'saveTeacher':       return saveTeacher(p);
    case 'saveTeachersBatch': return saveTeachersBatch(p.teachers);
    case 'deleteTeacher':     return deleteTeacher(p.id);

    case 'getSchedule':       return getSchedule(p.teacher);
    case 'importSchedule':    return importSchedule(p.rows);
    case 'syncSubjectsFromSchedule': return syncSubjectsFromSchedule();
    case 'dedupeSubjects':    return dedupeSubjects();
    case 'dedupeStudents':    return dedupeStudents();

    case 'bulkImportRoster':  assertAdmin_(p); return bulkImportRoster(p.rows);
    case 'renumberClass':     return renumberClass(p.classLevel);

    case 'getAttendance':     return getAttendance(p.date, p.classLevel, p.period);
    case 'getFlagAttendance': return getAttendance(p.date, p.classLevel, 0);
    case 'saveAttendance':    return saveAttendance(p);
    case 'getAttendanceSummary':   return getAttendanceSummary(p.classLevel, p.fromDate, p.toDate);
    case 'getAttendanceDashboard': return getAttendanceDashboard(p.date);
    case 'getAttendanceRegister':  return getAttendanceRegister(p.classLevel, p.fromDate, p.toDate);
    case 'getMonthlyGrid':         return getMonthlyGrid(p.classLevel, p.ym);

    default: throw new Error('ไม่พบ action: "' + action + '"');
  }
}


/** ====================== HELPERS ====================== */

function jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function ss() { return SpreadsheetApp.getActiveSpreadsheet(); }

// ---- แคชการอ่านชีต 2 ชั้น: ในคำขอเดียวกัน + ข้ามคำขอ (เฉพาะชีตที่เปลี่ยนไม่บ่อย 60 วิ) ----
var _readCache = {};
var CACHE_SHEETS = { Settings:1, Students:1, Subjects:1, Teachers:1, Schedule:1,
  TeacherAssign:1, GradeMapping:1, Calendar:1, SubmissionTypes:1, FormTemplates:1 };
function _sheetRaw_(name) {
  var s = ss().getSheetByName(name);
  if (!s) { s = ss().insertSheet(name); s.appendRow(SHEETS[name]); }
  return s;
}
function bustSheetCache_(name) {
  delete _readCache[name];
  if (CACHE_SHEETS[name]) { try { CacheService.getScriptCache().remove('sh_' + name); } catch (e) {} }
}
// ตัวเข้าถึงชีตสำหรับ "เขียน" — ล้างแคชของชีตนั้นอัตโนมัติทุกครั้ง (กันข้อมูลค้าง)
function sheet(name) {
  bustSheetCache_(name);
  markEdgeDirty_(name);   // ชีตสาธารณะถูกแก้ → ติดธงรอเผยแพร่ (มาตรฐาน Sheets+Edge ข้อ 2)
  return _sheetRaw_(name);
}

// สร้างชีตทั้งหมด + ใส่ค่าตั้งต้นเมื่อรันครั้งแรก
// แปลงหัวคอลัมน์ภาษาไทย/รูปแบบเดิม → คีย์อังกฤษที่ระบบใช้ (เพื่อให้ใช้ชีตจากระบบเดิมที่หัวเป็นไทยได้)
var HEADER_ALIASES = {
  Teachers: { 'ชื่อ': 'name', 'ชื่อ-สกุล': 'name', 'ชื่อ-นามสกุล': 'name', 'ชื่อครู': 'name', 'PIN': 'pin', 'รหัส': 'pin', 'รหัสผ่าน': 'pin', 'บทบาท': 'role', 'สิทธิ์': 'role', 'ห้องประจำชั้น': 'homeroomClass', 'ห้องที่ปรึกษา': 'homeroomClass', 'ครูประจำชั้น': 'homeroomClass' },
  Students: { 'รหัสนักเรียน': 'studentCode', 'เลขประจำตัว': 'studentCode', 'เลขประจำตัวนักเรียน': 'studentCode', 'เลขที่': 'number', 'ชื่อ-นามสกุล': 'fullName', 'ชื่อ-สกุล': 'fullName', 'ชื่อสกุล': 'fullName', 'ชื่อ - นามสกุล': 'fullName', 'ระดับชั้น/ห้อง': 'classLevel', 'ระดับชั้น': 'classLevel', 'ชั้น/ห้อง': 'classLevel', 'ชั้น': 'classLevel', 'ห้อง': 'classLevel', 'วันที่เพิ่ม': 'dateAdded', 'คำนำหน้า': 'prefix', 'ชื่อจริง': 'firstName', 'นามสกุล': 'lastName' },
  Settings: { 'คีย์': 'Key', 'ค่า': 'Value' }
};
function normalizeHeaders_() {
  Object.keys(HEADER_ALIASES).forEach(function (name) {
    var s = ss().getSheetByName(name); if (!s) return;
    var lastCol = s.getLastColumn(); if (lastCol < 1) return;
    var hdr = s.getRange(1, 1, 1, lastCol).getValues()[0];
    var alias = HEADER_ALIASES[name];
    var existing = {}; hdr.forEach(function (h) { existing[String(h).trim()] = true; });
    var changed = false;
    for (var c = 0; c < hdr.length; c++) {
      var key = String(hdr[c]).trim();
      var canon = alias[key];
      if (canon && key !== canon && !existing[canon]) { hdr[c] = canon; existing[canon] = true; changed = true; }
    }
    if (changed) s.getRange(1, 1, 1, hdr.length).setValues([hdr]);
  });

  // ซ่อมหัวคอลัมน์ชีต Subjects ให้ตรง schema มาตรฐาน (ข้อมูลถูกเขียนเรียงตามคอลัมน์มาตรฐานจากการนำเข้าตารางสอน)
  // กรณีชีตถูกสร้าง/แก้หัวจากระบบเดิม ทำให้ getSubjects อ่านคีย์ไม่ตรง (รายวิชาแสดงค่าว่าง)
  ['Subjects'].forEach(function (name) {
    var s = ss().getSheetByName(name); if (!s) return;
    var want = SHEETS[name];
    var lastCol = s.getLastColumn(); if (lastCol < 1) return;
    var cur = s.getRange(1, 1, 1, Math.max(lastCol, want.length)).getValues()[0];
    var same = want.every(function (h, i) { return String(cur[i]).trim() === h; });
    if (!same) s.getRange(1, 1, 1, want.length).setValues([want]);
  });

  // ซ่อมหัวคอลัมน์ชีต Attendance เมื่อคีย์หลักหาย (date/studentID/status) — ป้องกันบันทึกเช็คชื่อแล้วอ่านกลับไม่เจอ
  (function () {
    var s = ss().getSheetByName('Attendance'); if (!s) return;
    var want = SHEETS.Attendance;
    var lastCol = Math.max(s.getLastColumn(), want.length);
    var cur = s.getRange(1, 1, 1, lastCol).getValues()[0].map(function (h) { return String(h).trim(); });
    var hasKeys = cur.indexOf('date') >= 0 && cur.indexOf('studentID') >= 0 && cur.indexOf('status') >= 0;
    if (!hasKeys) s.getRange(1, 1, 1, want.length).setValues([want]);
  })();

  // ซ่อมชีต Scores จากโครงระบบเดิม (ภาคเรียน, SubjectID, StudentID, คะแนนหน่วยJSON, ...)
  // → โครงใหม่ (ID, studentID, subjectID, unitID, score, semester)
  // พร้อมกู้คืนแถวที่ระบบใหม่เคยบันทึกลงไปแบบตำแหน่งเพี้ยน (ID ขึ้นต้น SCR_)
  (function () {
    var s = ss().getSheetByName('Scores'); if (!s) return;
    var want = SHEETS.Scores;
    var lastCol = Math.max(s.getLastColumn(), want.length);
    var cur = s.getRange(1, 1, 1, lastCol).getValues()[0].map(function (h) { return String(h).trim(); });
    var same = want.every(function (h, i) { return cur[i] === h; });
    if (same) return;
    var n = s.getLastRow();
    var salvaged = [], legacy = [];
    if (n > 1) {
      s.getRange(2, 1, n - 1, lastCol).getValues().forEach(function (r) {
        if (String(r[0]).indexOf('SCR_') === 0) {
          salvaged.push([r[0], r[1], r[2], r[3], r[4], r[5]]);   // ค่าถูกลำดับอยู่แล้ว แค่หัวผิด
        } else if (r.join('') !== '') {
          legacy.push(r);
        }
      });
    }
    if (legacy.length) {   // เก็บข้อมูลระบบเดิมไว้ในชีตสำรอง ไม่ลบทิ้ง
      var bak = ss().getSheetByName('Scores_ระบบเดิม') || ss().insertSheet('Scores_ระบบเดิม');
      if (bak.getLastRow() === 0) bak.getRange(1, 1, 1, cur.length).setValues([cur.slice(0, lastCol)]);
      bak.getRange(bak.getLastRow() + 1, 1, legacy.length, lastCol).setValues(legacy);
    }
    s.clearContents();
    s.getRange(1, 1, 1, want.length).setValues([want]);
    if (salvaged.length) s.getRange(2, 1, salvaged.length, want.length).setValues(salvaged);
  })();

  // ซ่อมชีต GradeMapping จากโครงเดิม (ร้อยละต่ำสุด, ร้อยละสูงสุด, เกรด)
  // → (minPercent, maxPercent, grade51, letter51) — คงช่วงคะแนนเดิมของโรงเรียน เติมคำอธิบายระดับให้
  (function () {
    var s = ss().getSheetByName('GradeMapping'); if (!s) return;
    var want = SHEETS.GradeMapping;
    var lastCol = Math.max(s.getLastColumn(), want.length);
    var cur = s.getRange(1, 1, 1, lastCol).getValues()[0].map(function (h) { return String(h).trim(); });
    var same = want.every(function (h, i) { return cur[i] === h; });
    if (same) return;
    var idx = {}; cur.forEach(function (h, i) { if (h && idx[h] == null) idx[h] = i; });
    var iMin = (idx.minPercent != null) ? idx.minPercent : idx['ร้อยละต่ำสุด'];
    var iMax = (idx.maxPercent != null) ? idx.maxPercent : idx['ร้อยละสูงสุด'];
    var iGrd = (idx.grade51 != null) ? idx.grade51 : idx['เกรด'];
    var iLet = idx.letter51;
    var letterOf = {};
    DEFAULT_MAPPING.forEach(function (m) { letterOf[String(m.grade51)] = m.letter51 || ''; });
    var rows = [];
    var n = s.getLastRow();
    if (n > 1 && iMin != null && iMax != null && iGrd != null) {
      s.getRange(2, 1, n - 1, lastCol).getValues().forEach(function (r) {
        var mn = Number(r[iMin]), mx = Number(r[iMax]);
        var g = String(r[iGrd] == null ? '' : r[iGrd]).trim();
        if (isNaN(mn) || isNaN(mx) || g === '') return;
        rows.push([mn, mx, g, (iLet != null && r[iLet]) ? r[iLet] : (letterOf[g] || '')]);
      });
    }
    if (!rows.length) {
      rows = DEFAULT_MAPPING.map(function (m) { return [m.minPercent, m.maxPercent, m.grade51, m.letter51]; });
    }
    s.clearContents();
    s.getRange(1, 1, 1, want.length).setValues([want]);
    s.getRange(2, 1, rows.length, want.length).setValues(rows);
  })();

  // ซ่อมชีต Locks จากโครงเดิม (เช่น classLevel, semester, status, updatedAt, by)
  // → โครงใหม่ (classLevel, status, updatedAt, by, note) พร้อมย้ายข้อมูลเดิมตามชื่อหัวคอลัมน์
  (function () {
    var s = ss().getSheetByName('Locks'); if (!s) return;
    var want = SHEETS.Locks;
    var lastCol = Math.max(s.getLastColumn(), want.length);
    var cur = s.getRange(1, 1, 1, lastCol).getValues()[0].map(function (h) { return String(h).trim(); });
    var same = want.every(function (h, i) { return cur[i] === h; });
    if (same) return;
    var idx = {}; cur.forEach(function (h, i) { if (h && idx[h] == null) idx[h] = i; });
    var rows = [];
    var n = s.getLastRow();
    if (n > 1) {
      s.getRange(2, 1, n - 1, lastCol).getValues().forEach(function (r) {
        var cls = (idx.classLevel != null) ? r[idx.classLevel] : r[0];
        if (String(cls || '') === '') return;
        rows.push([cls,
          (idx.status != null) ? r[idx.status] : '',
          (idx.updatedAt != null) ? r[idx.updatedAt] : '',
          (idx.by != null) ? r[idx.by] : '',
          (idx.note != null) ? r[idx.note] : '']);
      });
    }
    s.clearContents();
    s.getRange(1, 1, 1, want.length).setValues([want]);
    if (rows.length) s.getRange(2, 1, rows.length, want.length).setValues(rows);
  })();
}

// รันติดตั้ง/ซ่อมโครงสร้างเฉพาะครั้งแรกหลัง Deploy เวอร์ชันใหม่ (เดิมสแกนทุกชีตทุกคำขอ ทำให้ช้า)
var SETUP_VERSION = 'v2026-07-21b';
function ensureSetup_() {
  var props = PropertiesService.getScriptProperties();
  if (props.getProperty('setupVersion') === SETUP_VERSION) return;
  setupSheets();
  props.setProperty('setupVersion', SETUP_VERSION);
}

function setupSheets() {
  Object.keys(SHEETS).forEach(function (name) {
    var s = ss().getSheetByName(name);
    if (!s) { s = ss().insertSheet(name); s.appendRow(SHEETS[name]); }
  });
  normalizeHeaders_(); // แปลงหัวคอลัมน์ไทย→อังกฤษ ให้ระบบอ่านข้อมูลเดิมได้
  // เติมหัวคอลัมน์ที่ "ต่อท้ายใหม่" (เช่น semester) ให้ชีตเดิมที่ยังไม่มี — ปลอดภัยเพราะเป็นการต่อท้ายเท่านั้น
  ['LearningUnits', 'Scores', 'Grades', 'AbilityDetail', 'Attendance', 'Locks', 'StudentProfiles'].forEach(function (name) {
    var s = ss().getSheetByName(name);
    if (!s) return;
    var lastCol = s.getLastColumn();
    var hdr = s.getRange(1, 1, 1, Math.max(lastCol, 1)).getValues()[0];
    var schema = SHEETS[name];
    if (hdr.length < schema.length) {
      s.getRange(1, 1, 1, schema.length).setValues([schema]); // เขียนหัวให้ครบตาม schema
    }
  });
  // ใส่ค่าตั้งต้น Settings
  var st = sheet('Settings');
  if (st.getLastRow() < 2) {
    Object.keys(DEFAULT_SETTINGS).forEach(function (k) {
      st.appendRow([k, DEFAULT_SETTINGS[k]]);
    });
  }
  // ใส่ตารางแปลงเกรดตั้งต้น
  var gm = sheet('GradeMapping');
  if (gm.getLastRow() < 2) {
    DEFAULT_MAPPING.forEach(function (m) {
      gm.appendRow([m.minPercent, m.maxPercent, m.grade51, m.letter51]);
    });
  }
  // seed บัญชีแอดมินเริ่มต้น (PIN: 1234 — เปลี่ยนได้ในหน้าจัดการครู)
  var tc = sheet('Teachers');
  if (tc.getLastRow() < 2) {
    tc.appendRow([genId('TCH'), 'แอดมิน', '1234', 'admin', '']);
  }
  SpreadsheetApp.flush();
}

// อ่านทั้งชีตเป็น array ของ object (แปลง Date เป็น string เสมอ)
function readAll(name) {
  if (_readCache[name]) return _readCache[name];
  if (CACHE_SHEETS[name]) {
    try {
      var hit = CacheService.getScriptCache().get('sh_' + name);
      if (hit) { var arr = JSON.parse(hit); _readCache[name] = arr; return arr; }
    } catch (e) {}
  }
  var s = _sheetRaw_(name);
  var values = s.getDataRange().getValues();
  if (values.length < 2) return [];
  var headers = values[0];
  var rows = [];
  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    if (row.join('') === '') continue; // ข้ามแถวว่าง
    var obj = {};
    for (var c = 0; c < headers.length; c++) {
      var v = row[c];
      if (v instanceof Date) v = Utilities.formatDate(v, 'GMT+7', 'yyyy-MM-dd HH:mm:ss');
      obj[headers[c]] = v;
    }
    obj._rowIndex = i + 1; // เลขแถวจริงในชีต (1-based)
    rows.push(obj);
  }
  _readCache[name] = rows;
  if (CACHE_SHEETS[name]) {
    try {
      var j = JSON.stringify(rows);
      if (j.length < 95000) CacheService.getScriptCache().put('sh_' + name, j, 60);
    } catch (e) {}
  }
  return rows;
}

function genId(prefix) {
  return (prefix || 'ID') + '_' + new Date().getTime() + '_' + Math.floor(Math.random() * 1000);
}

// แปลงค่าวันที่ (Date หรือ string) ให้เป็น 'yyyy-MM-dd' เสมอ (กันปัญหา Google Sheets แปลงเป็น Date)
function ymd(v) {
  if (v instanceof Date) return Utilities.formatDate(v, 'GMT+7', 'yyyy-MM-dd');
  return String(v == null ? '' : v).slice(0, 10);
}

function findRowById(name, id) {
  var s = sheet(name);
  var data = s.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) return i + 1; // 1-based row
  }
  return -1;
}


/** ====================== SETTINGS ====================== */

function getSettings() {
  var rows = readAll('Settings');
  var obj = {};
  rows.forEach(function (r) { obj[r.Key] = r.Value; });
  // เติมคีย์ที่อาจขาดด้วยค่าตั้งต้น
  Object.keys(DEFAULT_SETTINGS).forEach(function (k) {
    if (typeof obj[k] === 'undefined') obj[k] = DEFAULT_SETTINGS[k];
  });
  return obj;
}

function saveSettings(data) {
  var s = sheet('Settings');
  s.clearContents();
  var rows = [SHEETS.Settings];
  Object.keys(data).forEach(function (k) {
    if (k === 'action' || k === 'payload') return;
    rows.push([k, String(data[k] == null ? '' : data[k])]);
  });
  // บังคับคอลัมน์ Value เป็นข้อความ ป้องกัน Google Sheets แปลงวันที่/ตัวเลขอัตโนมัติ
  s.getRange(1, 2, rows.length, 1).setNumberFormat('@');
  s.getRange(1, 1, rows.length, 2).setValues(rows);
  SpreadsheetApp.flush();
  return getSettings();
}


/** ====================== GRADE MAPPING + เครื่องมือแปลงเกรด ====================== */

// ก้อนข้อมูลตั้งต้นทั้งหมดในคำขอเดียว (เดิมหน้าเว็บยิง ~7 คำขอตอนเปิด)
function getBootstrapFull(p) {
  p = p || {};
  var out = {
    settings: getSettings(),
    mapping: getGradeMapping(),
    students: getStudents(),
    subjects: getSubjects(),
    calendar: getCalendar(),
    locks: getLocks(),
    teachers: readAll('Teachers').map(function (t) {
      return { name: t.name, role: t.role, homeroomClass: t.homeroomClass || '' };
    })
  };
  var teacher = String(p.teacher || '').trim();
  if (teacher) {
    try { out.schedule = getSchedule(teacher); } catch (e) { out.schedule = []; }
    if (String(p.role || '') === 'teacher') {
      try { out.myTeach = getMyTeach({ teacher: teacher }); } catch (e) { out.myTeach = []; }
    }
  }
  return out;
}

function getGradeMapping() {
  var rows = readAll('GradeMapping').map(function (r) {
    return {
      minPercent: Number(r.minPercent),
      maxPercent: Number(r.maxPercent),
      grade51: r.grade51,
      letter51: r.letter51
    };
  }).filter(function (r) {
    return !isNaN(r.minPercent) && !isNaN(r.maxPercent) && String(r.grade51 == null ? '' : r.grade51) !== '';
  });
  // กันเหนียว: ถ้าชีตยังไม่พร้อม/อ่านไม่ได้ ใช้เกณฑ์มาตรฐาน 2551 เพื่อให้ตัดเกรดได้เสมอ
  return rows.length ? rows : DEFAULT_MAPPING.slice();
}

function saveGradeMapping(mapping) {
  var s = sheet('GradeMapping');
  s.clearContents();
  s.appendRow(SHEETS.GradeMapping);
  (mapping || []).forEach(function (m) {
    s.appendRow([m.minPercent, m.maxPercent, m.grade51, m.letter51]);
  });
  SpreadsheetApp.flush();
  return getGradeMapping();
}

// แกนกลาง: ร้อยละ -> {grade51, letter51} ตามหลักสูตร 2551
function percentToGrade(percent, mapping) {
  mapping = mapping || getGradeMapping();
  var p = Number(percent);
  if (isNaN(p)) return { percent: null, grade51: '', letter51: '' };
  for (var i = 0; i < mapping.length; i++) {
    if (p >= mapping[i].minPercent && p <= mapping[i].maxPercent) {
      return { percent: p, grade51: mapping[i].grade51, letter51: mapping[i].letter51 };
    }
  }
  return { percent: p, grade51: '', letter51: '' };
}


/** ====================== STUDENTS ====================== */

function getStudentsAll_() { ensureStudentCols_(); return readAll('Students'); }
// เฉพาะนักเรียนปัจจุบัน (ไม่รวมย้ายออก/จำหน่าย) — ทุกเมนูปกติใช้ตัวนี้
function getStudents() { return getStudentsAll_().filter(function (s) { return String(s.status || '') !== 'moved'; }); }
function getMovedStudents() {
  return getStudentsAll_().filter(function (s) { return String(s.status || '') === 'moved'; })
    .map(function (s) { s.movedDate = s.movedDate ? ymd(s.movedDate) : ''; return s; })
    .sort(function (a, b) { return String(b.movedDate).localeCompare(String(a.movedDate)); });
}
function assertAdmin_(p) { if (String(p._role || '') !== 'admin') throw new Error('ทำรายการนี้ได้เฉพาะแอดมิน'); }
// จัดเลขที่ใหม่ 1..n เฉพาะนักเรียนปัจจุบันของห้อง (เรียงตามเลขที่เดิม)
function renumberRoom_(room) {
  var s = sheet('Students');
  var data = s.getDataRange().getValues();
  var head = data[0]; var col = {}; head.forEach(function (h, i) { col[h] = i; });
  var rows = [];
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][col.classLevel]).trim() !== String(room).trim()) continue;
    if (String(data[i][col.status] || '') === 'moved') continue;
    rows.push({ row: i + 1, num: Number(data[i][col.number]) || 9999, name: String(data[i][col.fullName] || '') });
  }
  rows.sort(function (a, b) { return a.num - b.num || a.name.localeCompare(b.name); });
  rows.forEach(function (r, ix) { s.getRange(r.row, col.number + 1).setValue(ix + 1); });
  SpreadsheetApp.flush();
  return rows.length;
}
// ย้ายออก/จำหน่าย (เก็บประวัติทั้งหมด) + จัดเลขที่ห้องใหม่อัตโนมัติ
function moveOutStudent(p) {
  assertAdmin_(p);
  var s = sheet('Students');
  var row = findRowById('Students', p.studentID);
  if (row < 0) throw new Error('ไม่พบนักเรียน');
  var data = s.getDataRange().getValues();
  var head = data[0]; var col = {}; head.forEach(function (h, i) { col[h] = i; });
  var rec = data[row - 1];
  if (String(rec[col.status] || '') === 'moved') throw new Error('นักเรียนคนนี้ย้ายออกไปแล้ว');
  var today = Utilities.formatDate(new Date(), 'GMT+7', 'yyyy-MM-dd');
  s.getRange(row, col.status + 1).setValue('moved');
  s.getRange(row, col.movedDate + 1).setValue(today);
  s.getRange(row, col.movedNumber + 1).setValue(rec[col.number] || '');
  SpreadsheetApp.flush();
  var room = String(rec[col.classLevel] || '').trim();
  var remaining = renumberRoom_(room);
  return { name: rec[col.fullName], room: room, remaining: remaining };
}
// กู้คืน (กรณีย้ายผิดคน) — ต่อเลขที่ท้ายห้อง
function restoreStudent(p) {
  assertAdmin_(p);
  var s = sheet('Students');
  var row = findRowById('Students', p.studentID);
  if (row < 0) throw new Error('ไม่พบนักเรียน');
  var data = s.getDataRange().getValues();
  var head = data[0]; var col = {}; head.forEach(function (h, i) { col[h] = i; });
  var room = String(data[row - 1][col.classLevel] || '').trim();
  var maxN = 0;
  getStudents().forEach(function (st) {
    if (String(st.classLevel).trim() === room) maxN = Math.max(maxN, Number(st.number) || 0);
  });
  s.getRange(row, col.status + 1).setValue('');
  s.getRange(row, col.movedDate + 1).setValue('');
  s.getRange(row, col.number + 1).setValue(maxN + 1);
  SpreadsheetApp.flush();
  return { name: data[row - 1][col.fullName], room: room, number: maxN + 1 };
}
// ลบถาวร (เมนู NSR เท่านั้น): ลบนักเรียนที่ย้ายออกพร้อมข้อมูลทั้งหมด
function purgeMovedStudents() {
  var moved = getMovedStudents();
  if (!moved.length) return { purged: 0 };
  var ids = {}; moved.forEach(function (m) { ids[m.ID] = 1; });
  ['Scores', 'Grades', 'Attendance', 'AbilityDetail', 'Remediation'].forEach(function (sn) {
    var sh = ss().getSheetByName(sn); if (!sh) return;
    var d = sh.getDataRange().getValues(); if (d.length < 2) return;
    var c = d[0].indexOf('studentID'); if (c < 0) return;
    for (var i = d.length - 1; i >= 1; i--) { if (ids[d[i][c]]) sh.deleteRow(i + 1); }
  });
  var s = sheet('Students');
  var d2 = s.getDataRange().getValues();
  var idc = d2[0].indexOf('ID');
  for (var j = d2.length - 1; j >= 1; j--) { if (ids[d2[j][idc]]) s.deleteRow(j + 1); }
  SpreadsheetApp.flush();
  return { purged: moved.length };
}

// คำนำหน้าที่รู้จัก (เรียงยาว→สั้น เพื่อจับ "นางสาว" ก่อน "นาง")
var NAME_PREFIXES = ['เด็กชาย', 'เด็กหญิง', 'ด.ช.', 'ด.ญ.', 'นางสาว', 'นาง', 'นาย', 'ว่าที่ร้อยตรี', 'ว่าที่ร้อยตรีหญิง'];
function parseThaiName_(full) {
  full = String(full || '').trim();
  var prefix = '';
  for (var i = 0; i < NAME_PREFIXES.length; i++) {
    if (full.indexOf(NAME_PREFIXES[i]) === 0) { prefix = NAME_PREFIXES[i]; full = full.slice(NAME_PREFIXES[i].length).trim(); break; }
  }
  var parts = full.split(/\s+/).filter(String);
  var first = parts.shift() || '';
  return { prefix: prefix, first: first, last: parts.join(' ') };
}
function buildFullName_(prefix, first, last) {
  return ((prefix || '') + (first || '') + (last ? ' ' + last : '')).trim();
}
// เพิ่มคอลัมน์ใหม่ + แยกชื่อเดิม (ทำครั้งเดียว)
function ensureStudentCols_() {
  var s = sheet('Students');
  var lastCol = s.getLastColumn();
  var header = s.getRange(1, 1, 1, Math.max(lastCol, 1)).getValues()[0];
  // เติมคอลัมน์ใหม่สำหรับระบบย้ายออก (status/movedDate/movedNumber) ถ้ายังไม่มี
  var extra = ['status', 'movedDate', 'movedNumber'].filter(function (h) { return header.indexOf(h) < 0; });
  if (extra.length && header.indexOf('firstName') >= 0) {
    s.getRange(1, lastCol + 1, 1, extra.length).setValues([extra]);
  }
  if (header.indexOf('firstName') >= 0) return; // ย้ายข้อมูลแล้ว
  s.getRange(1, 1, 1, SHEETS.Students.length).setValues([SHEETS.Students]);
  var n = s.getLastRow();
  if (n < 2) return;
  var rng = s.getRange(2, 1, n - 1, SHEETS.Students.length);
  var vals = rng.getValues();
  vals.forEach(function (r) {
    if (!r[6] && !r[7] && !r[8]) {
      var p = parseThaiName_(r[3]);
      r[6] = p.prefix; r[7] = p.first; r[8] = p.last;
    }
  });
  rng.setValues(vals);
  SpreadsheetApp.flush();
}

function addStudent(p) {
  ensureStudentCols_();
  var s = sheet('Students');
  var id = genId('STD');
  var full = p.fullName || buildFullName_(p.prefix, p.firstName, p.lastName);
  var parts = (p.firstName || p.lastName || p.prefix) ? { prefix: p.prefix || '', first: p.firstName || '', last: p.lastName || '' } : parseThaiName_(full);
  s.appendRow([
    id, p.studentCode || '', p.number || '', full, p.classLevel || '',
    Utilities.formatDate(new Date(), 'GMT+7', 'yyyy-MM-dd HH:mm:ss'),
    parts.prefix, parts.first, parts.last
  ]);
  SpreadsheetApp.flush();
  return { id: id };
}

function updateStudent(id, p) {
  ensureStudentCols_();
  var row = findRowById('Students', id);
  if (row < 0) throw new Error('ไม่พบนักเรียน ID: ' + id);
  var s = sheet('Students');
  var parts = (p.firstName || p.lastName || p.prefix) ? { prefix: p.prefix || '', first: p.firstName || '', last: p.lastName || '' } : parseThaiName_(p.fullName);
  var full = buildFullName_(parts.prefix, parts.first, parts.last) || (p.fullName || '');
  s.getRange(row, 2, 1, 4).setValues([[p.studentCode || '', p.number || '', full, p.classLevel || '']]);
  s.getRange(row, 7, 1, 3).setValues([[parts.prefix, parts.first, parts.last]]);
  SpreadsheetApp.flush();
  return { id: id };
}

function deleteStudent(id) {
  var row = findRowById('Students', id);
  if (row < 0) throw new Error('ไม่พบนักเรียน ID: ' + id);
  // จำชั้นของนักเรียนคนนี้ไว้ก่อนลบ เพื่อเลื่อนเลขที่
  var s = sheet('Students');
  var classLevel = s.getRange(row, 5).getValue(); // คอลัมน์ classLevel
  s.deleteRow(row);
  // ลบคะแนน/เกรด/ความสามารถที่ผูกกับนักเรียนคนนี้ (ใช้ clearContent ปลอดภัยกว่า)
  clearRowsWhere('Scores', 'studentID', id);
  clearRowsWhere('Grades', 'studentID', id);
  clearRowsWhere('Attendance', 'studentID', id);
  SpreadsheetApp.flush();
  if (classLevel) renumberClass(classLevel); // เลื่อนเลขที่ด้านหลังขึ้นอัตโนมัติ
  return { id: id };
}

// เรียงเลขที่ของห้องให้ต่อเนื่อง 1..N (ตามลำดับเลขที่เดิม)
function renumberClass(classLevel) {
  var s = sheet('Students');
  var data = s.getDataRange().getValues();
  var items = [];
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][4]) === String(classLevel) && data[i].join('') !== '') {
      items.push({ row: i + 1, number: Number(data[i][2]) || 0 });
    }
  }
  items.sort(function (a, b) { return a.number - b.number; });
  items.forEach(function (it, idx) {
    s.getRange(it.row, 3).setValue(idx + 1); // คอลัมน์ number
  });
  SpreadsheetApp.flush();
  return { classLevel: classLevel, count: items.length };
}

// นำเข้ารายชื่อแบบ upsert + กันซ้ำในไฟล์ (จับคู่ด้วยรหัสนักเรียน หรือ ชื่อ+ชั้น)
function bulkImportStudents(rows) {
  if (!rows || !rows.length) return { count: 0, added: 0, updated: 0 };
  ensureStudentCols_();
  var s = sheet('Students');
  var data = s.getDataRange().getValues();
  var head = data[0]; var col = {}; head.forEach(function (h, i) { col[h] = i; });
  var now = Utilities.formatDate(new Date(), 'GMT+7', 'yyyy-MM-dd HH:mm:ss');
  var byCode = {}, byNameClass = {};
  for (var i = 1; i < data.length; i++) {
    var code = String(data[i][col.studentCode] || '').trim();
    var full = String(data[i][col.fullName] || '').trim();
    var cls = String(data[i][col.classLevel] || '').trim();
    if (code) byCode[code] = i + 1;
    if (full) byNameClass[full + '|' + cls] = i + 1;
  }
  var added = 0, updated = 0, seen = {};
  rows.forEach(function (r) {
    var code = String(r.studentCode || '').trim();
    var full = String(r.fullName || '').trim();
    var cls = String(r.classLevel || '').trim();
    if (!full && !code) return;
    var dkey = code ? ('C:' + code) : ('N:' + full + '|' + cls);
    if (seen[dkey]) return; seen[dkey] = true;          // กันซ้ำภายในไฟล์นำเข้า
    var parts = parseThaiName_(full);
    var rowIdx = (code && byCode[code]) || byNameClass[full + '|' + cls];
    if (rowIdx) {
      s.getRange(rowIdx, col.studentCode + 1).setValue(code);
      s.getRange(rowIdx, col.number + 1).setValue(r.number || '');
      s.getRange(rowIdx, col.fullName + 1).setValue(full);
      s.getRange(rowIdx, col.classLevel + 1).setValue(cls);
      s.getRange(rowIdx, col.prefix + 1, 1, 3).setValues([[parts.prefix, parts.first, parts.last]]);
      updated++;
    } else {
      var id = genId('STD');
      var newRow = head.map(function () { return ''; });
      newRow[col.ID] = id; newRow[col.studentCode] = code; newRow[col.number] = r.number || '';
      newRow[col.fullName] = full; newRow[col.classLevel] = cls; newRow[col.dateAdded] = now;
      newRow[col.prefix] = parts.prefix; newRow[col.firstName] = parts.first; newRow[col.lastName] = parts.last;
      s.appendRow(newRow);
      var last = s.getLastRow();
      if (code) byCode[code] = last;
      byNameClass[full + '|' + cls] = last;
      added++;
    }
  });
  SpreadsheetApp.flush();
  return { count: added + updated, added: added, updated: updated };
}

// ล้างนักเรียนซ้ำ (รหัสเดียวกัน หรือ ชื่อ+ชั้นเดียวกัน) — ย้ายคะแนน/เกรด/เช็คชื่อ ไปยังคนที่เก็บไว้
function dedupeStudents() {
  ensureStudentCols_();
  var s = sheet('Students');
  var data = s.getDataRange().getValues(); if (data.length < 2) return { removed: 0 };
  var head = data[0]; var col = {}; head.forEach(function (h, i) { col[h] = i; });
  var groups = {};
  for (var i = 1; i < data.length; i++) {
    var code = String(data[i][col.studentCode] || '').trim();
    var full = String(data[i][col.fullName] || '').trim();
    var cls = String(data[i][col.classLevel] || '').trim();
    if (!full && !code) continue;
    var key = code ? ('C:' + code) : ('N:' + full + '|' + cls);
    (groups[key] = groups[key] || []).push({ row: i + 1, id: data[i][col.ID] });
  }
  var idMap = {}, removeRows = [];
  Object.keys(groups).forEach(function (k) {
    var arr = groups[k]; if (arr.length < 2) return;
    var keeper = arr[0];
    arr.forEach(function (x) { if (x.id !== keeper.id) { idMap[x.id] = keeper.id; removeRows.push(x.row); } });
  });
  if (!removeRows.length) return { removed: 0 };
  ['Scores', 'Grades', 'AbilityDetail', 'Attendance'].forEach(function (sn) {
    var sh = ss().getSheetByName(sn); if (!sh) return;
    var d = sh.getDataRange().getValues(); if (d.length < 2) return;
    var c = d[0].indexOf('studentID'); if (c < 0) return;
    var changed = false;
    for (var i = 1; i < d.length; i++) { var old = d[i][c]; if (idMap[old]) { d[i][c] = idMap[old]; changed = true; } }
    if (changed) sh.getRange(1, 1, d.length, d[0].length).setValues(d);
  });
  removeRows.sort(function (a, b) { return b - a; }).forEach(function (r) { s.deleteRow(r); });
  SpreadsheetApp.flush();
  return { removed: removeRows.length };
}


/** ====================== SUBJECTS ====================== */

function getSubjects() { return readAll('Subjects'); }

function addSubject(p) {
  var s = sheet('Subjects');
  var id = genId('SUB');
  s.appendRow([
    id,
    p.code || '',
    p.name || '',
    p.learningArea || '',
    p.hours || '',
    p.type || 'พื้นฐาน',
    p.classLevel || '',
    p.sortOrder || (s.getLastRow())
  ]);
  SpreadsheetApp.flush();
  return { id: id };
}

function updateSubject(id, p) {
  var row = findRowById('Subjects', id);
  if (row < 0) throw new Error('ไม่พบรายวิชา ID: ' + id);
  var s = sheet('Subjects');
  s.getRange(row, 2, 1, 7).setValues([[
    p.code || '', p.name || '', p.learningArea || '', p.hours || '',
    p.type || 'พื้นฐาน', p.classLevel || '', p.sortOrder || ''
  ]]);
  SpreadsheetApp.flush();
  return { id: id };
}

function deleteSubject(id) {
  var row = findRowById('Subjects', id);
  if (row < 0) throw new Error('ไม่พบรายวิชา ID: ' + id);
  sheet('Subjects').deleteRow(row);
  clearRowsWhere('LearningUnits', 'subjectID', id);
  clearRowsWhere('Scores', 'subjectID', id);
  clearRowsWhere('Grades', 'subjectID', id);
  SpreadsheetApp.flush();
  return { id: id };
}


/** ====================== LEARNING UNITS (หน่วย/ผลลัพธ์ ตั้งชื่ออิสระ) ====================== */

function getUnits(subjectID, semester) {
  var sem = semester ? String(semester) : '';
  return readAll('LearningUnits')
    .filter(function (u) {
      if (String(u.subjectID) !== String(subjectID)) return false;
      if (sem && u.semester && String(u.semester) !== sem) return false;
      return true;
    })
    .sort(function (a, b) { return Number(a.sortOrder) - Number(b.sortOrder); });
}

// แทนที่หน่วยของรายวิชานี้ "เฉพาะภาคเรียนนั้น" ด้วยชุดใหม่
function saveUnits(subjectID, units, semester) {
  var sem = String(semester || '1');
  clearRowsWhere2('LearningUnits', 'subjectID', subjectID, 'semester', sem);
  var s = sheet('LearningUnits');
  (units || []).forEach(function (u, i) {
    s.appendRow([
      u.ID || genId('UNT'),
      subjectID,
      u.unitName || ('หน่วยที่ ' + (i + 1)),
      u.maxScore || 0,
      i + 1,
      sem
    ]);
  });
  SpreadsheetApp.flush();
  return getUnits(subjectID, sem);
}


/** ====================== SCORES (โหมด A) ====================== */

function getScores(subjectID, semester) {
  var sem = semester ? String(semester) : '';
  return readAll('Scores')
    .filter(function (r) {
      if (String(r.subjectID) !== String(subjectID)) return false;
      if (sem && r.semester && String(r.semester) !== sem) return false;
      return true;
    });
}

// บันทึกคะแนนรายหน่วยของรายวิชานี้ "เฉพาะภาคเรียนนั้น" (แทนที่) + คำนวณเกรดให้อัตโนมัติ
// scores: [{studentID, unitID, score}, ...]
function saveScores(subjectID, scores, semester) {
  var sem = String(semester || '1');
  assertNotLocked_((scores || []).map(function (r) { return r.studentID; }));
  clearRowsWhere2('Scores', 'subjectID', subjectID, 'semester', sem);
  var s = sheet('Scores');
  (scores || []).forEach(function (r) {
    s.appendRow([genId('SCR'), r.studentID, subjectID, r.unitID, r.score, sem]);
  });
  SpreadsheetApp.flush();
  return recomputeGrades(subjectID, sem);
}

// คำนวณเกรด (โหมด A) จากคะแนนรายหน่วยของรายวิชา+ภาคเรียน -> ร้อยละ -> เกรด 2551
function recomputeGrades(subjectID, semester) {
  var sem = String(semester || '1');
  var units = getUnits(subjectID, sem);
  var totalMax = units.reduce(function (sum, u) { return sum + Number(u.maxScore || 0); }, 0);
  var scores = getScores(subjectID, sem);
  var mapping = getGradeMapping();

  var byStudent = {};
  scores.forEach(function (sc) {
    var sid = sc.studentID;
    if (!byStudent[sid]) byStudent[sid] = 0;
    byStudent[sid] += Number(sc.score || 0);
  });

  var results = [];
  Object.keys(byStudent).forEach(function (sid) {
    var total = byStudent[sid];
    var percent = totalMax > 0 ? Math.round((total / totalMax) * 100) : 0;
    var g = percentToGrade(percent, mapping);
    writeGrade({
      studentID: sid, subjectID: subjectID, mode: 'A', semester: sem,
      percent: percent, grade51: g.grade51, letter51: g.letter51, note: ''
    });
    results.push({ studentID: sid, total: total, totalMax: totalMax, percent: percent,
                   grade51: g.grade51, letter51: g.letter51, semester: sem });
  });
  SpreadsheetApp.flush();
  return results;
}


/** ====================== GRADES ====================== */

function getGrades(semester) {
  var rows = readAll('Grades');
  if (!semester) return rows;
  var sem = String(semester);
  return rows.filter(function (g) { return !g.semester || String(g.semester) === sem; });
}

// บันทึกเกรด 1 รายการ (โหมด B: กรอกร้อยละเอง / override เกรด)
function saveGrade(p) {
  var mapping = getGradeMapping();
  var sem      = String(p.semester || '1');
  var percent  = (p.percent !== '' && p.percent != null) ? Number(p.percent) : null;
  var grade51  = p.grade51;
  var letter51 = p.letter51;

  if (percent != null && !p.manualOverride) {
    var g = percentToGrade(percent, mapping);
    grade51 = g.grade51; letter51 = g.letter51;
  }
  writeGrade({
    studentID: p.studentID, subjectID: p.subjectID, mode: p.mode || 'B', semester: sem,
    percent: percent, grade51: grade51, letter51: letter51, note: p.note || ''
  });
  SpreadsheetApp.flush();
  return { ok: true };
}

// บันทึกเกรดหลายคนในครั้งเดียว (โหมด B ทั้งห้อง)
function saveGradesBatch(grades, semester) {
  assertNotLocked_((grades || []).map(function (p) { return p.studentID; }));
  var mapping = getGradeMapping();
  (grades || []).forEach(function (p) {
    var sem      = String(p.semester || semester || '1');
    var percent  = (p.percent !== '' && p.percent != null) ? Number(p.percent) : null;
    var grade51  = p.grade51;
    var letter51 = p.letter51;
    if (percent != null && !p.manualOverride) {
      var g = percentToGrade(percent, mapping);
      grade51 = g.grade51; letter51 = g.letter51;
    }
    writeGrade({
      studentID: p.studentID, subjectID: p.subjectID, mode: p.mode || 'B', semester: sem,
      percent: percent, grade51: grade51, letter51: letter51, note: p.note || ''
    });
  });
  SpreadsheetApp.flush();
  return { count: (grades || []).length };
}

// upsert เกรดของ studentID+subjectID+semester
function writeGrade(g) {
  var s = sheet('Grades');
  var sem = String(g.semester || '1');
  var data = s.getDataRange().getValues();
  var found = -1;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][1]) === String(g.studentID) &&
        String(data[i][2]) === String(g.subjectID) &&
        String(data[i][9] || '1') === sem) {
      found = i + 1; break;
    }
  }
  var rowVals = [
    found > 0 ? data[found - 1][0] : genId('GRD'),
    g.studentID, g.subjectID, g.mode,
    (g.percent == null ? '' : g.percent),
    '', g.grade51, g.letter51, g.note || '', sem
  ];
  if (found > 0) s.getRange(found, 1, 1, SHEETS.Grades.length).setValues([rowVals]);
  else           s.appendRow(rowVals);
}


/** ====================== WP1: การประเมินละเอียด (รายตัวชี้วัด) ====================== */
// scale: 0=ไม่ผ่าน, 1=ผ่าน, 2=ดี, 3=ดีเยี่ยม
function scoreToLevel(avg) {
  if (avg >= 2.5) return 'ดีเยี่ยม';
  if (avg >= 1.5) return 'ดี';
  if (avg >= 0.5) return 'ผ่าน';
  return 'ไม่ผ่าน';
}

// ดึงผลประเมินละเอียดของนักเรียนทั้งห้อง (สำหรับหน้ากรอกแบบตาราง)
function getAbilityDetailClass(classLevel, semester) {
  var ids = {};
  getStudents().forEach(function (s) { if (String(s.classLevel) === String(classLevel)) ids[s.ID] = true; });
  var sem = semester ? String(semester) : '';
  return readAll('AbilityDetail').filter(function (r) {
    if (!ids[r.studentID]) return false;
    if (sem && r.semester && String(r.semester) !== sem) return false;
    return true;
  });
}

// บันทึกผลประเมินละเอียดแบบกลุ่ม "เฉพาะภาคเรียนนั้น"
// p: { studentIDs:[...], rows:[{studentID, group, item, value}], semester }
function saveAbilityDetailBatch(p) {
  assertNotLocked_(p.studentIDs || []);
  var sem = String(p.semester || '1');
  var ids = {}; (p.studentIDs || []).forEach(function (id) { ids[id] = true; });
  var s = sheet('AbilityDetail');
  var data = s.getDataRange().getValues();
  var semCol = data[0].indexOf('semester');
  var keep = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i].join('') === '') continue;
    var rowSem = semCol >= 0 ? String(data[i][semCol] || '1') : '1';
    // คงไว้ ถ้าไม่ใช่ของนักเรียนชุดนี้ หรือเป็นคนละภาคเรียน
    if (!ids[data[i][1]] || rowSem !== sem) keep.push(data[i]);
  }
  (p.rows || []).forEach(function (r) {
    if (r.value === '' || r.value == null) return;
    keep.push([genId('ABD'), r.studentID, r.group, r.item, r.value, sem]);
  });
  if (s.getLastRow() > 1) s.getRange(2, 1, s.getLastRow() - 1, SHEETS.AbilityDetail.length).clearContent();
  if (keep.length) s.getRange(2, 1, keep.length, SHEETS.AbilityDetail.length).setValues(keep);
  SpreadsheetApp.flush();
  return { count: (p.rows || []).length };
}

// สรุปผลการประเมินจากรายตัวชี้วัด → object {itemKey: result} สำหรับใบรายงาน
function summarizeAbilityDetail(rows) {
  var ab = {};
  var cVals = [], rVals = [];
  (rows || []).forEach(function (r) {
    if (r.group === 'characteristic') cVals.push(String(r.value || ''));
    else if (r.group === 'readthink') rVals.push(String(r.value || ''));
    else if (r.group === 'activity' || r.group === 'competency' || r.group === 'competency5') ab[r.item] = r.value;
  });
  // ผ่าน/ไม่ผ่าน: มีข้อใดไม่ผ่าน = ไม่ผ่าน, มีค่าครบและผ่านหมด = ผ่าน
  var passOf = function (vals) {
    var has = vals.some(function (v) { return v !== ''; });
    if (!has) return '';
    return vals.some(function (v) { return v === 'ไม่ผ่าน'; }) ? 'ไม่ผ่าน' : 'ผ่าน';
  };
  var c = passOf(cVals), r = passOf(rVals);
  if (c) ab['คุณลักษณะอันพึงประสงค์'] = c;
  if (r) ab['การอ่าน คิดวิเคราะห์ และเขียน'] = r;
  return ab;
}

/** ====================== แจ้งเตือนนักเรียนเสี่ยง ====================== */
// วันที่ที่ "เช็คชื่อหน้าเสาธง" แล้ว ของห้องหนึ่งในเดือนที่กำหนด → ['YYYY-MM-DD', ...]
// p: { classLevel, month:'YYYY-MM' }
// สถานะเช็คชื่อรายคาบทั้งเดือน (เทียบตารางสอน) — teacher ว่าง = รวมทุกคาบทุกห้อง (แอดมิน)
// คืน { days: { 'YYYY-MM-DD': { exp, done } } }
// p: { month, teacher?, classLevel?, subject? } — ถ้าระบุ classLevel+subject จะดูเฉพาะคาบของวิชานั้นในห้องนั้น (ทุกครู)
function getPeriodCheckedDates(p) {
  p = p || {};
  var month = String(p.month || '');
  if (!month) return { days: {} };
  var teacher = String(p.teacher || '').trim();
  var tc = teacher ? teacherCore(teacher) : '';
  var fCls = String(p.classLevel || '').trim(), fSub = String(p.subject || '').trim();
  var TH_DAYS = ['อาทิตย์', 'จันทร์', 'อังคาร', 'พุธ', 'พฤหัสบดี', 'ศุกร์', 'เสาร์'];
  var byWd = {};   // weekday -> [{cls, per}]
  readAll('Schedule').forEach(function (r) {
    var cls = String(r.classLevel || '').trim(), sub = String(r.subject || '').trim();
    if (fCls && fSub) { if (cls !== fCls || sub !== fSub) return; }
    else if (tc && teacherCore(r.teacher) !== tc) return;
    var d = String(r.day || '').trim(), per = Number(r.period);
    if (!d || !cls || !per) return;
    (byWd[d] = byWd[d] || []).push({ cls: cls, per: String(per) });
  });
  var chk = {};   // 'date|cls|per' -> true
  readAll('Attendance').forEach(function (a) {
    var per = Number(a.period);
    if (!per) return;
    var d = ymd(a.date);
    if (d.indexOf(month) !== 0) return;
    chk[d + '|' + String(a.classLevel).trim() + '|' + String(per)] = true;
  });
  var mp = month.split('-');
  var nDays = new Date(Number(mp[0]), Number(mp[1]), 0).getDate();
  var days = {};
  for (var d2 = 1; d2 <= nDays; d2++) {
    var iso = month + '-' + (d2 < 10 ? '0' : '') + d2;
    var wd = TH_DAYS[new Date(Number(mp[0]), Number(mp[1]) - 1, d2).getDay()];
    var slots = byWd[wd] || [];
    var done = 0, checked = [];
    slots.forEach(function (sl) {
      if (chk[iso + '|' + sl.cls + '|' + sl.per]) { done++; checked.push(sl.cls + '|' + sl.per); }
    });
    days[iso] = { exp: slots.length, done: done, checked: checked };
  }
  return { days: days };
}

function getFlagCheckedDates(p) {
  p = p || {};
  var cls = String(p.classLevel || ''); var month = String(p.month || '');
  if (!cls || !month) return [];
  var set = {};
  readAll('Attendance').forEach(function (r) {
    if (String(r.period) !== '0') return;
    if (String(r.classLevel) !== cls) return;
    var d = ymd(r.date);
    if (d.indexOf(month) === 0) set[d] = true;
  });
  return Object.keys(set);
}

/* ============================================================
   ส่งเอกสารการสอน (กำหนดการสอน + แผนการสอน) — ยึดตามตารางสอน
   docType: 'กำหนดการสอน' | 'แผนการสอน'
   ============================================================ */
var TEACHING_DOC_TYPES = ['กำหนดการสอน', 'แผนการสอน'];  // ค่าตั้งต้น (ระบบจริงใช้ SubmissionTypes)
var DEFAULT_SUB_TYPES = ['กำหนดการสอน', 'แผนการสอน', 'ข้อสอบกลางภาค', 'ข้อสอบปลายภาค', 'แบบวิเคราะห์ผู้เรียนรายบุคคล'];

// ประเภทงานที่ต้องส่ง (แอดมินกำหนดเอง) — seed ค่าตั้งต้น 4 ประเภทเมื่อยังว่าง
function getSubmissionTypes(p) {
  p = p || {};
  var sem = String(p.semester || getSettings().currentSemester || '1');
  var s = sheet('SubmissionTypes');
  if (s.getLastRow() < 2) {
    DEFAULT_SUB_TYPES.forEach(function (n) { s.appendRow([genId('STY'), n, '', '', '', '1']); });
    SpreadsheetApp.flush();
  } else {
    // เติมประเภทงานตั้งต้นที่เพิ่มใหม่ภายหลัง (ถ้ายังไม่มี)
    var have = {};
    readAll('SubmissionTypes').forEach(function (t) { have[String(t.name).trim()] = true; });
    var added = false;
    DEFAULT_SUB_TYPES.forEach(function (n) { if (!have[n]) { s.appendRow([genId('STY'), n, '', '', '', '1']); added = true; } });
    if (added) SpreadsheetApp.flush();
  }
  var today = Utilities.formatDate(new Date(), 'GMT+7', 'yyyy-MM-dd');
  return readAll('SubmissionTypes')
    .filter(function (t) { return String(t.active) !== '0' && (!t.semester || String(t.semester) === sem); })
    .map(function (t) {
      var start = t.startDate ? ymd(t.startDate) : '';
      var end = t.endDate ? ymd(t.endDate) : '';
      var open = (!start || today >= start) && (!end || today <= end);
      return { id: t.ID, name: t.name, semester: String(t.semester || ''), startDate: start, endDate: end, open: open };
    });
}
function getAllSubmissionTypes() {
  getSubmissionTypes({});  // ensure seeded
  return readAll('SubmissionTypes').map(function (t) {
    return { id: t.ID, name: t.name, semester: String(t.semester || ''), startDate: t.startDate ? ymd(t.startDate) : '', endDate: t.endDate ? ymd(t.endDate) : '', active: String(t.active) !== '0' };
  });
}
function saveSubmissionType(p) {
  var s = sheet('SubmissionTypes');
  if (!String(p.name || '').trim()) throw new Error('กรุณาระบุชื่องาน');
  var row = p.id ? findRowById('SubmissionTypes', p.id) : -1;
  var vals = [p.id || genId('STY'), String(p.name).trim(), String(p.semester || ''), p.startDate || '', p.endDate || '', p.active === false ? '0' : '1'];
  if (row > 0) s.getRange(row, 1, 1, vals.length).setValues([vals]);
  else s.appendRow(vals);
  SpreadsheetApp.flush();
  return getAllSubmissionTypes();
}
function deleteSubmissionType(p) {
  var row = findRowById('SubmissionTypes', p.id);
  if (row > 0) sheet('SubmissionTypes').deleteRow(row);
  SpreadsheetApp.flush();
  return getAllSubmissionTypes();
}

// ระดับชั้นจากห้อง: "ม.1/1" → "ม.1" (วิชาเดียวกันต่างห้องในชั้นเดียวกันใช้กำหนดการ/แผนชุดเดียว)
function gradeOf_(c) { return String(c || '').split('/')[0].trim(); }

// รายการที่ครูต้องส่ง = วิชา×ชั้น (ระดับชั้น ไม่แยกห้อง) ที่ปรากฏในตารางสอนของครูคนนั้น (distinct)
function getTeachingTasks(p) {
  p = p || {};
  var sem = String(p.semester || getSettings().currentSemester || '1');
  var teacher = p.teacher || '';
  var types = getSubmissionTypes({ semester: sem });
  if (!teacher) return { semester: sem, items: [], types: types };
  var core = teacherCore(teacher);
  var sched = readAll('Schedule').filter(function (r) {
    return teacherCore(r.teacher) === core || normName(r.teacher) === normName(teacher);
  });
  // distinct subject|ระดับชั้น (รวมห้อง)
  var seen = {}, items = [];
  sched.forEach(function (r) {
    var sub = String(r.subject || '').trim(); var grade = gradeOf_(r.classLevel);
    if (!sub || !grade) return;
    var key = sub + '|' + grade;
    if (seen[key]) return; seen[key] = true;
    items.push({ subject: sub, classLevel: grade });
  });
  // join สถานะการส่งจาก Submissions (เทียบที่ระดับชั้น รองรับข้อมูลเดิมที่เคยเก็บเป็นห้อง)
  var subs = readAll('Submissions').filter(function (s) {
    return (teacherCore(s.teacher) === core || normName(s.teacher) === normName(teacher)) && String(s.semester) === sem;
  });
  var smap = {};
  subs.forEach(function (s) { smap[s.subject + '|' + gradeOf_(s.classLevel) + '|' + s.docType] = s; });
  items.forEach(function (it) {
    it.docs = {};
    types.forEach(function (t) {
      var rec = smap[it.subject + '|' + it.classLevel + '|' + t.name];
      it.docs[t.name] = rec ? { id: rec.ID, fileName: rec.fileName, fileUrl: rec.fileUrl, fileId: rec.fileId, updatedAt: rec.updatedAt } : null;
    });
  });
  items.sort(function (a, b) { return String(a.classLevel).localeCompare(String(b.classLevel)) || String(a.subject).localeCompare(String(b.subject)); });
  return { semester: sem, teacher: teacher, items: items, types: types };
}

// อัปโหลด PDF 1 ไฟล์ → Drive แล้ว upsert แถวใน Submissions
// p: { teacher, subject, classLevel, docType, fileName, base64, semester }
function uploadTeachingDoc(p) {
  if (!p.base64) throw new Error('ไม่มีไฟล์');
  var sem = String(p.semester || getSettings().currentSemester || '1');
  var types = getSubmissionTypes({ semester: sem });
  var ty = types.filter(function (t) { return t.name === p.docType; })[0];
  if (!ty) throw new Error('ประเภทเอกสารไม่ถูกต้อง');
  // เส้นตาย: ครูส่งได้เฉพาะในช่วงที่กำหนด — แอดมินส่ง/ส่งแทนได้ตลอด
  if (!p.asAdmin && !ty.open) {
    throw new Error(ty.startDate && Utilities.formatDate(new Date(), 'GMT+7', 'yyyy-MM-dd') < ty.startDate
      ? 'ยังไม่ถึงช่วงเปิดรับ (' + ty.startDate + ' ถึง ' + (ty.endDate || 'ไม่กำหนด') + ')'
      : 'เกินกำหนดส่งแล้ว (สิ้นสุด ' + ty.endDate + ') — ติดต่อแอดมินหากต้องการส่งย้อนหลัง');
  }
  var grade = gradeOf_(p.classLevel);   // เก็บที่ระดับชั้น (ม.1) ไม่แยกห้อง
  var m = String(p.base64).match(/^data:(.+);base64,(.*)$/);
  var ct = m ? m[1] : 'application/pdf';
  var b64 = m ? m[2] : p.base64;
  var folder = getOrCreateFolder('เอกสารการสอน - ' + (getSettings().schoolName || ''));
  var blob = Utilities.newBlob(Utilities.base64Decode(b64), ct, p.fileName || ('teach_' + new Date().getTime() + '.pdf'));
  var file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  var fileId = file.getId();
  var fileUrl = 'https://drive.google.com/file/d/' + fileId + '/view';
  var now = Utilities.formatDate(new Date(), 'GMT+7', 'yyyy-MM-dd HH:mm:ss');

  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var s = sheet('Submissions');
    var data = s.getDataRange().getValues();
    var head = data[0];
    var col = {}; head.forEach(function (h, i) { col[h] = i; });
    var core = teacherCore(p.teacher);
    var foundRow = -1;
    for (var i = 1; i < data.length; i++) {
      if (teacherCore(data[i][col.teacher]) === core &&
        String(data[i][col.subject]).trim() === String(p.subject).trim() &&
        gradeOf_(data[i][col.classLevel]) === grade &&
        String(data[i][col.docType]) === String(p.docType) &&
        String(data[i][col.semester]) === sem) { foundRow = i + 1; break; }
    }
    if (foundRow > 0) {
      // ลบไฟล์เดิมทิ้ง (กันขยะ) แล้วอัปเดต
      var oldId = data[foundRow - 1][col.fileId];
      if (oldId) { try { DriveApp.getFileById(oldId).setTrashed(true); } catch (e) {} }
      s.getRange(foundRow, 1, 1, head.length).setValues([[
        data[foundRow - 1][col.ID], p.teacher, p.subject, grade, p.docType,
        p.fileName || '', fileId, fileUrl, 'ส่งแล้ว', now, sem
      ]]);
    } else {
      s.appendRow([genId('SUB'), p.teacher, p.subject, grade, p.docType,
        p.fileName || '', fileId, fileUrl, 'ส่งแล้ว', now, sem]);
    }
    SpreadsheetApp.flush();
  } finally { lock.releaseLock(); }
  return getTeachingTasks({ teacher: p.teacher, semester: sem });
}

function deleteTeachingDoc(p) {
  var row = findRowById('Submissions', p.id);
  if (row < 0) throw new Error('ไม่พบเอกสาร');
  var rec = readAll('Submissions').filter(function (x) { return x.ID === p.id; })[0];
  if (rec && rec.fileId) { try { DriveApp.getFileById(rec.fileId).setTrashed(true); } catch (e) {} }
  sheet('Submissions').deleteRow(row);
  SpreadsheetApp.flush();
  return getTeachingTasks({ teacher: (rec ? rec.teacher : p.teacher), semester: (rec ? rec.semester : p.semester) });
}

// ภาพรวมการส่งของครูทุกคน (แอดมิน/ผู้บริหาร)
function getTeachingOverview(p) {
  p = p || {};
  var sem = String(p.semester || getSettings().currentSemester || '1');
  var sched = readAll('Schedule');
  var subs = readAll('Submissions').filter(function (s) { return String(s.semester) === sem; });
  var smap = {};
  subs.forEach(function (s) { smap[teacherCore(s.teacher) + '|' + s.subject + '|' + gradeOf_(s.classLevel) + '|' + s.docType] = s; });

  // จัดกลุ่มงานตามครู (ระดับชั้น ไม่แยกห้อง)
  var byTeacher = {};
  sched.forEach(function (r) {
    var t = String(r.teacher || '').trim(); var sub = String(r.subject || '').trim(); var cls = gradeOf_(r.classLevel);
    if (!t || !sub || !cls) return;
    var core = teacherCore(t);
    if (!byTeacher[core]) byTeacher[core] = { teacher: t, items: {}, order: [] };
    var key = sub + '|' + cls;
    if (!byTeacher[core].items[key]) { byTeacher[core].items[key] = { subject: sub, classLevel: cls }; byTeacher[core].order.push(key); }
  });

  var types = getSubmissionTypes({ semester: sem });
  var typeNames = types.map(function (t) { return t.name; });

  var rows = Object.keys(byTeacher).map(function (core) {
    var g = byTeacher[core];
    var items = g.order.map(function (key) {
      var it = g.items[key];
      var rec = {};
      typeNames.forEach(function (dt) {
        var s = smap[core + '|' + it.subject + '|' + it.classLevel + '|' + dt];
        rec[dt] = s ? { fileUrl: s.fileUrl, updatedAt: s.updatedAt } : null;
      });
      return { subject: it.subject, classLevel: it.classLevel, docs: rec };
    });
    var total = items.length;
    var byType = {}, completeAll = total > 0;
    typeNames.forEach(function (dt) {
      var done = items.filter(function (x) { return x.docs[dt]; }).length;
      byType[dt] = { done: done, complete: total > 0 && done === total };
      if (!(total > 0 && done === total)) completeAll = false;
    });
    return { teacher: g.teacher, total: total, items: items, byType: byType, complete: completeAll };
  });
  rows.sort(function (a, b) { return String(a.teacher).localeCompare(String(b.teacher)); });

  // มุมมองรายวิชา: วิชาไหน มีครูคนไหนรับผิดชอบ ส่งแล้ว/ยังไม่ส่ง (รายประเภท)
  var bySubject = {};
  rows.forEach(function (r) {
    r.items.forEach(function (it) {
      var k = it.classLevel + '|' + it.subject;
      if (!bySubject[k]) bySubject[k] = { subject: it.subject, classLevel: it.classLevel, teachers: [] };
      var docs = {};
      typeNames.forEach(function (dt) { var d = it.docs[dt]; docs[dt] = d ? { url: d.fileUrl } : null; });
      bySubject[k].teachers.push({ teacher: r.teacher, docs: docs });
    });
  });
  var subjects = Object.keys(bySubject).map(function (k) { return bySubject[k]; })
    .sort(function (a, b) { return String(a.classLevel).localeCompare(String(b.classLevel)) || String(a.subject).localeCompare(String(b.subject)); });

  var tTotal = rows.reduce(function (a, r) { return a + r.total; }, 0);
  var totalsByType = {};
  typeNames.forEach(function (dt) {
    totalsByType[dt] = {
      itemsDone: rows.reduce(function (a, r) { return a + (r.byType[dt] ? r.byType[dt].done : 0); }, 0),
      teachersDone: rows.filter(function (r) { return r.byType[dt] && r.byType[dt].complete; }).length
    };
  });
  return {
    semester: sem, teachers: rows, subjects: subjects, types: types,
    totals: {
      teachers: rows.length, tasks: tTotal, byType: totalsByType,
      teachersComplete: rows.filter(function (r) { return r.complete; }).length,
      teachersIncomplete: rows.filter(function (r) { return !r.complete; }).length
    }
  };
}

/* ============================================================
   ระบบ 0/ร/มส/มผ + ติดตามสอบแก้ตัว (หลักสูตร 2551)
   - รายการรอแก้ไข = เกรดในชีต Grades ที่เป็น 0/ร/มส/มผ
   - แก้ไขแล้ว = บันทึกใน Remediation + เขียนเกรดใหม่ทับใน Grades
   เกณฑ์เกรดใหม่: 0 → ไม่เกิน 1 | ร → ตามจริง 0-4 | มส → ไม่เกิน 1 | มผ → ผ
   ============================================================ */
var SPECIAL_MARKS = ['0', 'ร', 'มส', 'มผ'];
function remedAllowed_(orig) {
  orig = String(orig);
  if (orig === 'ร') return ['4', '3.5', '3', '2.5', '2', '1.5', '1', '0'];
  if (orig === 'มผ') return ['ผ'];
  return ['1', '0'];   // 0 และ มส แก้ได้ไม่เกิน 1
}
function letterForGrade51_(g) {
  var rows = getGradeMapping().filter(function (m) { return String(m.grade51) === String(g); });
  return rows.length ? (rows[0].letter51 || '') : '';
}

// p: { semester?, classLevel?, teacher?, }  (teacher = จำกัดเฉพาะวิชาที่ครูคนนั้นสอน)
function getRemediation(p) {
  p = p || {};
  var sem = String(p.semester || getSettings().currentSemester || '1');
  var stBy = {}; getStudents().forEach(function (st) { stBy[st.ID] = st; });
  var suBy = {}; getSubjects().forEach(function (su) { suBy[su.ID] = su; });
  var allowSub = null;
  if (p.teacher) {
    var core = teacherCore(p.teacher), keys = {};
    readAll('Schedule').forEach(function (r) {
      if (teacherCore(r.teacher) === core) keys[String(r.subject).trim() + '|' + gradeOf_(r.classLevel)] = true;
    });
    allowSub = function (su) { return !!keys[String(su.name).trim() + '|' + String(su.classLevel || '').trim()]; };
  }
  var remBy = {};
  readAll('Remediation').forEach(function (r) { if (String(r.semester) === sem) remBy[r.studentID + '|' + r.subjectID] = r; });
  var out = [];
  var pass = function (st, su) {
    if (!st || !su) return false;
    if (allowSub && !allowSub(su)) return false;
    if (p.classLevel && String(st.classLevel) !== String(p.classLevel)) return false;
    return true;
  };
  // รอแก้ไข: เกรดพิเศษที่ยังค้างใน Grades
  readAll('Grades').forEach(function (g) {
    var gs = String(g.semester == null ? '' : g.semester);
    if (!(gs === sem || gs === '')) return;
    var mark = String(g.grade51);
    if (SPECIAL_MARKS.indexOf(mark) < 0) return;
    var st = stBy[g.studentID], su = suBy[g.subjectID];
    if (!pass(st, su)) return;
    var r = remBy[g.studentID + '|' + g.subjectID];
    out.push({ studentID: g.studentID, subjectID: g.subjectID, number: st.number, fullName: st.fullName,
      classLevel: st.classLevel, subjectCode: su.code || '', subject: su.name,
      origMark: mark, status: 'รอแก้ไข', newGrade: '', note: (r && r.status !== 'แก้ไขแล้ว') ? (r.note || '') : '' });
  });
  // แก้ไขแล้ว: จากทะเบียน Remediation (เกรดใน Grades ถูกแทนที่แล้ว)
  readAll('Remediation').forEach(function (r) {
    if (String(r.semester) !== sem || r.status !== 'แก้ไขแล้ว') return;
    var st = stBy[r.studentID], su = suBy[r.subjectID];
    if (!pass(st, su)) return;
    out.push({ studentID: r.studentID, subjectID: r.subjectID, number: st.number, fullName: st.fullName,
      classLevel: st.classLevel, subjectCode: su.code || '', subject: su.name,
      origMark: r.origMark, status: 'แก้ไขแล้ว', newGrade: r.newGrade, note: r.note || '',
      resolvedBy: r.resolvedBy || '', resolvedDate: r.resolvedDate ? ymd(r.resolvedDate) : '' });
  });
  out.sort(function (a, b) {
    return String(a.classLevel).localeCompare(String(b.classLevel)) || Number(a.number) - Number(b.number) || String(a.subject).localeCompare(String(b.subject));
  });
  var byMark = { '0': 0, 'ร': 0, 'มส': 0, 'มผ': 0 };
  out.forEach(function (x) { if (x.status === 'รอแก้ไข' && byMark[x.origMark] != null) byMark[x.origMark]++; });
  return { semester: sem, items: out,
    counts: { pending: out.filter(function (x) { return x.status === 'รอแก้ไข'; }).length,
      resolved: out.filter(function (x) { return x.status === 'แก้ไขแล้ว'; }).length, byMark: byMark } };
}

// p: { studentID, subjectID, semester, origMark, newGrade, note, by }
function resolveRemediation(p) {
  var allowed = remedAllowed_(p.origMark);
  if (allowed.indexOf(String(p.newGrade)) < 0) throw new Error('เกรดใหม่ต้องเป็น ' + allowed.join(' / ') + ' ตามเกณฑ์การแก้ ' + p.origMark);
  var sem = String(p.semester || getSettings().currentSemester || '1');
  var now = Utilities.formatDate(new Date(), 'GMT+7', 'yyyy-MM-dd HH:mm:ss');
  // 1) เขียนเกรดใหม่ทับใน Grades (mode R = จากการแก้ตัว)
  writeGrade({ studentID: p.studentID, subjectID: p.subjectID, mode: 'R', percent: '',
    grade51: String(p.newGrade), letter51: letterForGrade51_(p.newGrade),
    note: 'แก้ ' + p.origMark + (p.note ? (' • ' + p.note) : ''), semester: sem });
  // 2) upsert ทะเบียนสอบแก้ตัว
  var s = sheet('Remediation');
  var data = s.getDataRange().getValues();
  var found = -1;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][1]) === String(p.studentID) && String(data[i][2]) === String(p.subjectID) &&
        String(data[i][3]) === sem) { found = i + 1; break; }
  }
  var vals = [found > 0 ? data[found - 1][0] : genId('RMD'), p.studentID, p.subjectID, sem,
    String(p.origMark), String(p.newGrade), 'แก้ไขแล้ว', p.note || '', p.by || '', ymd(now.slice(0, 10)), now];
  if (found > 0) s.getRange(found, 1, 1, SHEETS.Remediation.length).setValues([vals]);
  else s.appendRow(vals);
  SpreadsheetApp.flush();
  return { ok: true };
}

/* ============================================================
   ประวัตินักเรียน (StudentProfiles) — ครูประจำชั้นกรอก / แอดมินนำเข้า DMC
   ============================================================ */
var PROFILE_FIELDS = ['nickname', 'citizenId', 'birthDate', 'religion', 'nationality', 'ethnicity',
  'bloodGroup', 'weight', 'height', 'address', 'fatherName', 'motherName',
  'guardianName', 'guardianRelation', 'guardianPhone',
  'guardianJob', 'fatherJob', 'motherJob', 'disadvantaged'];
var PHOTO_MAX_BYTES = 2 * 1024 * 1024;   // ขนาดรูปหลังบีบอัด ไม่เกิน 2MB

function isHomeroomOf_(teacherName, room) {
  var core = teacherCore(teacherName || '');
  var t = readAll('Teachers').filter(function (x) { return teacherCore(x.name) === core; })[0];
  return !!t && String(t.homeroomClass || '').trim() === String(room || '').trim();
}
function assertProfilePerm_(p, room) {
  if (String(p._role || '') === 'admin') return;
  if (!isHomeroomOf_(p._by || '', room))
    throw new Error('บันทึกประวัตินักเรียนได้เฉพาะครูประจำชั้นของห้องนั้น (หรือแอดมิน)');
}
function profileMap_() {
  var m = {};
  readAll('StudentProfiles').forEach(function (r) { m[String(r.studentID)] = r; });
  return m;
}
// รายชื่อ + ประวัติของห้อง (director อ่านได้)
function getClassProfiles(p) {
  var room = String((p || {}).classLevel || '').trim();
  if (!room) return [];
  var pm = profileMap_();
  return getStudents()
    .filter(function (st) { return String(st.classLevel).trim() === room; })
    .sort(function (a, b) { return (Number(a.number) || 0) - (Number(b.number) || 0); })
    .map(function (st) {
      var pr = pm[st.ID] || {};
      var out = { studentID: st.ID, number: st.number, fullName: st.fullName, classLevel: st.classLevel,
        studentCode: st.studentCode || '', photoUrl: pr.photoUrl || '' };
      PROFILE_FIELDS.forEach(function (k) { out[k] = (pr[k] == null ? '' : pr[k]); });
      out.birthDate = out.birthDate ? ymd(out.birthDate) : '';
      out.filled = PROFILE_FIELDS.filter(function (k) { return String(out[k] || '') !== ''; }).length;
      return out;
    });
}
function saveStudentProfile(p) {
  var st = getStudentsAll_().filter(function (x) { return x.ID === p.studentID; })[0];
  if (!st) throw new Error('ไม่พบนักเรียน');
  assertProfilePerm_(p, st.classLevel);
  upsertProfile_(p.studentID, p, p._by || '');
  return { ok: true };
}
function upsertProfile_(studentID, data, by) {
  var s = sheet('StudentProfiles');
  var rows = s.getDataRange().getValues();
  var head = rows[0]; var col = {}; head.forEach(function (h, i) { col[h] = i; });
  var row = -1;
  for (var i = 1; i < rows.length; i++) { if (String(rows[i][col.studentID]) === String(studentID)) { row = i + 1; break; } }
  var now = Utilities.formatDate(new Date(), 'GMT+7', 'yyyy-MM-dd HH:mm:ss');
  if (row < 0) {
    var newRow = head.map(function () { return ''; });
    newRow[col.studentID] = studentID;
    s.appendRow(newRow);
    row = s.getLastRow();
    rows.push(newRow);
  }
  var cur = s.getRange(row, 1, 1, head.length).getValues()[0];
  PROFILE_FIELDS.concat(['photoId', 'photoUrl']).forEach(function (k) {
    if (data[k] !== undefined && col[k] != null) cur[col[k]] = data[k];
  });
  cur[col.updatedAt] = now;
  cur[col.updatedBy] = by || '';
  s.getRange(row, 1, 1, head.length).setValues([cur]);
  SpreadsheetApp.flush();
}
function uploadStudentPhoto(p) {
  var st = getStudentsAll_().filter(function (x) { return x.ID === p.studentID; })[0];
  if (!st) throw new Error('ไม่พบนักเรียน');
  assertProfilePerm_(p, st.classLevel);
  if (!p.base64) throw new Error('ไม่มีไฟล์รูป');
  var m = String(p.base64).match(/^data:(.+);base64,(.*)$/);
  var ct = m ? m[1] : 'image/jpeg';
  var b64 = m ? m[2] : p.base64;
  if (b64.length * 0.75 > PHOTO_MAX_BYTES) throw new Error('รูปใหญ่เกิน 2MB หลังบีบอัด — เลือกรูปเล็กลง');
  var folder = getOrCreateFolder('รูปนักเรียน - ' + (getSettings().schoolName || ''));
  // ลบรูปเดิม
  var pm = profileMap_()[String(p.studentID)] || {};
  if (pm.photoId) { try { DriveApp.getFileById(pm.photoId).setTrashed(true); } catch (e) {} }
  var blob = Utilities.newBlob(Utilities.base64Decode(b64), ct, 'std_' + (st.studentCode || st.ID) + '.jpg');
  var file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  var url = 'https://drive.google.com/thumbnail?id=' + file.getId() + '&sz=w400';
  upsertProfile_(p.studentID, { photoId: file.getId(), photoUrl: url }, p._by || '');
  return { photoUrl: url };
}
// นำเข้า DMC (แอดมิน) — p.rows: [{studentCode?, citizenId?, fullName?, ...PROFILE_FIELDS}]
function importDMC(p) {
  assertAdmin_(p);
  var rows = p.rows || [];
  if (!rows.length) return { matched: 0, unmatched: [] };
  var students = getStudentsAll_();
  var byCode = {}, byCid = {}, byName = {};
  var pm = profileMap_();
  students.forEach(function (st) {
    if (st.studentCode) byCode[String(st.studentCode).trim()] = st;
    var pr = pm[st.ID];
    if (pr && pr.citizenId) byCid[String(pr.citizenId).replace(/\D/g, '')] = st;
    byName[normName(st.fullName)] = st;
  });
  var matched = 0, unmatched = [];
  rows.forEach(function (r) {
    var st = null;
    if (r.studentCode && byCode[String(r.studentCode).trim()]) st = byCode[String(r.studentCode).trim()];
    if (!st && r.citizenId && byCid[String(r.citizenId).replace(/\D/g, '')]) st = byCid[String(r.citizenId).replace(/\D/g, '')];
    if (!st && r.fullName && byName[normName(r.fullName)]) st = byName[normName(r.fullName)];
    if (!st) { unmatched.push(r.fullName || r.studentCode || r.citizenId || '?'); return; }
    var data = {};
    PROFILE_FIELDS.forEach(function (k) {
      var v = r[k];
      if (v !== undefined && v !== null && String(v).trim() !== '') data[k] = String(v).trim();
    });
    if (Object.keys(data).length) { upsertProfile_(st.ID, data, 'DMC (' + (p._by || 'admin') + ')'); matched++; }
  });
  return { matched: matched, unmatched: unmatched.slice(0, 30), unmatchedCount: unmatched.length };
}



/* ============================================================
   รายงานสารสนเทศ/SAR อัตโนมัติ — รวมสถิติทั้งระบบ
   p: { semester, scope:'semester'|'year' }
   ============================================================ */
function getSAR(p) {
  p = p || {};
  var set = getSettings();
  var sem = String(p.semester || set.currentSemester || '1');
  var scope = (p.scope === 'year') ? 'year' : 'semester';
  var semOk = function (v) {
    if (scope === 'year') return true;
    var s2 = String(v == null ? '' : v);
    return s2 === sem || s2 === '';
  };
  var students = getStudents();
  var roomCount = {}, roomStudents = {};
  students.forEach(function (st) {
    var r = String(st.classLevel || '').trim(); if (!r) return;
    roomCount[r] = (roomCount[r] || 0) + 1;
    (roomStudents[r] = roomStudents[r] || {})[st.ID] = 1;
  });
  var roomList = Object.keys(roomCount).sort();
  var stuRoom = {}; students.forEach(function (st) { stuRoom[st.ID] = String(st.classLevel || '').trim(); });
  var teachersN = readAll('Teachers').filter(function (t) { return String(t.role) !== 'admin' && t.name; }).length;

  // ---------- ผลสัมฤทธิ์รายวิชา + GPA รายห้อง ----------
  var suBy = {}; getSubjects().forEach(function (su) { suBy[su.ID] = su; });
  var DIST_KEYS = ['4', '3.5', '3', '2.5', '2', '1.5', '1', '0', 'ร', 'มส', 'มผ', 'ผ'];
  var bySub = {};
  var gpaAcc = {};   // room -> {sw, sg}
  readAll('Grades').forEach(function (g) {
    if (!semOk(g.semester)) return;
    var su = suBy[g.subjectID]; if (!su) return;
    var mark = String(g.grade51 == null ? '' : g.grade51);
    if (mark === '') return;
    var k = g.subjectID;
    if (!bySub[k]) {
      bySub[k] = { code: su.code || '', name: su.name, classLevel: su.classLevel || '', n: 0, sum: 0, numN: 0, good: 0, dist: {} };
      DIST_KEYS.forEach(function (d) { bySub[k].dist[d] = 0; });
    }
    var o = bySub[k];
    o.n++;
    if (o.dist[mark] != null) o.dist[mark]++;
    var num = parseFloat(mark);
    if (!isNaN(num)) {
      o.sum += num; o.numN++;
      if (num >= 3) o.good++;
      var room = stuRoom[g.studentID];
      if (room) {
        var h = Number(su.hours) || 1;
        var a = (gpaAcc[room] = gpaAcc[room] || { sw: 0, sg: 0 });
        a.sw += h; a.sg += num * h;
      }
    }
  });
  var subjects = Object.keys(bySub).map(function (k) {
    var o = bySub[k];
    o.mean = o.numN ? Math.round(o.sum / o.numN * 100) / 100 : null;
    o.pctGood = o.numN ? Math.round(o.good / o.numN * 100) : null;
    return o;
  }).sort(function (a, b) {
    return String(a.classLevel).localeCompare(String(b.classLevel)) || String(a.name).localeCompare(String(b.name));
  });
  var gpaRooms = roomList.map(function (r) {
    var a = gpaAcc[r];
    return { room: r, students: roomCount[r], gpa: a && a.sw ? Math.round(a.sg / a.sw * 100) / 100 : null };
  });

  // ---------- เวลาเรียน (หน้าเสาธง) ----------
  var attendance = { needDates: false, rooms: [] };
  var ranges = [];
  if (scope === 'semester') { var rr = semesterRange_(set, sem); if (rr) ranges.push(rr); }
  else { ['1', '2'].forEach(function (s2) { var r2 = semesterRange_(set, s2); if (r2) ranges.push(r2); }); }
  if (!ranges.length) { attendance.needDates = true; }
  else {
    var cal = calMap_();
    var today = Utilities.formatDate(new Date(), 'GMT+7', 'yyyy-MM-dd');
    var allDates = [];
    ranges.forEach(function (r2) { allDates = allDates.concat(datesBetween_(r2.from, r2.to)); });
    var schoolDates = {};
    allDates.forEach(function (iso) { if (isSchoolDay_(iso, cal) && iso <= today) schoolDates[iso] = 1; });
    var acc = {};   // room -> {records, present, days:{}}
    readAll('Attendance').forEach(function (a) {
      if (String(a.period) !== '0') return;
      var d = ymd(a.date);
      if (!schoolDates[d]) return;
      var r3 = String(a.classLevel).trim();
      var o2 = (acc[r3] = acc[r3] || { records: 0, present: 0, days: {} });
      o2.records++; o2.days[d] = 1;
      if (a.status === 'มา' || a.status === 'สาย') o2.present++;
    });
    attendance.schoolDays = Object.keys(schoolDates).length;
    attendance.rooms = roomList.map(function (r3) {
      var o2 = acc[r3] || { records: 0, present: 0, days: {} };
      return { room: r3, checkedDays: Object.keys(o2.days).length,
        pctPresent: o2.records ? Math.round(o2.present / o2.records * 100) : null };
    });
  }

  // ---------- 0/ร/มส/มผ ----------
  var remPending = { '0': 0, 'ร': 0, 'มส': 0, 'มผ': 0 };
  readAll('Grades').forEach(function (g) {
    if (!semOk(g.semester)) return;
    var m2 = String(g.grade51);
    if (remPending[m2] != null) remPending[m2]++;
  });
  var remResolved = readAll('Remediation').filter(function (r4) {
    return (scope === 'year' || String(r4.semester) === sem) && r4.status === 'แก้ไขแล้ว';
  }).length;

  // ---------- SDQ ----------
  var sdq = { assessed: 0, bands: { 'ปกติ': 0, 'เสี่ยง': 0, 'มีปัญหา': 0 }, rooms: {} };
  readAll('SDQ').forEach(function (r5) {
    if (scope === 'semester' && String(r5.semester) !== sem) return;
    var room = stuRoom[r5.studentID]; if (!room) return;
    sdq.assessed++;
    if (sdq.bands[r5.result] != null) sdq.bands[r5.result]++;
    var o3 = (sdq.rooms[room] = sdq.rooms[room] || { assessed: 0, risk: 0 });
    o3.assessed++;
    if (r5.result === 'เสี่ยง' || r5.result === 'มีปัญหา') o3.risk++;
  });

  // ---------- เยี่ยมบ้าน ----------
  var visited = {};
  readAll('HomeVisits').forEach(function (v) { visited[String(v.studentID)] = 1; });
  var hvRooms = roomList.map(function (r6) {
    var ids = Object.keys(roomStudents[r6] || {});
    var done = ids.filter(function (id) { return visited[id]; }).length;
    return { room: r6, done: done, total: ids.length };
  });

  // ---------- เอกสารครู + นิเทศ ----------
  var teaching = null;
  try { var tv = getTeachingOverview({ semester: sem }); teaching = tv.totals; teaching.types = (tv.types || []).map(function (t) { return t.name; }); } catch (e) {}
  var spv = { total: 0, bands: { 'ดีมาก': 0, 'ดี': 0, 'พอใช้': 0, 'ปรับปรุง': 0 } };
  readAll('Supervisions').forEach(function (r7) {
    if (scope === 'semester' && !(String(r7.semester) === sem || String(r7.semester || '') === '')) return;
    spv.total++;
    if (spv.bands[r7.rating] != null) spv.bands[r7.rating]++;
  });

  return {
    semester: sem, scope: scope,
    school: { name: set.schoolName || '', academicYear: set.academicYear || '', director: set.directorName || '', logoURL: set.logoURL || '' },
    basic: { students: students.length, teachers: teachersN,
      rooms: roomList.map(function (r8) { return { room: r8, students: roomCount[r8] }; }) },
    subjects: subjects, gpaRooms: gpaRooms, attendance: attendance,
    remediation: { pending: remPending, resolved: remResolved },
    sdq: sdq, homeVisits: hvRooms, teaching: teaching, supervision: spv
  };
}

/* ============================================================
   SDQ (ฉบับครูประเมินนักเรียน 25 ข้อ) + บันทึกเยี่ยมบ้าน
   ============================================================ */
var SDQ_REV = { 7: 1, 11: 1, 14: 1, 21: 1, 25: 1 };
var SDQ_SCALES = {
  emotional: [3, 8, 13, 16, 24], conduct: [5, 7, 12, 18, 22],
  hyper: [2, 10, 15, 21, 25], peer: [6, 11, 14, 19, 23], prosocial: [1, 4, 9, 17, 20]
};
function sdqScore_(answers) {
  var val = function (i) {
    var v = Number(answers[i - 1]);
    if (isNaN(v) || v < 0 || v > 2) v = 0;
    return SDQ_REV[i] ? (2 - v) : v;
  };
  var sc = {};
  Object.keys(SDQ_SCALES).forEach(function (k) {
    sc[k] = SDQ_SCALES[k].reduce(function (a, i) { return a + val(i); }, 0);
  });
  sc.total = sc.emotional + sc.conduct + sc.hyper + sc.peer;
  var band2 = function (v) { return v <= 3 ? 'ปกติ' : (v === 4 ? 'เสี่ยง' : 'มีปัญหา'); };       // อารมณ์/ความประพฤติ
  var band5 = function (v) { return v <= 5 ? 'ปกติ' : (v === 6 ? 'เสี่ยง' : 'มีปัญหา'); };       // สมาธิ/เพื่อน
  sc.bands = {
    emotional: band2(sc.emotional), conduct: band2(sc.conduct),
    hyper: band5(sc.hyper), peer: band5(sc.peer),
    prosocial: sc.prosocial >= 6 ? 'จุดแข็ง' : 'ควรส่งเสริม',
    total: sc.total <= 15 ? 'ปกติ' : (sc.total <= 17 ? 'เสี่ยง' : 'มีปัญหา')
  };
  return sc;
}
// รายชื่อทั้งห้อง + ผล SDQ ล่าสุดของภาคเรียน
function getSDQClass(p) {
  var room = String((p || {}).classLevel || '').trim();
  var sem = String(p.semester || getSettings().currentSemester || '1');
  if (!room) return [];
  var by = {};
  readAll('SDQ').forEach(function (r) { if (String(r.semester) === sem) by[String(r.studentID)] = r; });
  return getStudents()
    .filter(function (st) { return String(st.classLevel).trim() === room; })
    .sort(function (a, b) { return (Number(a.number) || 0) - (Number(b.number) || 0); })
    .map(function (st) {
      var r = by[st.ID];
      var out = { studentID: st.ID, number: st.number, fullName: st.fullName, assessed: !!r };
      if (r) {
        out.answers = String(r.answers || '');
        ['emotional', 'conduct', 'hyper', 'peer', 'prosocial', 'total'].forEach(function (k) { out[k] = Number(r[k]) || 0; });
        out.result = r.result || '';
        var sc = sdqScore_(String(r.answers || '').split(','));
        out.bands = sc.bands;
      }
      return out;
    });
}
// p: { studentID, answers:[25], semester }
function saveSDQ(p) {
  var st = getStudentsAll_().filter(function (x) { return x.ID === p.studentID; })[0];
  if (!st) throw new Error('ไม่พบนักเรียน');
  assertProfilePerm_(p, st.classLevel);
  var ans = (p.answers || []).map(function (v) { var n = Number(v); return (isNaN(n) || n < 0 || n > 2) ? 0 : n; });
  if (ans.length !== 25) throw new Error('คำตอบต้องครบ 25 ข้อ');
  var sem = String(p.semester || getSettings().currentSemester || '1');
  var sc = sdqScore_(ans);
  var now = Utilities.formatDate(new Date(), 'GMT+7', 'yyyy-MM-dd HH:mm:ss');
  var s = sheet('SDQ');
  var data = s.getDataRange().getValues();
  var row = -1;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][1]) === String(p.studentID) && String(data[i][2]) === sem) { row = i + 1; break; }
  }
  var vals = [row > 0 ? data[row - 1][0] : genId('SDQ'), p.studentID, sem, ans.join(','),
    sc.emotional, sc.conduct, sc.hyper, sc.peer, sc.prosocial, sc.total, sc.bands.total, p._by || '', now];
  if (row > 0) s.getRange(row, 1, 1, vals.length).setValues([vals]);
  else s.appendRow(vals);
  SpreadsheetApp.flush();
  return { ok: true, scores: sc };
}
// ---------- เยี่ยมบ้าน ----------
function getHomeVisits(p) {
  p = p || {};
  var rows = readAll('HomeVisits').map(function (r) { r.date = ymd(r.date); return r; });
  if (p.studentID) rows = rows.filter(function (r) { return String(r.studentID) === String(p.studentID); });
  return rows.sort(function (a, b) { return String(b.date).localeCompare(String(a.date)); });
}
// สรุปสถานะเยี่ยมบ้านของทั้งห้อง
function getHomeVisitClass(p) {
  var room = String((p || {}).classLevel || '').trim();
  if (!room) return [];
  var latest = {}, count = {};
  readAll('HomeVisits').forEach(function (r) {
    var k = String(r.studentID);
    count[k] = (count[k] || 0) + 1;
    var d = ymd(r.date);
    if (!latest[k] || d > latest[k]) latest[k] = d;
  });
  return getStudents()
    .filter(function (st) { return String(st.classLevel).trim() === room; })
    .sort(function (a, b) { return (Number(a.number) || 0) - (Number(b.number) || 0); })
    .map(function (st) {
      return { studentID: st.ID, number: st.number, fullName: st.fullName,
        visits: count[st.ID] || 0, lastDate: latest[st.ID] || '' };
    });
}
// p: { id?, studentID, date, visitors, found, living, needs, photoBase64?, semester }
function saveHomeVisit(p) {
  var st = getStudentsAll_().filter(function (x) { return x.ID === p.studentID; })[0];
  if (!st) throw new Error('ไม่พบนักเรียน');
  assertProfilePerm_(p, st.classLevel);
  var sem = String(p.semester || getSettings().currentSemester || '1');
  var now = Utilities.formatDate(new Date(), 'GMT+7', 'yyyy-MM-dd HH:mm:ss');
  var photoId = '', photoUrl = '';
  if (p.photoBase64) {
    var m = String(p.photoBase64).match(/^data:(.+);base64,(.*)$/);
    var b = Utilities.base64Decode(m ? m[2] : p.photoBase64);
    if (b.length > PHOTO_MAX_BYTES) throw new Error('รูปใหญ่เกิน 2MB');
    var folder = getOrCreateFolder('เยี่ยมบ้าน - ' + (getSettings().schoolName || ''));
    var file = folder.createFile(Utilities.newBlob(b, m ? m[1] : 'image/jpeg',
      'visit_' + (st.fullName || '') + '_' + new Date().getTime() + '.jpg'));
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    photoId = file.getId();
    photoUrl = 'https://drive.google.com/thumbnail?id=' + photoId + '&sz=w1000';
  }
  var s = sheet('HomeVisits');
  if (p.id) {
    var row = findRowById('HomeVisits', p.id);
    if (row < 0) throw new Error('ไม่พบบันทึกนี้');
    var old = readAll('HomeVisits').filter(function (x) { return x.ID === p.id; })[0] || {};
    s.getRange(row, 1, 1, SHEETS.HomeVisits.length).setValues([[p.id, p.studentID, ymd(p.date || now.slice(0, 10)),
      p.visitors || '', p.found || '', p.living || '', p.needs || '',
      photoId || old.photoId || '', photoUrl || old.photoUrl || '', sem, p._by || '', now]]);
  } else {
    s.appendRow([genId('HVS'), p.studentID, ymd(p.date || now.slice(0, 10)),
      p.visitors || '', p.found || '', p.living || '', p.needs || '', photoId, photoUrl, sem, p._by || '', now]);
  }
  SpreadsheetApp.flush();
  return { ok: true };
}
function deleteHomeVisit(p) {
  assertAdmin_(p);
  var rec = readAll('HomeVisits').filter(function (x) { return x.ID === p.id; })[0];
  if (rec && rec.photoId) { try { DriveApp.getFileById(rec.photoId).setTrashed(true); } catch (e) {} }
  var row = findRowById('HomeVisits', p.id);
  if (row > 0) sheet('HomeVisits').deleteRow(row);
  SpreadsheetApp.flush();
  return { ok: true };
}

/* ============================================================
   ตารางสอบ + กรรมการคุมสอบ (กันเวลาชน)
   ============================================================ */
function timesOverlap_(s1, e1, s2, e2) { return String(s1) < String(e2) && String(s2) < String(e1); }

function getExams(p) {
  p = p || {};
  var sem = String(p.semester || getSettings().currentSemester || '1');
  return readAll('ExamSchedule')
    .filter(function (r) {
      var rs = String(r.semester == null ? '' : r.semester);
      if (!(rs === sem || rs === '')) return false;
      if (p.examType && String(r.examType) !== String(p.examType)) return false;
      return true;
    })
    .map(function (r) { r.date = ymd(r.date); return r; })
    .sort(function (a, b) {
      return String(a.date).localeCompare(String(b.date)) ||
        String(a.startTime).localeCompare(String(b.startTime)) ||
        String(a.classLevel).localeCompare(String(b.classLevel));
    });
}

// p: { id?, examType, date, startTime, endTime, classLevel, subject, room, proctor1, proctor2 }
function saveExam(p) {
  assertAdmin_(p);
  var date = ymd(p.date), st = String(p.startTime || '').trim(), en = String(p.endTime || '').trim();
  var cls = String(p.classLevel || '').trim(), subject = String(p.subject || '').trim();
  if (!date || !st || !en || !cls || !subject) throw new Error('กรอก วันที่/เวลา/ห้อง/วิชา ให้ครบ');
  if (en <= st) throw new Error('เวลาสิ้นสุดต้องมากกว่าเวลาเริ่ม');
  var sem = String(p.semester || getSettings().currentSemester || '1');
  var rows = readAll('ExamSchedule').filter(function (r) { return r.ID !== p.id && ymd(r.date) === date; });
  // ชน 1: ห้อง(ชั้น)เดียวกัน เวลาทับกัน
  var c1 = rows.filter(function (r) {
    return String(r.classLevel).trim() === cls && timesOverlap_(st, en, r.startTime, r.endTime);
  })[0];
  if (c1) throw new Error('เวลาสอบชน: ' + cls + ' มีสอบ "' + c1.subject + '" เวลา ' + c1.startTime + '-' + c1.endTime + ' อยู่แล้ว');
  // ชน 2: กรรมการคุมสอบซ้อนเวลา
  [String(p.proctor1 || '').trim(), String(p.proctor2 || '').trim()].filter(Boolean).forEach(function (t) {
    var tc = teacherCore(t);
    var c2 = rows.filter(function (r) {
      return (teacherCore(r.proctor1) === tc || teacherCore(r.proctor2) === tc) &&
        timesOverlap_(st, en, r.startTime, r.endTime);
    })[0];
    if (c2) throw new Error('คุมสอบชน: ' + t + ' คุม ' + c2.classLevel + ' (' + c2.subject + ') เวลา ' + c2.startTime + '-' + c2.endTime + ' อยู่แล้ว');
  });
  if (String(p.proctor1 || '').trim() && teacherCore(p.proctor1) === teacherCore(p.proctor2 || '')) {
    throw new Error('กรรมการทั้งสองคนเป็นคนเดียวกัน');
  }
  var s = sheet('ExamSchedule');
  var now = Utilities.formatDate(new Date(), 'GMT+7', 'yyyy-MM-dd HH:mm:ss');
  var vals = [p.id || genId('EXM'), String(p.examType || 'กลางภาค'), date, st, en, cls, subject,
    String(p.room || '').trim(), String(p.proctor1 || '').trim(), String(p.proctor2 || '').trim(), sem, now];
  if (p.id) {
    var row = findRowById('ExamSchedule', p.id);
    if (row < 0) throw new Error('ไม่พบรายการสอบนี้');
    s.getRange(row, 1, 1, vals.length).setValues([vals]);
  } else {
    s.appendRow(vals);
  }
  SpreadsheetApp.flush();
  return getExams({ semester: sem });
}
function deleteExam(p) {
  assertAdmin_(p);
  var row = findRowById('ExamSchedule', p.id);
  var sem = '';
  if (row > 0) sheet('ExamSchedule').deleteRow(row);
  SpreadsheetApp.flush();
  return getExams({ semester: p.semester });
}

/* ============================================================
   แก้ไขตารางสอน (แอดมิน) — เพิ่ม/ย้าย/ลบคาบ + ตรวจคาบชน
   ============================================================ */
// p: { id?, day, period, classLevel, subject, teacher }
function saveScheduleSlot(p) {
  if (String(p._role || '') !== 'admin') throw new Error('แก้ไขตารางสอนได้เฉพาะแอดมิน');
  var day = String(p.day || '').trim();
  var period = Number(p.period);
  var room = String(p.classLevel || '').trim();
  var subject = String(p.subject || '').trim();
  var teacher = String(p.teacher || '').trim();
  if (!day || !period || !room || !subject) throw new Error('กรอก วัน/คาบ/ห้อง/วิชา ให้ครบ');

  var rows = readAll('Schedule');
  var slotId = p.id || '';
  // ชน 1: ห้องเดียวกัน วัน+คาบเดียวกัน
  var clashRoom = rows.filter(function (r) {
    return r.ID !== slotId && String(r.day).trim() === day && Number(r.period) === period &&
      String(r.classLevel).trim() === room;
  })[0];
  if (clashRoom) {
    if (String(clashRoom.subject).trim() === subject) {
      // วิชาเดียวกันในช่องเดิม → ถือเป็นการแก้ไขแถวเดิม (เช่น เติม/เปลี่ยนชื่อครู) ไม่ใช่คาบชน
      slotId = clashRoom.ID;
    } else {
      throw new Error('คาบชน: ห้อง ' + room + ' ' + day + ' คาบ ' + period + ' มีวิชา "' + clashRoom.subject + '" (' +
        (String(clashRoom.teacher || '').trim() || 'ยังไม่ระบุครู') + ') อยู่แล้ว — สลับมุมมอง "รายห้อง" เลือก ' + room + ' เพื่อแก้ไข/ลบคาบเดิม');
    }
  }
  // ชน 2: ครูคนเดียวกัน วัน+คาบเดียวกัน สอนห้องอื่นอยู่
  if (teacher) {
    var tc = teacherCore(teacher);
    var clashT = rows.filter(function (r) {
      return r.ID !== slotId && String(r.day).trim() === day && Number(r.period) === period &&
        teacherCore(r.teacher) === tc;
    })[0];
    if (clashT) throw new Error('คาบชน: ' + teacher + ' มีสอน ' + clashT.classLevel + ' (' + clashT.subject + ') ใน ' + day + ' คาบ ' + period + ' อยู่แล้ว');
  }

  var s = sheet('Schedule');
  if (slotId) {
    var row = findRowById('Schedule', slotId);
    if (row < 0) throw new Error('ไม่พบคาบสอนนี้ (อาจถูกลบไปแล้ว)');
    s.getRange(row, 2, 1, 5).setValues([[day, period, room, subject, teacher]]);
  } else {
    s.appendRow([genId('SCH'), day, period, room, subject, teacher]);
  }
  SpreadsheetApp.flush();
  // sync รายวิชา (ถ้าวิชาใหม่) + การจัดครู (เพิ่มอย่างเดียว)
  var addedSubject = 0;
  try { addedSubject = createSubjectsFromScheduleRows([{ subject: subject, classLevel: room }]) || 0; } catch (e) {}
  try { syncTeacherAssign(); } catch (e) {}
  return { rows: readAll('Schedule'), addedSubject: addedSubject };
}
function deleteScheduleSlot(p) {
  if (String(p._role || '') !== 'admin') throw new Error('แก้ไขตารางสอนได้เฉพาะแอดมิน');
  var row = findRowById('Schedule', p.id);
  if (row > 0) sheet('Schedule').deleteRow(row);
  SpreadsheetApp.flush();
  return { rows: readAll('Schedule') };
}

/* ============================================================
   จัดครูผู้สอน (วิชา × ห้อง) + สิทธิ์การบันทึก + ตรวจสุขภาพรายวิชา
   ============================================================ */
function gradeRoomsB_(grade) {
  var set = {};
  getStudents().forEach(function (st) {
    if (st.classLevel && gradeOf_(st.classLevel) === String(grade)) set[String(st.classLevel).trim()] = 1;
  });
  return Object.keys(set).sort();
}
// เติมจากตารางสอน (เพิ่มอย่างเดียว ไม่ทับที่ตั้งไว้แล้ว)
function syncTeacherAssign() {
  var byNC = {};
  getSubjects().forEach(function (su) { byNC[String(su.name).trim() + '|' + String(su.classLevel || '').trim()] = su.ID; });
  var existing = {};
  readAll('TeacherAssign').forEach(function (a) { existing[a.subjectID + '|' + String(a.room).trim()] = 1; });
  var s = sheet('TeacherAssign');
  var now = Utilities.formatDate(new Date(), 'GMT+7', 'yyyy-MM-dd HH:mm:ss');
  var added = 0;
  readAll('Schedule').forEach(function (r) {
    var name = String(r.subject || '').trim(), room = String(r.classLevel || '').trim(), t = String(r.teacher || '').trim();
    if (!name || !room || !t) return;
    var sid = byNC[name + '|' + gradeOf_(room)];
    if (!sid) return;
    var k = sid + '|' + room;
    if (existing[k]) return;
    existing[k] = 1;
    s.appendRow([genId('TAS'), sid, room, t, now]);
    added++;
  });
  if (added) SpreadsheetApp.flush();
  return { added: added };
}
function getTeacherAssign() {
  syncTeacherAssign();
  var asg = {};
  readAll('TeacherAssign').forEach(function (a) { asg[a.subjectID + '|' + String(a.room).trim()] = a.teacher || ''; });
  var subjects = getSubjects().sort(function (a, b) {
    return String(a.classLevel).localeCompare(String(b.classLevel)) || String(a.name).localeCompare(String(b.name));
  }).map(function (su) {
    var rooms = gradeRoomsB_(su.classLevel).map(function (rm) {
      return { room: rm, teacher: asg[su.ID + '|' + rm] || '' };
    });
    return { subjectID: su.ID, code: su.code || '', name: su.name, classLevel: su.classLevel || '', rooms: rooms };
  });
  var teachers = readAll('Teachers').filter(function (t) { return String(t.role) !== 'admin' && t.name; })
    .map(function (t) { return t.name; });
  return { subjects: subjects, teachers: teachers };
}
// p.items: [{subjectID, room, teacher}]
function saveTeacherAssign(p) {
  var items = p.items || [];
  var s = sheet('TeacherAssign');
  var data = s.getDataRange().getValues();
  var head = data[0]; var col = {}; head.forEach(function (h, i) { col[h] = i; });
  var rowOf = {};
  for (var i = 1; i < data.length; i++) rowOf[data[i][col.subjectID] + '|' + String(data[i][col.room]).trim()] = i + 1;
  var now = Utilities.formatDate(new Date(), 'GMT+7', 'yyyy-MM-dd HH:mm:ss');
  items.forEach(function (it) {
    var k = it.subjectID + '|' + String(it.room).trim();
    if (rowOf[k]) {
      s.getRange(rowOf[k], col.teacher + 1).setValue(it.teacher || '');
      s.getRange(rowOf[k], col.updatedAt + 1).setValue(now);
    } else {
      s.appendRow([genId('TAS'), it.subjectID, it.room, it.teacher || '', now]);
      rowOf[k] = s.getLastRow();
    }
  });
  SpreadsheetApp.flush();
  return { saved: items.length };
}
// วิชา×ห้อง ที่ครูคนหนึ่งสอน (จากเมนูจัดครู ∪ ตารางสอน)
function getMyTeach(p) {
  var teacher = p.teacher || '';
  if (!teacher) return [];
  var core = teacherCore(teacher);
  var suBy = {}; getSubjects().forEach(function (su) { suBy[su.ID] = su; });
  var byNC = {}; getSubjects().forEach(function (su) { byNC[String(su.name).trim() + '|' + String(su.classLevel || '').trim()] = su; });
  var out = {}, assignedKey = {};
  readAll('TeacherAssign').forEach(function (a) {
    var k = a.subjectID + '|' + String(a.room).trim();
    assignedKey[k] = teacherCore(a.teacher || '');   // ห้องที่ถูกกำหนดครูแล้ว (ใครก็ตาม)
    if (a.teacher && teacherCore(a.teacher) === core && suBy[a.subjectID]) out[k] = 1;
  });
  readAll('Schedule').forEach(function (r) {
    if (teacherCore(r.teacher) !== core) return;
    var su = byNC[String(r.subject || '').trim() + '|' + gradeOf_(r.classLevel)];
    if (!su) return;
    var k = su.ID + '|' + String(r.classLevel).trim();
    // ตารางสอนนับเฉพาะห้องที่ยังไม่ถูกกำหนดครูเป็นคนอื่นในเมนูจัดครู
    if (assignedKey[k] && assignedKey[k] !== core) return;
    out[k] = 1;
  });
  return Object.keys(out).map(function (k) {
    var pp = k.split('|'); var su = suBy[pp[0]];
    return { subjectID: pp[0], code: su.code || '', name: su.name, classLevel: su.classLevel || '', room: pp[1] };
  }).sort(function (a, b) { return String(a.room).localeCompare(String(b.room)) || String(a.name).localeCompare(String(b.name)); });
}
function canTeach_(teacher, subjectID, room) {
  if (!teacher) return false;
  return getMyTeach({ teacher: teacher }).some(function (x) {
    return x.subjectID === subjectID && (!room || x.room === String(room).trim());
  });
}
// ---- ตัวบังคับสิทธิ์ (แอดมินผ่านเสมอ) ----
function assertScorePerm_(p, subjectID, room) {
  if (String(p._role || '') === 'admin') return;
  var by = p._by || '';
  if (!by) throw new Error('ไม่ทราบผู้ใช้ — กรุณาเข้าสู่ระบบใหม่');
  if (!canTeach_(by, subjectID, room || '')) throw new Error('บันทึก/แก้ไขได้เฉพาะครูประจำวิชา (หรือแอดมิน)');
}
function roomOfStudent_(studentID) {
  var st = getStudents().filter(function (x) { return x.ID === studentID; })[0];
  return st ? String(st.classLevel).trim() : '';
}
function assertGradesBatchPerm_(p) {
  if (String(p._role || '') === 'admin') return;
  var seen = {};
  (p.grades || []).forEach(function (g) {
    if (seen[g.subjectID]) return; seen[g.subjectID] = 1;
    assertScorePerm_(p, g.subjectID, '');
  });
}
// ตรวจสุขภาพรายวิชา
function subjectHealthCheck() {
  var subs = getSubjects();
  var ids = {}, seen = {}, dups = [];
  subs.forEach(function (su) {
    ids[su.ID] = 1;
    var k = String(su.name).trim() + '|' + String(su.classLevel || '').trim();
    if (seen[k]) dups.push(k.replace('|', ' '));
    seen[k] = 1;
  });
  var orphans = {};
  ['LearningUnits', 'Scores', 'Grades', 'SubjectDocs'].forEach(function (sn) {
    var sh = ss().getSheetByName(sn); if (!sh) return;
    readAll(sn).forEach(function (r) {
      if (r.subjectID && !ids[r.subjectID]) orphans[sn] = (orphans[sn] || 0) + 1;
    });
  });
  var missing = {};
  readAll('Schedule').forEach(function (r) {
    var name = String(r.subject || '').trim(); if (!name) return;
    var k = name + '|' + gradeOf_(r.classLevel);
    if (!seen[k]) missing[k.replace('|', ' ')] = 1;
  });
  return { subjectCount: subs.length, duplicates: dups, orphans: orphans, missingFromSchedule: Object.keys(missing) };
}

/* ============================================================
   บันทึกการนิเทศ/เยี่ยมชั้นเรียน
   ============================================================ */
function getSupervisions(p) {
  p = p || {};
  var sem = String(p.semester || getSettings().currentSemester || '1');
  return readAll('Supervisions')
    .filter(function (r) {
      var rs = String(r.semester == null ? '' : r.semester);
      if (!(rs === sem || rs === '')) return false;
      if (p.classLevel && String(r.classLevel) !== String(p.classLevel)) return false;
      if (p.teacher && teacherCore(r.teacher) !== teacherCore(p.teacher)) return false;
      return true;
    })
    .map(function (r) { r.date = ymd(r.date); return r; })
    .sort(function (a, b) { return String(b.date).localeCompare(String(a.date)); });
}
function saveSupervision(p) {
  var s = sheet('Supervisions');
  var now = Utilities.formatDate(new Date(), 'GMT+7', 'yyyy-MM-dd HH:mm:ss');
  var sem = String(p.semester || getSettings().currentSemester || '1');
  var vals = [p.id || genId('SPV'), ymd(p.date || now.slice(0, 10)), p.classLevel || '', p.teacher || '',
    p.subject || '', p.topic || '', p.strengths || '', p.suggestions || '', p.rating || '', p.by || '', now, sem];
  if (p.id) {
    var row = findRowById('Supervisions', p.id);
    if (row > 0) { s.getRange(row, 1, 1, vals.length).setValues([vals]); SpreadsheetApp.flush(); return { ok: true }; }
  }
  s.appendRow(vals);
  SpreadsheetApp.flush();
  return { ok: true };
}
function deleteSupervision(p) {
  var row = findRowById('Supervisions', p.id);
  if (row > 0) sheet('Supervisions').deleteRow(row);
  SpreadsheetApp.flush();
  return { ok: true };
}

/* ============================================================
   คลังแบบฟอร์มโรงเรียน (แอดมินอัปโหลด ทุกคนดาวน์โหลด)
   ============================================================ */
function getFormTemplates() {
  return readAll('FormTemplates').map(function (r) { return r; })
    .sort(function (a, b) { return String(a.name).localeCompare(String(b.name)); });
}
function uploadFormTemplate(p) {
  if (!p.base64) throw new Error('ไม่มีไฟล์');
  var m = String(p.base64).match(/^data:(.+);base64,(.*)$/);
  var ct = m ? m[1] : 'application/octet-stream';
  var b64 = m ? m[2] : p.base64;
  var folder = getOrCreateFolder('แบบฟอร์มโรงเรียน - ' + (getSettings().schoolName || ''));
  var blob = Utilities.newBlob(Utilities.base64Decode(b64), ct, p.fileName || ('form_' + new Date().getTime()));
  var file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  var now = Utilities.formatDate(new Date(), 'GMT+7', 'yyyy-MM-dd HH:mm:ss');
  sheet('FormTemplates').appendRow([genId('FRM'), p.name || p.fileName || 'แบบฟอร์ม', p.fileName || '',
    file.getId(), 'https://drive.google.com/file/d/' + file.getId() + '/view', now]);
  SpreadsheetApp.flush();
  return getFormTemplates();
}
function deleteFormTemplate(p) {
  var rec = readAll('FormTemplates').filter(function (x) { return x.ID === p.id; })[0];
  if (rec && rec.fileId) { try { DriveApp.getFileById(rec.fileId).setTrashed(true); } catch (e) {} }
  var row = findRowById('FormTemplates', p.id);
  if (row > 0) sheet('FormTemplates').deleteRow(row);
  SpreadsheetApp.flush();
  return getFormTemplates();
}

/* ============================================================
   แดชบอร์ดผู้บริหาร — เช็คชื่อ (วัน/เดือน/ภาค/ปี) + สถานะบันทึกคะแนน
   ============================================================ */
function calMap_() {
  var m = {};
  readAll('Calendar').forEach(function (c) {
    var d = ymd(c.date); if (!d) return;
    m[d] = m[d] || {};
    if (c.type === 'holiday') { m[d].holiday = true; m[d].title = c.title || ''; }
    if (c.type === 'makeup') m[d].makeup = true;
  });
  return m;
}
function isSchoolDay_(iso, cal) {
  var ev = cal[iso] || {};
  if (ev.makeup) return true;
  if (ev.holiday) return false;
  var pp = iso.split('-');
  var wd = new Date(Number(pp[0]), Number(pp[1]) - 1, Number(pp[2])).getDay();
  return wd >= 1 && wd <= 5;
}
function datesBetween_(from, to) {
  var out = [];
  var pf = from.split('-'), pt = to.split('-');
  var d = new Date(Number(pf[0]), Number(pf[1]) - 1, Number(pf[2]));
  var end = new Date(Number(pt[0]), Number(pt[1]) - 1, Number(pt[2]));
  while (d <= end && out.length < 500) {
    out.push(Utilities.formatDate(d, 'GMT+7', 'yyyy-MM-dd'));
    d.setDate(d.getDate() + 1);
  }
  return out;
}

// p: { kind:'flag'|'period', range:'day'|'month'|'semester'|'year', date?, month?, semester? }
function getExecAttendance(p) {
  p = p || {};
  var kind = p.kind || 'flag';
  var range = p.range || 'day';
  var set = getSettings();
  var sem = String(p.semester || set.currentSemester || '1');
  var cal = calMap_();
  var students = getStudents();
  var roomCount = {};
  students.forEach(function (s) { if (s.classLevel) roomCount[s.classLevel] = (roomCount[s.classLevel] || 0) + 1; });
  var roomList = Object.keys(roomCount).sort();
  var today = Utilities.formatDate(new Date(), 'GMT+7', 'yyyy-MM-dd');

  // ---------- ช่วงวันที่ตาม range ----------
  function rangesFor() {
    var out = [];
    if (range === 'semester') { var rr = semesterRange_(set, sem); if (rr) out.push(rr); }
    else { ['1', '2'].forEach(function (s2) { var rr2 = semesterRange_(set, s2); if (rr2) out.push(rr2); }); }
    return out;
  }

  /* ================= เช็คชื่อรายคาบ (เทียบตารางสอน) ================= */
  if (kind === 'period') {
    var TH_DAYS = ['อาทิตย์', 'จันทร์', 'อังคาร', 'พุธ', 'พฤหัสบดี', 'ศุกร์', 'เสาร์'];
    var schedMap = {}, maxPer = 0;   // day -> room -> {per:true}
    readAll('Schedule').forEach(function (r) {
      var d = String(r.day || '').trim(), room = String(r.classLevel || '').trim();
      var per = Number(r.period);
      if (!d || !room || !per) return;
      ((schedMap[d] = schedMap[d] || {})[room] = schedMap[d][room] || {})[String(per)] = true;
      if (per > maxPer) maxPer = per;
    });
    maxPer = maxPer || 6;
    var chkP = {};   // date|room -> {per:{n,present}}
    readAll('Attendance').forEach(function (a) {
      var per = Number(a.period);
      if (!per) return;   // ข้าม period 0 (หน้าเสาธง)
      var k = ymd(a.date) + '|' + a.classLevel;
      var o = (chkP[k] = chkP[k] || {});
      var c = (o[String(per)] = o[String(per)] || { n: 0, present: 0 });
      c.n++;
      if (a.status === 'มา' || a.status === 'สาย') c.present++;
    });
    var wdOf = function (iso) { var pp = iso.split('-'); return TH_DAYS[new Date(Number(pp[0]), Number(pp[1]) - 1, Number(pp[2])).getDay()]; };
    var expFor = function (iso, room) { var m = schedMap[wdOf(iso)] || {}; return Object.keys(m[room] || {}); };
    var doneFor = function (iso, room) {
      var got = chkP[iso + '|' + room] || {};
      return expFor(iso, room).filter(function (x) { return !!got[x]; }).length;
    };

    if (range === 'day') {
      var date = p.date ? ymd(p.date) : today;
      var rows = roomList.map(function (rm) {
        var exp = expFor(date, rm);
        var got = chkP[date + '|' + rm] || {};
        var slots = [];
        for (var q = 1; q <= maxPer; q++) {
          var qq = String(q);
          slots.push({ period: q, scheduled: exp.indexOf(qq) >= 0, checked: !!got[qq], present: got[qq] ? got[qq].present : 0 });
        }
        return { room: rm, students: roomCount[rm], slots: slots, exp: exp.length,
          done: exp.filter(function (x) { return !!got[x]; }).length };
      });
      return { kind: 'period', range: 'day', date: date, isSchoolDay: isSchoolDay_(date, cal),
        holidayTitle: (cal[date] && cal[date].title) || '', maxPeriod: maxPer, rooms: rows };
    }

    if (range === 'month') {
      var month = p.month || today.slice(0, 7);
      var mp = month.split('-');
      var nDays = new Date(Number(mp[0]), Number(mp[1]), 0).getDate();
      var days = [];
      for (var d = 1; d <= nDays; d++) {
        var iso = month + '-' + (d < 10 ? '0' : '') + d;
        days.push({ d: d, iso: iso, school: isSchoolDay_(iso, cal) });
      }
      var rows2 = roomList.map(function (rm) {
        var cells = days.map(function (dd) {
          if (!dd.school) return { e: 0, dn: 0 };
          return { e: expFor(dd.iso, rm).length, dn: doneFor(dd.iso, rm) };
        });
        return { room: rm, students: roomCount[rm], cells: cells };
      });
      return { kind: 'period', range: 'month', month: month, days: days, rooms: rows2 };
    }

    // semester / year
    var ranges = rangesFor();
    if (!ranges.length) return { kind: 'period', range: range, needDates: true };
    var allDates = [];
    ranges.forEach(function (rr) { allDates = allDates.concat(datesBetween_(rr.from, rr.to)); });
    var schoolDates = allDates.filter(function (iso) { return isSchoolDay_(iso, cal) && iso <= today; });
    var rows3 = roomList.map(function (rm) {
      var expN = 0, doneN = 0;
      schoolDates.forEach(function (iso) { expN += expFor(iso, rm).length; doneN += doneFor(iso, rm); });
      return { room: rm, students: roomCount[rm], expected: expN, checked: doneN,
        pct: expN ? Math.round(doneN / expN * 100) : 0 };
    });
    return { kind: 'period', range: range, semester: (range === 'semester' ? sem : ''),
      from: ranges[0].from, to: ranges[ranges.length - 1].to, schoolDays: schoolDates.length, rooms: rows3 };
  }

  /* ================= เช็คชื่อหน้าเสาธง (period 0) ================= */
  var chk = {};   // 'date|room' -> {n, present}
  readAll('Attendance').forEach(function (a) {
    if (String(a.period) !== '0') return;
    var k = ymd(a.date) + '|' + a.classLevel;
    if (!chk[k]) chk[k] = { n: 0, present: 0 };
    chk[k].n++;
    if (a.status === 'มา' || a.status === 'สาย') chk[k].present++;
  });

  if (range === 'day') {
    var date2 = p.date ? ymd(p.date) : today;
    var rows4 = roomList.map(function (r) {
      var c = chk[date2 + '|' + r];
      return { room: r, students: roomCount[r], checked: !!c, present: c ? c.present : 0 };
    });
    return { kind: 'flag', range: 'day', date: date2, isSchoolDay: isSchoolDay_(date2, cal),
      holidayTitle: (cal[date2] && cal[date2].title) || '', rooms: rows4,
      checkedRooms: rows4.filter(function (x) { return x.checked; }).length, totalRooms: roomList.length };
  }

  if (range === 'month') {
    var month2 = p.month || today.slice(0, 7);
    var mp2 = month2.split('-');
    var nDays2 = new Date(Number(mp2[0]), Number(mp2[1]), 0).getDate();
    var days2 = [];
    for (var d2 = 1; d2 <= nDays2; d2++) {
      var iso2 = month2 + '-' + (d2 < 10 ? '0' : '') + d2;
      days2.push({ d: d2, iso: iso2, school: isSchoolDay_(iso2, cal) });
    }
    var rows5 = roomList.map(function (r) {
      var dates = [];
      days2.forEach(function (dd) { if (chk[dd.iso + '|' + r]) dates.push(dd.iso); });
      return { room: r, students: roomCount[r], checkedDates: dates };
    });
    var schoolN = days2.filter(function (x) { return x.school && x.iso <= today; }).length;
    return { kind: 'flag', range: 'month', month: month2, days: days2, rooms: rows5, schoolDaysToDate: schoolN };
  }

  var ranges2 = rangesFor();
  if (!ranges2.length) return { kind: 'flag', range: range, needDates: true };
  var allDates2 = [];
  ranges2.forEach(function (rr) { allDates2 = allDates2.concat(datesBetween_(rr.from, rr.to)); });
  var schoolDates2 = allDates2.filter(function (iso) { return isSchoolDay_(iso, cal) && iso <= today; });
  var rows6 = roomList.map(function (r) {
    var done = schoolDates2.filter(function (iso) { return !!chk[iso + '|' + r]; }).length;
    return { room: r, students: roomCount[r], schoolDays: schoolDates2.length, checkedDays: done,
      pct: schoolDates2.length ? Math.round(done / schoolDates2.length * 100) : 0 };
  });
  return { kind: 'flag', range: range, semester: (range === 'semester' ? sem : ''), from: ranges2[0].from,
    to: ranges2[ranges2.length - 1].to, schoolDays: schoolDates2.length, rooms: rows6 };
}

// สถานะการบันทึกคะแนนรายวิชา: ครูผู้สอน หน่วย/คะแนนเต็ม นักเรียนที่มีคะแนน ตัดเกรดแล้ว
function getExecScoreStatus(p) {
  p = p || {};
  var sem = String(p.semester || getSettings().currentSemester || '1');
  var students = getStudents();
  var subjects = getSubjects().sort(function (a, b) {
    return String(a.classLevel).localeCompare(String(b.classLevel)) || String(a.name).localeCompare(String(b.name));
  });
  var semOk = function (v) { var s2 = String(v == null ? '' : v); return s2 === sem || s2 === ''; };
  var unitsBy = {};
  readAll('LearningUnits').forEach(function (u) {
    if (!semOk(u.semester)) return;
    (unitsBy[u.subjectID] = unitsBy[u.subjectID] || []).push(u);
  });
  var scoreStudents = {}, scoreCells = {}, scoredByUnit = {};
  readAll('Scores').forEach(function (s2) {
    if (!semOk(s2.semester)) return;
    if (s2.score === '' || s2.score == null) return;
    (scoreStudents[s2.subjectID] = scoreStudents[s2.subjectID] || {})[s2.studentID] = true;
    scoreCells[s2.subjectID] = (scoreCells[s2.subjectID] || 0) + 1;
    (scoredByUnit[s2.unitID] = scoredByUnit[s2.unitID] || {})[s2.studentID] = true;
  });
  var gradeStudents = {};
  readAll('Grades').forEach(function (g) {
    if (!semOk(g.semester)) return;
    if (g.grade51 === '' && (g.percent === '' || g.percent == null)) return;
    (gradeStudents[g.subjectID] = gradeStudents[g.subjectID] || {})[g.studentID] = true;
  });
  // ครูผู้สอนจากตารางสอน: subject name|grade → รายชื่อครู
  var teach = {};
  readAll('Schedule').forEach(function (r) {
    var k = String(r.subject || '').trim() + '|' + gradeOf_(r.classLevel);
    var t = String(r.teacher || '').trim(); if (!t) return;
    (teach[k] = teach[k] || {})[t] = true;
  });
  var rows = subjects.map(function (su) {
    var expected = students.filter(function (st) { return !su.classLevel || classMatch(su.classLevel, st.classLevel); }).length;
    var us = unitsBy[su.ID] || [];
    var maxTotal = us.reduce(function (a, u) { return a + (Number(u.maxScore) || 0); }, 0);
    var unitDetail = us.sort(function (a, b) { return Number(a.sortOrder) - Number(b.sortOrder); })
      .map(function (u) { return { name: u.unitName, maxScore: u.maxScore, scored: Object.keys(scoredByUnit[u.ID] || {}).length }; });
    var tKey = String(su.name).trim() + '|' + String(su.classLevel || '').trim();
    return {
      classLevel: su.classLevel || '', code: su.code || '', name: su.name,
      teachers: Object.keys(teach[tKey] || {}),
      units: us.length, maxTotal: maxTotal,
      scoredStudents: Object.keys(scoreStudents[su.ID] || {}).length,
      scoreCells: scoreCells[su.ID] || 0, unitDetail: unitDetail,
      gradedStudents: Object.keys(gradeStudents[su.ID] || {}).length,
      expected: expected
    };
  });
  return { semester: sem, subjects: rows };
}

/* ============================================================
   ภาพรวมรายห้อง (สำหรับผู้บริหาร) — สรุปความคืบหน้าแต่ละห้อง
   p: { semester? }
   ============================================================ */
function getDirectorOverview(p) {
  p = p || {};
  var sem = String(p.semester || getSettings().currentSemester || '1');
  var students = getStudents();
  var subjects = getSubjects();
  var grades = readAll('Grades').filter(function (g) {
    var gs = String(g.semester == null ? '' : g.semester);
    return gs === sem || gs === '';
  });
  var today = Utilities.formatDate(new Date(), 'GMT+7', 'yyyy-MM-dd');
  var att = readAll('Attendance').filter(function (a) { return ymd(a.date) === today && String(a.period) === '0'; });

  var gradeKey = {};
  grades.forEach(function (g) { gradeKey[g.studentID + '|' + g.subjectID] = g; });

  var byClass = {};
  students.forEach(function (st) { (byClass[st.classLevel] = byClass[st.classLevel] || []).push(st); });
  var classes = Object.keys(byClass).filter(Boolean).sort();

  var rows = classes.map(function (cls) {
    var sts = byClass[cls];
    var subs = subjects.filter(function (su) { return !su.classLevel || classMatch(su.classLevel, cls); });
    var expected = sts.length * subs.length, filled = 0, risk = 0;
    sts.forEach(function (st) {
      var hasFail = false;
      subs.forEach(function (su) {
        var g = gradeKey[st.ID + '|' + su.ID];
        if (g && g.grade51 !== '' && g.grade51 != null) {
          filled++;
          if (String(g.grade51) === '0') hasFail = true;
        } else if (g && g.percent !== '' && g.percent != null && Number(g.percent) < 50) {
          hasFail = true;
        }
      });
      if (hasFail) risk++;
    });
    var attRows = att.filter(function (a) { return String(a.classLevel) === cls; });
    var present = attRows.filter(function (a) { return a.status === 'มา' || a.status === 'สาย'; }).length;
    var homeroom = readAll('Teachers').filter(function (t) { return String(t.homeroomClass) === cls; }).map(function (t) { return t.name; });
    return {
      classLevel: cls, students: sts.length, subjects: subs.length,
      gradePct: expected ? Math.round(filled / expected * 100) : 0,
      atRisk: risk, attChecked: attRows.length > 0, attPresent: present, attTotal: sts.length,
      homeroom: homeroom
    };
  });

  return {
    semester: sem, classes: rows,
    totals: {
      classes: rows.length, students: students.length,
      teachers: getTeachers().filter(function (t) { return t.role !== 'admin' && t.role !== 'director'; }).length,
      atRisk: rows.reduce(function (a, r) { return a + r.atRisk; }, 0)
    }
  };
}

/* ============================================================
   ระบบติดตามนักเรียนเสี่ยง (รายวิชา) — ณ วันปัจจุบัน
   1) เสี่ยงผลการเรียน: % สะสมของวิชา (คะแนนที่กรอกแล้ว / เต็มของหน่วยที่กรอกแล้ว) < เกณฑ์ (ค่าตั้งต้น 50)
   2) เสี่ยงติด มส.: เวลาเรียนรายวิชา (คาบที่มา/คาบที่ต้องเรียนตามตารางสอน) < เกณฑ์ (ค่าตั้งต้น 80)
   p: { classLevel?, semester?, gradeMin=50, attMin=80, critical=40 }
   ============================================================ */
function getRiskStudents(p) {
  p = p || {};
  var thAtt   = (p.attMin   != null) ? Number(p.attMin)   : 80; // % เวลาเรียนขั้นต่ำ (มส.)
  var thGrade = (p.gradeMin != null) ? Number(p.gradeMin) : 50; // % ผลการเรียนขั้นต่ำ
  var thCrit  = (p.critical != null) ? Number(p.critical) : 40; // % วิกฤต
  var semester = p.semester ? String(p.semester) : '';          // '' = ทุกภาคเรียน

  var students = getStudents();
  if (p.classLevel) students = students.filter(function (s) { return String(s.classLevel) === String(p.classLevel); });
  students.sort(function (a, b) { return String(a.classLevel).localeCompare(String(b.classLevel)) || Number(a.number) - Number(b.number); });

  // ---------- ผลการเรียนสะสม (รายคน-รายวิชา) ----------
  var subjects = getSubjects();
  var unitMax = {};
  readAll('LearningUnits').forEach(function (u) {
    if (semester && u.semester && String(u.semester) !== semester) return;
    unitMax[u.ID] = Number(u.maxScore || 0);
  });
  var scAgg = {}; // sid|subjectID -> {got, max}
  readAll('Scores').forEach(function (sc) {
    if (semester && sc.semester && String(sc.semester) !== semester) return;
    var mx = unitMax[sc.unitID]; if (mx == null) return;       // หน่วยไม่อยู่ในภาคนี้/ไม่มีคะแนนเต็ม
    var k = sc.studentID + '|' + sc.subjectID;
    var a = scAgg[k] || (scAgg[k] = { got: 0, max: 0 });
    a.got += Number(sc.score || 0);
    a.max += mx;                                                // นับเฉพาะหน่วยที่ "ทำแล้ว" = ณ วันนี้
  });

  // ---------- เวลาเรียนรายคาบ (จากตารางสอน + การเช็คชื่อ) ----------
  var att = readAll('Attendance');
  var range = semesterRange_(getSettings(), semester);
  if (range) att = att.filter(function (r) { var d = ymd(r.date); return d >= range.from && d <= range.to; });
  var look = {}, flag = {}, datesByClass = {};
  att.forEach(function (r) {
    var c = String(r.classLevel), d = ymd(r.date);
    look[c + '|' + d + '|' + r.period + '|' + r.studentID] = r.status;
    if (String(r.period) === '0') { flag[c + '|' + d + '|' + r.studentID] = r.status; (datesByClass[c] = datesByClass[c] || {})[d] = true; }
  });
  var schByClassDay = {}; // [class][weekday][period] = subjectName
  readAll('Schedule').forEach(function (s) {
    var c = String(s.classLevel);
    (schByClassDay[c] = schByClassDay[c] || {});
    (schByClassDay[c][s.day] = schByClassDay[c][s.day] || {})[Number(s.period)] = s.subject;
  });

  var out = [];
  students.forEach(function (st) {
    var c = String(st.classLevel);

    // (1) ผลการเรียนต่ำกว่าเกณฑ์ — รายวิชา
    var lowSubjects = [];
    subjects.forEach(function (su) {
      if (su.classLevel && !classMatch(su.classLevel, c)) return;
      var a = scAgg[st.ID + '|' + su.ID];
      if (!a || a.max <= 0) return;                              // ยังไม่มีคะแนน = ยังไม่ประเมิน
      var pct = Math.round(a.got / a.max * 100);
      if (pct < thGrade) lowSubjects.push({ subject: su.name || su.code || '-', percent: pct });
    });

    // (2) เวลาเรียนต่ำกว่าเกณฑ์ (เสี่ยง มส.) — รายวิชา ตามคาบในตารางสอน
    var msSubjects = [];
    var bySubj = {}; // subjectName -> {exp, pres}
    var sch = schByClassDay[c] || {};
    Object.keys(datesByClass[c] || {}).forEach(function (d) {
      var wd = WD_TH[new Date(d + 'T00:00:00').getDay()];
      var periods = sch[wd] || {};
      Object.keys(periods).forEach(function (per) {
        var nm = periods[per];
        var b = bySubj[nm] || (bySubj[nm] = { exp: 0, pres: 0 });
        b.exp++;
        var status = look[c + '|' + d + '|' + per + '|' + st.ID] || flag[c + '|' + d + '|' + st.ID];
        if (status === 'มา' || status === 'สาย') b.pres++;
      });
    });
    Object.keys(bySubj).forEach(function (nm) {
      var b = bySubj[nm];
      if (b.exp <= 0) return;
      var pct = Math.round(b.pres / b.exp * 100);
      if (pct < thAtt) {
        msSubjects.push({ subject: nm, percent: pct, present: b.pres, expected: b.exp });
      }
    });

    if (lowSubjects.length || msSubjects.length) {
      var minPct = lowSubjects.length ? lowSubjects.reduce(function (m, x) { return Math.min(m, x.percent); }, 100) : 100;
      var level = (lowSubjects.length && minPct < thCrit) ? 'วิกฤต' : 'เสี่ยง';
      out.push({
        studentID: st.ID, number: st.number, fullName: st.fullName, classLevel: st.classLevel,
        level: level, minPercent: lowSubjects.length ? minPct : null,
        lowSubjects: lowSubjects, msSubjects: msSubjects
      });
    }
  });

  // เรียง: วิกฤตก่อน แล้วตาม % ต่ำสุด
  out.sort(function (a, b) {
    if (a.level !== b.level) return a.level === 'วิกฤต' ? -1 : 1;
    return (a.minPercent == null ? 999 : a.minPercent) - (b.minPercent == null ? 999 : b.minPercent);
  });
  var counts = { grade: 0, ms: 0, critical: 0, total: out.length };
  out.forEach(function (r) { if (r.lowSubjects.length) counts.grade++; if (r.msSubjects.length) counts.ms++; if (r.level === 'วิกฤต') counts.critical++; });
  return { thresholds: { gradeMin: thGrade, attMin: thAtt, critical: thCrit }, semester: semester, counts: counts, total: out.length, students: out };
}

// ช่วงวันที่ของภาคเรียนจาก Settings (sem1Start/sem1End/sem2Start/sem2End) — คืน null ถ้าไม่ได้ตั้ง
/* ============================================================
   สรุปเวลาเรียนรายวิชา — จากตารางสอนจริง × วันเปิดเรียน × เช็คชื่อรายคาบ
   กติกา: คาบที่ไม่ได้เช็ค ใช้ผลหน้าเสาธงของวันนั้น (ทำเครื่องหมาย flag) / ไม่มีทั้งคู่ = ถือว่ามา (none)
   p: { classLevel, subject, semester }
   ============================================================ */
function getSubjectAttendance(p) {
  p = p || {};
  var room = String(p.classLevel || '').trim(), subject = String(p.subject || '').trim();
  if (!room || !subject) throw new Error('ระบุวิชาและห้อง');
  var set = getSettings();
  var sem = String(p.semester || set.currentSemester || '1');
  var range = semesterRange_(set, sem);
  var today = Utilities.formatDate(new Date(), 'GMT+7', 'yyyy-MM-dd');
  // ตารางสอนของวิชานี้ในห้องนี้: weekday -> [period]
  var TH_DAYS = ['อาทิตย์', 'จันทร์', 'อังคาร', 'พุธ', 'พฤหัสบดี', 'ศุกร์', 'เสาร์'];
  var byWd = {}, teacher = '';
  readAll('Schedule').forEach(function (r) {
    if (String(r.classLevel || '').trim() !== room || String(r.subject || '').trim() !== subject) return;
    var d = String(r.day || '').trim(), per = Number(r.period);
    if (!d || !per) return;
    (byWd[d] = byWd[d] || []).push(per);
    if (!teacher && r.teacher) teacher = String(r.teacher).trim();
  });
  Object.keys(byWd).forEach(function (k) { byWd[k].sort(function (a, b) { return a - b; }); });
  var su = getSubjects().filter(function (x) { return String(x.name).trim() === subject && classMatch(x.classLevel, room); })[0] || {};
  var out = { subject: subject, code: su.code || '', classLevel: room, semester: sem, teacher: teacher,
    hours: Number(su.hours) || 0, needDates: !range, dates: [], students: [], summary: {} };
  if (!range) return out;
  var cal = calMap_();
  var to = range.to < today ? range.to : today;
  var dates = [];
  datesBetween_(range.from, to).forEach(function (iso) {
    if (!isSchoolDay_(iso, cal)) return;
    var wd = TH_DAYS[new Date(Number(iso.slice(0, 4)), Number(iso.slice(5, 7)) - 1, Number(iso.slice(8, 10))).getDay()];
    (byWd[wd] || []).forEach(function (per) { dates.push({ date: iso, period: String(per) }); });
  });
  out.dates = dates;
  // เช็คชื่อของห้องนี้ในช่วง: key date|period -> {studentID: status}
  var att = {};
  readAll('Attendance').forEach(function (a) {
    if (String(a.classLevel).trim() !== room) return;
    var d = ymd(a.date);
    if (d < range.from || d > to) return;
    var k = d + '|' + String(Number(a.period) || 0);
    (att[k] = att[k] || {})[String(a.studentID)] = String(a.status || '');
  });
  dates.forEach(function (dt) {
    var kP = dt.date + '|' + dt.period, kF = dt.date + '|0';
    dt.source = att[kP] ? 'period' : (att[kF] ? 'flag' : 'none');
  });
  var students = getStudents().filter(function (st) { return String(st.classLevel).trim() === room; })
    .sort(function (a, b) { return (Number(a.number) || 0) - (Number(b.number) || 0); });
  var below80 = 0;
  out.students = students.map(function (st) {
    var marks = {}, c = { 'มา': 0, 'ขาด': 0, 'ลา': 0, 'สาย': 0 };
    dates.forEach(function (dt) {
      var kP = dt.date + '|' + dt.period, kF = dt.date + '|0';
      var v = (att[kP] && att[kP][st.ID]) || (att[kF] && att[kF][st.ID]) || 'มา';
      if (c[v] == null) v = 'มา';
      marks[kP] = v; c[v]++;
    });
    var total = dates.length;
    var attended = c['มา'] + c['สาย'];
    var pct = total ? Math.round(attended / total * 1000) / 10 : null;
    if (pct != null && pct < 80) below80++;
    return { studentID: st.ID, number: st.number, fullName: st.fullName, marks: marks,
      present: c['มา'], absent: c['ขาด'], leave: c['ลา'], late: c['สาย'], attended: attended, total: total, pct: pct };
  });
  out.summary = {
    periodsTotal: dates.length,
    checked: dates.filter(function (d) { return d.source === 'period'; }).length,
    flagOnly: dates.filter(function (d) { return d.source === 'flag'; }).length,
    none: dates.filter(function (d) { return d.source === 'none'; }).length,
    below80: below80, from: range.from, to: to
  };
  return out;
}

// ตรวจเวลาเรียนแบบเบา (ไม่ส่งรายคาบ) สำหรับหน้าบันทึกคะแนน: p{classLevel, subject, semester}
function getSubjectTimeCheck(p) {
  var d = getSubjectAttendance(p);
  var out = { needDates: !!d.needDates, periodsTotal: (d.summary && d.summary.periodsTotal) || 0, students: {}, below: [] };
  (d.students || []).forEach(function (st) {
    out.students[st.studentID] = { pct: st.pct, total: st.total, attended: st.attended, absent: st.absent };
    if (st.pct != null && st.pct < 80) out.below.push(st.studentID);
  });
  return out;
}

/* ============================================================
   งานค้างของครู (หน้าแรก): คาบยังไม่เช็ค • วันหน้าเสาธงยังไม่เช็ค • วิชายังไม่ตัดเกรด • เอกสารค้างส่ง
   p: { teacher, role, homeroomClass, semester }
   ============================================================ */
function getMyTodo(p) {
  p = p || {};
  var teacher = String(p.teacher || '').trim();
  if (!teacher) return {};
  var set = getSettings();
  var sem = String(p.semester || set.currentSemester || '1');
  var range = semesterRange_(set, sem);
  var today = Utilities.formatDate(new Date(), 'GMT+7', 'yyyy-MM-dd');
  var out = { semester: sem, needDates: !range, uncheckedPeriods: 0, uncheckedFlagDays: 0, ungradedSubjects: [], docsMissing: 0, docsTotal: 0 };
  var tc = teacherCore(teacher);
  var TH_DAYS = ['อาทิตย์', 'จันทร์', 'อังคาร', 'พุธ', 'พฤหัสบดี', 'ศุกร์', 'เสาร์'];
  if (range) {
    var to = range.to < today ? range.to : today;
    var cal = calMap_();
    var days = datesBetween_(range.from, to).filter(function (iso) { return isSchoolDay_(iso, cal); });
    // คาบสอนของครู
    var byWd = {};
    readAll('Schedule').forEach(function (r) {
      if (teacherCore(r.teacher) !== tc) return;
      var d = String(r.day || '').trim(), cls = String(r.classLevel || '').trim(), per = Number(r.period);
      if (!d || !cls || !per) return;
      (byWd[d] = byWd[d] || []).push(cls + '|' + per);
    });
    var chk = {}, flagDays = {};
    readAll('Attendance').forEach(function (a) {
      var d = ymd(a.date);
      if (d < range.from || d > to) return;
      var per = Number(a.period) || 0;
      if (per === 0) { flagDays[d + '|' + String(a.classLevel).trim()] = 1; return; }
      chk[d + '|' + String(a.classLevel).trim() + '|' + per] = 1;
    });
    var homeroom = String(p.homeroomClass || '').trim();
    days.forEach(function (iso) {
      var wd = TH_DAYS[new Date(Number(iso.slice(0, 4)), Number(iso.slice(5, 7)) - 1, Number(iso.slice(8, 10))).getDay()];
      (byWd[wd] || []).forEach(function (k) { if (!chk[iso + '|' + k]) out.uncheckedPeriods++; });
      if (homeroom && !flagDays[iso + '|' + homeroom]) out.uncheckedFlagDays++;
    });
  }
  // วิชาที่รับผิดชอบแต่ยังไม่มีเกรดเลยในภาคเรียนนี้
  try {
    var mine = getMyTeach({ teacher: teacher });
    var graded = {};
    readAll('Grades').forEach(function (g) {
      if (String(g.semester || '1') !== sem) return;
      if (String(g.grade51 == null ? '' : g.grade51) === '') return;
      graded[String(g.subjectID)] = 1;
    });
    var seen = {};
    (mine || []).forEach(function (m) {
      var key = String(m.subjectID);
      if (seen[key] || graded[key]) return;
      seen[key] = 1;
      out.ungradedSubjects.push({ subjectID: m.subjectID, name: m.name || '', room: m.room || '' });
    });
  } catch (e) {}
  // เอกสารการสอนค้างส่ง
  try {
    var tv = getTeachingOverview({ semester: sem });
    (tv.teachers || []).forEach(function (t) {
      if (teacherCore(t.teacher) !== tc) return;
      out.docsTotal += Number(t.total) || 0;
      var done = 0;
      Object.keys(t.byType || {}).forEach(function (k) { done += Number(t.byType[k].done) || 0; });
      out.docsMissing += Math.max(0, (Number(t.total) || 0) - done);
    });
  } catch (e) {}
  return out;
}

function semesterRange_(set, semester) {
  if (!semester) return null;
  var s = set['sem' + semester + 'Start'], e = set['sem' + semester + 'End'];
  if (!s || !e) return null;
  return { from: String(s), to: String(e) };
}

/** ====================== ล็อก/อนุมัติผลการเรียน (ต่อห้อง) ====================== */
function getLocks() { return readAll('Locks'); }
// สถานะ workflow: open(กำลังกรอก) → submitted(ส่งตรวจ) → locked(อนุมัติ/ล็อก) | returned(ตีกลับให้แก้)
// p: { classLevel, status?, note?, by, locked? (รองรับแบบเก่า true/false) }
function setLock(p) {
  var s = sheet('Locks');
  var status = p.status;
  if (status == null && p.locked != null) status = p.locked ? 'locked' : 'open';
  if (['open', 'submitted', 'locked', 'returned'].indexOf(String(status)) < 0) throw new Error('สถานะไม่ถูกต้อง');
  var data = s.getDataRange().getValues(); var row = -1;
  for (var i = 1; i < data.length; i++) { if (String(data[i][0]) === String(p.classLevel)) { row = i + 1; break; } }
  var now = Utilities.formatDate(new Date(), 'GMT+7', 'yyyy-MM-dd HH:mm:ss');
  var rec = [p.classLevel, status, now, p.by || '', p.note || ''];
  if (row < 0) s.appendRow(rec); else s.getRange(row, 1, 1, 5).setValues([rec]);
  SpreadsheetApp.flush();
  return getLocks();
}
function lockedRoomsSet_() {
  // ห้ามแก้ไขทั้งห้องที่ "อนุมัติ/ล็อกแล้ว" และห้องที่ "ส่งตรวจแล้ว" (รอผู้บริหารตรวจ)
  var o = {}; readAll('Locks').forEach(function (r) {
    var st = String(r.status);
    if (st === 'locked' || st === 'submitted') o[String(r.classLevel)] = true;
  });
  return o;
}
// ตรวจว่ามีนักเรียนอยู่ในห้องที่ถูกล็อกหรือไม่ → ถ้ามี โยน error
function assertNotLocked_(studentIDs) {
  var locked = lockedRoomsSet_();
  if (!Object.keys(locked).length) return;
  var room = {}; getStudents().forEach(function (s) { room[s.ID] = String(s.classLevel); });
  var hit = {};
  (studentIDs || []).forEach(function (id) { var r = room[id]; if (r && locked[r]) hit[r] = true; });
  var rooms = Object.keys(hit);
  if (rooms.length) throw new Error('ห้อง ' + rooms.join(', ') + ' ถูกล็อก (อนุมัติผลแล้ว) — ให้แอดมินปลดล็อกก่อนจึงจะแก้ไขได้');
}

/** ====================== ปฏิทินวิชาการ ====================== */
function getCalendar() {
  return readAll('Calendar').map(function (r) { r.date = ymd(r.date); return r; })
    .sort(function (a, b) { return String(a.date).localeCompare(String(b.date)); });
}
function addCalendarEvent(p) {
  if (!p.date || !p.type) throw new Error('ระบุวันที่และประเภท');
  var s = sheet('Calendar');
  s.appendRow([genId('CAL'), '', p.type, p.title || '', p.note || '']);
  // เขียน date เป็นข้อความ (กัน Sheets แปลงเป็น Date)
  var row = s.getLastRow();
  s.getRange(row, 2).setNumberFormat('@').setValue(ymd(p.date));
  SpreadsheetApp.flush();
  return getCalendar();
}
function deleteCalendarEvent(id) {
  var row = findRowById('Calendar', id);
  if (row < 0) throw new Error('ไม่พบรายการ');
  sheet('Calendar').deleteRow(row);
  SpreadsheetApp.flush();
  return getCalendar();
}

/** ====================== เล่มสรุปผลการเรียน (รวมทุกหน้า) ====================== */
function getPP5Book(p) {
  p = p || {};
  var classLevel = p.classLevel;
  var subjectMode = p.scope === 'subject';
  var students = getStudents().filter(function (s) { return String(s.classLevel) === String(classLevel); });
  students.sort(function (a, b) { return Number(a.number) - Number(b.number); });

  var subjects = getSubjects().filter(function (sub) { return !sub.classLevel || classMatch(sub.classLevel, classLevel); })
    .sort(function (a, b) { return Number(a.sortOrder) - Number(b.sortOrder); });
  if (subjectMode && p.subjectID) subjects = subjects.filter(function (s) { return String(s.ID) === String(p.subjectID); });

  var gradeBy = {};
  getGrades().forEach(function (g) { (gradeBy[g.studentID] = gradeBy[g.studentID] || {})[g.subjectID] = g; });

  var docsBy = {};
  subjects.forEach(function (sub) { docsBy[sub.ID] = getSubjectDocs(sub.ID); });

  var from = p.from || '2000-01-01', to = p.to || '2999-12-31';
  var register = getAttendanceRegister(classLevel, from, to);
  var grids = register.months.map(function (m) { return getMonthlyGrid(classLevel, m.ym); });

  var detail = readAll('AbilityDetail');
  var byStu = {};
  detail.forEach(function (r) { (byStu[r.studentID] = byStu[r.studentID] || []).push(r); });
  var abilityRaw = students.map(function (st) { return { studentID: st.ID, rows: byStu[st.ID] || [] }; });

  return {
    settings: getSettings(),
    classLevel: classLevel,
    scope: subjectMode ? 'subject' : 'class',
    cover: getPP5Cover(classLevel),
    stats: getGradeStats({ classLevel: classLevel }),
    students: students.map(function (s) {
      return { ID: s.ID, studentCode: s.studentCode, number: s.number, fullName: s.fullName,
        prefix: s.prefix, firstName: s.firstName, lastName: s.lastName, classLevel: s.classLevel };
    }),
    subjects: subjects.map(function (s) { return { ID: s.ID, code: s.code, name: s.name, classLevel: s.classLevel }; }),
    docs: docsBy,
    grades: students.map(function (st) {
      return { studentID: st.ID, bysubject: subjects.map(function (sub) {
        var g = (gradeBy[st.ID] || {})[sub.ID] || {};
        return { subjectID: sub.ID, grade51: g.grade51, percent: g.percent };
      }) };
    }),
    register: register,
    grids: grids,
    abilityRaw: abilityRaw
  };
}

/** ====================== WP3: ปกเล่ม สรุปผลรายห้อง ====================== */

function getPP5Cover(classLevel) {
  var settings = getSettings();
  var students = getStudents().filter(function (s) { return String(s.classLevel) === String(classLevel); });
  var ids = {}; students.forEach(function (s) { ids[s.ID] = true; });

  // นับชาย/หญิง จากคำนำหน้า
  var male = 0, female = 0;
  students.forEach(function (s) {
    var n = String(s.fullName || '');
    if (/^(เด็กชาย|ด\.ช\.|นาย)/.test(n)) male++;
    else if (/^(เด็กหญิง|ด\.ญ\.|นางสาว|นาง)/.test(n)) female++;
  });

  // การกระจายผลการเรียน (เกรด 2551) รายวิชา
  var GR = ['4', '3.5', '3', '2.5', '2', '1.5', '1', '0'];
  var subs = getSubjects().filter(function (sub) {
    return !sub.classLevel || classMatch(sub.classLevel, classLevel);
  }).sort(function (a, b) { return Number(a.sortOrder) - Number(b.sortOrder); });
  var gradeBy = {};
  readAll('Grades').forEach(function (g) { if (ids[g.studentID]) (gradeBy[g.studentID] = gradeBy[g.studentID] || {})[g.subjectID] = g; });
  var subjects = subs.map(function (sub) {
    var counts = {}; GR.forEach(function (k) { counts[k] = 0; });
    students.forEach(function (st) {
      var g = (gradeBy[st.ID] || {})[sub.ID];
      if (g && g.grade51 !== '' && g.grade51 != null && counts[String(g.grade51)] !== undefined) counts[String(g.grade51)]++;
    });
    return { code: sub.code || '', name: sub.name, counts: GR.map(function (k) { return counts[k]; }) };
  });

  // สรุปผลประเมินจากรายตัวชี้วัด
  var detail = readAll('AbilityDetail').filter(function (r) { return ids[r.studentID]; });
  var byStu = {};
  detail.forEach(function (r) { (byStu[r.studentID] = byStu[r.studentID] || []).push(r); });
  var LV = ['ผ่าน', 'ไม่ผ่าน'];
  var charCounts = { 'ผ่าน': 0, 'ไม่ผ่าน': 0 };
  var rtCounts = { 'ผ่าน': 0, 'ไม่ผ่าน': 0 };
  var actPass = 0, actFail = 0;
  students.forEach(function (st) {
    var rows = byStu[st.ID] || [];
    var ab = summarizeAbilityDetail(rows);
    if (ab['คุณลักษณะอันพึงประสงค์']) charCounts[ab['คุณลักษณะอันพึงประสงค์']]++;
    if (ab['การอ่าน คิดวิเคราะห์ และเขียน']) rtCounts[ab['การอ่าน คิดวิเคราะห์ และเขียน']]++;
    var acts = rows.filter(function (r) { return r.group === 'activity' && r.item !== 'ชื่อชุมนุม'; });
    if (acts.length) {
      var fail = acts.some(function (r) { return r.value === 'ไม่ผ่าน'; });
      if (fail) actFail++; else actPass++;
    }
  });

  // ครูประจำชั้น (จากหน้าจัดการครู)
  var homeroom = readAll('Teachers')
    .filter(function (t) { return String(t.homeroomClass) === String(classLevel); })
    .map(function (t) { return t.name; });

  return {
    settings: settings, classLevel: classLevel,
    totals: { all: students.length, male: male, female: female },
    gradeLevels: GR, subjects: subjects,
    levels: LV,
    charCounts: LV.map(function (k) { return charCounts[k]; }),
    rtCounts: LV.map(function (k) { return rtCounts[k]; }),
    activity: { pass: actPass, fail: actFail },
    homeroomTeachers: homeroom
  };
}


/** ====================== REPORT DATA ====================== */


// ชั้นตรงกันไหม รองรับ "ป.3", "ป.3/1" ฯลฯ (จับคู่ตามตัวเลขชั้น)
function classMatch(subjectClass, studentClass) {
  if (!subjectClass || !studentClass) return false;
  var a = String(subjectClass).replace(/[^0-9]/g, '').charAt(0);
  var b = String(studentClass).replace(/[^0-9]/g, '').charAt(0);
  return a && b && a === b;
}

// รวมข้อมูลนักเรียนทั้งชั้น สำหรับพิมพ์ใบรายงานทีเดียว (เรียก API ครั้งเดียว)
function getReportBatch(classLevel, onlyStudentID) {
  // onlyStudentID: สร้างรายงานรายคน (รวมนักเรียนที่ย้ายออก — พิมพ์ ปพ.6 ย้อนหลัง)
  var students = onlyStudentID
    ? getStudentsAll_().filter(function (st) { return String(st.ID) === String(onlyStudentID); })
    : getStudents();
  if (classLevel) {
    if (String(classLevel).indexOf('/') >= 0) {
      // ระบุห้องเต็ม เช่น ป.6/1 → เฉพาะห้องนั้น
      students = students.filter(function (s) { return String(s.classLevel) === String(classLevel); });
    } else {
      students = students.filter(function (s) { return classMatch(classLevel, s.classLevel); });
    }
  }
  students.sort(function (a, b) {
    return String(a.classLevel).localeCompare(String(b.classLevel)) || Number(a.number) - Number(b.number);
  });

  var settings = getSettings();
  var allSubjects = getSubjects();
  var allGrades = getGrades();
  var allAbilDetail = readAll('AbilityDetail');
  var allTeachers = readAll('Teachers');

  var gradeBy = {};
  allGrades.forEach(function (g) { (gradeBy[g.studentID] = gradeBy[g.studentID] || {})[g.subjectID] = g; });
  // เกรดแยกภาคเรียน (ปพ.6 แสดง ภาค 1 / ภาค 2)
  var g1By = {}, g2By = {};
  readAll('Grades').forEach(function (g) {
    var sm = String(g.semester == null ? '' : g.semester) || '1';
    var tgt = (sm === '2') ? g2By : g1By;
    (tgt[g.studentID] = tgt[g.studentID] || {})[g.subjectID] = g;
  });
  var abilDetailBy = {};
  allAbilDetail.forEach(function (a) { (abilDetailBy[a.studentID] = abilDetailBy[a.studentID] || []).push(a); });
  // ครูประจำชั้นตามห้อง (ห้อง → รายชื่อครู)
  var homeroomBy = {};
  allTeachers.forEach(function (t) {
    if (t.homeroomClass) (homeroomBy[t.homeroomClass] = homeroomBy[t.homeroomClass] || []).push(t.name);
  });

  // สรุปเวลาเรียน (หน้าเสาธง period 0) ต่อนักเรียน
  var attRec = readAll('Attendance').filter(function (r) { return String(r.period) === '0'; });
  var schoolDaysByClass = {}, attLook = {};
  attRec.forEach(function (r) {
    var d = ymd(r.date), c = String(r.classLevel);
    (schoolDaysByClass[c] = schoolDaysByClass[c] || {})[d] = true;
    attLook[c + '|' + d + '|' + r.studentID] = r.status;
  });

  var out = students.map(function (student) {
    var subs = allSubjects.filter(function (sub) {
      return !sub.classLevel || classMatch(sub.classLevel, student.classLevel);
    }).sort(function (a, b) { return Number(a.sortOrder) - Number(b.sortOrder); });
    var rows = subs.map(function (sub) {
      var gg = (gradeBy[student.ID] || {})[sub.ID] || {};
      var gg1 = (g1By[student.ID] || {})[sub.ID] || {};
      var gg2 = (g2By[student.ID] || {})[sub.ID] || {};
      return {
        code: sub.code, name: sub.name, learningArea: sub.learningArea, hours: sub.hours, type: sub.type,
        percent: gg.percent, grade51: gg.grade51, letter51: gg.letter51 || '',
        g1: (gg1.grade51 == null ? '' : gg1.grade51), g2: (gg2.grade51 == null ? '' : gg2.grade51)
      };
    });
    // GPA รายภาค + ตลอดปี (นับเฉพาะเกรดตัวเลข ถ่วงน้ำหนักด้วยชั่วโมงเรียน)
    var gpaOf = function (key) {
      var sw = 0, sg = 0;
      rows.forEach(function (r) {
        var v = parseFloat(r[key]); var h = Number(r.hours) || 1;
        if (!isNaN(v)) { sw += h; sg += v * h; }
      });
      return sw ? Math.round(sg / sw * 100) / 100 : null;
    };
    var gpa1 = gpaOf('g1'), gpa2 = gpaOf('g2');
    var swY = 0, sgY = 0;
    rows.forEach(function (r) {
      ['g1', 'g2'].forEach(function (k) {
        var v = parseFloat(r[k]); var h = Number(r.hours) || 1;
        if (!isNaN(v)) { swY += h; sgY += v * h; }
      });
    });
    var gpaYear = swY ? Math.round(sgY / swY * 100) / 100 : null;
    var abObj = summarizeAbilityDetail(abilDetailBy[student.ID] || []);
    var abilities = Object.keys(abObj).map(function (k) { return { itemKey: k, result: abObj[k] }; });
    // เวลาเรียนของนักเรียนคนนี้
    var c = String(student.classLevel);
    var days = Object.keys(schoolDaysByClass[c] || {});
    var present = 0, absent = 0, leave = 0, late = 0;
    days.forEach(function (d) {
      var s = attLook[c + '|' + d + '|' + student.ID];
      if (_attendedDay(s)) present++;
      if (s === 'ขาด') absent++; else if (s === 'ลา') leave++; else if (s === 'สาย') late++;
    });
    var attendance = { schoolDays: days.length, present: present, absent: absent, leave: leave, late: late,
      percent: days.length ? Math.round(present / days.length * 100) : 0 };
    return { student: student, subjects: rows, abilities: abilities, abilityGroups: [],
      gpa1: gpa1, gpa2: gpa2, gpaYear: gpaYear,
      homeroomTeachers: homeroomBy[student.classLevel] || [], attendance: attendance };
  });

  return { settings: settings, students: out };
}


/** ====================== WP3: สรุปผลการเรียนรายชั้น (หน้าปก) ====================== */
function getClassGradeSummary(classLevel, semester) {
  var settings = getSettings();
  var sem = semester ? String(semester) : '';
  var students = getStudents().filter(function (s) { return String(s.classLevel) === String(classLevel); });
  students.sort(function (a, b) { return Number(a.number) - Number(b.number); });
  var subjects = getSubjects().filter(function (sub) {
    return !sub.classLevel || classMatch(sub.classLevel, classLevel);
  }).sort(function (a, b) { return Number(a.sortOrder) - Number(b.sortOrder); });

  var gradeBy = {};
  readAll('Grades').filter(function (g) { return !sem || String(g.semester || '1') === sem; })
    .forEach(function (g) { (gradeBy[g.studentID] = gradeBy[g.studentID] || {})[g.subjectID] = g; });
  var abBy = {};
  readAll('AbilityDetail').filter(function (a) { return !sem || String(a.semester || '1') === sem; })
    .forEach(function (a) { (abBy[a.studentID] = abBy[a.studentID] || []).push(a); });

  var LEVELS = ['4', '3.5', '3', '2.5', '2', '1.5', '1', '0'];
  var subjRows = subjects.map(function (sub) {
    var counts = {}; LEVELS.forEach(function (l) { counts[l] = 0; });
    students.forEach(function (st) {
      var g = (gradeBy[st.ID] || {})[sub.ID];
      if (g && g.grade51 !== '' && g.grade51 != null) {
        var k = String(g.grade51);
        if (counts[k] !== undefined) counts[k]++;
      }
    });
    return { code: sub.code, name: sub.name, counts: counts };
  });

  var charCount = { 'ดีเยี่ยม': 0, 'ดี': 0, 'ผ่าน': 0, 'ไม่ผ่าน': 0 };
  var rtCount = { 'ดีเยี่ยม': 0, 'ดี': 0, 'ผ่าน': 0, 'ไม่ผ่าน': 0 };
  var actPass = 0, actFail = 0;
  students.forEach(function (st) {
    var ab = summarizeAbilityDetail(abBy[st.ID] || []);
    var c = ab['คุณลักษณะอันพึงประสงค์']; if (c && charCount[c] !== undefined) charCount[c]++;
    var r = ab['การอ่าน คิดวิเคราะห์ และเขียน']; if (r && rtCount[r] !== undefined) rtCount[r]++;
    var acts = Object.keys(ab).filter(function (k) { return /กิจกรรม|ลูกเสือ|ชุมนุม|ชมรม|เนตรนารี/.test(k); });
    if (acts.length) { if (acts.every(function (k) { return ab[k] === 'ผ่าน'; })) actPass++; else actFail++; }
  });

  var male = 0, female = 0;
  students.forEach(function (st) {
    var n = String(st.fullName || '');
    if (/^(เด็กชาย|นาย|ด\.ช)/.test(n)) male++;
    else if (/^(เด็กหญิง|นางสาว|นาง|ด\.ญ)/.test(n)) female++;
  });

  var homeroom = readAll('Teachers').filter(function (t) {
    return t.homeroomClass && String(t.homeroomClass) === String(classLevel);
  }).map(function (t) { return t.name; });

  return {
    settings: settings, classLevel: classLevel, subjects: subjRows,
    characteristic: charCount, readthink: rtCount, activity: { pass: actPass, fail: actFail },
    total: students.length, male: male, female: female, homeroomTeachers: homeroom
  };
}


/** ====================== STATS ====================== */

/** ====================== แผนภูมิสรุปผลรายชั้น-รายวิชา ====================== */
function getGradeStats(p) {
  p = p || {};
  var GR = ['4', '3.5', '3', '2.5', '2', '1.5', '1', '0'];
  var students = getStudents();
  if (p.classLevel) {
    if (String(p.classLevel).indexOf('/') >= 0)
      students = students.filter(function (s) { return String(s.classLevel) === String(p.classLevel); });
    else
      students = students.filter(function (s) { return classMatch(p.classLevel, s.classLevel); });
  }
  var inScope = {}, roomOf = {};
  students.forEach(function (s) { inScope[s.ID] = true; roomOf[s.ID] = String(s.classLevel); });

  var subjects = getSubjects();
  var subMap = {}; subjects.forEach(function (s) { subMap[s.ID] = s; });
  var sem = p.semester ? String(p.semester) : '';
  var grades = readAll('Grades').filter(function (g) {
    return inScope[g.studentID] && (!sem || String(g.semester || '1') === sem);
  });

  // ต่อรายวิชา: นับเกรด + เฉลี่ย
  var bySub = {};
  // ต่อห้อง: เฉลี่ยรวม
  var byRoom = {};
  var overall = {}; GR.forEach(function (k) { overall[k] = 0; });
  grades.forEach(function (g) {
    var k = String(g.grade51 == null ? '' : g.grade51);
    if (GR.indexOf(k) < 0) return;
    var num = Number(k);
    var sub = subMap[g.subjectID]; if (!sub) return;
    var b = bySub[g.subjectID] = bySub[g.subjectID] || { code: sub.code, name: sub.name, classLevel: sub.classLevel, counts: {}, sum: 0, n: 0 };
    b.counts[k] = (b.counts[k] || 0) + 1; b.sum += num; b.n++;
    overall[k]++;
    var rm = roomOf[g.studentID];
    var r = byRoom[rm] = byRoom[rm] || { sum: 0, n: 0 };
    r.sum += num; r.n++;
  });

  var subjectsOut = Object.keys(bySub).map(function (id) {
    var b = bySub[id];
    return { code: b.code, name: b.name, classLevel: b.classLevel,
      counts: GR.map(function (k) { return b.counts[k] || 0; }),
      avg: b.n ? Number((b.sum / b.n).toFixed(2)) : 0, n: b.n };
  }).sort(function (a, b) { return String(a.classLevel).localeCompare(String(b.classLevel)) || String(a.code).localeCompare(String(b.code)); });

  var roomsOut = Object.keys(byRoom).sort().map(function (rm) {
    var r = byRoom[rm];
    return { room: rm, avg: r.n ? Number((r.sum / r.n).toFixed(2)) : 0, n: r.n };
  });

  return { gradeLevels: GR, subjects: subjectsOut, rooms: roomsOut, overall: GR.map(function (k) { return overall[k]; }) };
}

function getStats() {
  var students = getStudents();
  var grades = getGrades();
  var byLevel = {};
  grades.forEach(function (g) {
    var k = (g.grade51 === '' || g.grade51 == null) ? 'ยังไม่ประเมิน' : ('เกรด ' + g.grade51);
    byLevel[k] = (byLevel[k] || 0) + 1;
  });
  return {
    studentCount: students.length,
    subjectCount: getSubjects().length,
    gradeCount: grades.length,
    byLevel: byLevel
  };
}


/** ====================== EXPORT SCHOOL MIS ====================== */

// คืนข้อมูลเกรดสำหรับนำไปลง School MIS (frontend แปลงเป็น Excel/CSV)
// classLevel ว่าง = ทั้งหมด
// ส่งออกแบบ wide ตรงรูปแบบ SchoolMIS: 1 แถว/คน, คอลัมน์วิชา=เกรด 2551, กิจกรรม=ผ/มผ
function exportSchoolMISWide(classLevel) {
  var students = getStudents();
  if (classLevel) students = students.filter(function (s) { return String(s.classLevel) === String(classLevel); });
  students.sort(function (a, b) { return Number(a.number) - Number(b.number); });

  // วิชาของชั้นนี้ (เรียงตาม sortOrder)
  var subjects = getSubjects().filter(function (sub) {
    return !sub.classLevel || classMatch(sub.classLevel, classLevel);
  }).sort(function (a, b) { return Number(a.sortOrder) - Number(b.sortOrder); });

  // กิจกรรมพัฒนาผู้เรียน (SchoolMIS ใช้ ผ/มผ)
  var ACTS = ['กิจกรรมแนะแนว', 'ลูกเสือ – เนตรนารี', 'ชมรม/ชุมนุม', 'กิจกรรมเพื่อสังคมและสาธารณประโยชน์'];
  var ACT_LABEL = { 'กิจกรรมแนะแนว': 'แนะแนว', 'ลูกเสือ – เนตรนารี': 'ลูกเสือ-เนตรนารี', 'ชมรม/ชุมนุม': 'ชุมนุม', 'กิจกรรมเพื่อสังคมและสาธารณประโยชน์': 'กิจกรรมเพื่อสังคมและสาธารณประโยชน์' };

  // เกรดต่อคนต่อวิชา
  var gradeBy = {};
  getGrades().forEach(function (g) { (gradeBy[g.studentID] = gradeBy[g.studentID] || {})[g.subjectID] = g; });

  // กิจกรรมต่อคน (จาก AbilityDetail group=activity)
  var actBy = {};
  readAll('AbilityDetail').forEach(function (r) {
    if (r.group === 'activity') (actBy[r.studentID] = actBy[r.studentID] || {})[r.item] = r.value;
  });

  // หัวตาราง
  var header = ['#', 'รหัสนักเรียน', 'ชื่อ-สกุล'];
  subjects.forEach(function (sub) { header.push(((sub.code || '') + ' ' + sub.name).trim()); });
  ACTS.forEach(function (a) { header.push(ACT_LABEL[a]); });

  var rows = students.map(function (st) {
    var row = [st.number, st.studentCode, st.fullName];
    subjects.forEach(function (sub) {
      var g = (gradeBy[st.ID] || {})[sub.ID];
      row.push(g && g.grade51 != null && g.grade51 !== '' ? g.grade51 : '');
    });
    ACTS.forEach(function (a) {
      var v = (actBy[st.ID] || {})[a];
      row.push(v === 'ผ่าน' ? 'ผ' : (v === 'ไม่ผ่าน' ? 'มผ' : ''));
    });
    return row;
  });

  return { header: header, rows: rows };
}

function exportSchoolMIS(classLevel) {
  var students = getStudents();
  if (classLevel) {
    students = students.filter(function (s) { return classMatch(classLevel, s.classLevel); });
  }
  var subjects = getSubjects();
  var subMap = {};
  subjects.forEach(function (s) { subMap[s.ID] = s; });
  var grades = getGrades();

  var out = [];
  grades.forEach(function (g) {
    var st = students.filter(function (s) { return String(s.ID) === String(g.studentID); })[0];
    if (!st) return;
    var sub = subMap[g.subjectID];
    if (!sub) return;
    out.push({
      studentCode: st.studentCode,
      number: st.number,
      fullName: st.fullName,
      classLevel: st.classLevel,
      subjectCode: sub.code,
      subjectName: sub.name,
      percent: g.percent,
      grade51: g.grade51
    });
  });
  return out;
}


/** ====================== BACKUP / IMPORT / CLEAR ====================== */

function backupAllData() {
  var dump = {};
  Object.keys(SHEETS).forEach(function (name) { bustSheetCache_(name); dump[name] = readAll(name); });
  dump._meta = { exportedAt: new Date().toString(), version: 1 };
  return dump;
}

/** ====================== เผยแพร่ข้อมูลนิ่งขึ้น GitHub (มาตรฐาน Sheets + Edge) ======================
 * หลักการ: ข้อมูลกลางที่เปลี่ยนไม่บ่อย (ตั้งค่า/รายวิชา/ตารางสอน/เกณฑ์เกรด/ปฏิทิน/รายชื่อครูไม่มีรหัสผ่าน)
 * เผยแพร่เป็นไฟล์ JSON นิ่งบน GitHub Pages ให้หน้าเว็บโหลดก่อน (Static-First)
 * ⚠️ ข้อมูลนักเรียน/คะแนน/ประวัติ ไม่เผยแพร่เด็ดขาด — อ่านผ่าน GAS เท่านั้น
 * ==================================================================================== */
var EDGE_PUBLIC_SHEETS = { Settings: 1, GradeMapping: 1, Subjects: 1, Schedule: 1, Teachers: 1,
  Calendar: 1, SubmissionTypes: 1, FormTemplates: 1, TeacherAssign: 1 };
var _edgeDirtyMarked = false;
function markEdgeDirty_(name) {
  if (!EDGE_PUBLIC_SHEETS[name]) return;
  if (_edgeDirtyMarked) return;
  _edgeDirtyMarked = true;
  try { PropertiesService.getScriptProperties().setProperty('EDGE_DIRTY', '1'); } catch (e) {}
}
function edgeConfig_() {
  var pr = PropertiesService.getScriptProperties();
  return {
    token: pr.getProperty('GH_TOKEN') || '',
    repo: pr.getProperty('GH_REPO') || 'nsrschool/record_grades_m1-3',
    branch: pr.getProperty('GH_BRANCH') || 'main',
    path: pr.getProperty('GH_PATH') || 'data/boot_public.json'
  };
}
function saveEdgeConfig(p) {
  var pr = PropertiesService.getScriptProperties();
  if (p.repo != null && String(p.repo).trim()) pr.setProperty('GH_REPO', String(p.repo).trim());
  if (p.branch != null && String(p.branch).trim()) pr.setProperty('GH_BRANCH', String(p.branch).trim());
  if (p.path != null && String(p.path).trim()) pr.setProperty('GH_PATH', String(p.path).trim());
  if (p.token) pr.setProperty('GH_TOKEN', String(p.token).trim());   // เก็บอย่างเดียว ไม่ส่งกลับ
  return getEdgeStatus();
}
// ก้อนข้อมูลสาธารณะ (ไม่มีข้อมูลนักเรียน/รหัสผ่าน)
function buildEdgeSnapshot_() {
  return {
    publishedAt: Utilities.formatDate(new Date(), 'GMT+7', "yyyy-MM-dd HH:mm:ss"),
    settings: getSettings(),
    mapping: getGradeMapping(),
    subjects: getSubjects(),
    calendar: getCalendar(),
    submissionTypes: getAllSubmissionTypes(),
    schedule: readAll('Schedule'),
    teacherAssign: readAll('TeacherAssign'),
    teachers: readAll('Teachers').map(function (t) {
      return { name: t.name, role: t.role, homeroomClass: t.homeroomClass || '' };
    }),
    formTemplates: (function () { try { return getFormTemplates(); } catch (e) { return []; } })()
  };
}
function publishEdgeSnapshot_() {
  var cfg = edgeConfig_();
  if (!cfg.token) throw new Error('ยังไม่ได้ตั้งค่า GitHub Token — ไปที่ 🛠️ แอดมิน → 🚀 เผยแพร่/สำรอง');
  var json = JSON.stringify(buildEdgeSnapshot_());
  var b64 = Utilities.base64Encode(Utilities.newBlob(json, 'application/json').getBytes());
  var base = 'https://api.github.com/repos/' + cfg.repo + '/contents/' + cfg.path;
  var headers = { Authorization: 'Bearer ' + cfg.token, Accept: 'application/vnd.github+json' };
  var sha = '';
  var g = UrlFetchApp.fetch(base + '?ref=' + cfg.branch, { headers: headers, muteHttpExceptions: true });
  if (g.getResponseCode() === 200) { try { sha = JSON.parse(g.getContentText()).sha || ''; } catch (e) {} }
  var body = { message: 'edge: publish ' + Utilities.formatDate(new Date(), 'GMT+7', 'yyyy-MM-dd HH:mm'), content: b64, branch: cfg.branch };
  if (sha) body.sha = sha;
  var res = UrlFetchApp.fetch(base, { method: 'put', contentType: 'application/json', headers: headers, payload: JSON.stringify(body), muteHttpExceptions: true });
  var code = res.getResponseCode();
  if (code < 200 || code >= 300) throw new Error('เผยแพร่ไม่สำเร็จ (HTTP ' + code + '): ' + res.getContentText().slice(0, 180));
  var pr = PropertiesService.getScriptProperties();
  pr.setProperty('EDGE_LAST_PUBLISH', Utilities.formatDate(new Date(), 'GMT+7', 'yyyy-MM-dd HH:mm:ss'));
  pr.deleteProperty('EDGE_LAST_ERROR');
  pr.setProperty('EDGE_DIRTY', '0');
  return { ok: true, bytes: json.length };
}
// เรียกจาก trigger ทุก 5 นาที — เผยแพร่เฉพาะเมื่อมีธงค้าง (ประหยัดโควตา)
function edgePublishTick() {
  var pr = PropertiesService.getScriptProperties();
  if (pr.getProperty('EDGE_DIRTY') !== '1') return;
  try { publishEdgeSnapshot_(); }
  catch (e) { pr.setProperty('EDGE_LAST_ERROR', String((e && e.message) || e).slice(0, 300)); }
}
function publishEdgeNow(p) { publishEdgeSnapshot_(); return getEdgeStatus(); }
function setEdgeAuto(p) {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'edgePublishTick') ScriptApp.deleteTrigger(t);
  });
  if (p && p.enable) ScriptApp.newTrigger('edgePublishTick').timeBased().everyMinutes(5).create();
  PropertiesService.getScriptProperties().setProperty('EDGE_AUTO_ON', (p && p.enable) ? '1' : '0');
  return getEdgeStatus();
}
function getEdgeStatus() {
  var pr = PropertiesService.getScriptProperties();
  var cfg = edgeConfig_();
  var has = ScriptApp.getProjectTriggers().some(function (t) { return t.getHandlerFunction() === 'edgePublishTick'; });
  return {
    enabled: has,
    dirty: pr.getProperty('EDGE_DIRTY') === '1',
    lastPublish: pr.getProperty('EDGE_LAST_PUBLISH') || '',
    lastError: pr.getProperty('EDGE_LAST_ERROR') || '',
    repo: cfg.repo, branch: cfg.branch, path: cfg.path, tokenSet: !!cfg.token
  };
}

// ---------- เมนู NSR: เผยแพร่ Edge ----------
function menuEdgePublishNow() {
  var ui = SpreadsheetApp.getUi();
  try {
    var r = publishEdgeSnapshot_();
    ui.alert('🚀 เผยแพร่ Edge', 'เผยแพร่ขึ้น GitHub สำเร็จ ✅\nขนาดไฟล์ ' + Math.round(r.bytes / 1024) + ' KB\nมีผลบนเว็บภายใน ~1 นาที', ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('🚀 เผยแพร่ Edge', '❌ ' + ((e && e.message) || e), ui.ButtonSet.OK);
  }
}
function menuEdgeAutoOn() {
  var st = setEdgeAuto({ enable: true });
  SpreadsheetApp.getUi().alert('🔁 เผยแพร่อัตโนมัติ', 'เปิดแล้ว ✅ ระบบจะตรวจทุก 5 นาที และเผยแพร่เฉพาะเมื่อมีการแก้ไขข้อมูลกลาง' + (st.tokenSet ? '' : '\n\n⚠️ อย่าลืมตั้งค่า GitHub Token ก่อน'), SpreadsheetApp.getUi().ButtonSet.OK);
}
function menuEdgeAutoOff() {
  setEdgeAuto({ enable: false });
  SpreadsheetApp.getUi().alert('⏸️ เผยแพร่อัตโนมัติ', 'ปิดแล้ว — เผยแพร่ได้ด้วยปุ่ม "เผยแพร่ Edge เดี๋ยวนี้"', SpreadsheetApp.getUi().ButtonSet.OK);
}
function menuEdgeSetToken() {
  var ui = SpreadsheetApp.getUi();
  var r = ui.prompt('🔑 GitHub Token',
    'วาง Fine-grained Token (สิทธิ์ Contents: Read and write เฉพาะ repo เว็บโรงเรียน)\nToken จะถูกเก็บฝั่งเซิร์ฟเวอร์เท่านั้น:', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  var tk = String(r.getResponseText() || '').trim();
  if (!tk) { ui.alert('ไม่ได้บันทึก — ช่องว่าง'); return; }
  saveEdgeConfig({ token: tk });
  ui.alert('🔑 GitHub Token', 'บันทึกแล้ว ✅ ทดสอบด้วยเมนู "เผยแพร่ Edge เดี๋ยวนี้" ได้เลย', ui.ButtonSet.OK);
}
function menuEdgeStatus() {
  var e = getEdgeStatus();
  var b = getBackupStatus();
  var lines = [
    '— เผยแพร่ Edge —',
    (e.lastError ? '🔴 ผิดพลาดล่าสุด: ' + e.lastError : (e.dirty ? '🟡 มีการแก้ไขรอเผยแพร่' : '🟢 ข้อมูลบนเว็บเป็นปัจจุบัน')),
    'อัตโนมัติ: ' + (e.enabled ? 'เปิด (ทุก 5 นาที)' : 'ปิด'),
    'เผยแพร่ล่าสุด: ' + (e.lastPublish || '—'),
    'Token: ' + (e.tokenSet ? 'ตั้งค่าแล้ว ✅' : 'ยังไม่ได้ตั้ง ⚠️'),
    'ปลายทาง: ' + e.repo + '/' + e.path,
    '',
    '— สำรองข้อมูล Drive —',
    'อัตโนมัติ: ' + (b.enabled ? ('เปิด (ทุกวัน ' + b.hour + ':00 น.)') : 'ปิด'),
    'สำรองล่าสุด: ' + (b.lastBackup || '—'),
    'ไฟล์ล่าสุด: ' + ((b.files && b.files[0] && b.files[0].name) || '—')
  ];
  SpreadsheetApp.getUi().alert('📊 สถานะเผยแพร่/สำรอง', lines.join('\n'), SpreadsheetApp.getUi().ButtonSet.OK);
}

/** ====================== สำรองข้อมูลอัตโนมัติลง Drive ====================== */
var AUTO_BACKUP_FOLDER = 'สำรองข้อมูล - ระบบรายงานผลการเรียน';
var AUTO_BACKUP_KEEP = 30; // เก็บไฟล์ย้อนหลังสูงสุด

// สร้างไฟล์สำรอง (เรียกจากปุ่มหรือ trigger) → คืนข้อมูลไฟล์ล่าสุด
function runBackupToDrive() {
  var dump = backupAllData();
  var folder = getOrCreateFolder(AUTO_BACKUP_FOLDER);
  var stamp = Utilities.formatDate(new Date(), 'GMT+7', 'yyyy-MM-dd_HHmm');
  var name = 'backup_' + stamp + '.json';
  var file = folder.createFile(Utilities.newBlob(JSON.stringify(dump), 'application/json', name));
  // ลบไฟล์เก่าเกินจำนวนที่กำหนด
  pruneOldBackups_(folder);
  PropertiesService.getScriptProperties().setProperty('LAST_BACKUP', new Date().toISOString());
  return { name: name, id: file.getId(), at: new Date().toString() };
}
function pruneOldBackups_(folder) {
  var files = [];
  var it = folder.getFilesByType('application/json');
  while (it.hasNext()) { var f = it.next(); files.push({ f: f, t: f.getDateCreated().getTime() }); }
  files.sort(function (a, b) { return b.t - a.t; });
  for (var i = AUTO_BACKUP_KEEP; i < files.length; i++) { try { files[i].f.setTrashed(true); } catch (e) {} }
}

// เปิด/ปิด trigger สำรองอัตโนมัติรายวัน
function setAutoBackup(p) {
  var enable = p && p.enable;
  var hour = (p && p.hour != null) ? Number(p.hour) : 22;
  // ลบ trigger เดิมของ runBackupToDrive ก่อนเสมอ
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'runBackupToDrive') ScriptApp.deleteTrigger(t);
  });
  if (enable) {
    ScriptApp.newTrigger('runBackupToDrive').timeBased().everyDays(1).atHour(hour).create();
  }
  PropertiesService.getScriptProperties().setProperty('AUTO_BACKUP_ON', enable ? '1' : '0');
  PropertiesService.getScriptProperties().setProperty('AUTO_BACKUP_HOUR', String(hour));
  return getBackupStatus();
}

function getBackupStatus() {
  var pr = PropertiesService.getScriptProperties();
  var on = pr.getProperty('AUTO_BACKUP_ON') === '1';
  // ตรวจ trigger จริง
  var has = ScriptApp.getProjectTriggers().some(function (t) { return t.getHandlerFunction() === 'runBackupToDrive'; });
  var folder = null, list = [];
  try {
    var it = DriveApp.getFoldersByName(AUTO_BACKUP_FOLDER);
    if (it.hasNext()) {
      folder = it.next();
      var fit = folder.getFilesByType('application/json'), arr = [];
      while (fit.hasNext()) { var f = fit.next(); arr.push({ name: f.getName(), id: f.getId(), at: f.getDateCreated().getTime() }); }
      arr.sort(function (a, b) { return b.at - a.at; });
      list = arr.slice(0, 10).map(function (x) {
        return { name: x.name, id: x.id, at: Utilities.formatDate(new Date(x.at), 'GMT+7', 'yyyy-MM-dd HH:mm') };
      });
    }
  } catch (e) {}
  return {
    enabled: on && has,
    hour: Number(pr.getProperty('AUTO_BACKUP_HOUR') || 22),
    lastBackup: pr.getProperty('LAST_BACKUP') || '',
    files: list,
    folderId: folder ? folder.getId() : ''
  };
}

// กู้คืนจากไฟล์สำรองใน Drive
function restoreBackup(p) {
  var file = DriveApp.getFileById(p.id);
  var json = file.getBlob().getDataAsString();
  // มาตรฐานข้อ 7: งานอันตรายต้อง non-destructive — สำรองสภาพปัจจุบันก่อนทับเสมอ
  var safety = runBackupToDrive();
  var r = importBackupData(json);
  r.safetyBackup = safety.name;
  return r;
}

function importBackupData(json) {
  var data = (typeof json === 'string') ? JSON.parse(json) : json;
  Object.keys(SHEETS).forEach(function (name) {
    if (!data[name]) return;
    var s = sheet(name);
    s.clearContents();
    s.appendRow(SHEETS[name]);
    var rows = data[name].map(function (obj) {
      return SHEETS[name].map(function (h) { return typeof obj[h] === 'undefined' ? '' : obj[h]; });
    });
    if (rows.length) s.getRange(2, 1, rows.length, SHEETS[name].length).setValues(rows);
  });
  SpreadsheetApp.flush();
  return { ok: true };
}

function clearAllData() {
  // ล้างทุกชีตยกเว้นหัวตาราง (ไม่แตะ Settings/GradeMapping ที่เป็นค่าตั้งค่า)
  ['Students', 'Subjects', 'LearningUnits', 'Scores', 'Grades'].forEach(function (name) {
    var s = sheet(name);
    var last = s.getLastRow();
    if (last > 1) s.getRange(2, 1, last - 1, s.getLastColumn()).clearContent();
  });
  SpreadsheetApp.flush();
  return { ok: true };
}

// ล้างเฉพาะแถวที่ตรงเงื่อนไข (clearContent ปลอดภัยกว่า deleteRows)
function clearRowsWhere(name, field, value) {
  var s = sheet(name);
  var data = s.getDataRange().getValues();
  if (data.length < 2) return;
  var col = data[0].indexOf(field);
  if (col < 0) return;
  // เก็บแถวที่ไม่ตรงเงื่อนไขไว้ แล้วเขียนทับใหม่
  var keep = [];
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][col]) !== String(value) && data[i].join('') !== '') keep.push(data[i]);
  }
  var cols = data[0].length;
  if (s.getLastRow() > 1) s.getRange(2, 1, s.getLastRow() - 1, cols).clearContent();
  if (keep.length) s.getRange(2, 1, keep.length, cols).setValues(keep);
}


/** ====================== UPLOAD IMAGE TO DRIVE ====================== */

function uploadImageToDrive(base64, filename) {
  if (!base64) throw new Error('ไม่มีข้อมูลรูปภาพ');
  // รองรับทั้งแบบมี prefix "data:image/...;base64," และไม่มี
  var matches = String(base64).match(/^data:(.+);base64,(.*)$/);
  var contentType = 'image/png';
  var b64 = base64;
  if (matches) { contentType = matches[1]; b64 = matches[2]; }

  var folder = getOrCreateFolder(DRIVE_FOLDER_NAME);
  var blob = Utilities.newBlob(Utilities.base64Decode(b64), contentType, filename || ('img_' + new Date().getTime()));
  var file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  var id = file.getId();
  // URL ที่ฝังใน <img> ได้
  return { url: 'https://drive.google.com/thumbnail?id=' + id + '&sz=w1000', fileId: id };
}

function getOrCreateFolder(name) {
  var it = DriveApp.getFoldersByName(name);
  return it.hasNext() ? it.next() : DriveApp.createFolder(name);
}

/** ====================== คำอธิบายรายวิชา (รูป/PDF) ====================== */
function getSubjectDocs(subjectID) {
  return readAll('SubjectDocs').filter(function (r) { return String(r.subjectID) === String(subjectID); });
}
// คืนไฟล์ Drive เป็น base64 เพื่อให้ frontend (pdf.js) เรนเดอร์ฝังในเล่ม (เลี่ยงปัญหา CORS)
function getDocBase64(fileId) {
  var f = DriveApp.getFileById(fileId);
  var blob = f.getBlob();
  return { base64: Utilities.base64Encode(blob.getBytes()), mimeType: blob.getContentType(), name: f.getName() };
}
function uploadSubjectDoc(p) {
  if (!p.base64) throw new Error('ไม่มีไฟล์');
  var m = String(p.base64).match(/^data:(.+);base64,(.*)$/);
  var ct = m ? m[1] : 'application/octet-stream';
  var b64 = m ? m[2] : p.base64;
  var folder = getOrCreateFolder('คำอธิบายรายวิชา - ' + (getSettings().schoolName || ''));
  var blob = Utilities.newBlob(Utilities.base64Decode(b64), ct, p.fileName || ('doc_' + new Date().getTime()));
  var isPdf = (ct.indexOf('pdf') >= 0);

  if (isPdf) {
    // แปลง PDF → Google Slides → ส่ง PNG ทีละหน้า
    var pdfPages = convertPdfToImages_(blob, folder, p.fileName || 'doc');
    pdfPages.forEach(function (imgFile) {
      sheet('SubjectDocs').appendRow([
        genId('DOC'), p.subjectID, imgFile.getName(), imgFile.getId(), 'image',
        Utilities.formatDate(new Date(), 'GMT+7', 'yyyy-MM-dd HH:mm:ss')
      ]);
    });
  } else {
    var file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    sheet('SubjectDocs').appendRow([genId('DOC'), p.subjectID, p.fileName || '', file.getId(), 'image',
      Utilities.formatDate(new Date(), 'GMT+7', 'yyyy-MM-dd HH:mm:ss')]);
  }
  SpreadsheetApp.flush();
  return getSubjectDocs(p.subjectID);
}

// แปลง PDF blob → PNG รายหน้า ผ่าน Google Slides API
// คืน array ของ DriveFile (image/png)
function convertPdfToImages_(pdfBlob, folder, baseName) {
  // บันทึก PDF ลง Drive ชั่วคราว แล้วคัดลอกแปลงเป็น Slides (ผ่าน REST — ไม่ต้องเปิด Advanced Service)
  var pdfFile = folder.createFile(pdfBlob.setName(baseName + '_src.pdf'));
  var slideFileId = '';
  try {
    slideFileId = driveCopyConvert_(pdfFile.getId(), baseName + '_tmp', 'application/vnd.google-apps.presentation');
  } catch (e) {
    try { pdfFile.setTrashed(true); } catch (e2) {}
    throw e;
  }

  // รอ conversion เสร็จ (poll สูงสุด ~18 วิ — PDF หน้าเยอะใช้เวลานานกว่า 3 วิ)
  var pres = null, slides = [];
  for (var w = 0; w < 12; w++) {
    Utilities.sleep(1500);
    try {
      pres = SlidesApp.openById(slideFileId);
      slides = pres.getSlides();
      if (slides.length > 0) break;
    } catch (ePoll) {}
  }
  if (!slides.length) {
    try { DriveApp.getFileById(slideFileId).setTrashed(true); } catch (e2) {}
    try { pdfFile.setTrashed(true); } catch (e2) {}
    throw new Error('แปลง PDF ยังไม่เสร็จ (ไฟล์ใหญ่/หน้าเยอะ) — กรุณาลองใหม่อีกครั้ง');
  }
  var imgs = [];
  slides.forEach(function (slide, i) {
    var url = 'https://docs.google.com/presentation/d/' + slideFileId +
      '/export/png?id=' + slideFileId + '&pageid=' + slide.getObjectId();
    var response = UrlFetchApp.fetch(url, {
      headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
      muteHttpExceptions: true
    });
    if (response.getResponseCode() !== 200) {
      Utilities.sleep(1200);
      response = UrlFetchApp.fetch(url, {
        headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
        muteHttpExceptions: true
      });
    }
    var imgBlob = response.getBlob().setName(baseName + '_p' + (i + 1) + '.png');
    var imgFile = folder.createFile(imgBlob);
    imgFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    imgs.push(imgFile);
  });

  // ลบไฟล์ชั่วคราว (PDF ต้นฉบับ + Slides)
  try { DriveApp.getFileById(slideFileId).setTrashed(true); } catch (e) {}
  try { pdfFile.setTrashed(true); } catch (e) {}
  return imgs;
}
function deleteSubjectDoc(id) {
  var row = findRowById('SubjectDocs', id);
  if (row < 0) throw new Error('ไม่พบไฟล์');
  var fileId = sheet('SubjectDocs').getRange(row, 4).getValue();
  try { DriveApp.getFileById(fileId).setTrashed(true); } catch (e) {}
  sheet('SubjectDocs').deleteRow(row);
  SpreadsheetApp.flush();
  return { id: id };
}


/** ====================== PART 5: ล็อกอิน / ครู ====================== */

function getTeachers() { return readAll('Teachers'); }

// ---- การจับคู่ชื่อครูแบบยืดหยุ่น (รองรับคำนำหน้า + ช่องว่างซ้อน) ----
function normName(s) { return String(s == null ? '' : s).replace(/\s+/g, ' ').trim(); }
function teacherCore(name) {
  var n = normName(name).replace(/^(นางสาว|นาง|นาย|ครู|เด็กชาย|เด็กหญิง|ด\.ช\.|ด\.ญ\.|ว่าที่ ?ร\.?ต\.?|ว่าที่)/, '').trim();
  return n.split(' ')[0];
}

function login(name, pin) {
  var key = normName(name);
  var t = getTeachers().filter(function (x) { return normName(x.name) === key; })[0];
  if (!t) throw new Error('ไม่พบชื่อครูในระบบ');
  if (String(t.pin).trim() !== String(pin).trim()) throw new Error('PIN ไม่ถูกต้อง');
  return { id: t.ID, name: t.name, role: t.role || 'teacher', homeroomClass: t.homeroomClass || '' };
}

// เปลี่ยนรหัสผ่าน (PIN) ของครูที่ล็อกอินอยู่ — ยืนยันด้วยรหัสเดิม
// p: { name, oldPin, newPin }
function changeMyPassword(p) {
  var key = normName(p.name);
  var rows = getTeachers();
  var t = null, rowIndex = -1;
  for (var i = 0; i < rows.length; i++) {
    if (normName(rows[i].name) === key) { t = rows[i]; rowIndex = rows[i]._rowIndex; break; }
  }
  if (!t) throw new Error('ไม่พบบัญชีผู้ใช้');
  if (String(t.pin).trim() !== String(p.oldPin).trim()) throw new Error('รหัสผ่านเดิมไม่ถูกต้อง');
  var newPin = String(p.newPin || '').trim();
  if (newPin.length < 4) throw new Error('รหัสผ่านใหม่ต้องมีอย่างน้อย 4 ตัว');
  // คอลัมน์ pin = ลำดับที่ 3 ของชีต Teachers (ID,name,pin,role,homeroomClass)
  sheet('Teachers').getRange(rowIndex, 3).setValue(newPin);
  SpreadsheetApp.flush();
  return { ok: true };
}

function getSchedule(teacher) {
  var rows = readAll('Schedule');
  if (teacher) {
    var core = teacherCore(teacher);
    rows = rows.filter(function (r) { return teacherCore(r.teacher) === core || normName(r.teacher) === normName(teacher); });
  }
  return rows;
}

function saveTeacher(p) {
  var s = sheet('Teachers');
  if (p.id) {
    var row = findRowById('Teachers', p.id);
    if (row < 0) throw new Error('ไม่พบครู ID: ' + p.id);
    s.getRange(row, 2, 1, 4).setValues([[p.name || '', p.pin || '', p.role || 'teacher', p.homeroomClass || '']]);
  } else {
    s.appendRow([genId('TCH'), p.name || '', p.pin || '1234', p.role || 'teacher', p.homeroomClass || '']);
  }
  SpreadsheetApp.flush();
  return getTeachers();
}

// บันทึกครูทั้งหมดในครั้งเดียว (เขียนทับทั้งชีต โดยคงรหัส ID เดิม)
function saveTeachersBatch(teachers) {
  var s = sheet('Teachers');
  s.clearContents();
  s.appendRow(SHEETS.Teachers);
  var out = (teachers || []).map(function (t) {
    return [t.ID || genId('TCH'), t.name || '', t.pin || '1234', t.role || 'teacher', t.homeroomClass || ''];
  });
  if (out.length) s.getRange(2, 1, out.length, SHEETS.Teachers.length).setValues(out);
  SpreadsheetApp.flush();
  return getTeachers();
}

function deleteTeacher(id) {
  var row = findRowById('Teachers', id);
  if (row < 0) throw new Error('ไม่พบครู ID: ' + id);
  sheet('Teachers').deleteRow(row);
  SpreadsheetApp.flush();
  return { id: id };
}


/** ====================== PART 5: ตารางสอน ====================== */

// นำเข้าตารางสอน (แทนที่ทั้งหมด) + สร้างบัญชีครูอัตโนมัติจากชื่อที่พบ
// rows: [{day, period, classLevel, subject, teacher}, ...] (frontend กรอง ม.1-3 มาแล้ว)
function importSchedule(rows) {
  var sc = sheet('Schedule');
  sc.clearContents();
  sc.appendRow(SHEETS.Schedule);
  var out = [];
  (rows || []).forEach(function (r) {
    out.push([genId('SCH'), r.day, r.period, r.classLevel, r.subject, r.teacher]);
  });
  if (out.length) sc.getRange(2, 1, out.length, SHEETS.Schedule.length).setValues(out);

  // สร้างบัญชีครูอัตโนมัติ (role=teacher, PIN=1234) สำหรับชื่อที่ยังไม่มี (จับคู่แบบยืดหยุ่น)
  var existing = {};
  getTeachers().forEach(function (t) { existing[teacherCore(t.name)] = true; });
  var tcSheet = sheet('Teachers');
  var added = 0;
  (rows || []).forEach(function (r) {
    var nm = (r.teacher || '').trim();
    var core = teacherCore(nm);
    if (nm && core && !existing[core]) { tcSheet.appendRow([genId('TCH'), nm, '1234', 'teacher', '']); existing[core] = true; added++; }
  });
  SpreadsheetApp.flush();
  var subjectsAdded = createSubjectsFromScheduleRows(rows);
  return { count: out.length, teachersAdded: added, subjectsAdded: subjectsAdded };
}

// กิจกรรมที่ไม่ใช่รายวิชาที่ตัดเกรด (ไม่สร้างเป็นรายวิชา)
var NON_SUBJECT_RE = /(ชุมนุม|ชมรม|แนะแนว|ลูกเสือ|เนตรนารี|ยุวกาชาด|กิจกรรมพัฒนา)/;

// สร้างรายวิชาในระบบเกรดจากตารางสอน (1 วิชาต่อระดับชั้น เช่น "คณิตศาสตร์ ป.5") เฉพาะที่ยังไม่มี
function createSubjectsFromScheduleRows(rows) {
  var subjSheet = sheet('Subjects');
  var existing = {};
  getSubjects().forEach(function (su) { existing[su.name + '|' + su.classLevel] = true; });
  var added = 0, seen = {};
  (rows || []).forEach(function (r) {
    if (!r.subject || NON_SUBJECT_RE.test(r.subject)) return;
    var digit = String(r.classLevel).replace(/[^0-9]/g, '').charAt(0);
    if (!digit) return;
    var grade = 'ม.' + digit;
    var key = r.subject + '|' + grade;
    if (existing[key] || seen[key]) return;
    seen[key] = true;
    subjSheet.appendRow([genId('SUB'), '', r.subject, '', '', 'พื้นฐาน', grade, subjSheet.getLastRow()]);
    added++;
  });
  SpreadsheetApp.flush();
  return added;
}

// สร้างรายวิชาจากตารางสอนที่นำเข้าไว้แล้ว (ไม่ต้องนำเข้าไฟล์ใหม่)
function syncSubjectsFromSchedule() {
  var rows = readAll('Schedule').filter(function (s) { return /^ม\.[1-3]/.test(String(s.classLevel)); });
  var added = createSubjectsFromScheduleRows(rows);
  var dd = dedupeSubjects();        // กันกรณีมีวิชาซ้ำค้างอยู่
  return { added: added, removed: dd.removed };
}

// ล้างรายวิชาซ้ำ (ชื่อ+ชั้นเดียวกัน) — เก็บแถวแรกที่มีรหัสวิชา แล้วย้าย grade/คะแนน/หน่วย/ไฟล์ ไปอ้างวิชาที่เก็บไว้
function dedupeSubjects() {
  var s = sheet('Subjects');
  var data = s.getDataRange().getValues();
  if (data.length < 2) return { removed: 0 };
  var head = data[0]; var col = {}; head.forEach(function (h, i) { col[h] = i; });
  var groups = {};
  for (var i = 1; i < data.length; i++) {
    var name = String(data[i][col.name]).trim();
    var cls = String(data[i][col.classLevel]).trim();
    if (!name) continue;
    var key = name + '|' + cls;
    (groups[key] = groups[key] || []).push({ row: i + 1, id: data[i][col.ID], code: String(data[i][col.code] || '') });
  }
  var idMap = {}, removeRows = [];
  Object.keys(groups).forEach(function (k) {
    var arr = groups[k]; if (arr.length < 2) return;
    var keeper = arr.filter(function (x) { return x.code; })[0] || arr[0];
    arr.forEach(function (x) { if (x.id !== keeper.id) { idMap[x.id] = keeper.id; removeRows.push(x.row); } });
  });
  if (!removeRows.length) return { removed: 0 };
  // ย้ายการอ้างอิง subjectID ของแถวที่จะลบ → วิชาที่เก็บไว้
  ['Scores', 'Grades', 'LearningUnits', 'SubjectDocs'].forEach(function (sn) {
    var sh = ss().getSheetByName(sn); if (!sh) return;
    var d = sh.getDataRange().getValues(); if (d.length < 2) return;
    var c = d[0].indexOf('subjectID'); if (c < 0) return;
    var changed = false;
    for (var i = 1; i < d.length; i++) { var old = d[i][c]; if (idMap[old]) { d[i][c] = idMap[old]; changed = true; } }
    if (changed) sh.getRange(1, 1, d.length, d[0].length).setValues(d);
  });
  // ลบแถวซ้ำจากล่างขึ้นบน
  removeRows.sort(function (a, b) { return b - a; }).forEach(function (r) { s.deleteRow(r); });
  SpreadsheetApp.flush();
  return { removed: removeRows.length };
}


/** ====================== PART 5: นำเข้ารายชื่อ (A-F) + เลื่อนเลขที่ ====================== */

// แทนที่รายชื่อทั้งหมดด้วยไฟล์นำเข้า แล้วเรียงเลขที่แต่ละห้องใหม่
// rows: [{classLevel, number, studentCode, prefix, firstName, lastName}, ...]
function bulkImportRoster(rows) {
  if (!rows || !rows.length) return { count: 0, added: 0, updated: 0 };
  ensureStudentCols_();
  var s = sheet('Students');
  var now = Utilities.formatDate(new Date(), 'GMT+7', 'yyyy-MM-dd HH:mm:ss');
  var W = SHEETS.Students.length; // 9 คอลัมน์
  var nrm = function (x) { return String(x || '').replace(/\s+/g, '').trim(); };

  // อ่านข้อมูลเดิม (รวมหัวตาราง) แล้วทำดัชนีตามชื่อ
  var data = s.getDataRange().getValues();
  var body = [];
  for (var i = 1; i < data.length; i++) {
    var r = data[i].slice(0, W);
    while (r.length < W) r.push('');
    if (String(r[0] || '') === '' && String(r[3] || '') === '') continue; // ข้ามแถวว่าง
    body.push(r);
  }
  // คอลัมน์: 0=ID 1=studentCode 2=number 3=fullName 4=classLevel 5=dateAdded 6=prefix 7=firstName 8=lastName
  var byName = {}, byNameClass = {};
  body.forEach(function (r, idx) {
    var nk = nrm(r[7]) + '|' + nrm(r[8]);
    (byName[nk] = byName[nk] || []).push(idx);
    byNameClass[r[4] + '|' + nk] = idx;
  });

  var added = 0, updated = 0;
  rows.forEach(function (r) {
    var prefix = r.prefix || '', first = r.firstName || '', last = r.lastName || '';
    var cls = r.classLevel || '';
    var nk = nrm(first) + '|' + nrm(last);
    var idx = -1;
    if (byNameClass[cls + '|' + nk] != null) idx = byNameClass[cls + '|' + nk];      // ชื่อ+ชั้นตรง
    else if ((byName[nk] || []).length >= 1) idx = byName[nk][0];                     // ชื่อตรง (คงไว้)
    if (idx >= 0) {
      // ชื่อซ้ำ → คงข้อมูลเดิม อัปเดตเฉพาะเลขที่ + ชั้น (และเลขประจำตัวถ้ามีในไฟล์)
      body[idx][2] = (r.number != null && r.number !== '') ? r.number : body[idx][2];
      body[idx][4] = cls || body[idx][4];
      if (r.studentCode) body[idx][1] = r.studentCode;
      updated++;
    } else {
      // ชื่อใหม่ → เพิ่ม
      var full = buildFullName_(prefix, first, last);
      var nr = [genId('STD'), r.studentCode || '', r.number || '', full, cls, now, prefix, first, last];
      var ni = body.push(nr) - 1;
      (byName[nk] = byName[nk] || []).push(ni);
      byNameClass[cls + '|' + nk] = ni;
      added++;
    }
  });

  // เขียนกลับทั้งหมด (ไม่ renumber — ใช้เลขที่ตามไฟล์)
  s.clearContents();
  s.getRange(1, 1, 1, W).setValues([SHEETS.Students]);
  if (body.length) s.getRange(2, 1, body.length, W).setValues(body);
  SpreadsheetApp.flush();
  return { count: rows.length, added: added, updated: updated };
}


/** ====================== PART 5: เช็คชื่อ ====================== */

// period: 0 = หน้าเสาธง, 1-6 = รายคาบ
function getAttendance(date, classLevel, period) {
  return readAll('Attendance').filter(function (r) {
    return ymd(r.date) === ymd(date) &&
           String(r.classLevel) === String(classLevel) &&
           String(r.period) === String(period);
  });
}

// บันทึกการเช็คชื่อ (แทนที่ของ วันที่+ห้อง+คาบ นั้น)
// p: {date, classLevel, period, checkedBy, records:[{studentID, status}]}
// ล้างสถานะ "สาย" ปลอมที่ระบบเคยสร้างเองตอนเช็ครายคาบ (บั๊กก่อน 13 ก.ค. 69)
// เงื่อนไขระบุตัว: กลุ่มหน้าเสาธง (period 0) ของ วัน+ห้อง ใดเป็น "สาย" ทั้งชุด,
// updatedAt เหมือนกันทั้งชุด และตรงกับ updatedAt ของการเช็ครายคาบ (period ≥ 1) วัน+ห้องเดียวกัน → แปลงเป็น "มา"
function fixPhantomLate() {
  var s = sheet('Attendance');
  var data = s.getDataRange().getValues();
  var head = data[0]; var col = {}; head.forEach(function (h, i) { col[h] = i; });
  var groups = {};
  for (var i = 1; i < data.length; i++) {
    var r = data[i]; if (!r[col.studentID]) continue;
    var key = ymd(r[col.date]) + '|' + String(r[col.classLevel]);
    var g = groups[key] = groups[key] || { flag: [], perUps: {} };
    if (String(r[col.period]) === '0') g.flag.push({ row: i + 1, st: String(r[col.status]), up: String(r[col.updatedAt]) });
    else g.perUps[String(r[col.updatedAt])] = true;
  }
  var fixed = 0, rooms = [];
  var now = Utilities.formatDate(new Date(), 'GMT+7', 'yyyy-MM-dd HH:mm:ss');
  Object.keys(groups).forEach(function (key) {
    var g = groups[key];
    if (!g.flag.length) return;
    if (!g.flag.every(function (f) { return f.st === 'สาย'; })) return;   // ต้องสายทั้งชุด
    if (g.flag.length < 3) return;                                          // กันเคสเล็กผิดปกติ
    if (!Object.keys(g.perUps).length) return;                              // วันนั้นห้องนั้นต้องมีการเช็ครายคาบ (ต้นเหตุของบั๊ก)
    g.flag.forEach(function (f) {
      s.getRange(f.row, col.status + 1).setValue('มา');
      s.getRange(f.row, col.updatedAt + 1).setValue(now);
      fixed++;
    });
    rooms.push(key.replace('|', ' '));
  });
  SpreadsheetApp.flush();
  return { fixed: fixed, groups: rooms };
}

function saveAttendance(p) {
  // ---- สิทธิ์: หน้าเสาธง = ครูประจำชั้นห้องนั้น | รายคาบ = ครูที่สอนห้องนั้น | แอดมินผ่านเสมอ ----
  if (String(p._role || '') !== 'admin') {
    var byName = p._by || p.checkedBy || '';
    if (!byName) throw new Error('ไม่ทราบผู้ใช้ — กรุณาเข้าสู่ระบบใหม่');
    var byCore = teacherCore(byName);
    if (String(p.period) === '0') {
      var hr = readAll('Teachers').filter(function (t) { return teacherCore(t.name) === byCore; })[0];
      if (!hr || String(hr.homeroomClass || '').trim() !== String(p.classLevel).trim())
        throw new Error('เช็คชื่อหน้าเสาธงของห้อง ' + p.classLevel + ' ได้เฉพาะครูประจำชั้น (หรือแอดมิน)');
    } else {
      var teachesRoom = getMyTeach({ teacher: byName }).some(function (x) { return x.room === String(p.classLevel).trim(); });
      if (!teachesRoom) throw new Error('เช็คชื่อรายคาบของห้อง ' + p.classLevel + ' ได้เฉพาะครูที่สอนห้องนี้ (หรือแอดมิน)');
    }
  }
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var s = sheet('Attendance');
    var HEAD = SHEETS.Attendance; // ['ID','date','classLevel','period','studentID','status','checkedBy','updatedAt','semester']
    // อ่านตามชื่อหัวคอลัมน์ (ทนต่อชีตเดิมที่คอลัมน์ไม่ตรงตำแหน่ง/มีคอลัมน์เกิน)
    var rows = readAll('Attendance');
    var keep = [];
    rows.forEach(function (r) {
      var same = ymd(r.date) === ymd(p.date) &&
                 String(r.classLevel) === String(p.classLevel) &&
                 String(r.period) === String(p.period);
      if (!same) keep.push([r.ID || genId('ATT'), ymd(r.date), r.classLevel, r.period, r.studentID,
                            r.status || '', r.checkedBy || '', r.updatedAt || '', (r.semester == null ? '' : r.semester)]);
    });
    var now = Utilities.formatDate(new Date(), 'GMT+7', 'yyyy-MM-dd HH:mm:ss');
    var sem = String(p.semester || getSettings().currentSemester || '1');
    (p.records || []).forEach(function (r) {
      keep.push([genId('ATT'), ymd(p.date), p.classLevel, p.period, r.studentID, r.status || 'มา', p.checkedBy || '', now, sem]);
    });

    // กฎ (ปรับ 13 ก.ค. 69): เช็ครายคาบแล้วนักเรียน "มา" จะปรับหน้าเสาธงเป็น "สาย"
    // เฉพาะกรณีที่ "มีบันทึกหน้าเสาธงอยู่แล้วและเป็น ขาด" เท่านั้น (มาสายหลังเข้าแถว)
    // ห้ามสร้างบันทึกหน้าเสาธงใหม่เด็ดขาด — เคยทำให้เกิด "สายทั้งห้อง" เมื่อเช็ครายคาบก่อนหน้าเสาธง
    if (String(p.period) !== '0') {
      var dymd = ymd(p.date), cls = String(p.classLevel);
      var flagIdx = {};
      for (var k = 0; k < keep.length; k++) {
        if (String(keep[k][3]) === '0' && ymd(keep[k][1]) === dymd && String(keep[k][2]) === cls) {
          flagIdx[String(keep[k][4])] = k;
        }
      }
      (p.records || []).forEach(function (r) {
        if (r.status !== 'มา') return;
        var sid = String(r.studentID);
        if (flagIdx[sid] == null) return;               // ไม่มีบันทึกหน้าเสาธง → ไม่ทำอะไร
        var row = keep[flagIdx[sid]];
        if (row[5] === 'ขาด') { row[5] = 'สาย'; row[7] = now; }
      });
    }

    // เขียนกลับ: ล้างเต็มความกว้างชีต (กันคอลัมน์เกินจากระบบเดิมค้าง) แล้วเขียนตาม schema มาตรฐาน
    var lastRow = s.getLastRow(), lastCol = Math.max(s.getLastColumn(), HEAD.length);
    if (lastRow > 1) s.getRange(2, 1, lastRow - 1, lastCol).clearContent();
    if (keep.length) {
      s.getRange(2, 2, keep.length, 1).setNumberFormat('@'); // คอลัมน์ date เก็บเป็นข้อความ
      s.getRange(2, 1, keep.length, HEAD.length).setValues(keep);
    }
    SpreadsheetApp.flush();
    return { count: (p.records || []).length };
  } finally { lock.releaseLock(); }
}


/** ====================== PART 8: รายงานการมาเรียน ====================== */

var WD_TH = ['อาทิตย์', 'จันทร์', 'อังคาร', 'พุธ', 'พฤหัสบดี', 'ศุกร์', 'เสาร์'];

// สรุปการมาเรียนรายคน (ช่วงวันที่) — รายวันจากหน้าเสาธง + เวลาเรียนรายคาบ (มีกฎ fallback ใช้ผลหน้าเสาธง)
function getAttendanceSummary(classLevel, fromDate, toDate) {
  var students = getStudents().filter(function (s) { return String(s.classLevel) === String(classLevel); });
  students.sort(function (a, b) { return Number(a.number) - Number(b.number); });

  var att = readAll('Attendance').filter(function (r) {
    return String(r.classLevel) === String(classLevel) &&
           ymd(r.date) >= String(fromDate) && ymd(r.date) <= String(toDate);
  });

  // คาบที่แต่ละชั้นมีในแต่ละวัน (จากตารางสอน)
  var periodsByDay = {};
  readAll('Schedule').filter(function (s) { return String(s.classLevel) === String(classLevel); })
    .forEach(function (s) {
      var d = s.day; if (!periodsByDay[d]) periodsByDay[d] = {};
      periodsByDay[d][Number(s.period)] = true;
    });

  // วันที่ที่มีการเช็คหน้าเสาธง (ถือเป็นวันเรียนที่มีข้อมูล)
  var datesWithFlag = {};
  var look = {}; // look[date|period|studentID] = status
  att.forEach(function (r) {
    var d = ymd(r.date);
    look[d + '|' + r.period + '|' + r.studentID] = r.status;
    if (String(r.period) === '0') datesWithFlag[d] = true;
  });
  var dates = Object.keys(datesWithFlag).sort();

  var rows = students.map(function (st) {
    var ma = 0, khad = 0, la = 0, sai = 0, expP = 0, presP = 0;
    dates.forEach(function (date) {
      var flag = look[date + '|0|' + st.ID];
      if (flag === 'มา') ma++; else if (flag === 'ขาด') khad++; else if (flag === 'ลา') la++; else if (flag === 'สาย') sai++;
      var wd = WD_TH[new Date(date + 'T00:00:00').getDay()];
      var ps = periodsByDay[wd] ? Object.keys(periodsByDay[wd]) : [];
      ps.forEach(function (p) {
        expP++;
        var s = look[date + '|' + p + '|' + st.ID] || flag; // กฎ fallback: ไม่มีผลรายคาบ → ใช้หน้าเสาธง
        if (s === 'มา' || s === 'สาย') presP++;
      });
    });
    return {
      studentID: st.ID, number: st.number, fullName: st.fullName, classLevel: st.classLevel,
      ma: ma, khad: khad, la: la, sai: sai, schoolDays: dates.length,
      expectedPeriods: expP, presentPeriods: presP,
      percent: expP > 0 ? Math.round(presP / expP * 100) : 0
    };
  });
  return { classLevel: classLevel, from: fromDate, to: toDate, schoolDays: dates.length, students: rows };
}

// ภาพรวมการมาเรียนของทั้งโรงเรียนในวันหนึ่ง (จากหน้าเสาธง)
function getAttendanceDashboard(date) {
  var students = getStudents();
  var classCount = {};
  students.forEach(function (s) { if (s.classLevel) classCount[s.classLevel] = (classCount[s.classLevel] || 0) + 1; });

  var att = readAll('Attendance').filter(function (r) {
    return ymd(r.date) === ymd(date) && String(r.period) === '0';
  });
  var byClass = {};
  att.forEach(function (r) {
    var c = byClass[r.classLevel] = byClass[r.classLevel] || { checked: 0, present: 0, absent: 0, leave: 0, late: 0 };
    c.checked++;
    if (r.status === 'มา') c.present++; else if (r.status === 'ขาด') c.absent++;
    else if (r.status === 'ลา') c.leave++; else if (r.status === 'สาย') c.late++;
  });

  var rows = Object.keys(classCount).sort().map(function (c) {
    var s = byClass[c] || { checked: 0, present: 0, absent: 0, leave: 0, late: 0 };
    return { classLevel: c, total: classCount[c], taken: s.checked > 0,
      present: s.present, absent: s.absent, leave: s.leave, late: s.late };
  });
  var tot = { total: 0, present: 0, absent: 0, leave: 0, late: 0, classes: rows.length, classesTaken: 0 };
  rows.forEach(function (r) {
    tot.total += r.total; tot.present += r.present; tot.absent += r.absent; tot.leave += r.leave; tot.late += r.late;
    if (r.taken) tot.classesTaken++;
  });
  return { date: date, classes: rows, totals: tot };
}

/** ====================== WP2: บันทึกเวลาเรียน ====================== */
var TH_MONTH_ABBR = ['', 'ม.ค.', 'ก.พ.', 'มี.ค.', 'เม.ย.', 'พ.ค.', 'มิ.ย.', 'ก.ค.', 'ส.ค.', 'ก.ย.', 'ต.ค.', 'พ.ย.', 'ธ.ค.'];
function thaiMonthLabel(ym) {
  var p = String(ym).split('-'); var m = Number(p[1]); var by = (Number(p[0]) + 543) % 100;
  return (TH_MONTH_ABBR[m] || ym) + ' ' + (by < 10 ? '0' + by : by);
}

// นับเป็น "มาเรียน" = มา + สาย
function _attendedDay(s) { return s === 'มา' || s === 'สาย'; }

// สรุปเวลาเรียนรายเดือน → ทั้งปี (จากหน้าเสาธง period 0)
function getAttendanceRegister(classLevel, fromDate, toDate) {
  var students = getStudents().filter(function (s) { return String(s.classLevel) === String(classLevel); });
  students.sort(function (a, b) { return Number(a.number) - Number(b.number); });
  var att = readAll('Attendance').filter(function (r) {
    return String(r.classLevel) === String(classLevel) && String(r.period) === '0' &&
           ymd(r.date) >= String(fromDate) && ymd(r.date) <= String(toDate);
  });
  var schoolDaysByYM = {}, look = {};
  att.forEach(function (r) {
    var d = ymd(r.date); var ym = d.slice(0, 7);
    (schoolDaysByYM[ym] = schoolDaysByYM[ym] || {})[d] = true;
    look[d + '|' + r.studentID] = r.status;
  });
  var yms = Object.keys(schoolDaysByYM).sort();
  var months = yms.map(function (ym) { return { ym: ym, label: thaiMonthLabel(ym), schoolDays: Object.keys(schoolDaysByYM[ym]).length }; });
  var totalSchool = months.reduce(function (a, m) { return a + m.schoolDays; }, 0);

  var rows = students.map(function (st) {
    var byMonth = {}, totalPresent = 0, ab = 0, la = 0, sa = 0;
    yms.forEach(function (ym) {
      var dates = Object.keys(schoolDaysByYM[ym]); var p = 0;
      dates.forEach(function (d) {
        var s = look[d + '|' + st.ID];
        if (_attendedDay(s)) p++;
        if (s === 'ขาด') ab++; else if (s === 'ลา') la++; else if (s === 'สาย') sa++;
      });
      byMonth[ym] = { present: p, total: dates.length };
      totalPresent += p;
    });
    return {
      studentID: st.ID, number: st.number, fullName: st.fullName, byMonth: byMonth,
      totalPresent: totalPresent, absent: ab, leave: la, late: sa,
      percent: totalSchool > 0 ? Math.round(totalPresent / totalSchool * 100) : 0
    };
  });
  return { classLevel: classLevel, months: months, totalSchoolDays: totalSchool, students: rows };
}

// กริดเช็คชื่อรายวันของ 1 เดือน (ym = 'YYYY-MM')
function getMonthlyGrid(classLevel, ym) {
  var students = getStudents().filter(function (s) { return String(s.classLevel) === String(classLevel); });
  students.sort(function (a, b) { return Number(a.number) - Number(b.number); });
  var att = readAll('Attendance').filter(function (r) {
    return String(r.classLevel) === String(classLevel) && String(r.period) === '0' && ymd(r.date).slice(0, 7) === String(ym);
  });
  var datesSet = {}, look = {};
  att.forEach(function (r) { var d = ymd(r.date); datesSet[d] = true; look[d + '|' + r.studentID] = r.status; });
  var dates = Object.keys(datesSet).sort();
  var rows = students.map(function (st) {
    var byDate = {}, p = 0, ab = 0, la = 0, sa = 0;
    dates.forEach(function (d) {
      var s = look[d + '|' + st.ID] || '';
      byDate[d] = s;
      if (_attendedDay(s)) p++;
      if (s === 'ขาด') ab++; else if (s === 'ลา') la++; else if (s === 'สาย') sa++;
    });
    return { studentID: st.ID, number: st.number, fullName: st.fullName, byDate: byDate, present: p, absent: ab, leave: la, late: sa };
  });
  return { classLevel: classLevel, ym: ym, label: thaiMonthLabel(ym), dates: dates, students: rows };
}



function INIT() {
  setupSheets();
  Logger.log('สร้างชีตเรียบร้อย: ' + Object.keys(SHEETS).join(', '));
}
