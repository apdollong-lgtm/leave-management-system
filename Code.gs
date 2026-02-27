const SHEET_NAME = 'LeaveRequests';

/**
 * แสดงหน้าเว็บแอป
 */
function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('ระบบลางานออนไลน์')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * include HTML partial
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * สร้างชีตเริ่มต้นถ้ายังไม่มี
 */
function setupSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow([
      'รหัสคำขอ',
      'ชื่อพนักงาน',
      'แผนก',
      'ประเภทการลา',
      'วันที่เริ่มลา',
      'วันที่สิ้นสุด',
      'จำนวนวัน',
      'เหตุผล',
      'สถานะ',
      'หมายเหตุผู้อนุมัติ',
      'วันที่สร้าง',
      'วันที่อัปเดต'
    ]);
  }

  return sheet;
}

/**
 * โหลดข้อมูลเริ่มต้นสำหรับแดชบอร์ด
 */
function getDashboardData() {
  const sheet = setupSheet_();
  const values = sheet.getDataRange().getValues();

  if (values.length <= 1) {
    return {
      stats: {
        pending: 0,
        approvedToday: 0,
        onLeave: 0
      },
      requests: []
    };
  }

  const headers = values[0];
  const rows = values.slice(1);
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');

  const requests = rows
    .filter((row) => row[0])
    .map((row) => {
      const item = {};
      headers.forEach((header, idx) => {
        item[header] = row[idx];
      });
      return item;
    })
    .sort((a, b) => new Date(b['วันที่สร้าง']) - new Date(a['วันที่สร้าง']));

  const pending = requests.filter((r) => r['สถานะ'] === 'รออนุมัติ').length;
  const approvedToday = requests.filter((r) => {
    if (r['สถานะ'] !== 'อนุมัติ' || !r['วันที่อัปเดต']) return false;
    const updated = Utilities.formatDate(new Date(r['วันที่อัปเดต']), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    return updated === today;
  }).length;

  const onLeave = requests.filter((r) => {
    if (r['สถานะ'] !== 'อนุมัติ') return false;
    const start = new Date(r['วันที่เริ่มลา']);
    const end = new Date(r['วันที่สิ้นสุด']);
    const now = new Date();
    start.setHours(0, 0, 0, 0);
    end.setHours(23, 59, 59, 999);
    return now >= start && now <= end;
  }).length;

  return {
    stats: {
      pending,
      approvedToday,
      onLeave
    },
    requests
  };
}

/**
 * บันทึกคำขอลาใหม่ลงชีต
 */
function submitLeaveRequest(form) {
  const sheet = setupSheet_();

  if (!form.name || !form.leaveType || !form.startDate || !form.endDate || !form.reason) {
    throw new Error('กรุณากรอกข้อมูลให้ครบถ้วน');
  }

  const now = new Date();
  const id = `LR-${Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyyMMddHHmmss')}`;
  const start = new Date(form.startDate);
  const end = new Date(form.endDate);

  if (end < start) {
    throw new Error('วันที่สิ้นสุดต้องไม่น้อยกว่าวันที่เริ่มลา');
  }

  const diffMs = end.getTime() - start.getTime();
  const duration = Math.floor(diffMs / (1000 * 60 * 60 * 24)) + 1;

  sheet.appendRow([
    id,
    form.name,
    form.department || '-',
    form.leaveType,
    form.startDate,
    form.endDate,
    duration,
    form.reason,
    'รออนุมัติ',
    '',
    now,
    now
  ]);

  return { success: true, message: 'ส่งคำขอลาสำเร็จ', requestId: id };
}

/**
 * อนุมัติ/ปฏิเสธคำขอลา
 */
function updateLeaveStatus(requestId, status, managerNote) {
  const sheet = setupSheet_();
  const values = sheet.getDataRange().getValues();

  for (let i = 1; i < values.length; i++) {
    if (values[i][0] === requestId) {
      sheet.getRange(i + 1, 9).setValue(status);
      sheet.getRange(i + 1, 10).setValue(managerNote || '');
      sheet.getRange(i + 1, 12).setValue(new Date());
      return { success: true, message: `อัปเดตสถานะเป็น ${status} แล้ว` };
    }
  }

  throw new Error('ไม่พบรหัสคำขอที่ต้องการอัปเดต');
}

/**
 * สร้างข้อมูลตัวอย่าง
 */
function seedSampleData() {
  const sheet = setupSheet_();
  if (sheet.getLastRow() > 1) return 'มีข้อมูลอยู่แล้ว';

  const sample = [
    ['LR-20261001090001', 'ศิริพร ใจดี', 'Design', 'ลาป่วย', '2026-02-26', '2026-02-27', 2, 'มีไข้สูงต้องพบแพทย์', 'รออนุมัติ', '', new Date(), new Date()],
    ['LR-20261001090002', 'อนุชา วิริยะ', 'Engineering', 'ลาพักร้อน', '2026-03-03', '2026-03-07', 5, 'เดินทางต่างจังหวัดกับครอบครัว', 'อนุมัติ', 'อนุมัติให้ลาได้', new Date(), new Date()],
    ['LR-20261001090003', 'นภัสสร พงษ์ไทย', 'Product', 'ลากิจ', '2026-03-10', '2026-03-10', 1, 'ติดต่อธุระทางราชการ', 'รออนุมัติ', '', new Date(), new Date()]
  ];

  sheet.getRange(2, 1, sample.length, sample[0].length).setValues(sample);
  return 'เพิ่มข้อมูลตัวอย่างเรียบร้อย';
}
