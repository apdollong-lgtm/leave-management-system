/**
 * ระบบใบลาออนไลน์ + Dashboard
 * ------------------------------------------------------------
 * ไฟล์ในโปรเจกต์ Apps Script:
 *   Code.gs        (ไฟล์นี้)
 *   Dashboard.html (หน้าเว็บทั้งหมด: เข้าสู่ระบบ / ยื่นใบลา / ใบลาของฉัน / Dashboard / อนุมัติ / ตั้งค่า)
 *
 * สิทธิ์การใช้งาน (กำหนดโดยผู้ดูแลในเมนู "ตั้งค่า"):
 *   employee  พนักงาน     เห็นเฉพาะเมนู ยื่นใบลา และ ใบลาของฉัน
 *   approver  ผู้อนุมัติ    + หน้าแรก อนุมัติใบลา ปฏิทิน ข้อมูลพนักงาน รายงาน
 *   admin     ผู้ดูแลระบบ  + ตั้งค่าผู้ใช้งานและสิทธิ์
 *
 * เป็นโปรเจกต์แบบ standalone (จัดการด้วย clasp):
 *   1) clasp push แล้ว clasp deploy (ค่า Web app อยู่ใน appsscript.json)
 *   2) เปิดโปรเจกต์ใน Apps Script แล้ว Run ฟังก์ชัน setup 1 ครั้ง เพื่ออนุญาตสิทธิ์
 *      ระบบจะสร้าง Google Sheets สำหรับเก็บข้อมูลให้อัตโนมัติ
 *      (หรือใส่ SPREADSHEET_ID ด้านล่าง เพื่อใช้ชีตที่มีอยู่แล้ว)
 *   3) เข้าเว็บด้วยชื่อผู้ใช้ "admin" + รหัสผ่านผู้ดูแล แล้วเพิ่มพนักงานในเมนู "ตั้งค่า"
 *   ตั้ง ADMIN_PASSWORD ที่ Project Settings > Script Properties แล้วรัน setup
 */

// ===== ตั้งค่า =====
const SHEET_NAME = 'LeaveRequests';
const USERS_SHEET_NAME = 'Users';
// ตั้ง ADMIN_PASSWORD ใน Script Properties แล้วรัน setup เพื่อเก็บเป็น hash
const SUPER_ADMIN_ID = 'admin'; // รหัสอ้างอิงผู้ดูแลหลัก (ชื่อเข้าสู่ระบบตั้งผ่าน ADMIN_USERNAME)

// ไอดีของ Google Sheets ที่ใช้เก็บข้อมูล (เว้นว่าง = สร้างไฟล์ใหม่ให้อัตโนมัติครั้งแรก)
const SPREADSHEET_ID = '';
const SPREADSHEET_TITLE = 'ระบบการลาออนไลน์ - ข้อมูลใบลา';

const ORG_NAME    = 'ระบบลา';
const ORG_TAGLINE = 'People · Process · Better Tomorrow';

const LEAVE_TYPES = ['ลาพักร้อน', 'ลาป่วย', 'ลากิจ', 'ลาคลอด', 'ลาอื่นๆ'];
const DEPARTMENTS = ['QC', 'QA', 'Production', 'Maintenance', 'Packing', 'คลังสินค้า', 'บุคคล', 'บัญชี'];

// สิทธิ์วันลาต่อปี (ประเภทที่ไม่อยู่ในนี้ = ไม่จำกัด / ไม่แสดงวันคงเหลือ)
const LEAVE_QUOTA = { 'ลาพักร้อน': 10, 'ลาป่วย': 30, 'ลากิจ': 5 };
const HOLIDAYS = []; // วันหยุดบริษัท รูปแบบ YYYY-MM-DD

// จำนวนพนักงานทั้งหมด (0 = นับจากผู้ใช้งานในชีต Users หรือรายชื่อที่เคยยื่นใบลา)
const TOTAL_EMPLOYEES = 0;

const ROLES = { employee: 'พนักงาน', approver: 'ผู้อนุมัติ', admin: 'ผู้ดูแลระบบ' };
const STAFF = ['approver', 'admin'];
const SESSION_SECONDS = 6 * 60 * 60; // อายุการเข้าสู่ระบบ (สูงสุดของ CacheService คือ 6 ชม.)

const STATUS = { PENDING: 'รออนุมัติ', APPROVED: 'อนุมัติ', REJECTED: 'ไม่อนุมัติ' };

const HEADERS = [
  'รหัสคำขอ', 'วันเวลาที่ส่ง', 'ชื่อ-นามสกุล', 'รหัสพนักงาน', 'แผนก',
  'ประเภทการลา', 'วันที่เริ่ม', 'วันที่สิ้นสุด', 'จำนวนวัน', 'เหตุผล',
  'สถานะ', 'วันที่ดำเนินการ', 'ผู้อนุมัติ', 'หมายเหตุ'
];
// ตำแหน่งคอลัมน์ (index เริ่มที่ 0)
const C = { ID:0, TS:1, NAME:2, EMP:3, DEPT:4, TYPE:5, START:6, END:7, DAYS:8, REASON:9, STATUS:10, ACTED:11, APPROVER:12, NOTE:13 };

const USER_HEADERS = ['รหัสพนักงาน', 'ชื่อ-นามสกุล', 'แผนก', 'สิทธิ์', 'Password hash', 'ใช้งาน', 'อัปเดตล่าสุด', 'ชื่อผู้ใช้'];
const U = { EMP:0, NAME:1, DEPT:2, ROLE:3, PASSWORD:4, ACTIVE:5, UPDATED:6, USERNAME:7 };
const PASSWORD_ITERATIONS = 600000;

// ===== Web App =====
function doGet() {
  const tpl = HtmlService.createTemplateFromFile('Dashboard');
  tpl.orgName = ORG_NAME;
  tpl.orgTagline = ORG_TAGLINE;
  return tpl.evaluate()
    .setTitle('ระบบการลาออนไลน์')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** รันครั้งแรกครั้งเดียว เพื่ออนุญาตสิทธิ์ และสร้างชีตข้อมูล */
function setup() {
  adminPasswordHash_();
  const sh = getSheet_();
  getUsersSheet_();
  Logger.log('พร้อมใช้งาน: ' + sh.getParent().getUrl());
}

/** เปิดชีตข้อมูล: SPREADSHEET_ID > Script Property "SPREADSHEET_ID" > สร้างใหม่แล้วจำไอดีไว้ */
function getSpreadsheet_() {
  const props = PropertiesService.getScriptProperties();
  const id = SPREADSHEET_ID || props.getProperty('SPREADSHEET_ID');
  if (id) return SpreadsheetApp.openById(id);

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const again = props.getProperty('SPREADSHEET_ID');
    if (again) return SpreadsheetApp.openById(again);
    const ss = SpreadsheetApp.create(SPREADSHEET_TITLE);
    ss.setSpreadsheetTimeZone('Asia/Bangkok');
    ss.getSheets()[0].setName(SHEET_NAME);
    props.setProperty('SPREADSHEET_ID', ss.getId());
    return ss;
  } finally {
    lock.releaseLock();
  }
}

function getSheet_() {
  const ss = getSpreadsheet_();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  if (sh.getLastRow() === 0) {
    sh.setFrozenRows(1);
    sh.getRange('B:B').setNumberFormat('yyyy-mm-dd hh:mm');
    sh.getRange('G:H').setNumberFormat('yyyy-mm-dd');
    sh.getRange('D:D').setNumberFormat('@');
    sh.getRange('L:L').setNumberFormat('yyyy-mm-dd hh:mm');
    sh.setColumnWidth(3, 180);
    sh.setColumnWidth(10, 260);
  }
  if (sh.getLastRow() > 0) {
    const width = Math.min(sh.getLastColumn(), HEADERS.length);
    const actual = sh.getRange(1, 1, 1, width).getValues()[0];
    if (actual.some((value, i) => String(value) !== HEADERS[i])) {
      throw new Error('โครงสร้างชีต LeaveRequests ไม่ตรงกับเวอร์ชันนี้ กรุณาสำรองข้อมูลและใช้ชีตใหม่ตามคู่มือ');
    }
  }
  // เติมหัวคอลัมน์ที่เพิ่มใหม่ (ผู้อนุมัติ / หมายเหตุ) ให้ชีตเดิมโดยอัตโนมัติ
  if (sh.getLastColumn() < HEADERS.length) {
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS])
      .setFontWeight('bold').setBackground('#1F3A68').setFontColor('#FFFFFF');
  }
  return sh;
}

function getUsersSheet_() {
  const ss = getSpreadsheet_();
  let sh = ss.getSheetByName(USERS_SHEET_NAME);
  if (!sh) sh = ss.insertSheet(USERS_SHEET_NAME);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, USER_HEADERS.length).setValues([USER_HEADERS])
      .setFontWeight('bold').setBackground('#1F3A68').setFontColor('#FFFFFF');
    sh.setFrozenRows(1);
    sh.getRange('A:A').setNumberFormat('@'); // รหัสพนักงานเป็นข้อความ (กันเลข 0 นำหน้าหาย)
    sh.getRange('H:H').setNumberFormat('@');
    sh.getRange('G:G').setNumberFormat('yyyy-mm-dd hh:mm');
    sh.setColumnWidth(2, 180);
  }
  if (sh.getLastColumn() < USER_HEADERS.length) {
    sh.getRange(1, U.USERNAME + 1).setValue(USER_HEADERS[U.USERNAME]);
    sh.getRange('H:H').setNumberFormat('@');
  }
  if (String(sh.getRange(1, U.PASSWORD + 1).getValue()) === 'PIN (เข้ารหัส)') {
    sh.getRange(1, U.PASSWORD + 1).setValue(USER_HEADERS[U.PASSWORD]);
  }
  return sh;
}

// ===== ผู้ใช้งานและการเข้าสู่ระบบ =====
function adminUsername_() {
  const username = PropertiesService.getScriptProperties().getProperty('ADMIN_USERNAME') || 'admin';
  validateUsername_(username);
  return username;
}

function adminPasswordHash_() {
  const props = PropertiesService.getScriptProperties();
  const configured = props.getProperty('ADMIN_PASSWORD');
  if (configured) {
    validatePassword_(configured);
    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      const latest = props.getProperty('ADMIN_PASSWORD');
      if (latest) {
        validatePassword_(latest);
        props.setProperty('ADMIN_PASSWORD_HASH', hashPassword_(latest));
        props.deleteProperty('ADMIN_PASSWORD');
      }
    } finally { lock.releaseLock(); }
  }
  const hash = props.getProperty('ADMIN_PASSWORD_HASH');
  if (!hash || !hash.startsWith('pbkdf2-sha256$')) throw new Error('กรุณาตั้ง ADMIN_PASSWORD ใน Script Properties แล้วรัน setup');
  return hash;
}

function passwordBytes_(text) {
  return new Uint8Array(Utilities.newBlob(String(text)).getBytes().map(value => value & 255));
}

function hashPassword_(password) {
  validatePassword_(password);
  const salt = Utilities.getUuid();
  const bytes = PasswordCrypto.derive(passwordBytes_(password), passwordBytes_(salt), PASSWORD_ITERATIONS);
  return 'pbkdf2-sha256$' + PASSWORD_ITERATIONS + '$' + salt + '$' + Utilities.base64Encode(Array.from(bytes, value => value > 127 ? value - 256 : value));
}

function verifyPassword_(password, encoded) {
  const parts = String(encoded || '').split('$');
  if (parts.length !== 4 || parts[0] !== 'pbkdf2-sha256' || Number(parts[1]) !== PASSWORD_ITERATIONS || !/^[a-f0-9-]{36}$/i.test(parts[2])) return false;
  if (String(password).length > 128) return false;
  const bytes = PasswordCrypto.derive(passwordBytes_(password), passwordBytes_(parts[2]), PASSWORD_ITERATIONS);
  const actual = Utilities.base64Encode(Array.from(bytes, value => value > 127 ? value - 256 : value));
  if (actual.length !== parts[3].length) return false;
  let difference = 0;
  for (let i = 0; i < actual.length; i++) difference |= actual.charCodeAt(i) ^ parts[3].charCodeAt(i);
  return difference === 0;
}

function validateUsername_(username) {
  if (!/^[A-Za-z0-9_.-]{3,30}$/.test(String(username || ''))) throw new Error('ชื่อผู้ใช้ต้องมี 3-30 ตัว ใช้ A-Z, 0-9, _ - .');
}

function validatePassword_(password) {
  const value = String(password || '');
  if (value.length < 12 || value.length > 128 || /^\s+$/.test(value) || /^\d+$/.test(value)) {
    throw new Error('รหัสผ่านต้องยาว 12-128 ตัวอักษร และไม่เป็นตัวเลขล้วน');
  }
}

function readUsers_() {
  const sh = getUsersSheet_();
  if (sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, USER_HEADERS.length).getValues()
    .map((r, i) => ({
      row: i + 2,
      empId: String(r[U.EMP]).trim(),
      name: String(r[U.NAME]).trim(),
      dept: String(r[U.DEPT]).trim(),
      role: Object.prototype.hasOwnProperty.call(ROLES, r[U.ROLE]) ? String(r[U.ROLE]) : 'employee',
      passwordHash: String(r[U.PASSWORD]),
      username: String(r[U.USERNAME] || r[U.EMP]).trim(),
      active: r[U.ACTIVE] === true || String(r[U.ACTIVE]).toUpperCase() === 'TRUE'
    }))
    .filter(u => u.empId);
}

function findUser_(empId) {
  const k = String(empId).trim().toLowerCase();
  return readUsers_().find(u => u.empId.toLowerCase() === k) || null;
}

const publicUser_ = u => ({
  empId: u.empId, username: u.username, name: u.name, dept: u.dept, role: u.role, roleName: ROLES[u.role],
  isSuper: u.empId === SUPER_ADMIN_ID
});

function superAdmin_() {
  return { empId: SUPER_ADMIN_ID, username: adminUsername_(), passwordHash: adminPasswordHash_(), name: 'ผู้ดูแลระบบ', dept: '', role: 'admin' };
}

function login(username, password) {
  username = String(username || '').trim();
  password = String(password || '');
  if (!username || !password || username.length > 30 || password.length > 128) throw new Error('กรุณากรอกชื่อผู้ใช้และรหัสผ่าน');

  const cache = CacheService.getScriptCache();
  const failKey = 'fail:' + username.toLowerCase();
  const fails = Number(cache.get(failKey) || 0);
  if (fails >= 5) throw new Error('ใส่รหัสผ่านผิดหลายครั้ง กรุณารอ 10 นาทีแล้วลองใหม่');

  let user = null;
  if (username.toLowerCase() === adminUsername_().toLowerCase()) {
    const admin = superAdmin_();
    if (verifyPassword_(password, admin.passwordHash)) user = admin;
  } else {
    const u = readUsers_().find(account => account.username.toLowerCase() === username.toLowerCase());
    if (u && u.active && verifyPassword_(password, u.passwordHash)) user = u;
  }
  if (!user) {
    cache.put(failKey, String(fails + 1), 600);
    throw new Error('ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง');
  }
  cache.remove(failKey);
  const token = Utilities.getUuid();
  cache.put('s:' + token, JSON.stringify({empId: user.empId, fingerprint: sessionFingerprint_(user)}), SESSION_SECONDS);
  return { token: token, user: publicUser_(user) };
}

function logout(token) {
  if (token) CacheService.getScriptCache().remove('s:' + token);
  return true;
}

/** ตรวจ token และสิทธิ์ (อ่านสิทธิ์ล่าสุดจากชีตทุกครั้ง เพื่อให้การเปลี่ยนสิทธิ์/ปิดบัญชีมีผลทันที) */
function session_(token, roles) {
  const cached = token ? CacheService.getScriptCache().get('s:' + token) : null;
  let stored;
  try { stored = JSON.parse(cached || 'null'); } catch (error) { stored = null; }
  const empId = stored && stored.empId;
  if (!empId) throw new Error('SESSION_EXPIRED: กรุณาเข้าสู่ระบบใหม่');
  let user;
  if (empId === SUPER_ADMIN_ID) user = superAdmin_();
  else {
    user = findUser_(empId);
    if (!user || !user.active) throw new Error('SESSION_EXPIRED: บัญชีนี้ถูกปิดการใช้งาน');
  }
  if (stored.fingerprint !== sessionFingerprint_(user)) throw new Error('SESSION_EXPIRED: ข้อมูลเข้าสู่ระบบเปลี่ยนแล้ว กรุณาเข้าสู่ระบบใหม่');
  if (roles && roles.indexOf(user.role) < 0) throw new Error('คุณไม่มีสิทธิ์ใช้งานส่วนนี้');
  return user;
}

function sessionFingerprint_(user) {
  return Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,
    user.empId + '|' + user.username + '|' + user.passwordHash, Utilities.Charset.UTF_8));
}

function whoami(token) {
  return publicUser_(session_(token));
}

function getAppConfig(token) {
  session_(token);
  return { leaveTypes: LEAVE_TYPES, departments: DEPARTMENTS, quota: LEAVE_QUOTA, roles: ROLES, holidays: HOLIDAYS };
}

function changeMyPassword(token, oldPassword, newPassword) {
  const user = session_(token);
  validatePassword_(newPassword);
  if (!verifyPassword_(String(oldPassword || ''), user.passwordHash)) throw new Error('รหัสผ่านเดิมไม่ถูกต้อง');
  const hash = hashPassword_(newPassword);
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const latest = session_(token);
    if (latest.passwordHash !== user.passwordHash) throw new Error('ข้อมูลเข้าสู่ระบบเปลี่ยนแล้ว กรุณาเข้าสู่ระบบใหม่');
    if (user.empId === SUPER_ADMIN_ID) PropertiesService.getScriptProperties().setProperty('ADMIN_PASSWORD_HASH', hash);
    else {
      getUsersSheet_().getRange(user.row, U.PASSWORD + 1).setValue(hash);
      getUsersSheet_().getRange(user.row, U.UPDATED + 1).setValue(new Date());
    }
  } finally { lock.releaseLock(); }
  return true;
}

// ===== จัดการผู้ใช้งาน (admin) =====
function listUsers(token) {
  session_(token, ['admin']);
  return readUsers_().map(u => Object.assign(publicUser_(u), { active: u.active, hasPassword: u.passwordHash.startsWith('pbkdf2-sha256$') }))
    .sort((a, b) => a.empId.localeCompare(b.empId));
}

function saveUser(token, data) {
  data = data || {};
  const me = session_(token, ['admin']);
  const empId = String(data.empId || '').trim();
  const name = String(data.name || '').trim().slice(0, 80);
  const dept = String(data.dept || '').trim();
  const role = String(data.role || 'employee');
  const active = data.active !== false;
  const password = String(data.password || '');
  const username = String(data.username || data.empId || '').trim();
  const ownCredentialsChanged = me.empId.toLowerCase() === empId.toLowerCase() && (!!password || username !== me.username);
  validateUsername_(username);
  if (username.toLowerCase() === adminUsername_().toLowerCase()) throw new Error('ชื่อผู้ใช้นี้สงวนไว้สำหรับผู้ดูแลหลัก');

  if (!/^[A-Za-z0-9_\-.]{1,30}$/.test(empId)) throw new Error('รหัสพนักงานใช้ได้เฉพาะ A-Z, 0-9, _ - . (ไม่เกิน 30 ตัว)');
  if (empId.toLowerCase() === SUPER_ADMIN_ID) throw new Error('รหัส "' + SUPER_ADMIN_ID + '" สงวนไว้สำหรับผู้ดูแลหลัก');
  if (!name) throw new Error('กรุณากรอกชื่อ-นามสกุล');
  if (DEPARTMENTS.indexOf(dept) < 0) throw new Error('กรุณาเลือกแผนก');
  if (!Object.prototype.hasOwnProperty.call(ROLES, role)) throw new Error('สิทธิ์ไม่ถูกต้อง');
  if (password) validatePassword_(password);
  if (me.empId.toLowerCase() === empId.toLowerCase() && (role !== 'admin' || !active)) {
    throw new Error('ลดสิทธิ์หรือปิดบัญชีของตัวเองไม่ได้');
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = getUsersSheet_();
    const existing = findUser_(empId);
    if (readUsers_().some(u => u.username.toLowerCase() === username.toLowerCase() && u.empId.toLowerCase() !== empId.toLowerCase())) throw new Error('ชื่อผู้ใช้นี้ถูกใช้แล้ว');
    if ((!existing || !existing.passwordHash.startsWith('pbkdf2-sha256$')) && !password) throw new Error('ต้องตั้งรหัสผ่านใหม่สำหรับบัญชีนี้');
    const passwordHash = password ? hashPassword_(password) : existing.passwordHash;
    const row = [existing ? existing.empId : empId, sheetText_(name), dept, role, passwordHash, active, new Date(), username];
    if (existing) sh.getRange(existing.row, 1, 1, row.length).setValues([row]);
    else sh.appendRow(row);
  } finally {
    lock.releaseLock();
  }
  // Credentials were saved successfully; the caller must sign in again with them.
  return ownCredentialsChanged ? [] : listUsers(token);
}

// ===== ใบลา =====
function submitLeave(token, form) {
  form = form || {};
  const user = session_(token);
  if (user.empId === SUPER_ADMIN_ID) throw new Error('บัญชีผู้ดูแลหลักยื่นใบลาไม่ได้ กรุณาใช้บัญชีพนักงาน');
  const type   = String(form.type || '').trim();
  const reason = String(form.reason || '').trim().slice(0, 500);
  if (LEAVE_TYPES.indexOf(type) < 0) throw new Error('กรุณาเลือกประเภทการลา');
  if (!reason) throw new Error('กรุณากรอกเหตุผลการลา');

  const start = parseDate_(form.start);
  const end   = parseDate_(form.end);
  if (!start || !end) throw new Error('กรุณาเลือกวันที่เริ่มและวันที่สิ้นสุด');
  if (end < start)    throw new Error('วันที่สิ้นสุดต้องไม่ก่อนวันที่เริ่ม');
  if (start.getFullYear() !== end.getFullYear()) throw new Error('กรุณาแยกใบลาตามปีปฏิทิน');

  let days = countWorkdays_(start, end);
  if (form.halfDay === true) {
    if (start.getTime() !== end.getTime()) throw new Error('ลาครึ่งวันต้องเลือกวันเดียว');
    if (days > 0) days = 0.5;
  }
  if (days <= 0) throw new Error('ช่วงวันที่เลือกเป็นวันหยุดเสาร์-อาทิตย์ทั้งหมด');

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = getSheet_();
    validateLeavePolicy_(user.empId, type, start, end, days);
    const id = 'LV' + Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyMMdd') + '-' +
               Utilities.getUuid();
    sh.appendRow([id, new Date(), sheetText_(user.name), user.empId, user.dept, type, start, end, days, sheetText_(reason), STATUS.PENDING, '', '', '']);
    return { ok: true, id: id, days: days };
  } finally {
    lock.releaseLock();
  }
}

/** อ่านใบลาทั้งหมดในชีต แปลงเป็น object ที่ส่งให้หน้าเว็บได้ */
function readLeaves_() {
  const tz = 'Asia/Bangkok';
  const sh = getSheet_();
  const values = sh.getLastRow() > 1
    ? sh.getRange(2, 1, sh.getLastRow() - 1, HEADERS.length).getValues()
    : [];
  const toDate = v => v instanceof Date ? v : new Date(v);
  const fmt = (d, p) => (d instanceof Date && !isNaN(d)) ? Utilities.formatDate(d, tz, p) : '';
  return values.filter(r => r[C.ID]).map(r => {
    const start = toDate(r[C.START]);
    return {
      id: String(r[C.ID]),
      submitted: fmt(toDate(r[C.TS]), 'yyyy-MM-dd HH:mm'),
      name: String(r[C.NAME]),
      empId: String(r[C.EMP]),
      dept: String(r[C.DEPT]),
      type: String(r[C.TYPE]),
      start: fmt(start, 'yyyy-MM-dd'),
      end: fmt(toDate(r[C.END]), 'yyyy-MM-dd'),
      days: Number(r[C.DAYS]) || 0,
      reason: String(r[C.REASON] || ''),
      status: String(r[C.STATUS] || STATUS.PENDING),
      approver: String(r[C.APPROVER] || ''),
      note: String(r[C.NOTE] || ''),
      _y: start.getFullYear(), _m: start.getMonth()
    };
  });
}

const yearsOf_ = rows => Array.from(new Set(rows.map(r => r._y).filter(y => !isNaN(y)))).sort((a, b) => b - a);
const strip_ = r => { const o = Object.assign({}, r); delete o._y; delete o._m; return o; };

/** ใบลาของผู้ใช้ที่เข้าสู่ระบบ (ทุกสิทธิ์) */
function getMyLeaves(token, year) {
  const user = session_(token);
  const mine = readLeaves_().filter(r => r.empId.toLowerCase() === user.empId.toLowerCase());
  const thisYear = new Date().getFullYear();
  const years = yearsOf_(mine);
  if (years.indexOf(thisYear) < 0) years.unshift(thisYear);
  const y = Number(year) || thisYear;
  return { year: y, years: years.sort((a, b) => b - a), rows: mine.filter(r => r._y === y).map(strip_) };
}

/**
 * ส่งใบลาทั้งหมดของปีที่เลือก (และ ธ.ค. ปีก่อนหน้า สำหรับเทียบเดือนก่อน) — เฉพาะผู้อนุมัติ/ผู้ดูแล
 * การกรองเดือน/แผนก และการสรุปผลทำฝั่งหน้าเว็บ เพื่อให้สลับตัวกรองได้ทันที
 */
function getDashboardData(token, year) {
  session_(token, STAFF);
  const all = readLeaves_();
  const years = yearsOf_(all);
  const thisYear = new Date().getFullYear();
  const y = Number(year) || (years.indexOf(thisYear) >= 0 ? thisYear : years[0]) || thisYear;
  const rows = all.filter(r => r._y === y || (r._y === y - 1 && r._m === 11)).map(strip_);

  const activeUsers = readUsers_().filter(u => u.active).length;
  const empSet = new Set(all.map(r => r.empId + '|' + r.name));

  return {
    year: y,
    years: years.length ? years : [y],
    rows: rows,
    config: {
      departments: DEPARTMENTS,
      leaveTypes: LEAVE_TYPES,
      quota: LEAVE_QUOTA,
      totalEmployees: TOTAL_EMPLOYEES || activeUsers || empSet.size,
      status: STATUS
    },
    sheetUrl: getSpreadsheet_().getUrl()
  };
}

function updateStatus(token, id, status, note) {
  const user = session_(token, STAFF);
  if ([STATUS.APPROVED, STATUS.REJECTED].indexOf(status) < 0) throw new Error('สถานะไม่ถูกต้อง');
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = getSheet_();
    const n = Math.max(sh.getLastRow() - 1, 1);
    const data = sh.getRange(2, 1, n, HEADERS.length).getValues();
    const i = data.findIndex(r => String(r[C.ID]) === String(id));
    if (i < 0) throw new Error('ไม่พบใบลา ' + id);
    if (String(data[i][C.STATUS]) !== STATUS.PENDING) throw new Error('ใบลานี้ถูกดำเนินการไปแล้ว');
    if (String(data[i][C.EMP]).toLowerCase() === user.empId.toLowerCase()) {
      throw new Error('อนุมัติใบลาของตัวเองไม่ได้');
    }
    sh.getRange(i + 2, C.STATUS + 1, 1, 4).setValues([[
      status, new Date(), sheetText_(user.name), sheetText_(String(note || '').trim().slice(0, 300))
    ]]);
  } finally {
    lock.releaseLock();
  }
  return { ok: true, id: id, status: status };
}

// ===== Helpers =====
function sheetText_(value) {
  const text = String(value || '');
  return /^[=+@-]/.test(text) ? "'" + text : text;
}

function validateLeavePolicy_(empId, type, start, end, days) {
  const rows = readLeaves_().filter(r => r.empId.toLowerCase() === empId.toLowerCase() && r.status !== STATUS.REJECTED);
  const a = Utilities.formatDate(start, 'Asia/Bangkok', 'yyyy-MM-dd');
  const b = Utilities.formatDate(end, 'Asia/Bangkok', 'yyyy-MM-dd');
  if (rows.some(r => r.start <= b && r.end >= a)) throw new Error('ช่วงวันลาซ้อนกับใบลาที่รออนุมัติหรืออนุมัติแล้ว');
  const used = rows.filter(r => r._y === start.getFullYear() && r.type === type).reduce((sum, r) => sum + r.days, 0);
  if (LEAVE_QUOTA[type] != null && used + days > LEAVE_QUOTA[type]) {
    throw new Error('วันลาเกินสิทธิ์ประจำปี (รวมใบลาที่รออนุมัติ) เหลือ ' + Math.max(LEAVE_QUOTA[type] - used, 0) + ' วัน');
  }
}

function parseDate_(s) {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(s || ''));
  if (!m) return null;
  const y = Number(m[1]), month = Number(m[2]) - 1, day = Number(m[3]);
  if (y < 2000 || y > 2100) return null;
  const date = new Date(y, month, day);
  return date.getFullYear() === y && date.getMonth() === month && date.getDate() === day ? date : null;
}

function countWorkdays_(start, end) {
  let n = 0;
  for (let d = new Date(start); d <= end; d.setDate(d.getDate() + 1)) {
    const w = d.getDay();
    if (w !== 0 && w !== 6 && HOLIDAYS.indexOf(Utilities.formatDate(d, 'Asia/Bangkok', 'yyyy-MM-dd')) < 0) n++;
  }
  return n;
}

/** (ทางเลือก) สร้างข้อมูลตัวอย่าง 120 รายการ เพื่อทดลอง Dashboard */
function seedSampleData_() {
  const sh = getSheet_();
  const names = ['สมชาย ใจดี','สุรีย์ วัฒนา','กมลพรรณ ศรีสุข','วิทยา ใจมั่น','นภา รุ่งเรือง','อนันต์ มีชัย','ประยุทธ์ ทองดี','จันทร์เพ็ญ วงศ์ใหญ่'];
  const approvers = ['หัวหน้าแผนก', 'ผู้จัดการ'];
  const rows = [];
  const year = new Date().getFullYear();
  for (let i = 0; i < 120; i++) {
    const n = Math.floor(Math.random() * names.length);
    const start = new Date(year, Math.floor(Math.random() * 12), 1 + Math.floor(Math.random() * 26));
    const end = new Date(start); end.setDate(end.getDate() + Math.floor(Math.random() * 3));
    const days = countWorkdays_(start, end) || 1;
    const st = [STATUS.APPROVED, STATUS.APPROVED, STATUS.APPROVED, STATUS.PENDING, STATUS.REJECTED][Math.floor(Math.random() * 5)];
    const decided = st !== STATUS.PENDING;
    rows.push(['LVDEMO-' + String(i + 1).padStart(4, '0'), new Date(start.getTime() - 86400000 * 3),
      names[n], 'E' + String(100 + n), DEPARTMENTS[n % DEPARTMENTS.length],
      LEAVE_TYPES[Math.floor(Math.random() * 4)], start, end, days, 'ข้อมูลตัวอย่าง', st,
      decided ? new Date() : '', decided ? approvers[i % 2] : '',
      st === STATUS.REJECTED ? 'ไม่มีใบรับรองแพทย์' : '']);
  }
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, HEADERS.length).setValues(rows);
}
