const {test} = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const crypto = require('node:crypto');

function app() {
  const rows = [], users = [], properties = new Map([['ADMIN_PIN','83927461'],['PIN_SALT','test-salt']]), cache = new Map();
  const sheet = values => ({
    getLastRow: () => values.length + 1,
    getRange: (row, col, count = 1, width = 1) => ({
      getValues: () => values.slice(row-2,row-2+count).map(r => r.slice(col-1,col-1+width)),
      setValues: newRows => newRows.forEach((r,i) => { r.forEach((v,j) => values[row-2+i][col-1+j]=v); }),
      setValue: value => {values[row-2][col-1] = value;},
    }),
    appendRow: row => values.push(row),
  });
  const ctx = vm.createContext({
    PropertiesService: {getScriptProperties: () => ({getProperty: k => properties.get(k),setProperty: (k,v) => properties.set(k,v)})},
    CacheService: {getScriptCache: () => ({get: k => cache.get(k),put: (k,v) => cache.set(k,v),remove: k => cache.delete(k)})},
    LockService: {getScriptLock: () => ({waitLock(){},releaseLock(){}})},
    Utilities: {
      getUuid: () => crypto.randomUUID(),
      formatDate: (date,tz,pattern) => {
        const p = new Intl.DateTimeFormat('en-CA',{timeZone:tz,year:'numeric',month:'2-digit',day:'2-digit'}).formatToParts(date);
        const v = Object.fromEntries(p.map(x => [x.type,x.value]));
        return pattern === 'yyMMdd' ? v.year.slice(-2)+v.month+v.day : `${v.year}-${v.month}-${v.day}`;
      },
      DigestAlgorithm:{SHA_256:'sha256'}, Charset:{UTF_8:'utf8'},
      computeDigest: (_,text) => crypto.createHash('sha256').update(text).digest(),
      base64Encode: value => Buffer.from(value).toString('base64'),
    },
  });
  vm.runInContext(fs.readFileSync('Code.gs','utf8'),ctx);
  ctx.getSheet_ = () => sheet(rows);
  ctx.getUsersSheet_ = () => sheet(users);
  const admin = ctx.login('admin','83927461').token;
  ctx.saveUser(admin,{empId:'E001',name:'Employee',dept:'QC',role:'employee',pin:'675839'});
  ctx.saveUser(admin,{empId:'A001',name:'Approver',dept:'QA',role:'approver',pin:'786594'});
  const employee = ctx.login('E001','675839').token;
  const approver = ctx.login('A001','786594').token;
  return {ctx,rows,users,properties,admin,employee,approver};
}
const form = extra => ({type:'ลาพักร้อน',start:'2026-10-06',end:'2026-10-06',reason:'พักผ่อน',...extra});

test('server and browser scripts parse', () => {
  new vm.Script(fs.readFileSync('Code.gs','utf8'));
  for (const m of fs.readFileSync('Dashboard.html','utf8').matchAll(/<script\b[^>]*>([\s\S]*?)<\/script>/gi)) if(m[1].trim()) new vm.Script(m[1]);
});
test('rejects invalid calendar dates and accepts leap day', () => {
  const {ctx} = app();
  for(const date of ['2026-02-31','2026-02-29','2026-13-01','2026-00-01','2026-01-00','2101-01-01','bad']) assert.equal(ctx.parseDate_(date),null);
  assert.ok(ctx.parseDate_('2028-02-29'));
});
test('rejects weekend and holiday half-days', () => {
  const {ctx,employee} = app();
  assert.throws(() => ctx.submitLeave(employee,form({start:'2026-10-10',end:'2026-10-10',halfDay:true})),/วันหยุด/);
  vm.runInContext("HOLIDAYS.push('2026-10-06')",ctx);
  assert.throws(() => ctx.submitLeave(employee,form({halfDay:true})),/วันหยุด/);
});
test('valid half-day is stored, duplicate is rejected', () => {
  const {ctx,employee,rows} = app();
  assert.equal(ctx.submitLeave(employee,form({halfDay:true})).days,0.5);
  assert.equal(rows.length,1);
  assert.throws(() => ctx.submitLeave(employee,form()),/ซ้อน/);
});
test('requires reason and valid range; refuses multi-day half-day and cross-year request', () => {
  const {ctx,employee} = app();
  for(const extra of [{reason:''},{start:'2026-10-07'},{end:'2026-10-07',halfDay:true},{start:'2026-12-31',end:'2027-01-04'}]) assert.throws(() => ctx.submitLeave(employee,form(extra)));
});
test('pending leaves reserve quota, rejected leaves release it', () => {
  const {ctx,employee,approver,rows} = app();
  const first = ctx.submitLeave(employee,form({start:'2026-10-05',end:'2026-10-16'}));
  assert.equal(first.days,10);
  assert.throws(() => ctx.submitLeave(employee,form({start:'2026-10-19',end:'2026-10-19'})),/เกินสิทธิ์/);
  ctx.updateStatus(approver,first.id,'ไม่อนุมัติ','แก้ไขวันลา');
  assert.equal(ctx.submitLeave(employee,form({start:'2026-10-19',end:'2026-10-19'})).days,1);
  assert.equal(rows.length,2);
});
test('approved leaves block overlap and decisions are final', () => {
  const {ctx,employee,approver} = app();
  const result = ctx.submitLeave(employee,form());
  ctx.updateStatus(approver,result.id,'อนุมัติ','');
  assert.throws(() => ctx.submitLeave(employee,form()),/ซ้อน/);
  assert.throws(() => ctx.updateStatus(approver,result.id,'ไม่อนุมัติ',''),/ดำเนินการ/);
});
test('employees cannot approve or list users; approvers cannot approve own leave', () => {
  const {ctx,employee,approver} = app();
  const result = ctx.submitLeave(approver,form());
  assert.throws(() => ctx.updateStatus(employee,result.id,'อนุมัติ',''),/สิทธิ์/);
  assert.throws(() => ctx.listUsers(employee),/สิทธิ์/);
  assert.throws(() => ctx.getDashboardData(employee),/สิทธิ์/);
  assert.throws(() => ctx.updateStatus(approver,result.id,'อนุมัติ',''),/ตัวเอง/);
});
test('admin employees also cannot approve own leave; super admin cannot submit', () => {
  const {ctx,admin} = app();
  ctx.saveUser(admin,{empId:'HR1',name:'HR',dept:'บุคคล',role:'admin',pin:'786594'});
  const token=ctx.login('HR1','786594').token;
  const result=ctx.submitLeave(token,form());
  assert.throws(() => ctx.updateStatus(token,result.id,'อนุมัติ',''),/ตัวเอง/);
  assert.throws(() => ctx.submitLeave(admin,form()),/ผู้ดูแลหลัก/);
});
test('inactive accounts and logout invalidate sessions', () => {
  const {ctx,employee,admin} = app();
  ctx.saveUser(admin,{empId:'E001',name:'Employee',dept:'QC',role:'employee',active:false});
  assert.throws(() => ctx.whoami(employee),/SESSION_EXPIRED/);
  ctx.logout(admin);
  assert.throws(() => ctx.listUsers(admin),/SESSION_EXPIRED/);
});
test('own leave queries exclude other users, users expose no hashes', () => {
  const {ctx,employee,approver,admin} = app();
  ctx.submitLeave(employee,form()); ctx.submitLeave(approver,form());
  assert.equal(ctx.getMyLeaves(employee,2026).rows.length,1);
  assert.equal(ctx.getMyLeaves(employee,2026).rows[0].empId,'E001');
  assert.equal('pinHash' in ctx.listUsers(admin)[0],false);
});
test('invalid role, weak admin PIN, and brute force are rejected', () => {
  const {ctx,admin,properties} = app();
  assert.throws(() => ctx.saveUser(admin,{empId:'bad',name:'Bad',dept:'QC',role:'constructor',pin:'675839'}),/สิทธิ์/);
  for(let i=0;i<5;i++) assert.throws(() => ctx.login('E001','0000'),/ไม่ถูกต้อง/);
  assert.throws(() => ctx.login('E001','675839'),/10 นาที/);
  properties.set('ADMIN_PIN','1234'); assert.throws(() => ctx.adminPin_(),/ADMIN_PIN/);
});
test('formula-like text is stored as plain text', () => {
  const {ctx,employee,rows} = app();
  ctx.submitLeave(employee,form({reason:'=IMPORTXML("https://example.com", "x")'}));
  assert.ok(rows[0][9].startsWith("'="));
});
test('changing PIN invalidates prior sessions and accepts the new PIN', () => {
  const {ctx,employee}=app();
  ctx.changeMyPin(employee,'675839','938475');
  assert.throws(()=>ctx.whoami(employee),/SESSION_EXPIRED/);
  assert.throws(()=>ctx.login('E001','675839'),/ไม่ถูกต้อง/);
  assert.ok(ctx.login('E001','938475').token);
});
