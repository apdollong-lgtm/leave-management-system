# FactoryStudio - ระบบลางานออนไลน์

เวอร์ชัน 2.2.0 ใช้ Google Apps Script + Google Sheets ติดตั้งแยกในบัญชี Google ของแต่ละองค์กร

## ติดตั้ง

1. เปิด https://script.google.com สร้างโปรเจกต์ใหม่
2. วาง `Code.gs` และ `PasswordCrypto.gs` เป็นไฟล์สคริปต์แยก แล้วสร้างไฟล์ HTML ชื่อ `Dashboard` แล้ววาง `Dashboard.html`
3. Project Settings: เปิดแสดง `appsscript.json` และวาง manifest จากแพ็กเกจ
4. Script Properties: เพิ่ม `ADMIN_PASSWORD` เป็นรหัสผ่าน 12-128 ตัวอักษรที่ไม่เป็นตัวเลขล้วน ไม่มีรหัสผ่านเริ่มต้น
5. แก้ `ORG_NAME`, `ORG_TAGLINE`, `DEPARTMENTS`, `LEAVE_QUOTA`, `HOLIDAYS` ใน `Code.gs` ตามองค์กร
6. รัน `setup` หนึ่งครั้ง ตรวจสิทธิ์และอนุญาตด้วยบัญชีเจ้าของ ระบบจะสร้างชีตและแสดง URL ใน Execution log
7. Deploy > New deployment > Web app: Execute as Me, Access Anyone (หรือขอบเขตที่ Google Workspace อนุญาต)
8. เปิด URL `/exec` เข้าด้วย Username `admin` และรหัสผ่านที่ตั้งเอง เพิ่มชื่อผู้ใช้ รหัสพนักงาน และรหัสผ่านของพนักงาน/ผู้อนุมัติในเมนูตั้งค่า
9. ทดสอบด้วยข้อมูลทดสอบขององค์กร: ยื่นใบลา อนุมัติด้วยอีกบัญชี ตรวจสถานะ และตรวจแถวข้อมูลใน Google Sheets

คู่มือฉบับพิมพ์: `docs/installation-guide.html` และ `release-assets/installation-guide-th.pdf`

## กติกาวันลา

- นับวันจันทร์-ศุกร์ ไม่รวมวันใน `HOLIDAYS` (YYYY-MM-DD)
- ครึ่งวันเลือกได้เฉพาะวันทำงานวันเดียว ไม่แยกช่วงเช้า/บ่าย; ถ้ามีใบลาในวันนั้นอยู่แล้วถือว่าซ้อน
- ห้ามช่วงวันลาซ้อนกับคำขอรออนุมัติหรืออนุมัติแล้ว
- โควตาประจำปีหักทั้งคำขอรออนุมัติและอนุมัติแล้ว คืนโควตาเมื่อไม่อนุมัติ
- ใบลาข้ามปีให้แยกเป็นสองคำขอ ประเภทที่ไม่มีใน `LEAVE_QUOTA` ไม่จำกัดโควตา
- ผู้ใช้ทุกสิทธิ์อนุมัติใบลาของตนเองไม่ได้ ผู้ดูแลหลัก `admin` ยื่นใบลาไม่ได้
- เก็บชีตเป็นส่วนตัว จัดการผู้ใช้ผ่านหน้าเว็บเพื่อรักษาการตรวจสิทธิ์
- ระบบนี้ยังไม่มีการลารายชั่วโมง ตารางกะ การแนบใบรับรองแพทย์ หรือการแจ้งเตือนอีเมล

## การอัปเดต

สำรองชีตก่อนอัปเดต วางไฟล์เวอร์ชันใหม่ทั้ง 4 ไฟล์ แล้ว Deploy > Manage deployments > Edit > New version เพื่อคง URL เดิม อย่าลบ `SPREADSHEET_ID` และ `ADMIN_PASSWORD_HASH` ใน Script Properties

การอัปเดตแพ็กเกจบนแค็ตตาล็อกทำผ่าน GitHub Actions ทุกครั้งที่ push เข้า `main` และผ่านการทดสอบ ดู `docs/catalog-sync.md` การส่งแพ็กเกจไปแค็ตตาล็อกไม่ได้เปลี่ยนโปรเจกต์ Google ของลูกค้าแต่ละราย

## ทดสอบสำหรับผู้พัฒนา

ต้องมี Node.js 22 ขึ้นไป จาก root รัน `npm test` ใช้บริการ Google จำลองในหน่วยทดสอบ ไม่เขียนข้อมูล HR จริง ตรวจเว็บจริงอีกครั้งก่อนส่งมอบ

## วิดีโอสอนติดตั้ง

ใน `installation-video`: รัน `npm ci` แล้ว `npx remotion studio --no-open` เปิด composition `InstallationGuide` ขนาด 1920x1080, 30 fps, ความยาว 114 วินาที เป็นข้อความภาษาไทยพร้อม Motion Graphic ไม่มีเสียงพากย์

เมื่อจะส่งออก: `npx remotion render src/index.ts InstallationGuide ../release-assets/installation-guide-th.mp4`

ฟอนต์ Prompt มาพร้อม SIL Open Font License ใน `installation-video/public/fonts/OFL.txt` การใช้งาน Remotion ให้เป็นไปตามใบอนุญาตของ Remotion

เอกสาร Google: [Web app deployment](https://developers.google.com/apps-script/guides/web), [Script Properties](https://developers.google.com/apps-script/guides/properties), [Service quotas](https://developers.google.com/apps-script/guides/services/quotas)

## ย้ายจากระบบ PIN เดิม

รหัสพนักงานและประวัติใบลายังคงเดิม เพิ่มคอลัมน์ชื่อผู้ใช้ใน Users อัตโนมัติ ชื่อผู้ใช้เดิมเริ่มจากรหัสพนักงาน ผู้ดูแลต้องตั้งรหัสผ่านใหม่ให้บัญชีเดิมทุกบัญชีในหน้า ตั้งค่า บัญชีที่ยังไม่ได้ตั้งรหัสผ่านใหม่จะเข้าสู่ระบบไม่ได้ ไม่รับ PIN เดิมเป็นรหัสผ่าน

ชื่อผู้ใช้มี 3-30 ตัว ใช้ A-Z, 0-9, _ - . ไม่แยกตัวพิมพ์ใหญ่เล็ก รหัสผ่านรักษาตัวพิมพ์และช่องว่าง เก็บด้วย PBKDF2-HMAC-SHA256 600,000 รอบและ salt แยกแต่ละบัญชี setup แปลง ADMIN_PASSWORD เป็น ADMIN_PASSWORD_HASH แล้วลบค่ารหัสผ่านตั้งต้นออกจาก Script Properties เปลี่ยนรหัสผ่านได้ในเมนูผู้ใช้ Session เก่าจะหมดสิทธิ์ทันที

หากต้องการชื่อผู้ดูแลหลักอื่น ให้ตั้ง ADMIN_USERNAME (ค่าเริ่มต้น admin) การกู้บัญชีผู้ดูแลให้ตั้ง ADMIN_PASSWORD ใหม่ใน Script Properties แล้วรัน setup อีกครั้ง
