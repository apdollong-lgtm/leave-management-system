# FactoryStudio - ระบบลางานออนไลน์

เวอร์ชัน 2.1.0 ใช้ Google Apps Script + Google Sheets ติดตั้งแยกในบัญชี Google ของแต่ละองค์กร

## ติดตั้ง

1. เปิด https://script.google.com สร้างโปรเจกต์ใหม่
2. วาง `Code.gs` และสร้างไฟล์ HTML ชื่อ `Dashboard` แล้ววาง `Dashboard.html`
3. Project Settings: เปิดแสดง `appsscript.json` และวาง manifest จากแพ็กเกจ
4. Script Properties: เพิ่ม `ADMIN_PIN` เป็นตัวเลข 6-12 หลักที่คาดเดายาก ห้ามใช้ตัวเลขซ้ำทั้งหมดหรือ 123456 ไม่มี PIN เริ่มต้น
5. แก้ `ORG_NAME`, `ORG_TAGLINE`, `DEPARTMENTS`, `LEAVE_QUOTA`, `HOLIDAYS` ใน `Code.gs` ตามองค์กร
6. รัน `setup` หนึ่งครั้ง ตรวจสิทธิ์และอนุญาตด้วยบัญชีเจ้าของ ระบบจะสร้างชีตและแสดง URL ใน Execution log
7. Deploy > New deployment > Web app: Execute as Me, Access Anyone (หรือขอบเขตที่ Google Workspace อนุญาต)
8. เปิด URL `/exec` เข้าด้วยรหัส `admin` และ PIN ที่ตั้งเอง เพิ่มบัญชีพนักงานและผู้อนุมัติในเมนูตั้งค่า
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

สำรองชีตก่อนอัปเดต วางไฟล์เวอร์ชันใหม่ แล้ว Deploy > Manage deployments > Edit > New version เพื่อคง URL เดิม อย่าลบ `SPREADSHEET_ID` และ `PIN_SALT` ใน Script Properties

การอัปเดตแพ็กเกจบนแค็ตตาล็อกทำผ่าน GitHub Actions ทุกครั้งที่ push เข้า `main` และผ่านการทดสอบ ดู `docs/catalog-sync.md` การส่งแพ็กเกจไปแค็ตตาล็อกไม่ได้เปลี่ยนโปรเจกต์ Google ของลูกค้าแต่ละราย

## ทดสอบสำหรับผู้พัฒนา

ต้องมี Node.js 22 ขึ้นไป จาก root รัน `npm test` ใช้บริการ Google จำลองในหน่วยทดสอบ ไม่เขียนข้อมูล HR จริง ตรวจเว็บจริงอีกครั้งก่อนส่งมอบ

## วิดีโอสอนติดตั้ง

ใน `installation-video`: รัน `npm ci` แล้ว `npx remotion studio --no-open` เปิด composition `InstallationGuide` ขนาด 1920x1080, 30 fps, ความยาว 114 วินาที เป็นข้อความภาษาไทยพร้อม Motion Graphic ไม่มีเสียงพากย์

เมื่อจะส่งออก: `npx remotion render src/index.ts InstallationGuide ../release-assets/installation-guide-th.mp4`

ฟอนต์ Prompt มาพร้อม SIL Open Font License ใน `installation-video/public/fonts/OFL.txt` การใช้งาน Remotion ให้เป็นไปตามใบอนุญาตของ Remotion

เอกสาร Google: [Web app deployment](https://developers.google.com/apps-script/guides/web), [Script Properties](https://developers.google.com/apps-script/guides/properties), [Service quotas](https://developers.google.com/apps-script/guides/services/quotas)
