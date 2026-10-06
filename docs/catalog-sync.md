# อัปเดตแพ็กเกจบน FactoryStudio Catalog

Repository: `apdollong-lgtm/leave-management-system`

Product slug: `leave-management-system`

GitHub Actions รันทดสอบก่อนส่งไฟล์ทุกครั้งที่ push เข้า `main` สามารถเรียกเองด้วย workflow_dispatch ได้ เวอร์ชันแพ็กเกจเริ่มจาก package.json และเพิ่ม commit SHA เพื่อระบุโค้ดที่ส่งจริง

เว็บรับ GitHub OIDC เฉพาะ repository ที่อยู่ใน allowlist และ main/semantic version tags ใช้ audience `https://catalog.factorystudio.pro` ไม่ต้องเก็บ secret ระยะยาวใน product repository

ขั้นตอน: ทดสอบ → สร้าง ZIP จากไฟล์ที่ commit → ขอ URL อัปโหลดชั่วคราว → PUT ไฟล์ไป R2 ส่วนตัว → PATCH ยืนยันและลงทะเบียนในแค็ตตาล็อก

`release-assets/` เก็บคู่มือ PDF ภาพหน้าปก และ MP4 ที่ส่งออกแล้ว ถ้าแก้ source วิดีโอ ให้ส่งออกใหม่ก่อน push จึงจะได้ไฟล์วิดีโอใหม่บนเว็บ รายละเอียดสินค้าฝั่งเนื้อหาจะอัปเดตจาก `catalog-product.json` ส่วนราคาและการเผยแพร่จัดการในหน้า Admin

แพ็กเกจสำหรับลูกค้ามี Code.gs, Dashboard.html, manifest, README, คู่มือ และ source วิดีโอ ไม่ส่ง `.clasp.json` ซึ่งเป็น Script ID ของเจ้าของ ไม่ส่งข้อมูลรับรอง โฟลเดอร์ .git ไฟล์ .env หรือข้อมูล HR

คำสั่งหลังแก้ไฟล์และทดสอบแล้ว:

```powershell
npm test
git add Code.gs Dashboard.html appsscript.json README.md docs tests installation-video release-assets
git commit -m "Update leave management"
git push origin main
```

การอัปเดตนี้เปลี่ยนไฟล์ดาวน์โหลดของสินค้า ไม่อัปเดตโปรเจกต์ Google ของลูกค้าโดยอัตโนมัติ ลูกค้าต้องเผยแพร่ New version ในบัญชีของตนเองตามคู่มือ
