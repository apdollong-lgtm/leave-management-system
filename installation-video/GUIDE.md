# วิดีโอสอนติดตั้งระบบลางาน

ภาษาไทย 16:9, 1920x1080, 30 fps, 114 วินาที ใช้ภาพจำลองพร้อม Motion Graphic ไม่มีเสียงพากย์

```powershell
npm ci
npm run dev
```

เปิด composition `InstallationGuide` ใน Remotion Studio แต่ละฉากมี composition ของตนเองใน `Scenes` จึงแก้ข้อความและเวลาของแต่ละขั้นตอนได้

ส่งออก MP4 ด้วย `npm run render` ไฟล์จะอยู่ใน `../release-assets/installation-guide-th.mp4` และถูกแนบไปแค็ตตาล็อกเมื่อ push เข้า main

ตรวจ source ด้วย `npm run lint` ฟอนต์ไทย Prompt อยู่ใน public/fonts พร้อมใบอนุญาต OFL ไม่ต้องโหลดฟอนต์จากอินเทอร์เน็ตตอนแสดงวิดีโอ
