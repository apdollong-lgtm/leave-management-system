import { Scene, Card, Row, Action } from "../Design";
export const Outro = () => (
  <Scene
    step="READY"
    kicker="HAND OVER"
    title={"ตรวจครบแล้ว\nส่งมอบให้ทีม"}
    lines={[
      "แชร์เฉพาะลิงก์เว็บกับพนักงาน",
      "เก็บชีตให้เข้าถึงเฉพาะผู้รับผิดชอบ",
      "สำรองข้อมูลและเก็บคู่มือติดตั้งไว้",
    ]}
    note="อัปเดตโค้ด: Deploy → Manage deployments → Edit → New version เพื่อรักษาลิงก์เดิม"
  >
    <Card title="Installation checklist">
      <Row label="ตั้งรหัสผ่าน + องค์กร" value="✓" />
      <Row label="สร้างชีต + ผู้ใช้งาน" value="✓" />
      <Row label="ทดสอบส่ง + อนุมัติ" value="✓" active />
      <Action>FactoryStudio · Better Tomorrow</Action>
    </Card>
  </Scene>
);
