import { Scene, Card, Row, Action, Cursor } from "../Design";
export const Security = () => (
  <Scene
    step="02"
    kicker="SECURE ADMIN"
    title={"ตั้งรหัสผ่านผู้ดูแล\nก่อนเปิดใช้งาน"}
    lines={[
      "Project Settings → Script Properties",
      "เพิ่มชื่อ ADMIN_PASSWORD",
      "ใช้รหัสผ่าน 12–128 ตัวอักษรที่คาดเดายาก",
    ]}
    note="setup จะแปลงรหัสผ่านเป็น hash และลบรหัสผ่านตั้งต้นออกจาก Script Properties"
  >
    <Card title="Project Settings / Script Properties">
      <Row label="Property" value="ADMIN_PASSWORD" active />
      <Row label="Value" value="● ● ● ● ● ● ● ●" />
      <Row label="Time zone" value="Asia/Bangkok" />
      <Action>Save script properties</Action>
    </Card>
    <Cursor x={380} y={420} />
  </Scene>
);
