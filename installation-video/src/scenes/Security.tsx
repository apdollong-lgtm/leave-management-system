import { Scene, Card, Row, Action, Cursor } from "../Design";
export const Security = () => (
  <Scene
    step="02"
    kicker="SECURE ADMIN"
    title={"ตั้ง PIN ผู้ดูแล\nก่อนเปิดใช้งาน"}
    lines={[
      "Project Settings → Script Properties",
      "เพิ่มชื่อ ADMIN_PIN",
      "ใช้ตัวเลข 6–12 หลักที่คาดเดายาก",
    ]}
    note="เก็บ PIN เป็นความลับ ไม่ใส่ใน GitHub หรือคู่มือที่ส่งให้คนอื่น"
  >
    <Card title="Project Settings / Script Properties">
      <Row label="Property" value="ADMIN_PIN" active />
      <Row label="Value" value="● ● ● ● ● ● ● ●" />
      <Row label="Time zone" value="Asia/Bangkok" />
      <Action>Save script properties</Action>
    </Card>
    <Cursor x={380} y={420} />
  </Scene>
);
