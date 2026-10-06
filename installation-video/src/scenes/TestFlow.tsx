import { Scene, Card, Row, Action } from "../Design";
export const TestFlow = () => (
  <Scene
    step="07"
    kicker="VERIFY END TO END"
    title={"ทดสอบหนึ่งใบลา\nตั้งแต่ต้นจนจบ"}
    lines={[
      "พนักงานเลือกวันทำงาน กรอกเหตุผล แล้วส่ง",
      "ผู้อนุมัติอีกบัญชีตรวจและอนุมัติ",
      "พนักงานตรวจสถานะ ผู้ดูแลตรวจชีตข้อมูล",
    ]}
    note="ทดสอบวันหยุด ใบลาซ้อน และโควตาเกินด้วย ระบบควรปฏิเสธคำขอเหล่านี้"
  >
    <Card title="ใบลาทดสอบ / E001">
      <Row label="1 · ส่งคำขอ" value="รออนุมัติ" />
      <Row label="2 · A001 ตรวจ" value="อนุมัติ" active />
      <Row label="3 · ข้อมูล" value="บันทึกใน Google Sheets" />
      <Action>ตรวจครบทั้ง 3 จุด</Action>
    </Card>
  </Scene>
);
