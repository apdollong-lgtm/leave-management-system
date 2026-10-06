import { Scene, Card, Row, Action, Cursor } from "../Design";
export const Setup = () => (
  <Scene
    step="04"
    kicker="INITIALIZE DATA"
    title={"รัน setup\nสร้างชีตข้อมูล"}
    lines={[
      "เลือกฟังก์ชัน setup → Run",
      "ตรวจสิทธิ์และอนุญาตด้วยบัญชีเจ้าของ",
      "เปิดลิงก์ Google Sheets ใน Execution log",
    ]}
    note="setup สร้าง LeaveRequests และ Users พร้อมจำ Spreadsheet ID อัตโนมัติ"
  >
    <Card title="Apps Script / Execution">
      <Row label="Function" value="setup" active />
      <Action>▶ Run</Action>
      <div style={{ marginTop: 30 }}>
        <Row label="Sheet 1" value="LeaveRequests" />
        <Row label="Sheet 2" value="Users" />
      </div>
    </Card>
    <Cursor x={178} y={222} />
  </Scene>
);
