import { Scene, Card, Row, Action } from "../Design";
export const Users = () => (
  <Scene
    step="06"
    kicker="ADD YOUR TEAM"
    title={"เพิ่มพนักงาน\nและกำหนดสิทธิ์"}
    lines={[
      "เปิด URL → รหัส admin และ PIN ที่ตั้งไว้",
      "เมนูตั้งค่า → เพิ่มรหัส ชื่อ แผนก และ PIN",
      "แยกบัญชีพนักงานและผู้อนุมัติสำหรับทดสอบ",
    ]}
    note="บัญชี admin หลักใช้จัดการระบบ บัญชีพนักงานใช้ยื่นใบลา"
  >
    <Card title="ตั้งค่า / ผู้ใช้งาน">
      <Row label="E001" value="พนักงาน" active />
      <Row label="A001" value="ผู้อนุมัติ" />
      <Row label="HR001" value="ผู้ดูแลระบบ" />
      <Action>บันทึกผู้ใช้งาน</Action>
    </Card>
  </Scene>
);
