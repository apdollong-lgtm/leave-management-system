import { Scene, Card, Row, Action } from "../Design";
export const Intro = () => (
  <Scene
    step="START"
    kicker="ติดตั้งด้วยบัญชี Google ของคุณ"
    title={"ระบบลางานออนไลน์\nเริ่มต้นอย่างเป็นขั้นตอน"}
    lines={[
      "ยื่นใบลา → อนุมัติ → ดูรายงาน",
      "ข้อมูลเก็บใน Google Sheets ขององค์กร",
      "คู่มือนี้พาไปตั้งค่าจนทดสอบได้",
    ]}
    note="เตรียม Code.gs, PasswordCrypto.gs, Dashboard.html และ appsscript.json จากแพ็กเกจสินค้า"
  >
    <Card title="Leave Management / FactoryStudio">
      <Row label="พนักงาน" value="ยื่นใบลาและติดตามสถานะ" />
      <Row label="ผู้อนุมัติ" value="ตรวจและอนุมัติคำขอ" />
      <Row label="ผู้ดูแล" value="จัดการผู้ใช้งานและสิทธิ์" />
      <Action>7 ขั้นตอนติดตั้ง</Action>
    </Card>
  </Scene>
);
