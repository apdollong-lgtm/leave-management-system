import { Scene, Card, Row, Action, Cursor } from "../Design";
export const Project = () => (
  <Scene
    step="01"
    kicker="CREATE PROJECT"
    title={"สร้างโปรเจกต์\nGoogle Apps Script"}
    lines={[
      "เปิด script.google.com → New project",
      "วาง Code.gs, PasswordCrypto.gs และ Dashboard",
      "เปิดแสดง appsscript.json ใน Project Settings",
    ]}
    note="คัดลอกเนื้อหาให้ครบทั้ง 4 ไฟล์ แล้วบันทึกโปรเจกต์"
  >
    <Card title="Apps Script / Files">
      <Row label="Script" value="Code.gs" active />
      <Row label="Crypto" value="PasswordCrypto.gs" />
      <Row label="HTML" value="Dashboard.html" />
      <Row label="Manifest" value="appsscript.json" />
      <Action>Save project</Action>
    </Card>
    <Cursor x={355} y={410} />
  </Scene>
);
