import { Scene, Card, Row, Action, Cursor } from "../Design";
export const Deploy = () => (
  <Scene
    step="05"
    kicker="DEPLOY WEB APP"
    title={"เผยแพร่เว็บ\nรับลิงก์ใช้งาน"}
    lines={[
      "Deploy → New deployment → Web app",
      "Execute as: Me / ผู้เผยแพร่",
      "Who has access: Anyone แล้วกด Deploy",
    ]}
    note="หากองค์กรจำกัดการเผยแพร่ ให้ใช้ตัวเลือกที่ผู้ดูแล Google Workspace อนุญาต"
  >
    <Card title="New deployment">
      <Row label="Type" value="Web app" />
      <Row label="Execute as" value="Me" active />
      <Row label="Who has access" value="Anyone" />
      <Action>Deploy → Copy Web app URL</Action>
    </Card>
    <Cursor x={375} y={420} />
  </Scene>
);
