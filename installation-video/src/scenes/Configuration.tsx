import { Scene, Card, Row } from "../Design";
export const Configuration = () => (
  <Scene
    step="03"
    kicker="ORGANIZATION SETTINGS"
    title={"ปรับข้อมูล\nให้ตรงกับองค์กร"}
    lines={[
      "แก้ ORG_NAME และ DEPARTMENTS ใน Code.gs",
      "ปรับ LEAVE_QUOTA ตามนโยบายองค์กร",
      "ใส่ HOLIDAYS เป็นวันที่ YYYY-MM-DD",
    ]}
    note="ระบบนับวันจันทร์–ศุกร์ หักวันหยุดบริษัท และให้แยกใบลาข้ามปี"
  >
    <Card title="Code.gs / Configuration">
      <Row label="ORG_NAME" value="ชื่อองค์กรของคุณ" active />
      <Row label="DEPARTMENTS" value="QC, QA, บุคคล …" />
      <Row label="LEAVE_QUOTA" value="พักร้อน / ป่วย / กิจ" />
      <Row label="HOLIDAYS" value="['2026-12-31']" />
    </Card>
  </Scene>
);
