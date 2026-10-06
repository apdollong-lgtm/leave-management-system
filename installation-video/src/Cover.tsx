import { AbsoluteFill } from "remotion";

export const Cover = () => (
  <AbsoluteFill
    style={{
      background: "#08152a",
      fontFamily: "Prompt",
      color: "#edf4ff",
      padding: "150px 92px",
      overflow: "hidden",
    }}
  >
    <div
      style={{
        position: "absolute",
        width: 1200,
        height: 1200,
        right: -400,
        top: -200,
        borderRadius: "50%",
        background: "radial-gradient(circle,#285fa855,transparent 68%)",
      }}
    />
    <div style={{ fontSize: 36, letterSpacing: 4, fontWeight: 600 }}>
      FACTORY<span style={{ color: "#63a4ff" }}>STUDIO</span>
    </div>
    <div
      style={{
        marginTop: 100,
        color: "#5fe0b6",
        fontSize: 28,
        letterSpacing: 4,
      }}
    >
      HUMAN RESOURCES / FS-HR-101
    </div>
    <div
      style={{ fontSize: 100, fontWeight: 600, lineHeight: 1.5, marginTop: 35 }}
    >
      ระบบลางาน
      <br />
      ออนไลน์
    </div>
    <div
      style={{ color: "#a8b9d3", fontSize: 40, lineHeight: 1.9, marginTop: 28 }}
    >
      ยื่นใบลา · อนุมัติ
      <br />
      ติดตามวันลาในองค์กร
    </div>
    <div
      style={{
        marginTop: 100,
        borderRadius: 30,
        border: "2px solid #345175",
        padding: 38,
        background: "#162943",
      }}
    >
      <div
        style={{
          fontSize: 28,
          color: "#a8b9d3",
          letterSpacing: 3,
          marginBottom: 30,
        }}
      >
        ONE CONNECTED WORKFLOW
      </div>
      <div
        style={{
          padding: "25px 28px",
          background: "#234572",
          borderRadius: 16,
          fontSize: 36,
          marginBottom: 18,
        }}
      >
        01 · พนักงานส่งใบลา
      </div>
      <div
        style={{
          padding: "25px 28px",
          background: "#1b3456",
          borderRadius: 16,
          fontSize: 36,
          marginBottom: 18,
        }}
      >
        02 · ผู้อนุมัติตรวจคำขอ
      </div>
      <div
        style={{
          padding: "25px 28px",
          background: "#1b3456",
          borderRadius: 16,
          fontSize: 36,
        }}
      >
        03 · ติดตามสถานะและโควตา
      </div>
    </div>
    <div
      style={{ fontSize: 32, color: "#5fe0b6", marginTop: 55, lineHeight: 1.8 }}
    >
      Google Apps Script + Sheets
      <br />
      คู่มือติดตั้งภาษาไทย
    </div>
    <div
      style={{
        position: "absolute",
        left: 92,
        bottom: 115,
        fontSize: 25,
        color: "#a8b9d3",
        letterSpacing: 3,
      }}
    >
      PEOPLE · PROCESS · BETTER TOMORROW
    </div>
  </AbsoluteFill>
);
