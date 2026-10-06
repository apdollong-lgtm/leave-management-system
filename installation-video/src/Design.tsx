import React from "react";
import {
  AbsoluteFill,
  Easing,
  interpolate,
  useCurrentFrame,
  useVideoConfig,
  staticFile,
  delayRender,
  continueRender,
} from "remotion";
import { loadFont } from "@remotion/fonts";
const handle = delayRender("Thai fonts");
Promise.all([
  loadFont({
    family: "Prompt",
    url: staticFile("fonts/Prompt-Regular.ttf"),
    weight: "400",
  }),
  loadFont({
    family: "Prompt",
    url: staticFile("fonts/Prompt-SemiBold.ttf"),
    weight: "600",
  }),
]).then(() => continueRender(handle));
export const Card: React.FC<React.PropsWithChildren<{ title: string }>> = ({
  title,
  children,
}) => (
  <div
    style={{
      width: 820,
      height: 620,
      border: "1px solid #324968",
      borderRadius: 28,
      background: "#13233d",
      overflow: "hidden",
      boxShadow: "0 35px 80px #0005",
    }}
  >
    <div
      style={{
        height: 66,
        padding: "18px 26px",
        background: "#1b3050",
        fontSize: 22,
        color: "#a8b9d3",
      }}
    >
      ● ● ● <span style={{ marginLeft: 20 }}>{title}</span>
    </div>
    <div style={{ padding: 32 }}>{children}</div>
  </div>
);
export const Row: React.FC<{
  label: string;
  value: string;
  active?: boolean;
}> = ({ label, value, active }) => {
  const f = useCurrentFrame();
  return (
    <div
      style={{
        padding: "18px 22px",
        marginBottom: 16,
        borderRadius: 16,
        border: `1px solid ${active ? "#63a4ff" : "#324968"}`,
        background: active ? "#234572" : "#182b47",
        opacity: interpolate(f, [12, 30], [0, 1], {
          extrapolateLeft: "clamp",
          extrapolateRight: "clamp",
        }),
        display: "flex",
        justifyContent: "space-between",
        gap: 20,
        fontSize: 26,
      }}
    >
      <span style={{ color: "#a8b9d3" }}>{label}</span>
      <b style={{ fontWeight: 600 }}>{value}</b>
    </div>
  );
};
export const Action: React.FC<React.PropsWithChildren> = ({ children }) => {
  const f = useCurrentFrame();
  return (
    <div
      style={{
        marginTop: 22,
        background: "#63a4ff",
        color: "#08162a",
        borderRadius: 14,
        padding: "20px 26px",
        fontSize: 30,
        fontWeight: 600,
        display: "inline-block",
        scale: interpolate(f, [25, 48], [0.92, 1], {
          extrapolateLeft: "clamp",
          extrapolateRight: "clamp",
          easing: Easing.bezier(0.16, 1, 0.3, 1),
        }),
      }}
    >
      {children}
    </div>
  );
};
export const Cursor: React.FC<{ x: number; y: number }> = ({ x, y }) => {
  const f = useCurrentFrame();
  return (
    <div
      style={{
        position: "absolute",
        left: x,
        top: y,
        translate: interpolate(f, [25, 65], ["140px 100px", "0px 0px"], {
          extrapolateLeft: "clamp",
          extrapolateRight: "clamp",
          easing: Easing.bezier(0.16, 1, 0.3, 1),
        }),
        opacity: interpolate(f, [20, 30], [0, 1], {
          extrapolateLeft: "clamp",
          extrapolateRight: "clamp",
        }),
      }}
    >
      <svg width="66" height="66" viewBox="0 0 40 40">
        <path
          d="M4 3 6 32l8-8 7 12 7-4-7-12 12-2Z"
          fill="#fff"
          stroke="#63a4ff"
          strokeWidth="2"
        />
      </svg>
      <div
        style={{
          position: "absolute",
          left: -24,
          top: -24,
          width: 70,
          height: 70,
          borderRadius: "50%",
          border: "3px solid #63a4ff",
          scale: interpolate(f, [65, 85], [0.2, 1.4], {
            extrapolateLeft: "clamp",
            extrapolateRight: "clamp",
          }),
          opacity: interpolate(f, [65, 85], [1, 0], {
            extrapolateLeft: "clamp",
            extrapolateRight: "clamp",
          }),
        }}
      />
    </div>
  );
};
export const Scene: React.FC<
  React.PropsWithChildren<{
    step: string;
    kicker: string;
    title: string;
    lines: string[];
    note: string;
  }>
> = ({ step, kicker, title, lines, note, children }) => {
  const f = useCurrentFrame();
  const { durationInFrames } = useVideoConfig();
  return (
    <AbsoluteFill
      style={{
        background: "#08152a",
        color: "#edf4ff",
        fontFamily: "Prompt",
        overflow: "hidden",
      }}
    >
      <div
        style={{
          position: "absolute",
          width: 1050,
          height: 1050,
          right: -340,
          top: -280,
          borderRadius: "50%",
          background: "radial-gradient(circle,#214b8540,transparent 68%)",
          translate: interpolate(f, [0, 420], ["0px 0px", "-80px 90px"]),
        }}
      />
      <svg
        style={{ position: "absolute", inset: 0, opacity: 0.18 }}
        width="1920"
        height="1080"
      >
        <defs>
          <pattern
            id="grid"
            width="70"
            height="70"
            patternUnits="userSpaceOnUse"
          >
            <path d="M70 0H0V70" fill="none" stroke="#6385ae" strokeWidth="1" />
          </pattern>
        </defs>
        <rect width="1920" height="1080" fill="url(#grid)" />
      </svg>
      <div
        style={{
          position: "absolute",
          left: 110,
          right: 110,
          top: 68,
          display: "flex",
          justifyContent: "space-between",
          fontSize: 24,
          color: "#a8b9d3",
          letterSpacing: 2,
        }}
      >
        <b style={{ color: "#edf4ff" }}>
          FACTORY<span style={{ color: "#63a4ff" }}>STUDIO</span>
        </b>
        <span>LEAVE MANAGEMENT · INSTALLATION GUIDE</span>
      </div>
      <div
        style={{
          position: "absolute",
          left: 110,
          top: 210,
          width: 800,
          opacity: interpolate(
            f,
            [0, 18, durationInFrames - 15, durationInFrames - 1],
            [0, 1, 1, 0],
            { extrapolateLeft: "clamp", extrapolateRight: "clamp" },
          ),
          translate: interpolate(f, [0, 28], ["0px 32px", "0px 0px"], {
            extrapolateRight: "clamp",
            easing: Easing.bezier(0.16, 1, 0.3, 1),
          }),
        }}
      >
        <div
          style={{
            color: "#5fe0b6",
            fontSize: 26,
            marginBottom: 24,
            letterSpacing: 3,
          }}
        >
          {step} / {kicker}
        </div>
        <div
          style={{
            fontSize: 76,
            fontWeight: 600,
            lineHeight: 1.3,
            whiteSpace: "pre-line",
            letterSpacing: -1,
            marginBottom: 36,
          }}
        >
          {title}
        </div>
        {lines.map((line, i) => (
          <div
            key={line}
            style={{
              fontSize: 32,
              lineHeight: 1.7,
              color: "#a8b9d3",
              opacity: interpolate(f, [20 + i * 12, 36 + i * 12], [0, 1], {
                extrapolateLeft: "clamp",
                extrapolateRight: "clamp",
              }),
              display: "flex",
              gap: 18,
              marginBottom: 12,
            }}
          >
            <span style={{ color: "#63a4ff" }}>→</span>
            <span>{line}</span>
          </div>
        ))}
      </div>
      <div
        style={{
          position: "absolute",
          left: 990,
          top: 218,
          opacity: interpolate(f, [8, 30], [0, 1], {
            extrapolateLeft: "clamp",
            extrapolateRight: "clamp",
          }),
          translate: interpolate(f, [8, 35], ["60px 0px", "0px 0px"], {
            extrapolateLeft: "clamp",
            extrapolateRight: "clamp",
            easing: Easing.bezier(0.16, 1, 0.3, 1),
          }),
        }}
      >
        {children}
      </div>
      <div
        style={{
          position: "absolute",
          bottom: 110,
          left: 110,
          right: 110,
          padding: "20px 26px",
          borderLeft: "4px solid #5fe0b6",
          background: "#142640",
          borderRadius: 12,
          fontSize: 27,
          color: "#bfd0e7",
        }}
      >
        {note}
      </div>
      <div
        style={{
          position: "absolute",
          left: 110,
          bottom: 64,
          fontSize: 18,
          color: "#a8b9d3",
        }}
      >
        ภาพจำลองเพื่ออธิบายขั้นตอน · Google Apps Script + Google Sheets
      </div>
      <div
        style={{
          position: "absolute",
          bottom: 0,
          left: 0,
          height: 5,
          width: `${interpolate(f, [0, durationInFrames], [0, 100])}%`,
          background: "#63a4ff",
        }}
      />
    </AbsoluteFill>
  );
};
