import { useState } from "react";
import { useTheme } from "./ThemeContext";
import { THEMES, THEME_KEYS } from "./themes";

const SWATCHES = {
  dark:   ["#030712", "#0b1220", "#2563eb", "#93c5fd"],
  light:  ["#f8fafc", "#ffffff", "#2563eb", "#1e40af"],
  ocean:  ["#0a1628", "#0c1f3d", "#0ea5e9", "#38bdf8"],
  forest: ["#0a1a0e", "#0d2414", "#16a34a", "#4ade80"],
  purple: ["#0f0a1e", "#160f2e", "#7c3aed", "#a78bfa"],
};

export default function ThemePicker() {
  const { theme, themeKey, setTheme } = useTheme();
  const [open, setOpen] = useState(false);

  return (
    <div style={{ position: "relative" }}>
      <button
        onClick={() => setOpen((v) => !v)}
        title="Change theme"
        style={{
          display: "flex", alignItems: "center", gap: 7,
          background: theme.cardBg, border: `1px solid ${theme.cardBorder}`,
          borderRadius: 8, padding: "6px 12px", cursor: "pointer",
          color: theme.pageSubText, fontSize: 12, fontWeight: 600,
          fontFamily: theme.font,
        }}
      >
        <Swatch colors={SWATCHES[themeKey]} size={14} />
        {THEMES[themeKey].name}
        <span style={{ fontSize: 10, opacity: 0.6 }}>{open ? "▲" : "▼"}</span>
      </button>

      {open && (
        <>
          {/* backdrop */}
          <div onClick={() => setOpen(false)}
            style={{ position: "fixed", inset: 0, zIndex: 98 }} />

          <div style={{
            position: "absolute", top: "calc(100% + 6px)", right: 0, zIndex: 99,
            background: theme.cardBg, border: `1px solid ${theme.cardBorder}`,
            borderRadius: 10, padding: 8, display: "flex", flexDirection: "column",
            gap: 4, minWidth: 160, boxShadow: "0 8px 32px #00000055",
          }}>
            {THEME_KEYS.map((key) => {
              const active = key === themeKey;
              return (
                <button key={key} onClick={() => { setTheme(key); setOpen(false); }}
                  style={{
                    display: "flex", alignItems: "center", gap: 10,
                    background: active ? theme.tabActiveBg : "none",
                    border: `1px solid ${active ? theme.accent : "transparent"}`,
                    borderRadius: 7, padding: "7px 10px", cursor: "pointer",
                    color: active ? theme.accentText : theme.pageSubText,
                    fontSize: 12, fontWeight: active ? 700 : 400,
                    fontFamily: THEMES[key].font, textAlign: "left", width: "100%",
                  }}
                >
                  <Swatch colors={SWATCHES[key]} size={16} />
                  {THEMES[key].name}
                  {active && <span style={{ marginLeft: "auto", fontSize: 10 }}>✓</span>}
                </button>
              );
            })}
          </div>
        </>
      )}
    </div>
  );
}

function Swatch({ colors, size }) {
  return (
    <span style={{ display: "inline-flex", borderRadius: 4, overflow: "hidden", flexShrink: 0 }}>
      {colors.map((c, i) => (
        <span key={i} style={{ width: size / 2, height: size, background: c }} />
      ))}
    </span>
  );
}
