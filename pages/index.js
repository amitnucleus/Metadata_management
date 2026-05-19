import MetadataTab from "../components/MetadataTab";
import ThemePicker from "../components/ThemePicker";
import HelpDocs from "../components/HelpDocs";
import { useTheme } from "../components/ThemeContext";

export default function Home() {
  const { theme } = useTheme();
  return (
    <div style={{ maxWidth: 1200, margin: "0 auto", padding: "32px 24px", fontFamily: theme.font }}>
      <div style={{ display: "flex", alignItems: "flex-start", justifyContent: "space-between", marginBottom: 24, gap: 16 }}>
        <div>
          <h1 style={{ fontSize: 22, fontWeight: 700, color: theme.pageText, marginBottom: 4 }}>
            Metadata Management
          </h1>
          <p style={{ fontSize: 13, color: theme.pageSubText }}>
            Upload, view, edit, and download data mapping specification files (XLSX, CSV, JSON).
          </p>
        </div>
        <div style={{ display: "flex", alignItems: "center", gap: 8, flexShrink: 0 }}>
          <HelpDocs />
          <ThemePicker />
        </div>
      </div>
      <MetadataTab />
    </div>
  );
}
