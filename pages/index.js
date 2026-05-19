import MetadataTab from "../components/MetadataTab";

export default function Home() {
  return (
    <div style={{ maxWidth: 1200, margin: "0 auto", padding: "32px 24px" }}>
      <div style={{ marginBottom: 24 }}>
        <h1 style={{ fontSize: 22, fontWeight: 700, color: "#e2e8f0", marginBottom: 4 }}>
          Metadata Management
        </h1>
        <p style={{ fontSize: 13, color: "#64748b" }}>
          Upload, view, edit, and download data mapping specification files (XLSX, CSV, JSON).
        </p>
      </div>
      <MetadataTab />
    </div>
  );
}
