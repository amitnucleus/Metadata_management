import { useState } from "react";
import { useTheme } from "./ThemeContext";

const SECTIONS = [
  {
    id: "upload",
    icon: "⬆",
    title: "File Upload",
    color: "#22c55e",
    items: [
      { label: "Supported formats", desc: "XLSX, XLS, ODS, CSV, and JSON files are accepted." },
      { label: "Drag & drop", desc: "Drag a file anywhere onto the upload zone, or click it to browse." },
      { label: "Multi-sheet files", desc: "Each sheet in an XLSX/ODS file becomes its own tab. CSV/JSON files create a single sheet named after the file." },
      { label: "Python API fallback", desc: "Files are first parsed by the Python API (if running). If unavailable, the browser parses locally via SheetJS." },
    ],
  },
  {
    id: "edit",
    icon: "✎",
    title: "Editing",
    color: "#60a5fa",
    items: [
      { label: "Inline cell editing", desc: "Click any cell to edit it. Press Enter or click away to confirm, Esc to cancel." },
      { label: "Enum dropdowns", desc: "Columns matched to an enum rule show a dropdown instead of a text field." },
      { label: "Integer fields", desc: "Columns matched to an integer rule show a number input." },
      { label: "Dirty tracking", desc: "Edited cells are highlighted in amber. A badge in the toolbar counts total unsaved edits." },
      { label: "Add row", desc: 'Click "+ Add Row" below the table to append a new empty row.' },
      { label: "Delete row", desc: "Click the ✕ button at the end of any row to remove it." },
    ],
  },
  {
    id: "columns",
    icon: "⊞",
    title: "Column Management",
    color: "#a78bfa",
    items: [
      { label: "Add column", desc: 'Click the "+" button in the table header to add a new column. You will be prompted for a name.' },
      { label: "Rename column", desc: "Double-click any column header to rename it inline. Press Enter to confirm, Esc to cancel." },
      { label: "Delete column", desc: "Hover over a column header and click the ✕ that appears to delete the column from all rows." },
    ],
  },
  {
    id: "undo",
    icon: "↩",
    title: "Undo / Redo",
    color: "#fbbf24",
    items: [
      { label: "Undo", desc: "Click ↩ Undo in the toolbar or press Ctrl+Z (Cmd+Z on Mac) to revert the last cell change." },
      { label: "Redo", desc: "Click ↪ Redo or press Ctrl+Y (Cmd+Shift+Z on Mac) to reapply an undone change." },
      { label: "History depth", desc: "Up to 50 steps of cell-edit history are kept. History resets when a new file is loaded." },
    ],
  },
  {
    id: "search",
    icon: "🔍",
    title: "Search & Filter",
    color: "#38bdf8",
    items: [
      { label: "Search bar", desc: "Type in the search box above the table to filter rows in real time. Matches any column value." },
      { label: "Match count", desc: 'When a filter is active, a "X of Y rows" counter appears below the table.' },
      { label: "Clear", desc: "Click the ✕ inside the search box or clear the text to show all rows again." },
    ],
  },
  {
    id: "validation",
    icon: "⚙",
    title: "Column Validation Rules",
    color: "#7c3aed",
    items: [
      { label: "Auto-detection", desc: "Rules are matched to columns by name pattern (e.g. any column containing \"data_type\" gets the Column Data Type enum rule)." },
      { label: "Rule types", desc: "enum — must match a list of allowed values; integer — must be a whole number; regex — must be a valid regular expression." },
      { label: "Edit enum options", desc: 'Expand the rules panel (⚙) and click "Edit options" on an enum rule to customise the allowed values.' },
      { label: "Explicit assignment", desc: 'Use "+ Assign column" on any rule to manually pin a column to that rule, overriding pattern matching.' },
      { label: "Reset", desc: '"↺ Reset to defaults" restores all rules and clears explicit assignments.' },
      { label: "Validation errors", desc: "Invalid cells are highlighted red. An ⚠ badge on the sheet tab and toolbar counts errors per sheet." },
    ],
  },
  {
    id: "schema",
    icon: "✓",
    title: "Schema Validation",
    color: "#34d399",
    items: [
      { label: "Open", desc: 'Click "✓ Schema" in the toolbar to open the schema checker for the active sheet.' },
      { label: "How to use", desc: "Paste or type the expected column names (one per line) into the text box and click Check Schema." },
      { label: "Results", desc: "Missing columns (in schema but not in file), extra columns (in file but not in schema), and matched columns are listed with colour coding." },
    ],
  },
  {
    id: "profile",
    icon: "∑",
    title: "Data Profiling",
    color: "#f472b6",
    items: [
      { label: "Profile modal", desc: 'Click "∑ Profile" to open a per-column statistics table for all sheets.' },
      { label: "Stats shown", desc: "Total row count, null count, null %, unique value count, top 3 most frequent values, numeric min and max." },
      { label: "Inline null badges", desc: 'Click "∅ Show Nulls" to toggle a small null-% badge on each column header directly in the table.' },
      { label: "High-null warning", desc: "Columns with > 50 % nulls are highlighted red in the profile modal and in the inline badge." },
    ],
  },
  {
    id: "diff",
    icon: "⊡",
    title: "Diff View",
    color: "#fb923c",
    items: [
      { label: "Open", desc: 'When there are unsaved edits, a "⊡ Diff (N)" button appears in the toolbar. Click it to review all changes.' },
      { label: "What it shows", desc: "Every dirty cell is listed with its sheet name, row number, column name, original value, and new value." },
      { label: "Grouped by sheet", desc: "Changes are grouped by sheet tab for easy navigation." },
    ],
  },
  {
    id: "audit",
    icon: "📋",
    title: "Audit Log",
    color: "#fbbf24",
    items: [
      { label: "Live recording", desc: "Every cell edit is recorded automatically with a timestamp, sheet name, row, column, old value, and new value." },
      { label: "Open", desc: 'Click "📋 Log (N)" in the toolbar to view the full audit history.' },
      { label: "Order", desc: "Events are shown newest-first." },
      { label: "Export", desc: 'Click "↓ Export XLSX" inside the audit log modal to download the log as a separate spreadsheet.' },
    ],
  },
  {
    id: "comments",
    icon: "💬",
    title: "Cell Comments",
    color: "#f59e0b",
    items: [
      { label: "Add a comment", desc: "Hover over any cell to reveal a 💬 button on the right side of the cell. Click it to open the comment editor." },
      { label: "Keyboard shortcut", desc: "Press Ctrl+Enter inside the comment textarea to save." },
      { label: "Visual indicator", desc: "An orange dot appears in the top-right corner of any cell that has a comment." },
      { label: "Edit / remove", desc: "Click the orange dot (or 💬) to re-open the editor. Use the Remove button to delete the comment." },
      { label: "Comment count", desc: "A badge in the toolbar shows the total number of comments across all sheets." },
    ],
  },
  {
    id: "export",
    icon: "↓",
    title: "Export",
    color: "#22c55e",
    items: [
      { label: "Save Changes", desc: "The green Save Changes button commits all edits and downloads the file. After saving, the dirty count resets." },
      { label: "↓ JSON", desc: "Downloads the current data (all sheets) as a formatted JSON file at any time, regardless of dirty state." },
      { label: "↓ XLSX", desc: "Downloads as an XLSX file. Available when SheetJS is loaded (should be immediate)." },
      { label: "Audit Export", desc: "Inside the Audit Log modal, export the full edit history as a separate XLSX file." },
      { label: "File naming", desc: 'Downloaded files are named "{original}_edited.xlsx / .json" and "{original}_audit.xlsx".' },
    ],
  },
  {
    id: "themes",
    icon: "🎨",
    title: "Themes",
    color: "#e879f9",
    items: [
      { label: "Available themes", desc: "Dark, Light, Ocean, Forest, Purple — select from the theme picker in the top-right corner." },
      { label: "Persistence", desc: "The selected theme is remembered in localStorage and restored on next visit." },
    ],
  },
];

export default function HelpDocs() {
  const { theme } = useTheme();
  const [open,           setOpen]           = useState(false);
  const [expandedSection, setExpandedSection] = useState(null);

  function toggle(id) {
    setExpandedSection((v) => (v === id ? null : id));
  }

  return (
    <>
      {/* Trigger button */}
      <button
        onClick={() => setOpen(true)}
        title="Feature documentation"
        style={{
          background: "none",
          border: `1px solid ${theme.cardBorder}`,
          borderRadius: 7,
          padding: "5px 12px",
          cursor: "pointer",
          color: theme.mutedText,
          fontSize: 12,
          fontWeight: 600,
          fontFamily: theme.font,
          display: "flex",
          alignItems: "center",
          gap: 6,
          transition: "border-color 0.15s, color 0.15s",
        }}
        onMouseEnter={(e) => {
          e.currentTarget.style.borderColor = `${theme.accent}66`;
          e.currentTarget.style.color = theme.accentText;
        }}
        onMouseLeave={(e) => {
          e.currentTarget.style.borderColor = theme.cardBorder;
          e.currentTarget.style.color = theme.mutedText;
        }}
      >
        ? Docs
      </button>

      {/* Modal overlay */}
      {open && (
        <div
          style={{
            position: "fixed", inset: 0, background: "#00000099", zIndex: 2000,
            display: "flex", alignItems: "flex-start", justifyContent: "center",
            paddingTop: 40, paddingBottom: 40, overflowY: "auto",
          }}
          onClick={() => setOpen(false)}
        >
          <div
            style={{
              background: theme.cardBg,
              border: `1px solid ${theme.cardBorder}`,
              borderRadius: 14,
              padding: "28px 28px 32px",
              maxWidth: 700,
              width: "90vw",
              display: "flex",
              flexDirection: "column",
              gap: 20,
              fontFamily: theme.font,
              boxShadow: "0 24px 80px #00000088",
            }}
            onClick={(e) => e.stopPropagation()}
          >
            {/* Header */}
            <div style={{ display: "flex", alignItems: "flex-start", justifyContent: "space-between", gap: 12 }}>
              <div>
                <h2 style={{ margin: 0, fontSize: 18, fontWeight: 700, color: theme.pageText }}>
                  Metadata Management — Feature Guide
                </h2>
                <p style={{ margin: "4px 0 0", fontSize: 12, color: theme.mutedText }}>
                  Click any section to expand its details.
                </p>
              </div>
              <button
                onClick={() => setOpen(false)}
                style={{ background: "none", border: `1px solid ${theme.cardBorder}`, borderRadius: 6,
                  cursor: "pointer", color: theme.mutedText, fontSize: 13, padding: "3px 10px",
                  fontFamily: theme.font, flexShrink: 0 }}
              >✕ Close</button>
            </div>

            {/* Section grid — 2 columns summary chips */}
            <div style={{ display: "flex", flexWrap: "wrap", gap: 6 }}>
              {SECTIONS.map((s) => (
                <button
                  key={s.id}
                  onClick={() => toggle(s.id)}
                  style={{
                    background: expandedSection === s.id ? `${s.color}18` : "none",
                    border: `1px solid ${expandedSection === s.id ? s.color + "55" : theme.cardBorder}`,
                    borderRadius: 6, padding: "4px 10px", cursor: "pointer",
                    color: expandedSection === s.id ? s.color : theme.mutedText,
                    fontSize: 11, fontWeight: 600, fontFamily: theme.font,
                    display: "flex", alignItems: "center", gap: 5,
                    transition: "all 0.15s",
                  }}
                >
                  <span>{s.icon}</span> {s.title}
                </button>
              ))}
            </div>

            {/* Expanded section detail */}
            {expandedSection && (() => {
              const s = SECTIONS.find((x) => x.id === expandedSection);
              if (!s) return null;
              return (
                <div style={{
                  background: theme.surfaceBg,
                  border: `1px solid ${s.color}33`,
                  borderLeft: `3px solid ${s.color}`,
                  borderRadius: 8,
                  padding: "16px 20px",
                  display: "flex",
                  flexDirection: "column",
                  gap: 12,
                }}>
                  <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                    <span style={{ fontSize: 18 }}>{s.icon}</span>
                    <span style={{ fontSize: 14, fontWeight: 700, color: s.color }}>{s.title}</span>
                  </div>
                  <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
                    {s.items.map((item, i) => (
                      <div key={i} style={{ display: "flex", gap: 12, alignItems: "flex-start" }}>
                        <span style={{
                          flexShrink: 0, minWidth: 160, fontSize: 12, fontWeight: 600,
                          color: theme.accentText, paddingTop: 1,
                        }}>
                          {item.label}
                        </span>
                        <span style={{ fontSize: 12, color: theme.cellText, lineHeight: 1.6 }}>
                          {item.desc}
                        </span>
                      </div>
                    ))}
                  </div>
                </div>
              );
            })()}

            {/* Quick-reference keyboard shortcuts */}
            <div style={{
              background: theme.surfaceBg,
              border: `1px solid ${theme.cardBorder}`,
              borderRadius: 8,
              padding: "14px 18px",
            }}>
              <div style={{ fontSize: 11, fontWeight: 700, color: theme.mutedText,
                textTransform: "uppercase", letterSpacing: 0.5, marginBottom: 10 }}>
                Keyboard Shortcuts
              </div>
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "6px 24px" }}>
                {[
                  ["Ctrl + Z",       "Undo last cell edit"],
                  ["Ctrl + Y",       "Redo last undone edit"],
                  ["Enter",          "Confirm cell edit"],
                  ["Esc",            "Cancel cell edit"],
                  ["Ctrl + Enter",   "Save cell comment"],
                  ["Double-click header", "Rename column"],
                ].map(([key, desc]) => (
                  <div key={key} style={{ display: "flex", gap: 8, alignItems: "center" }}>
                    <kbd style={{
                      background: theme.cardBg, border: `1px solid ${theme.cardBorder}`,
                      borderRadius: 4, padding: "1px 6px", fontSize: 10,
                      color: theme.accentText, fontFamily: "monospace", whiteSpace: "nowrap",
                    }}>{key}</kbd>
                    <span style={{ fontSize: 11, color: theme.mutedText }}>{desc}</span>
                  </div>
                ))}
              </div>
            </div>

          </div>
        </div>
      )}
    </>
  );
}
