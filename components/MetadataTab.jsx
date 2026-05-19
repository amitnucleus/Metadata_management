import { useCallback, useEffect, useMemo, useRef, useState } from "react";

// ── SheetJS from CDN ──────────────────────────────────────────────────────────
function useXLSX() {
  const [xlsx, setXlsx] = useState(null);
  useEffect(() => {
    if (window.XLSX) { setXlsx(window.XLSX); return; }
    const script = document.createElement("script");
    script.src = "https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js";
    script.async = true;
    script.onload = () => setXlsx(window.XLSX);
    document.head.appendChild(script);
  }, []);
  return xlsx;
}

// ── Predefined column rules ───────────────────────────────────────────────────
export const DEFAULT_RULES = [
  { id:"file_format",   label:"File Format",          pattern:/file[\s_-]?format/i,
    type:"enum",    color:"#60a5fa",
    options:["CSV","TSV","JSON","XML","Parquet","Avro","ORC","Delimited","Fixed","XLSX"] },
  { id:"file_encoding", label:"File Encoding",        pattern:/file[\s_-]?encoding|encoding/i,
    type:"enum",    color:"#a78bfa",
    options:["UTF-8","UTF-16","UTF-32","ASCII","ISO-8859-1","LATIN-1"] },
  { id:"col_data_type", label:"Column Data Type",     pattern:/column[\s_-]?data[\s_-]?type|data[\s_-]?type/i,
    type:"enum",    color:"#34d399",
    options:["STRING","INTEGER","BIGINT","SMALLINT","TINYINT","FLOAT","DOUBLE","DECIMAL",
             "BOOLEAN","DATE","TIMESTAMP","BINARY","ARRAY","MAP","STRUCT","VARCHAR","CHAR"] },
  { id:"action_type",   label:"Action Type",          pattern:/action[\s_-]?type/i,
    type:"enum",    color:"#fbbf24",
    options:["VIEW","INSERT","UPDATE","DELETE","ADD","MODIFY","UPSERT"] },
  { id:"feed_type",     label:"Feed / File Type",     pattern:/feed[\s_-]?file[\s_-]?type|feed[\s_-]?type/i,
    type:"enum",    color:"#f472b6",
    options:["Delimited","Fixed","JSON","XML","Binary","Multi-record"] },
  { id:"mandatory",     label:"Mandatory",            pattern:/mandatory|required/i,
    type:"enum",    color:"#22c55e",
    options:["YES","NO"] },
  { id:"derived_logic", label:"Derived Logic",        pattern:/derived[\s_-]?logic|transformation[\s_-]?logic/i,
    type:"enum",    color:"#fb923c",
    options:["Direct Mapping","Transformation","Lookup","Calculated","Derived",
             "Concatenation","Split","Default Value","Aggregation","Filter"] },
  { id:"integer_pos",   label:"Integer / Position",  pattern:/\b(position|index|order|seq|sequence|count|row[\s_-]?count|num|number)\b/i,
    type:"integer", color:"#94a3b8" },
  { id:"col_format",    label:"Column Format (regex)",pattern:/column[\s_-]?format|format[\s_-]?string|pattern/i,
    type:"regex",   color:"#e879f9" },
];

// ── Rule matching ─────────────────────────────────────────────────────────────
// columnAssignments: { [ruleId]: Set<colName> } — explicit overrides beat patterns
function matchRule(colName, rules, columnAssignments = {}) {
  for (const rule of rules) {
    if (columnAssignments[rule.id]?.has(colName)) return rule;
  }
  return rules.find((r) => r.pattern.test(colName)) || null;
}

function validateValue(val, rule) {
  if (!rule || val == null || val === "") return null;
  const s = String(val);
  if (rule.type === "enum")    return rule.options.includes(s) ? null : `Must be one of: ${rule.options.join(", ")}`;
  if (rule.type === "integer") {
    if (!/^-?\d+$/.test(s.trim())) return "Must be an integer";
    const n = parseInt(s, 10);
    if (rule.min != null && n < rule.min) return `Min value is ${rule.min}`;
    if (rule.max != null && n > rule.max) return `Max value is ${rule.max}`;
    return null;
  }
  if (rule.type === "regex") {
    try { new RegExp(s); return null; } catch { return "Invalid regular expression"; }
  }
  return null;
}

// ── Parse JSON ────────────────────────────────────────────────────────────────
function parseJsonMapping(text) {
  const raw = JSON.parse(text);
  if (typeof raw !== "object" || Array.isArray(raw) || raw === null)
    throw new Error("Expected a JSON object with sheet names as keys.");
  const parsed = {};
  for (const [name, value] of Object.entries(raw)) {
    if (!Array.isArray(value)) { parsed[name] = { headers: [], rows: [] }; continue; }
    const headerSet = new Map();
    value.forEach((row) => {
      if (row && typeof row === "object")
        Object.keys(row).forEach((k) => { if (!headerSet.has(k)) headerSet.set(k, true); });
    });
    const headers = [...headerSet.keys()];
    const rows = value.map((row) => {
      const obj = {};
      headers.forEach((h) => { obj[h] = row?.[h] ?? null; });
      return obj;
    });
    parsed[name] = { headers, rows };
  }
  return parsed;
}

// ── Export helpers ────────────────────────────────────────────────────────────
function exportJson(sheets, fileName) {
  const out = {};
  Object.entries(sheets).forEach(([n, { rows }]) => { out[n] = rows; });
  const blob = new Blob([JSON.stringify(out, null, 2)], { type: "application/json" });
  triggerDownload(blob, fileName.replace(/\.[^.]+$/, "") + "_edited.json");
}
function exportXlsx(xlsx, sheets, fileName) {
  const wb = xlsx.utils.book_new();
  Object.entries(sheets).forEach(([name, { headers, rows }]) => {
    const aoa = [headers, ...rows.map((r) => headers.map((h) => r[h] ?? ""))];
    xlsx.utils.book_append_sheet(wb, xlsx.utils.aoa_to_sheet(aoa), name.slice(0, 31));
  });
  xlsx.writeFile(wb, fileName.replace(/\.[^.]+$/, "") + "_edited.xlsx");
}
function triggerDownload(blob, name) {
  const url = URL.createObjectURL(blob);
  const a = Object.assign(document.createElement("a"), { href: url, download: name });
  document.body.appendChild(a); a.click();
  document.body.removeChild(a); URL.revokeObjectURL(url);
}

// ── Styles ────────────────────────────────────────────────────────────────────
const S = {
  container: { display: "flex", flexDirection: "column", gap: 20 },
  card:      { background: "#0b1220", border: "1px solid #1e293b", borderRadius: 10, padding: "16px 20px" },
  dropzone: (active) => ({
    border: `2px dashed ${active ? "#2563eb" : "#1e293b"}`,
    borderRadius: 12, padding: "40px 32px", textAlign: "center", cursor: "pointer",
    background: active ? "#0f172a" : "#060c18", transition: "border-color 0.2s, background 0.2s",
  }),
  uploadIcon:  { fontSize: 32, marginBottom: 8, color: "#2563eb" },
  uploadTitle: { fontSize: 15, fontWeight: 700, color: "#e2e8f0", marginBottom: 4 },
  uploadSub:   { fontSize: 12, color: "#475569" },
  badge: (color) => ({
    display: "inline-block", padding: "2px 8px", borderRadius: 4, fontSize: 10,
    fontWeight: 700, background: `${color}22`, border: `1px solid ${color}44`, color,
  }),
  fileInfo: {
    display: "flex", alignItems: "center", gap: 8, padding: "10px 16px",
    background: "#0b1220", border: "1px solid #1e293b", borderRadius: 8, flexWrap: "wrap",
  },
  fileName: { fontSize: 13, fontWeight: 600, color: "#93c5fd", flex: 1, minWidth: 0 },
  btn: (variant = "ghost") => ({
    background: variant === "primary" ? "#1e3a5f" : variant === "success" ? "#14532d"
              : variant === "danger"  ? "#450a0a" : "none",
    border: `1px solid ${variant === "primary" ? "#2563eb55" : variant === "success" ? "#16a34a55"
                        : variant === "danger"  ? "#ef444455" : "#334155"}`,
    borderRadius: 6, padding: "4px 11px", cursor: "pointer",
    color: variant === "primary" ? "#93c5fd" : variant === "success" ? "#4ade80"
         : variant === "danger"  ? "#f87171" : "#64748b",
    fontSize: 11, fontWeight: 600, whiteSpace: "nowrap",
  }),
  sheetNav: { display: "flex", gap: 6, flexWrap: "wrap", borderBottom: "1px solid #1e293b", paddingBottom: 8 },
  sheetBtn: (active) => ({
    background: active ? "#1e293b" : "none",
    border: `1px solid ${active ? "#2563eb" : "#1e293b"}`,
    borderRadius: 6, padding: "5px 12px", cursor: "pointer",
    color: active ? "#93c5fd" : "#64748b", fontSize: 12, fontWeight: active ? 700 : 400,
    whiteSpace: "nowrap",
  }),
  tableWrap: { overflowX: "auto", borderRadius: 8, border: "1px solid #1e293b" },
  table:     { width: "100%", borderCollapse: "collapse", fontSize: 12 },
  th: (ruleColor) => ({
    background: "#0f172a", padding: "9px 12px", textAlign: "left",
    color: ruleColor || "#93c5fd", fontWeight: 700, fontSize: 11,
    textTransform: "uppercase", letterSpacing: 0.5, whiteSpace: "nowrap",
    borderBottom: `2px solid ${ruleColor ? ruleColor + "55" : "#1e293b"}`,
  }),
  thDelete: { background: "#0f172a", padding: "9px 8px", textAlign: "center",
              borderBottom: "2px solid #1e293b", width: 32 },
  td: (i, dirty, invalid) => ({
    padding: 0, fontSize: 12, borderBottom: "1px solid #0f172a",
    background: invalid ? "#2d0a0a" : dirty ? "#1c1a0a" : i % 2 === 0 ? "#060c18" : "#080e1a",
    transition: "background 0.15s",
    outline: invalid ? "1px solid #ef444455" : "none",
  }),
  errorBox:     { color: "#ef4444", fontSize: 12, padding: "8px 12px", background: "#ef444411", borderRadius: 6, border: "1px solid #ef444433" },
  dirtyBadge:   { fontSize: 10, fontWeight: 700, padding: "2px 7px", borderRadius: 4, background: "#fbbf2422", border: "1px solid #fbbf2444", color: "#fbbf24" },
  invalidBadge: { fontSize: 10, fontWeight: 700, padding: "2px 7px", borderRadius: 4, background: "#ef444422", border: "1px solid #ef444444", color: "#f87171" },
  inputBase:    { width: "100%", boxSizing: "border-box", background: "#0f2040", border: "none", color: "#e2e8f0", fontSize: 12, padding: "7px 12px" },
};

// ── Editable cell ─────────────────────────────────────────────────────────────
function EditableCell({ value, rowIdx, colKey, dirty, rule, invalid, validationMsg, onChange }) {
  const [editing, setEditing] = useState(false);
  const [draft,   setDraft]   = useState("");
  const inputRef = useRef();

  function start() { setDraft(value != null ? String(value) : ""); setEditing(true); }
  useEffect(() => { if (editing) inputRef.current?.focus(); }, [editing]);
  function commit() { setEditing(false); onChange(rowIdx, colKey, draft === "" ? null : draft); }
  function cancel() { setEditing(false); }

  const displayVal = value != null ? String(value) : "";
  const isYesNo    = ["YES","NO"].includes(displayVal.toUpperCase());
  const yesNoColor = displayVal.toUpperCase() === "YES" ? "#22c55e" : "#ef4444";
  const outlineColor = rule ? rule.color : "#2563eb";

  return (
    <td style={S.td(rowIdx, dirty, invalid)} title={validationMsg || "Click to edit"}
        onClick={() => { if (!editing) start(); }}>
      {editing ? (
        rule?.type === "enum" ? (
          <select ref={inputRef} value={draft}
            onChange={(e) => setDraft(e.target.value)}
            onBlur={commit}
            onKeyDown={(e) => { if (e.key === "Escape") cancel(); }}
            style={{ ...S.inputBase, cursor: "pointer", outline: `2px solid ${outlineColor}` }}>
            <option value="">— select —</option>
            {rule.options.map((o) => <option key={o} value={o}>{o}</option>)}
          </select>
        ) : rule?.type === "integer" ? (
          <input ref={inputRef} type="number" value={draft}
            onChange={(e) => setDraft(e.target.value)} onBlur={commit}
            onKeyDown={(e) => { if (e.key === "Enter") commit(); if (e.key === "Escape") cancel(); }}
            style={{ ...S.inputBase, outline: `2px solid ${outlineColor}` }}
            min={rule.min} max={rule.max} />
        ) : (
          <input ref={inputRef} value={draft}
            onChange={(e) => setDraft(e.target.value)} onBlur={commit}
            onKeyDown={(e) => { if (e.key === "Enter") commit(); if (e.key === "Escape") cancel(); }}
            style={{ ...S.inputBase, outline: `2px solid ${outlineColor}` }}
            placeholder={rule?.type === "regex" ? "e.g. ^[\\w\\s]+$" : ""} />
        )
      ) : (
        <div style={{ padding: "7px 12px", cursor: "text", minHeight: 32,
          display: "flex", alignItems: "center", gap: 6,
          color: invalid ? "#f87171" : dirty ? "#fbbf24" : "#cbd5e1" }}>
          {isYesNo && !dirty && !invalid
            ? <span style={S.badge(yesNoColor)}>{displayVal}</span>
            : displayVal || <span style={{ color: "#1e293b" }}>—</span>}
          {dirty && !invalid && <span style={{ fontSize: 9, color: "#fbbf2488" }}>✎</span>}
          {invalid && <span title={validationMsg} style={{ fontSize: 9, color: "#ef4444", marginLeft: "auto", cursor: "help" }}>⚠</span>}
        </div>
      )}
    </td>
  );
}

// ── Column header ─────────────────────────────────────────────────────────────
function ColHeader({ name, rule, source }) {
  const [tip, setTip] = useState(false);
  const srcLabel = source === "explicit" ? "explicit" : source === "pattern" ? "auto" : null;
  const srcColor = source === "explicit" ? "#fbbf24" : "#64748b";

  return (
    <th style={S.th(rule?.color)}>
      <div style={{ display: "flex", alignItems: "center", gap: 5 }}>
        {name}
        {rule && (
          <span
            style={{ ...S.badge(rule.color), fontSize: 8, padding: "1px 5px", cursor: "help", position: "relative" }}
            title={`Rule: ${rule.label} (${rule.type})`}
            onMouseEnter={() => setTip(true)} onMouseLeave={() => setTip(false)}
          >
            {rule.type === "enum" ? "▼" : rule.type === "integer" ? "123" : ".*"}
            {srcLabel && <span style={{ marginLeft: 3, color: srcColor, fontSize: 7 }}>{srcLabel}</span>}
            {tip && (
              <span style={{
                position: "absolute", top: "calc(100% + 4px)", left: 0, zIndex: 99,
                background: "#0f172a", border: `1px solid ${rule.color}44`,
                borderRadius: 6, padding: "8px 10px", whiteSpace: "nowrap",
                fontSize: 11, color: "#e2e8f0", fontWeight: 400,
                textTransform: "none", letterSpacing: 0, boxShadow: "0 4px 16px #00000066",
              }}>
                <strong style={{ color: rule.color }}>{rule.label}</strong>
                {srcLabel && <span style={{ marginLeft: 6, fontSize: 9, color: srcColor }}>({srcLabel})</span>}
                <br />Type: <em>{rule.type}</em><br />
                {rule.type === "enum"    && <span style={{ color: "#64748b" }}>{rule.options.slice(0,6).join(" · ")}{rule.options.length>6?` +${rule.options.length-6} more`:""}</span>}
                {rule.type === "integer" && <span style={{ color: "#64748b" }}>Whole numbers only</span>}
                {rule.type === "regex"   && <span style={{ color: "#64748b" }}>Must be a valid regex</span>}
              </span>
            )}
          </span>
        )}
      </div>
    </th>
  );
}

// ── Sheet table ───────────────────────────────────────────────────────────────
function SheetTable({ sheetName, sheet, dirtySet, validationMap, onCellChange, onAddRow, onDeleteRow, customRules, columnAssignments }) {
  const { headers, rows } = sheet;

  // For each header: compute rule + whether it came from explicit assignment or pattern
  const headerMeta = headers.map((h) => {
    const explicitRule = customRules.find((r) => columnAssignments[r.id]?.has(h));
    if (explicitRule) return { rule: explicitRule, source: "explicit" };
    const patternRule = customRules.find((r) => r.pattern.test(h));
    if (patternRule) return { rule: patternRule, source: "pattern" };
    return { rule: null, source: "none" };
  });

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
      <div style={S.tableWrap}>
        <table style={S.table}>
          <thead>
            <tr>
              {headers.map((h, i) => (
                <ColHeader key={i} name={h} rule={headerMeta[i].rule} source={headerMeta[i].source} />
              ))}
              <th style={S.thDelete} />
            </tr>
          </thead>
          <tbody>
            {rows.length === 0 ? (
              <tr>
                <td colSpan={headers.length + 1}
                  style={{ padding: "24px", textAlign: "center", color: "#334155", fontSize: 13, fontStyle: "italic" }}>
                  No rows — click "+ Add Row" to start
                </td>
              </tr>
            ) : rows.map((row, ri) => (
              <tr key={ri}>
                {headers.map((h, hi) => {
                  const key = `${sheetName}:${ri}:${h}`;
                  const msg = validationMap.get(key) || null;
                  return (
                    <EditableCell key={h} value={row[h]} rowIdx={ri} colKey={h}
                      dirty={dirtySet.has(key)} rule={headerMeta[hi].rule}
                      invalid={!!msg} validationMsg={msg} onChange={onCellChange} />
                  );
                })}
                <td style={{ ...S.td(ri, false, false), textAlign: "center", padding: "0 4px" }}>
                  <button onClick={() => onDeleteRow(ri)}
                    style={{ background: "none", border: "none", cursor: "pointer",
                      color: "#ef444466", fontSize: 14, lineHeight: 1, padding: "4px 6px" }}
                    title="Delete row">✕</button>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
        <button style={S.btn("success")} onClick={onAddRow}>+ Add Row</button>
        {validationMap.size > 0 && (
          <span style={S.invalidBadge}>⚠ {validationMap.size} validation error{validationMap.size !== 1 ? "s" : ""}</span>
        )}
      </div>
    </div>
  );
}

// ── Column chip ───────────────────────────────────────────────────────────────
function ColChip({ name, color, onRemove, isAuto }) {
  return (
    <span style={{
      display: "inline-flex", alignItems: "center", gap: 4,
      padding: "2px 8px", borderRadius: 4, fontSize: 10, fontWeight: 600,
      background: isAuto ? "#1e293b" : `${color}22`,
      border: `1px solid ${isAuto ? "#334155" : color + "55"}`,
      color: isAuto ? "#64748b" : color,
    }}>
      {name}
      {isAuto && <span style={{ fontSize: 9, color: "#475569" }}>auto</span>}
      {!isAuto && (
        <button onClick={onRemove} style={{
          background: "none", border: "none", cursor: "pointer",
          color: color, fontSize: 11, lineHeight: 1, padding: 0, marginLeft: 2,
        }} title="Remove assignment">×</button>
      )}
    </span>
  );
}

// ── Rules panel ───────────────────────────────────────────────────────────────
function RulesPanel({ customRules, columnAssignments, allColumns, onUpdateRule, onResetRules, onAssignColumn, onUnassignColumn }) {
  const [expanded,    setExpanded]    = useState(false);
  const [editingId,   setEditingId]   = useState(null);
  const [editOptions, setEditOptions] = useState("");
  const [addingFor,   setAddingFor]   = useState(null); // ruleId currently showing column picker
  const [colSearch,   setColSearch]   = useState("");

  function saveOptions(id) {
    const opts = editOptions.split(",").map((s) => s.trim()).filter(Boolean);
    onUpdateRule(id, { options: opts });
    setEditingId(null);
  }

  // For a given rule, the columns currently matched (auto + explicit)
  function ruleColumns(rule) {
    const explicit = [...(columnAssignments[rule.id] || new Set())];
    const auto     = allColumns.filter((c) =>
      !explicit.includes(c) &&
      rule.pattern.test(c) &&
      !customRules.some((r) => r.id !== rule.id && columnAssignments[r.id]?.has(c))
    );
    return { explicit, auto };
  }

  // Columns not yet assigned or pattern-matched to ANY rule
  function availableColumns(ruleId) {
    const alreadyUsed = new Set(
      customRules.flatMap((r) => {
        const { explicit, auto } = ruleColumns(r);
        return [...explicit, ...auto];
      })
    );
    return allColumns.filter((c) =>
      !alreadyUsed.has(c) || (columnAssignments[ruleId]?.has(c))
    ).filter((c) =>
      !c || c.toLowerCase().includes(colSearch.toLowerCase())
    );
  }

  return (
    <div style={{ background: "#080e1a", border: "1px solid #1e293b", borderRadius: 8 }}>
      {/* Header */}
      <button onClick={() => setExpanded((v) => !v)} style={{
        width: "100%", background: "none", border: "none", cursor: "pointer",
        padding: "10px 14px", display: "flex", alignItems: "center", gap: 8,
        color: "#64748b", fontSize: 12, fontWeight: 600, textAlign: "left",
      }}>
        <span style={{ color: "#7c3aed", fontSize: 14 }}>⚙</span>
        Column Validation Rules
        <span style={{ marginLeft: "auto", fontSize: 10 }}>{expanded ? "▲" : "▼"}</span>
        <span style={S.badge("#7c3aed")}>{customRules.length} rules</span>
        {allColumns.length > 0 && (
          <span style={S.badge("#475569")}>
            {allColumns.filter((c) => customRules.some((r) => {
              const { explicit, auto } = ruleColumns(r);
              return explicit.includes(c) || auto.includes(c);
            })).length} / {allColumns.length} columns covered
          </span>
        )}
      </button>

      {expanded && (
        <div style={{ padding: "0 14px 14px", display: "flex", flexDirection: "column", gap: 10 }}>
          <div style={{ fontSize: 11, color: "#334155" }}>
            <strong style={{ color: "#475569" }}>auto</strong> — matched by name pattern.&nbsp;
            <strong style={{ color: "#60a5fa" }}>explicit</strong> — manually pinned. Explicit assignments override patterns.
            {allColumns.length === 0 && <span style={{ color: "#ef444466", marginLeft: 8 }}>Upload a file to see available columns.</span>}
          </div>

          {customRules.map((rule) => {
            const { explicit, auto } = ruleColumns(rule);
            const isAddingHere = addingFor === rule.id;
            const avail = isAddingHere ? availableColumns(rule.id).filter((c) => !explicit.includes(c)) : [];

            return (
              <div key={rule.id} style={{
                background: "#0b1220", border: `1px solid ${rule.color}33`,
                borderLeft: `3px solid ${rule.color}`, borderRadius: 6, padding: "10px 12px",
                display: "flex", flexDirection: "column", gap: 8,
              }}>
                {/* Row 1: name + type + pattern + enum editor */}
                <div style={{ display: "flex", alignItems: "center", gap: 8, flexWrap: "wrap" }}>
                  <span style={{ fontSize: 12, fontWeight: 700, color: rule.color, minWidth: 130 }}>{rule.label}</span>
                  <span style={S.badge("#475569")}>{rule.type}</span>
                  <span style={{ fontSize: 10, color: "#334155", flex: 1 }}>
                    matches: <code style={{ color: "#475569" }}>{rule.pattern.toString()}</code>
                  </span>
                  {rule.type === "enum" && (
                    editingId === rule.id ? (
                      <div style={{ display: "flex", gap: 6, alignItems: "center", width: "100%" }}>
                        <input value={editOptions} onChange={(e) => setEditOptions(e.target.value)}
                          placeholder="Comma-separated values…"
                          style={{ flex: 1, background: "#0f172a", border: "1px solid #2563eb55",
                            borderRadius: 5, padding: "4px 8px", color: "#e2e8f0", fontSize: 11 }} />
                        <button style={S.btn("primary")} onClick={() => saveOptions(rule.id)}>Save</button>
                        <button style={S.btn()} onClick={() => setEditingId(null)}>Cancel</button>
                      </div>
                    ) : (
                      <>
                        <span style={{ fontSize: 10, color: "#475569" }}>
                          {rule.options.slice(0,5).join(", ")}{rule.options.length>5 ? ` +${rule.options.length-5}` : ""}
                        </span>
                        <button style={{ ...S.btn(), fontSize: 10, padding: "2px 8px" }}
                          onClick={() => { setEditingId(rule.id); setEditOptions(rule.options.join(", ")); }}>
                          Edit options
                        </button>
                      </>
                    )
                  )}
                  {rule.type === "integer" && <span style={{ fontSize: 10, color: "#475569" }}>Whole numbers only</span>}
                  {rule.type === "regex"   && <span style={{ fontSize: 10, color: "#475569" }}>Must be a valid regex</span>}
                </div>

                {/* Row 2: applied columns */}
                <div style={{ display: "flex", alignItems: "flex-start", gap: 8, flexWrap: "wrap" }}>
                  <span style={{ fontSize: 10, color: "#475569", fontWeight: 600, paddingTop: 2, whiteSpace: "nowrap" }}>
                    Applied to:
                  </span>
                  <div style={{ display: "flex", gap: 4, flexWrap: "wrap", flex: 1 }}>
                    {auto.map((c) => (
                      <ColChip key={c} name={c} color={rule.color} isAuto />
                    ))}
                    {explicit.map((c) => (
                      <ColChip key={c} name={c} color={rule.color} isAuto={false}
                        onRemove={() => onUnassignColumn(rule.id, c)} />
                    ))}
                    {auto.length === 0 && explicit.length === 0 && (
                      <span style={{ fontSize: 10, color: "#334155", fontStyle: "italic" }}>No columns matched yet</span>
                    )}
                  </div>

                  {/* + Assign column button */}
                  {allColumns.length > 0 && (
                    isAddingHere ? (
                      <div style={{ display: "flex", flexDirection: "column", gap: 4, minWidth: 200 }}>
                        <input
                          autoFocus
                          value={colSearch}
                          onChange={(e) => setColSearch(e.target.value)}
                          placeholder="Search columns…"
                          style={{ background: "#0f172a", border: `1px solid ${rule.color}55`,
                            borderRadius: 5, padding: "4px 8px", color: "#e2e8f0", fontSize: 11 }}
                        />
                        <div style={{ background: "#0f172a", border: `1px solid ${rule.color}33`,
                          borderRadius: 5, maxHeight: 140, overflowY: "auto" }}>
                          {avail.length === 0
                            ? <div style={{ padding: "8px 10px", fontSize: 11, color: "#334155", fontStyle: "italic" }}>No columns available</div>
                            : avail.map((c) => (
                              <button key={c} onClick={() => { onAssignColumn(rule.id, c); setColSearch(""); }}
                                style={{ display: "block", width: "100%", textAlign: "left",
                                  background: "none", border: "none", cursor: "pointer",
                                  padding: "5px 10px", color: "#e2e8f0", fontSize: 11,
                                  borderBottom: "1px solid #1e293b" }}
                                onMouseEnter={(e) => e.currentTarget.style.background = "#1e293b"}
                                onMouseLeave={(e) => e.currentTarget.style.background = "none"}
                              >{c}</button>
                            ))
                          }
                        </div>
                        <button style={{ ...S.btn(), fontSize: 10 }}
                          onClick={() => { setAddingFor(null); setColSearch(""); }}>Cancel</button>
                      </div>
                    ) : (
                      <button
                        style={{ ...S.btn(), fontSize: 10, padding: "2px 8px",
                          borderColor: rule.color + "44", color: rule.color }}
                        onClick={() => { setAddingFor(rule.id); setColSearch(""); }}
                        title="Pin a column to this rule"
                      >+ Assign column</button>
                    )
                  )}
                </div>
              </div>
            );
          })}

          <button style={S.btn("danger")} onClick={onResetRules}>↺ Reset to defaults</button>
        </div>
      )}
    </div>
  );
}

// ── File helpers ──────────────────────────────────────────────────────────────
const ACCEPT = ".xlsx,.xls,.ods,.csv,.json";
const isJson = (n) => /\.json$/i.test(n);
const isXlsx = (n) => /\.(xlsx|xls|ods|csv)$/i.test(n);

// ── Main component ────────────────────────────────────────────────────────────
export default function MetadataTab() {
  const xlsx = useXLSX();

  const [sheets,          setSheets]          = useState(null);
  const [activeSheet,     setActiveSheet]      = useState(null);
  const [fileName,        setFileName]         = useState(null);
  const [fileType,        setFileType]         = useState(null);
  const [dragOver,        setDragOver]         = useState(false);
  const [error,           setError]            = useState(null);
  const [dirtySet,        setDirtySet]         = useState(new Set());
  const [savedFlash,      setSavedFlash]       = useState(false);
  const [validationMap,   setValidationMap]    = useState(new Map());
  const [customRules,     setCustomRules]      = useState(() =>
    DEFAULT_RULES.map((r) => ({ ...r, options: r.options ? [...r.options] : undefined }))
  );
  // columnAssignments: { [ruleId]: Set<colName> } — explicit column→rule pinning
  const [columnAssignments, setColumnAssignments] = useState({});
  const inputRef = useRef();

  // All unique column names across all sheets
  const allColumns = useMemo(() => {
    if (!sheets) return [];
    const cols = new Set();
    Object.values(sheets).forEach(({ headers }) => headers.forEach((h) => cols.add(h)));
    return [...cols].sort();
  }, [sheets]);

  // Revalidate whenever data, rules, or assignments change
  useEffect(() => {
    if (!sheets) return;
    const map = new Map();
    Object.entries(sheets).forEach(([sname, { headers, rows }]) => {
      headers.forEach((h) => {
        const rule = matchRule(h, customRules, columnAssignments);
        if (!rule) return;
        rows.forEach((row, ri) => {
          const val = row[h];
          if (val == null || val === "") return;
          const msg = validateValue(val, rule);
          if (msg) map.set(`${sname}:${ri}:${h}`, msg);
        });
      });
    });
    setValidationMap(map);
  }, [sheets, customRules, columnAssignments]);

  const applyParsed = useCallback((parsed, name, type) => {
    setSheets(parsed); setActiveSheet(Object.keys(parsed)[0] || null);
    setFileName(name); setFileType(type); setError(null);
    setDirtySet(new Set()); setColumnAssignments({});
  }, []);

  const parseFile = useCallback((file) => {
    if (!file) return;
    setError(null);
    if (isJson(file.name)) {
      const r = new FileReader();
      r.onload = (e) => {
        try {
          const p = parseJsonMapping(e.target.result);
          if (!Object.keys(p).length) throw new Error("No sheets found.");
          applyParsed(p, file.name, "json");
        } catch (err) { setError("JSON parse error: " + err.message); }
      };
      r.readAsText(file); return;
    }
    if (isXlsx(file.name)) {
      if (!xlsx) { setError("Parser not ready — try again in a moment."); return; }
      const r = new FileReader();
      r.onload = (e) => {
        try {
          const wb = xlsx.read(e.target.result, { type: "array" });
          const parsed = {};
          wb.SheetNames.forEach((name) => {
            const ws = wb.Sheets[name];
            const raw = xlsx.utils.sheet_to_json(ws, { header: 1, defval: null });
            if (!raw.length) { parsed[name] = { headers: [], rows: [] }; return; }
            const headers = (raw[0] || []).map((h) => (h != null ? String(h) : ""));
            const rows    = raw.slice(1).map((r2) => {
              const obj = {};
              headers.forEach((h, i) => { obj[h] = r2[i] ?? null; });
              return obj;
            });
            parsed[name] = { headers, rows };
          });
          applyParsed(parsed, file.name, "xlsx");
        } catch (err) { setError("Parse error: " + err.message); }
      };
      r.readAsArrayBuffer(file); return;
    }
    setError("Unsupported file type. Upload XLSX, XLS, ODS, CSV, or JSON.");
  }, [xlsx, applyParsed]);

  const onDrop        = useCallback((e) => { e.preventDefault(); setDragOver(false); parseFile(e.dataTransfer.files[0]); }, [parseFile]);
  const onInputChange = useCallback((e) => { parseFile(e.target.files[0]); e.target.value = ""; }, [parseFile]);
  const clear         = () => {
    setSheets(null); setActiveSheet(null); setFileName(null); setFileType(null);
    setError(null); setDirtySet(new Set()); setValidationMap(new Map()); setColumnAssignments({});
  };

  const onCellChange = useCallback((rowIdx, colKey, newVal) => {
    setSheets((prev) => {
      const sheet = prev[activeSheet];
      return { ...prev, [activeSheet]: { ...sheet, rows: sheet.rows.map((r, i) => i === rowIdx ? { ...r, [colKey]: newVal } : r) } };
    });
    setDirtySet((prev) => { const n = new Set(prev); n.add(`${activeSheet}:${rowIdx}:${colKey}`); return n; });
  }, [activeSheet]);

  const onAddRow = useCallback(() => {
    setSheets((prev) => {
      const sheet = prev[activeSheet];
      const empty = Object.fromEntries(sheet.headers.map((h) => [h, null]));
      return { ...prev, [activeSheet]: { ...sheet, rows: [...sheet.rows, empty] } };
    });
  }, [activeSheet]);

  const onDeleteRow = useCallback((rowIdx) => {
    setSheets((prev) => {
      const sheet = prev[activeSheet];
      return { ...prev, [activeSheet]: { ...sheet, rows: sheet.rows.filter((_, i) => i !== rowIdx) } };
    });
    setDirtySet((prev) => {
      const next = new Set();
      prev.forEach((key) => {
        const parts = key.split(":");
        const sn = parts[0], ri = parseInt(parts[1], 10), col = parts.slice(2).join(":");
        if (sn !== activeSheet) { next.add(key); return; }
        if (ri < rowIdx) next.add(key);
        if (ri > rowIdx) next.add(`${sn}:${ri - 1}:${col}`);
      });
      return next;
    });
  }, [activeSheet]);

  const onSave = useCallback(() => {
    setDirtySet(new Set()); setSavedFlash(true);
    setTimeout(() => setSavedFlash(false), 2000);
    if (fileType === "json") exportJson(sheets, fileName);
    else if (xlsx)           exportXlsx(xlsx, sheets, fileName);
  }, [sheets, fileName, fileType, xlsx]);

  const onUpdateRule    = useCallback((id, patch) => setCustomRules((prev) => prev.map((r) => r.id === id ? { ...r, ...patch } : r)), []);
  const onResetRules    = useCallback(() => {
    setCustomRules(DEFAULT_RULES.map((r) => ({ ...r, options: r.options ? [...r.options] : undefined })));
    setColumnAssignments({});
  }, []);
  const onAssignColumn  = useCallback((ruleId, colName) => {
    setColumnAssignments((prev) => {
      const next = { ...prev };
      next[ruleId] = new Set([...(prev[ruleId] || []), colName]);
      return next;
    });
  }, []);
  const onUnassignColumn = useCallback((ruleId, colName) => {
    setColumnAssignments((prev) => {
      const next = { ...prev };
      const s = new Set(prev[ruleId] || []);
      s.delete(colName); next[ruleId] = s;
      return next;
    });
  }, []);

  const sheetNames   = sheets ? Object.keys(sheets) : [];
  const currentSheet = sheets && activeSheet ? sheets[activeSheet] : null;
  const hasDirty     = dirtySet.size > 0;
  const formatColor  = fileType === "json" ? "#f59e0b" : "#22c55e";
  const formatLabel  = fileType === "json" ? "JSON" : fileType ? "XLSX" : "";
  const totalErrors  = validationMap.size;

  return (
    <div style={S.container}>
      <div style={{ fontSize: 14, fontWeight: 700, color: "#e2e8f0" }}>Data Mapping Specification</div>

      {!sheets ? (
        <div style={S.card}>
          <div style={S.dropzone(dragOver)} onClick={() => inputRef.current?.click()}
            onDragOver={(e) => { e.preventDefault(); setDragOver(true); }}
            onDragLeave={() => setDragOver(false)} onDrop={onDrop}>
            <div style={S.uploadIcon}>⬆</div>
            <div style={S.uploadTitle}>Upload XLSX / XLS / CSV / JSON</div>
            <div style={S.uploadSub}>{xlsx ? "Drag & drop or click to browse" : "Loading parser…"}</div>
            <div style={{ display: "flex", gap: 8, justifyContent: "center", marginTop: 14 }}>
              <span style={S.badge("#22c55e")}>XLSX / XLS</span>
              <span style={S.badge("#22c55e")}>CSV</span>
              <span style={S.badge("#f59e0b")}>JSON</span>
            </div>
          </div>
          <input ref={inputRef} type="file" accept={ACCEPT} style={{ display: "none" }} onChange={onInputChange} />
          {error && <div style={{ ...S.errorBox, marginTop: 10 }}>{error}</div>}
        </div>
      ) : (
        <>
          {/* File info bar */}
          <div style={S.fileInfo}>
            <span style={{ fontSize: 16 }}>{fileType === "json" ? "{ }" : "📄"}</span>
            <span style={S.fileName}>{fileName}</span>
            {formatLabel && <span style={S.badge(formatColor)}>{formatLabel}</span>}
            <span style={{ fontSize: 11, color: "#475569" }}>{sheetNames.length} sheet{sheetNames.length !== 1 ? "s" : ""}</span>
            {totalErrors > 0 && <span style={S.invalidBadge}>⚠ {totalErrors} error{totalErrors !== 1 ? "s" : ""}</span>}
            {hasDirty && (
              <button onClick={onSave} style={{
                background: "linear-gradient(135deg,#16a34a,#15803d)", border: "1px solid #22c55e88",
                borderRadius: 7, padding: "5px 14px", cursor: "pointer", color: "#fff",
                fontSize: 12, fontWeight: 700, boxShadow: "0 0 10px #16a34a55",
                display: "flex", alignItems: "center", gap: 6,
              }} title="Commit edits and download">✓ Save Changes</button>
            )}
            {savedFlash && !hasDirty && <span style={{ fontSize: 11, color: "#22c55e", fontWeight: 600 }}>✓ Saved</span>}
            {hasDirty && <span style={S.dirtyBadge}>✎ {dirtySet.size} edit{dirtySet.size !== 1 ? "s" : ""}</span>}
            <button style={S.btn("success")} onClick={() => exportJson(sheets, fileName)}>↓ JSON</button>
            {xlsx && <button style={S.btn("primary")} onClick={() => exportXlsx(xlsx, sheets, fileName)}>↓ XLSX</button>}
            <button style={S.btn()} onClick={clear}>Clear</button>
          </div>

          {/* Hint */}
          <div style={{ fontSize: 11, color: "#334155", fontStyle: "italic" }}>
            Click any cell to edit · Enter or click away to confirm · Esc to cancel
            {hasDirty && <span style={{ color: "#fbbf2488", marginLeft: 8 }}>· <strong style={{ color: "#fbbf24" }}>Save Changes</strong> to commit &amp; download</span>}
          </div>

          {/* Rules panel */}
          <RulesPanel
            customRules={customRules}
            columnAssignments={columnAssignments}
            allColumns={allColumns}
            onUpdateRule={onUpdateRule}
            onResetRules={onResetRules}
            onAssignColumn={onAssignColumn}
            onUnassignColumn={onUnassignColumn}
          />

          {/* Sheet tabs */}
          <div style={S.sheetNav}>
            {sheetNames.map((name) => {
              const sheetDirty = [...dirtySet].some((k) => k.startsWith(name + ":"));
              const sheetErrs  = [...validationMap.keys()].filter((k) => k.startsWith(name + ":")).length;
              return (
                <button key={name} style={S.sheetBtn(activeSheet === name)} onClick={() => setActiveSheet(name)}>
                  {name}
                  {sheetDirty && <span style={{ marginLeft: 5, color: "#fbbf24", fontSize: 10 }}>✎</span>}
                  {sheetErrs > 0 && <span style={{ marginLeft: 4, color: "#ef4444", fontSize: 10 }}>⚠{sheetErrs}</span>}
                </button>
              );
            })}
          </div>

          {/* Editable table */}
          {currentSheet && (
            <SheetTable
              sheetName={activeSheet}
              sheet={currentSheet}
              dirtySet={dirtySet}
              validationMap={validationMap}
              onCellChange={onCellChange}
              onAddRow={onAddRow}
              onDeleteRow={onDeleteRow}
              customRules={customRules}
              columnAssignments={columnAssignments}
            />
          )}
        </>
      )}
    </div>
  );
}
