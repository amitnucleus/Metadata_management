import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { useTheme } from "./ThemeContext";

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
function exportAuditXlsx(xlsx, auditLog, fileName) {
  const wb = xlsx.utils.book_new();
  const headers = ["Timestamp", "Sheet", "Row", "Column", "Old Value", "New Value"];
  const rows = auditLog.map((e) => [e.ts, e.sheet, e.rowIdx + 1, e.col, e.oldVal ?? "", e.newVal ?? ""]);
  xlsx.utils.book_append_sheet(wb, xlsx.utils.aoa_to_sheet([headers, ...rows]), "Audit Log");
  xlsx.writeFile(wb, fileName.replace(/\.[^.]+$/, "") + "_audit.xlsx");
}
function triggerDownload(blob, name) {
  const url = URL.createObjectURL(blob);
  const a = Object.assign(document.createElement("a"), { href: url, download: name });
  document.body.appendChild(a); a.click();
  document.body.removeChild(a); URL.revokeObjectURL(url);
}

// ── Python API parser ─────────────────────────────────────────────────────────
async function parseViaApi(file) {
  const fd = new FormData();
  fd.append("file", file);
  const res = await fetch("/api/parse", { method: "POST", body: fd });
  if (!res.ok) {
    const err = await res.json().catch(() => ({}));
    throw new Error(err.detail || err.error || `API error ${res.status}`);
  }
  return res.json();
}

// ── File helpers ──────────────────────────────────────────────────────────────
const ACCEPT = ".xlsx,.xls,.ods,.csv,.json";
const isJson = (n) => /\.json$/i.test(n);
const isXlsx = (n) => /\.(xlsx|xls|ods|csv)$/i.test(n);

// ── Data profiling ────────────────────────────────────────────────────────────
function profileColumn(rows, col) {
  const vals = rows.map((r) => r[col]);
  const total = vals.length;
  const nullCount = vals.filter((v) => v == null || v === "").length;
  const nonNull = vals.filter((v) => v != null && v !== "").map((v) => String(v));
  const unique = new Set(nonNull).size;
  const freq = {};
  nonNull.forEach((v) => { freq[v] = (freq[v] || 0) + 1; });
  const topEntries = Object.entries(freq).sort((a, b) => b[1] - a[1]).slice(0, 3);
  const nums = nonNull.map(Number).filter((n) => !isNaN(n));
  const min = nums.length ? Math.min(...nums) : null;
  const max = nums.length ? Math.max(...nums) : null;
  return { total, nullCount, nullPct: total ? Math.round((nullCount / total) * 100) : 0, unique, topEntries, min, max, numericCount: nums.length };
}

// ── File comparison helpers ───────────────────────────────────────────────────
// Parse any supported file entirely client-side (xlsx obj passed in)
function parseFileClientSide(file, xlsxLib) {
  return new Promise((resolve, reject) => {
    const type = isJson(file.name) ? "json" : isXlsx(file.name) ? "xlsx" : null;
    if (!type) { reject(new Error("Unsupported file type")); return; }

    if (type === "json") {
      const r = new FileReader();
      r.onload = (e) => {
        try { resolve(parseJsonMapping(e.target.result)); }
        catch (err) { reject(err); }
      };
      r.onerror = reject;
      r.readAsText(file);
    } else {
      if (!xlsxLib) { reject(new Error("Parser not ready")); return; }
      const r = new FileReader();
      r.onload = (e) => {
        try {
          const wb = xlsxLib.read(e.target.result, { type: "array" });
          const parsed = {};
          wb.SheetNames.forEach((name) => {
            const ws  = wb.Sheets[name];
            const raw = xlsxLib.utils.sheet_to_json(ws, { header: 1, defval: null });
            if (!raw.length) { parsed[name] = { headers: [], rows: [] }; return; }
            const headers = (raw[0] || []).map((h) => (h != null ? String(h) : ""));
            const rows    = raw.slice(1).map((r2) => {
              const obj = {};
              headers.forEach((h, i) => { obj[h] = r2[i] ?? null; });
              return obj;
            });
            parsed[name] = { headers, rows };
          });
          resolve(parsed);
        } catch (err) { reject(err); }
      };
      r.onerror = reject;
      r.readAsArrayBuffer(file);
    }
  });
}

// Compute row-level diff between two sheet objects.
// Strategy: match rows by position (row index). Returns array of diff entries.
function diffSheets(sheetA, sheetB) {
  const allHeaders = [...new Set([...(sheetA?.headers || []), ...(sheetB?.headers || [])])];
  const rowsA = sheetA?.rows || [];
  const rowsB = sheetB?.rows || [];
  const maxLen = Math.max(rowsA.length, rowsB.length);
  const results = [];

  for (let i = 0; i < maxLen; i++) {
    const rA = rowsA[i] || null;
    const rB = rowsB[i] || null;
    if (!rA) { results.push({ type: "added",   idx: i, rowA: null, rowB: rB, diffs: [] }); continue; }
    if (!rB) { results.push({ type: "removed", idx: i, rowA: rA,  rowB: null, diffs: [] }); continue; }
    const cellDiffs = allHeaders.filter((h) => String(rA[h] ?? "") !== String(rB[h] ?? ""));
    results.push({ type: cellDiffs.length ? "changed" : "same", idx: i, rowA: rA, rowB: rB, diffs: cellDiffs });
  }
  return { allHeaders, results };
}

// ── File comparison modal ─────────────────────────────────────────────────────
function FileCmpModal({ xlsxLib, onClose }) {
  const { theme } = useTheme();
  const S = useMemo(() => makeStyles(theme), [theme]);

  const [fileA,       setFileA]       = useState(null);  // { name, sheets }
  const [fileB,       setFileB]       = useState(null);
  const [loadingA,    setLoadingA]    = useState(false);
  const [loadingB,    setLoadingB]    = useState(false);
  const [errA,        setErrA]        = useState(null);
  const [errB,        setErrB]        = useState(null);
  const [activeSheet, setActiveSheet] = useState(null);
  const [showSame,    setShowSame]    = useState(false);
  const [filterCol,   setFilterCol]   = useState("");
  const inputARef = useRef();
  const inputBRef = useRef();

  async function loadFile(file, side) {
    const setLoading = side === "A" ? setLoadingA : setLoadingB;
    const setErr     = side === "A" ? setErrA     : setErrB;
    const setFile    = side === "A" ? setFileA    : setFileB;
    setLoading(true); setErr(null);
    try {
      // Try API first, fall back to client-side
      let parsed;
      try {
        parsed = await parseViaApi(file);
        if (!Object.keys(parsed).length) throw new Error("empty");
      } catch {
        parsed = await parseFileClientSide(file, xlsxLib);
      }
      setFile({ name: file.name, sheets: parsed });
      // Auto-select first common sheet when both are loaded
      setActiveSheet((prev) => prev || Object.keys(parsed)[0] || null);
    } catch (e) {
      setErr(e.message || "Parse error");
    } finally {
      setLoading(false);
    }
  }

  function onDropSide(e, side) {
    e.preventDefault();
    const file = e.dataTransfer.files[0];
    if (file) loadFile(file, side);
  }

  // All sheet names present in either file
  const allSheetNames = useMemo(() => {
    const names = new Set([
      ...Object.keys(fileA?.sheets || {}),
      ...Object.keys(fileB?.sheets || {}),
    ]);
    return [...names];
  }, [fileA, fileB]);

  // Diff for the active sheet
  const { allHeaders, results } = useMemo(() => {
    if (!fileA && !fileB) return { allHeaders: [], results: [] };
    const sA = fileA?.sheets?.[activeSheet] || null;
    const sB = fileB?.sheets?.[activeSheet] || null;
    return diffSheets(sA, sB);
  }, [fileA, fileB, activeSheet]);

  const visibleResults = useMemo(() => {
    return results.filter((r) => {
      if (!showSame && r.type === "same") return false;
      if (filterCol && !r.diffs.includes(filterCol)) return false;
      return true;
    });
  }, [results, showSame, filterCol]);

  const stats = useMemo(() => ({
    added:   results.filter((r) => r.type === "added").length,
    removed: results.filter((r) => r.type === "removed").length,
    changed: results.filter((r) => r.type === "changed").length,
    same:    results.filter((r) => r.type === "same").length,
  }), [results]);

  const changedCols = useMemo(() => {
    const s = new Set();
    results.forEach((r) => r.diffs.forEach((c) => s.add(c)));
    return [...s].sort();
  }, [results]);

  function rowBg(type) {
    if (type === "added")   return `${theme.btnSuccessText}11`;
    if (type === "removed") return `${theme.invalidColor}11`;
    if (type === "changed") return `${theme.dirtyColor}11`;
    return undefined;
  }
  function cellColor(type, col, diffs) {
    if (type === "changed" && diffs.includes(col)) return theme.dirtyColor;
    if (type === "added")   return "#22c55e";
    if (type === "removed") return theme.invalidColor;
    return theme.cellText;
  }

  // Upload slot UI
  function UploadSlot({ side, fileInfo, loading, err }) {
    const ref = side === "A" ? inputARef : inputBRef;
    const label = side === "A" ? "File A" : "File B";
    const color = side === "A" ? theme.accentText : "#f472b6";
    const [drag, setDrag] = useState(false);

    return (
      <div style={{ flex: 1, minWidth: 0, display: "flex", flexDirection: "column", gap: 6 }}>
        <div style={{ fontSize: 11, fontWeight: 700, color, textTransform: "uppercase", letterSpacing: 0.5 }}>
          {label} {fileInfo && <span style={{ fontWeight: 400, color: theme.mutedText, textTransform: "none" }}>— {fileInfo.name}</span>}
        </div>
        <div
          onClick={() => ref.current?.click()}
          onDragOver={(e) => { e.preventDefault(); setDrag(true); }}
          onDragLeave={() => setDrag(false)}
          onDrop={(e) => { setDrag(false); onDropSide(e, side); }}
          style={{
            border: `2px dashed ${drag ? color : theme.cardBorder}`,
            borderRadius: 8, padding: "18px 12px", textAlign: "center",
            cursor: loading ? "wait" : "pointer",
            background: drag ? `${color}11` : theme.surfaceBg,
            transition: "all 0.15s",
          }}
        >
          {loading
            ? <span style={{ fontSize: 12, color: theme.mutedText }}>Parsing…</span>
            : fileInfo
              ? <span style={{ fontSize: 12, color }}>
                  {Object.keys(fileInfo.sheets).length} sheet{Object.keys(fileInfo.sheets).length !== 1 ? "s" : ""}
                  &nbsp;·&nbsp;{Object.values(fileInfo.sheets).reduce((s, sh) => s + sh.rows.length, 0)} rows
                  &nbsp;<span style={{ color: theme.mutedText, fontWeight: 400 }}>— click to replace</span>
                </span>
              : <span style={{ fontSize: 12, color: theme.mutedText }}>Drag or click to upload</span>
          }
        </div>
        <input ref={ref} type="file" accept={ACCEPT} style={{ display: "none" }}
          onChange={(e) => { if (e.target.files[0]) loadFile(e.target.files[0], side); e.target.value = ""; }} />
        {err && <div style={{ ...S.errorBox, padding: "5px 10px" }}>{err}</div>}
      </div>
    );
  }

  return (
    <div style={{ position: "fixed", inset: 0, background: "#00000099", zIndex: 1000,
      display: "flex", alignItems: "flex-start", justifyContent: "center",
      paddingTop: 32, paddingBottom: 32, overflowY: "auto" }}
      onClick={onClose}>
      <div style={{ background: theme.cardBg, border: `1px solid ${theme.cardBorder}`, borderRadius: 14,
        padding: "24px", width: "96vw", maxWidth: 1200,
        display: "flex", flexDirection: "column", gap: 18, fontFamily: theme.font }}
        onClick={(e) => e.stopPropagation()}>

        {/* Header */}
        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 12 }}>
          <span style={{ fontSize: 14, fontWeight: 700, color: theme.pageText }}>⇄ File Comparison</span>
          <button onClick={onClose} style={{ ...S.btn(), fontSize: 13, padding: "3px 10px" }}>✕ Close</button>
        </div>

        {/* Upload slots */}
        <div style={{ display: "flex", gap: 16, flexWrap: "wrap" }}>
          <UploadSlot side="A" fileInfo={fileA} loading={loadingA} err={errA} />
          <div style={{ display: "flex", alignItems: "center", color: theme.mutedText, fontSize: 18, flexShrink: 0 }}>⇄</div>
          <UploadSlot side="B" fileInfo={fileB} loading={loadingB} err={errB} />
        </div>

        {(fileA || fileB) && (
          <>
            {/* Sheet tabs */}
            {allSheetNames.length > 0 && (
              <div style={S.sheetNav}>
                {allSheetNames.map((name) => {
                  const inA = !!fileA?.sheets?.[name];
                  const inB = !!fileB?.sheets?.[name];
                  const label = !inA ? "+ B only" : !inB ? "− A only" : null;
                  const labelColor = !inA ? "#22c55e" : !inB ? theme.invalidColor : null;
                  return (
                    <button key={name} style={S.sheetBtn(activeSheet === name)}
                      onClick={() => setActiveSheet(name)}>
                      {name}
                      {label && <span style={{ marginLeft: 5, color: labelColor, fontSize: 9 }}>{label}</span>}
                    </button>
                  );
                })}
              </div>
            )}

            {/* Summary stats */}
            {fileA && fileB && activeSheet && (
              <div style={{ display: "flex", gap: 10, flexWrap: "wrap", alignItems: "center" }}>
                {stats.added   > 0 && <span style={S.badge("#22c55e")}>+{stats.added} added</span>}
                {stats.removed > 0 && <span style={S.invalidBadge}>−{stats.removed} removed</span>}
                {stats.changed > 0 && <span style={S.dirtyBadge}>≠{stats.changed} changed</span>}
                {stats.same    > 0 && <span style={S.badge(theme.mutedText)}>{stats.same} identical</span>}

                {/* Filter by column */}
                {changedCols.length > 0 && (
                  <select value={filterCol} onChange={(e) => setFilterCol(e.target.value)}
                    style={{ background: theme.inputBg, border: `1px solid ${theme.cardBorder}`,
                      borderRadius: 5, color: theme.pageText, fontSize: 11, padding: "3px 8px",
                      fontFamily: theme.font, marginLeft: 4 }}>
                    <option value="">All columns</option>
                    {changedCols.map((c) => <option key={c} value={c}>{c}</option>)}
                  </select>
                )}

                <label style={{ display: "flex", alignItems: "center", gap: 5, fontSize: 11,
                  color: theme.mutedText, cursor: "pointer", marginLeft: 4 }}>
                  <input type="checkbox" checked={showSame} onChange={(e) => setShowSame(e.target.checked)}
                    style={{ accentColor: theme.accent }} />
                  Show identical rows
                </label>

                <span style={{ fontSize: 11, color: theme.dimText, marginLeft: "auto" }}>
                  Showing {visibleResults.length} / {results.length} rows
                </span>
              </div>
            )}

            {/* Diff table */}
            {fileA && fileB && activeSheet && visibleResults.length > 0 && (
              <div style={{ overflowX: "auto", border: `1px solid ${theme.cardBorder}`, borderRadius: 8 }}>
                <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 11 }}>
                  <thead>
                    <tr style={{ background: theme.headerBg }}>
                      <th style={{ padding: "7px 10px", color: theme.mutedText, fontWeight: 700,
                        fontSize: 10, textTransform: "uppercase", letterSpacing: 0.5,
                        borderBottom: `2px solid ${theme.headerBorder}`, whiteSpace: "nowrap", width: 50 }}>Row</th>
                      <th style={{ padding: "7px 10px", color: theme.mutedText, fontWeight: 700,
                        fontSize: 10, textTransform: "uppercase", letterSpacing: 0.5,
                        borderBottom: `2px solid ${theme.headerBorder}`, width: 70 }}>Status</th>
                      {allHeaders.map((h) => (
                        <th key={h} colSpan={2} style={{
                          padding: "7px 10px", color: changedCols.includes(h) ? theme.dirtyColor : theme.headerText,
                          fontWeight: 700, fontSize: 10, textTransform: "uppercase", letterSpacing: 0.5,
                          borderBottom: `2px solid ${changedCols.includes(h) ? theme.dirtyColor + "55" : theme.headerBorder}`,
                          textAlign: "center", whiteSpace: "nowrap", borderLeft: `1px solid ${theme.cardBorder}`,
                        }}>{h}</th>
                      ))}
                    </tr>
                    {/* Sub-header: A / B labels */}
                    <tr style={{ background: theme.surfaceBg }}>
                      <td colSpan={2} style={{ borderBottom: `1px solid ${theme.cardBorder}` }} />
                      {allHeaders.map((h) => (
                        <>
                          <td key={h + "_a"} style={{ padding: "3px 8px", fontSize: 9, fontWeight: 700,
                            color: theme.accentText, borderBottom: `1px solid ${theme.cardBorder}`,
                            borderLeft: `1px solid ${theme.cardBorder}`, textAlign: "center" }}>A</td>
                          <td key={h + "_b"} style={{ padding: "3px 8px", fontSize: 9, fontWeight: 700,
                            color: "#f472b6", borderBottom: `1px solid ${theme.cardBorder}`, textAlign: "center" }}>B</td>
                        </>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {visibleResults.map((r) => (
                      <tr key={r.idx} style={{ background: rowBg(r.type), borderTop: `1px solid ${theme.rowBorder}` }}>
                        <td style={{ padding: "5px 10px", color: theme.mutedText, fontSize: 11 }}>{r.idx + 1}</td>
                        <td style={{ padding: "5px 10px", fontSize: 10, fontWeight: 700, whiteSpace: "nowrap",
                          color: r.type === "added" ? "#22c55e" : r.type === "removed" ? theme.invalidColor
                            : r.type === "changed" ? theme.dirtyColor : theme.dimText }}>
                          {r.type === "added" ? "+ added" : r.type === "removed" ? "− removed"
                            : r.type === "changed" ? `≠ ${r.diffs.length} col${r.diffs.length !== 1 ? "s" : ""}` : "= same"}
                        </td>
                        {allHeaders.map((h) => {
                          const valA = r.rowA?.[h] ?? null;
                          const valB = r.rowB?.[h] ?? null;
                          const changed = r.diffs.includes(h);
                          return (
                            <>
                              <td key={h + "_a"} style={{ padding: "5px 8px", borderLeft: `1px solid ${theme.cardBorder}`,
                                background: changed && r.type === "changed" ? `${theme.invalidColor}18` : undefined,
                                color: r.type === "removed" ? theme.invalidColor : changed ? theme.invalidColor : theme.cellText,
                                maxWidth: 160, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}
                                title={valA != null ? String(valA) : ""}>
                                {valA != null ? String(valA) : <span style={{ color: theme.dimText }}>—</span>}
                              </td>
                              <td key={h + "_b"} style={{ padding: "5px 8px",
                                background: changed && r.type === "changed" ? `${"#22c55e"}18` : undefined,
                                color: r.type === "added" ? "#22c55e" : changed ? "#22c55e" : theme.cellText,
                                maxWidth: 160, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}
                                title={valB != null ? String(valB) : ""}>
                                {valB != null ? String(valB) : <span style={{ color: theme.dimText }}>—</span>}
                              </td>
                            </>
                          );
                        })}
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )}

            {/* Only-in-A / Only-in-B sheets notice */}
            {activeSheet && (!fileA?.sheets?.[activeSheet] || !fileB?.sheets?.[activeSheet]) && (
              <div style={{ ...S.errorBox, background: `${theme.dirtyColor}11`, borderColor: `${theme.dirtyColor}33`,
                color: theme.dirtyColor }}>
                Sheet "{activeSheet}" exists only in {!fileA?.sheets?.[activeSheet] ? "File B" : "File A"} — no comparison possible.
              </div>
            )}

            {fileA && fileB && activeSheet && visibleResults.length === 0 && (
              <div style={{ padding: "24px", textAlign: "center", color: theme.mutedText,
                fontSize: 13, fontStyle: "italic", border: `1px solid ${theme.cardBorder}`, borderRadius: 8 }}>
                {results.length === 0 ? "Both sheets are empty." : "No differences found (all rows are identical)."}
              </div>
            )}
          </>
        )}

        {!fileA && !fileB && (
          <div style={{ fontSize: 12, color: theme.mutedText, fontStyle: "italic", textAlign: "center", padding: "12px 0" }}>
            Upload two files above to compare them side by side.
          </div>
        )}
      </div>
    </div>
  );
}

// ── Undo/Redo history hook ────────────────────────────────────────────────────
function useHistory(initial) {
  const [state,  setState]  = useState(initial);
  const [past,   setPast]   = useState([]);
  const [future, setFuture] = useState([]);

  const set = useCallback((updater) => {
    setState((prev) => {
      const next = typeof updater === "function" ? updater(prev) : updater;
      setPast((p) => [...p.slice(-49), prev]);
      setFuture([]);
      return next;
    });
  }, []);

  const undo = useCallback(() => {
    setPast((p) => {
      if (!p.length) return p;
      const prev = p[p.length - 1];
      setState((cur) => { setFuture((f) => [cur, ...f]); return prev; });
      return p.slice(0, -1);
    });
  }, []);

  const redo = useCallback(() => {
    setFuture((f) => {
      if (!f.length) return f;
      const next = f[0];
      setState((cur) => { setPast((p) => [...p, cur]); return next; });
      return f.slice(1);
    });
  }, []);

  const reset = useCallback((next) => {
    setState(next); setPast([]); setFuture([]);
  }, []);

  return { state, set, undo, redo, reset, canUndo: past.length > 0, canRedo: future.length > 0 };
}

// ── Theme-aware style factory ─────────────────────────────────────────────────
function makeStyles(t) {
  return {
    container: { display: "flex", flexDirection: "column", gap: 20, fontFamily: t.font },
    card:      { background: t.cardBg, border: `1px solid ${t.cardBorder}`, borderRadius: 10, padding: "16px 20px" },
    dropzone: (active) => ({
      border: `2px dashed ${active ? t.accent : t.cardBorder}`,
      borderRadius: 12, padding: "40px 32px", textAlign: "center", cursor: "pointer",
      background: active ? t.dropzoneActiveBg : t.dropzoneBg,
      transition: "border-color 0.2s, background 0.2s",
    }),
    uploadIcon:  { fontSize: 32, marginBottom: 8, color: t.accent },
    uploadTitle: { fontSize: 15, fontWeight: 700, color: t.pageText, marginBottom: 4 },
    uploadSub:   { fontSize: 12, color: t.mutedText },
    badge: (color) => ({
      display: "inline-block", padding: "2px 8px", borderRadius: 4, fontSize: 10,
      fontWeight: 700, background: `${color}22`, border: `1px solid ${color}44`, color,
    }),
    fileInfo: {
      display: "flex", alignItems: "center", gap: 8, padding: "10px 16px",
      background: t.cardBg, border: `1px solid ${t.cardBorder}`, borderRadius: 8, flexWrap: "wrap",
    },
    fileName: { fontSize: 13, fontWeight: 600, color: t.accentText, flex: 1, minWidth: 0 },
    btn: (variant = "ghost") => ({
      background: variant === "primary" ? t.btnPrimaryBg : variant === "success" ? t.btnSuccessBg
                : variant === "danger"  ? t.btnDangerBg  : "none",
      border: `1px solid ${variant === "primary" ? t.btnPrimaryBorder : variant === "success" ? t.btnSuccessBorder
                           : variant === "danger" ? t.btnDangerBorder  : t.btnGhost}`,
      borderRadius: 6, padding: "4px 11px", cursor: "pointer",
      color: variant === "primary" ? t.btnPrimaryText : variant === "success" ? t.btnSuccessText
           : variant === "danger"  ? t.btnDangerText  : t.btnGhostText,
      fontSize: 11, fontWeight: 600, whiteSpace: "nowrap", fontFamily: t.font,
    }),
    sheetNav: { display: "flex", gap: 6, flexWrap: "wrap", borderBottom: `1px solid ${t.cardBorder}`, paddingBottom: 8 },
    sheetBtn: (active) => ({
      background: active ? t.tabActiveBg : "none",
      border: `1px solid ${active ? t.tabActiveBorder : t.cardBorder}`,
      borderRadius: 6, padding: "5px 12px", cursor: "pointer",
      color: active ? t.tabActiveText : t.tabInactiveText,
      fontSize: 12, fontWeight: active ? 700 : 400, whiteSpace: "nowrap", fontFamily: t.font,
    }),
    tableWrap: { overflowX: "auto", borderRadius: 8, border: `1px solid ${t.cardBorder}` },
    table:     { width: "100%", borderCollapse: "collapse", fontSize: 12 },
    th: (ruleColor) => ({
      background: t.headerBg, padding: "9px 12px", textAlign: "left",
      color: ruleColor || t.headerText, fontWeight: 700, fontSize: 11,
      textTransform: "uppercase", letterSpacing: 0.5, whiteSpace: "nowrap",
      borderBottom: `2px solid ${ruleColor ? ruleColor + "55" : t.headerBorder}`,
    }),
    thDelete: {
      background: t.headerBg, padding: "9px 8px", textAlign: "center",
      borderBottom: `2px solid ${t.headerBorder}`, width: 32,
    },
    td: (i, dirty, invalid) => ({
      padding: 0, fontSize: 12, borderBottom: `1px solid ${t.rowBorder}`,
      background: invalid ? t.rowInvalid : dirty ? t.rowDirty : i % 2 === 0 ? t.rowEven : t.rowOdd,
      transition: "background 0.15s",
      outline: invalid ? `1px solid ${t.invalidColor}55` : "none",
    }),
    errorBox:     { color: t.invalidColor, fontSize: 12, padding: "8px 12px", background: `${t.invalidColor}11`, borderRadius: 6, border: `1px solid ${t.invalidColor}33` },
    dirtyBadge:   { fontSize: 10, fontWeight: 700, padding: "2px 7px", borderRadius: 4, background: `${t.dirtyColor}22`, border: `1px solid ${t.dirtyColor}44`, color: t.dirtyColor },
    invalidBadge: { fontSize: 10, fontWeight: 700, padding: "2px 7px", borderRadius: 4, background: `${t.invalidColor}22`, border: `1px solid ${t.invalidColor}44`, color: t.invalidColor },
    inputBase:    { width: "100%", boxSizing: "border-box", background: t.inputBg, border: "none", color: t.pageText, fontSize: 12, padding: "7px 12px", fontFamily: t.font },
  };
}

// ── Editable cell (with comment support) ─────────────────────────────────────
function EditableCell({ value, rowIdx, colKey, dirty, rule, invalid, validationMsg, onChange, comment, onSetComment }) {
  const { theme } = useTheme();
  const S = useMemo(() => makeStyles(theme), [theme]);
  const [editing,        setEditing]        = useState(false);
  const [draft,          setDraft]          = useState("");
  const [commentEditing, setCommentEditing] = useState(false);
  const [commentDraft,   setCommentDraft]   = useState("");
  const inputRef    = useRef();
  const commentRef  = useRef();

  function start() { setDraft(value != null ? String(value) : ""); setEditing(true); }
  useEffect(() => { if (editing) inputRef.current?.focus(); }, [editing]);
  function commit() { setEditing(false); onChange(rowIdx, colKey, draft === "" ? null : draft); }
  function cancel() { setEditing(false); }

  function openComment(e) {
    e.stopPropagation();
    setCommentDraft(comment || "");
    setCommentEditing(true);
  }
  useEffect(() => { if (commentEditing) commentRef.current?.focus(); }, [commentEditing]);
  function saveComment() {
    setCommentEditing(false);
    onSetComment(rowIdx, colKey, commentDraft.trim() || null);
  }

  const displayVal = value != null ? String(value) : "";
  const isYesNo    = ["YES","NO"].includes(displayVal.toUpperCase());
  const yesNoColor = displayVal.toUpperCase() === "YES" ? "#22c55e" : theme.invalidColor;
  const outlineColor = rule ? rule.color : theme.accent;

  return (
    <td style={{ ...S.td(rowIdx, dirty, invalid), position: "relative" }}
        title={validationMsg || "Click to edit"}
        onClick={() => { if (!editing && !commentEditing) start(); }}>
      {/* Comment indicator dot */}
      {comment && !editing && !commentEditing && (
        <span
          onClick={openComment}
          title={comment}
          style={{ position: "absolute", top: 3, right: 3, width: 6, height: 6,
            borderRadius: "50%", background: "#f59e0b", cursor: "pointer", zIndex: 2 }}
        />
      )}
      {commentEditing ? (
        <div style={{ padding: "4px 8px" }} onClick={(e) => e.stopPropagation()}>
          <textarea
            ref={commentRef}
            value={commentDraft}
            onChange={(e) => setCommentDraft(e.target.value)}
            onKeyDown={(e) => { if (e.key === "Escape") setCommentEditing(false); if (e.key === "Enter" && e.ctrlKey) saveComment(); }}
            placeholder="Add a note… (Ctrl+Enter to save)"
            style={{ width: "100%", boxSizing: "border-box", minHeight: 60, background: theme.inputBg,
              border: `1px solid ${theme.accentBorder}`, borderRadius: 4, color: theme.pageText,
              fontSize: 11, padding: "4px 6px", fontFamily: theme.font, resize: "vertical" }}
          />
          <div style={{ display: "flex", gap: 4, marginTop: 4 }}>
            <button style={{ ...S.btn("primary"), fontSize: 10, padding: "2px 8px" }} onClick={saveComment}>Save</button>
            <button style={{ ...S.btn(), fontSize: 10, padding: "2px 8px" }} onClick={() => setCommentEditing(false)}>Cancel</button>
            {comment && <button style={{ ...S.btn("danger"), fontSize: 10, padding: "2px 8px" }}
              onClick={() => { setCommentEditing(false); onSetComment(rowIdx, colKey, null); }}>Remove</button>}
          </div>
        </div>
      ) : editing ? (
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
          color: invalid ? theme.invalidColor : dirty ? theme.dirtyColor : theme.cellText }}>
          {isYesNo && !dirty && !invalid
            ? <span style={S.badge(yesNoColor)}>{displayVal}</span>
            : displayVal || <span style={{ color: theme.placeholder }}>—</span>}
          {dirty && !invalid && <span style={{ fontSize: 9, color: `${theme.dirtyColor}88` }}>✎</span>}
          {invalid && <span title={validationMsg} style={{ fontSize: 9, color: theme.invalidColor, marginLeft: "auto", cursor: "help" }}>⚠</span>}
          {/* Add comment button on hover via CSS trick workaround (inline) */}
          {!comment && (
            <button
              onClick={openComment}
              title="Add comment"
              style={{ background: "none", border: "none", cursor: "pointer", color: theme.dimText,
                fontSize: 9, padding: "0 2px", marginLeft: "auto", opacity: 0,
                transition: "opacity 0.15s", lineHeight: 1 }}
              onMouseEnter={(e) => e.currentTarget.style.opacity = 1}
              onMouseLeave={(e) => e.currentTarget.style.opacity = 0}
            >💬</button>
          )}
        </div>
      )}
    </td>
  );
}

// ── Column header (rename + delete + profile) ─────────────────────────────────
function ColHeader({ name, rule, source, onRename, onDeleteCol, profile, showProfile }) {
  const { theme } = useTheme();
  const S = useMemo(() => makeStyles(theme), [theme]);
  const [tip,      setTip]      = useState(false);
  const [renaming, setRenaming] = useState(false);
  const [draft,    setDraft]    = useState("");
  const inputRef = useRef();
  const srcLabel = source === "explicit" ? "explicit" : source === "pattern" ? "auto" : null;
  const srcColor = source === "explicit" ? theme.dirtyColor : theme.mutedText;

  function startRename() { setDraft(name); setRenaming(true); }
  useEffect(() => { if (renaming) inputRef.current?.focus(); }, [renaming]);
  function commitRename() {
    setRenaming(false);
    const trimmed = draft.trim();
    if (trimmed && trimmed !== name) onRename(name, trimmed);
  }

  return (
    <th style={S.th(rule?.color)}>
      <div style={{ display: "flex", alignItems: "center", gap: 5 }}>
        {renaming ? (
          <input
            ref={inputRef}
            value={draft}
            onChange={(e) => setDraft(e.target.value)}
            onBlur={commitRename}
            onKeyDown={(e) => { if (e.key === "Enter") commitRename(); if (e.key === "Escape") setRenaming(false); }}
            style={{ background: theme.inputBg, border: `1px solid ${theme.accent}`, borderRadius: 4,
              color: theme.pageText, fontSize: 11, padding: "2px 6px", width: 100, fontFamily: theme.font }}
          />
        ) : (
          <span
            onDoubleClick={startRename}
            title="Double-click to rename"
            style={{ cursor: "text", flex: 1 }}
          >{name}</span>
        )}
        {rule && !renaming && (
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
                background: theme.headerBg, border: `1px solid ${rule.color}44`,
                borderRadius: 6, padding: "8px 10px", whiteSpace: "nowrap",
                fontSize: 11, color: theme.pageText, fontWeight: 400,
                textTransform: "none", letterSpacing: 0, boxShadow: "0 4px 16px #00000066",
                fontFamily: theme.font,
              }}>
                <strong style={{ color: rule.color }}>{rule.label}</strong>
                {srcLabel && <span style={{ marginLeft: 6, fontSize: 9, color: srcColor }}>({srcLabel})</span>}
                <br />Type: <em>{rule.type}</em><br />
                {rule.type === "enum"    && <span style={{ color: theme.mutedText }}>{rule.options.slice(0,6).join(" · ")}{rule.options.length>6?` +${rule.options.length-6} more`:""}</span>}
                {rule.type === "integer" && <span style={{ color: theme.mutedText }}>Whole numbers only</span>}
                {rule.type === "regex"   && <span style={{ color: theme.mutedText }}>Must be a valid regex</span>}
              </span>
            )}
          </span>
        )}
        {/* Profile badge */}
        {profile && showProfile && !renaming && (
          <span style={{ fontSize: 9, color: profile.nullPct > 50 ? theme.invalidColor : theme.mutedText,
            whiteSpace: "nowrap", cursor: "default" }}
            title={`${profile.unique} unique · ${profile.nullPct}% null${profile.numericCount ? ` · min ${profile.min} max ${profile.max}` : ""}`}>
            {profile.nullPct}%∅
          </span>
        )}
        {!renaming && (
          <button
            onClick={(e) => { e.stopPropagation(); onDeleteCol(name); }}
            title="Delete column"
            style={{ background: "none", border: "none", cursor: "pointer", color: theme.dimText,
              fontSize: 10, lineHeight: 1, padding: "0 2px",
              opacity: 0, transition: "opacity 0.15s" }}
            onMouseEnter={(e) => e.currentTarget.style.opacity = 1}
            onMouseLeave={(e) => e.currentTarget.style.opacity = 0}
          >✕</button>
        )}
      </div>
    </th>
  );
}

// ── Data profile modal ────────────────────────────────────────────────────────
function ProfileModal({ sheets, activeSheet, onClose }) {
  const { theme } = useTheme();
  const S = useMemo(() => makeStyles(theme), [theme]);
  const [selectedSheet, setSelectedSheet] = useState(activeSheet);
  const sheet = sheets[selectedSheet];
  const profiles = useMemo(() => {
    if (!sheet) return {};
    const p = {};
    sheet.headers.forEach((h) => { p[h] = profileColumn(sheet.rows, h); });
    return p;
  }, [sheet]);

  return (
    <div style={{
      position: "fixed", inset: 0, background: "#00000099", zIndex: 1000,
      display: "flex", alignItems: "center", justifyContent: "center",
    }} onClick={onClose}>
      <div style={{
        background: theme.cardBg, border: `1px solid ${theme.cardBorder}`, borderRadius: 12,
        padding: "24px", maxWidth: 780, width: "90vw", maxHeight: "82vh",
        overflowY: "auto", display: "flex", flexDirection: "column", gap: 16,
        fontFamily: theme.font,
      }} onClick={(e) => e.stopPropagation()}>
        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", flexWrap: "wrap", gap: 8 }}>
          <span style={{ fontSize: 14, fontWeight: 700, color: theme.pageText }}>Data Profile</span>
          <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
            {Object.keys(sheets).map((n) => (
              <button key={n} style={S.sheetBtn(selectedSheet === n)} onClick={() => setSelectedSheet(n)}>{n}</button>
            ))}
          </div>
          <button onClick={onClose} style={{ ...S.btn(), fontSize: 13, padding: "3px 10px" }}>✕</button>
        </div>
        {sheet ? (
          <div style={{ overflowX: "auto" }}>
            <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
              <thead>
                <tr style={{ background: theme.headerBg }}>
                  {["Column","Total","Nulls","Null %","Unique","Top Values","Min","Max"].map((h) => (
                    <th key={h} style={{ padding: "7px 12px", color: theme.headerText, fontWeight: 700,
                      fontSize: 11, textAlign: "left", borderBottom: `2px solid ${theme.headerBorder}`,
                      textTransform: "uppercase", letterSpacing: 0.5, whiteSpace: "nowrap" }}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {sheet.headers.map((h, i) => {
                  const p = profiles[h];
                  return (
                    <tr key={h} style={{ borderTop: `1px solid ${theme.rowBorder}`,
                      background: i % 2 === 0 ? theme.rowEven : theme.rowOdd }}>
                      <td style={{ padding: "7px 12px", color: theme.accentText, fontWeight: 600 }}>{h}</td>
                      <td style={{ padding: "7px 12px", color: theme.cellText }}>{p.total}</td>
                      <td style={{ padding: "7px 12px", color: p.nullCount > 0 ? theme.invalidColor : theme.mutedText }}>{p.nullCount}</td>
                      <td style={{ padding: "7px 12px" }}>
                        <span style={{ color: p.nullPct > 50 ? theme.invalidColor : p.nullPct > 20 ? theme.dirtyColor : theme.mutedText }}>
                          {p.nullPct}%
                        </span>
                      </td>
                      <td style={{ padding: "7px 12px", color: theme.cellText }}>{p.unique}</td>
                      <td style={{ padding: "7px 12px", color: theme.mutedText, maxWidth: 200 }}>
                        {p.topEntries.map(([v, c]) => `${v} (${c})`).join(" · ") || "—"}
                      </td>
                      <td style={{ padding: "7px 12px", color: theme.mutedText }}>{p.min != null ? p.min : "—"}</td>
                      <td style={{ padding: "7px 12px", color: theme.mutedText }}>{p.max != null ? p.max : "—"}</td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        ) : (
          <div style={{ color: theme.mutedText, fontStyle: "italic" }}>No data.</div>
        )}
      </div>
    </div>
  );
}

// ── Diff modal ────────────────────────────────────────────────────────────────
function DiffModal({ sheets, originalSheets, dirtySet, onClose }) {
  const { theme } = useTheme();
  const S = useMemo(() => makeStyles(theme), [theme]);
  const diffs = [];
  dirtySet.forEach((key) => {
    const parts  = key.split(":");
    const sheet  = parts[0];
    const rowIdx = parseInt(parts[1], 10);
    const col    = parts.slice(2).join(":");
    const newVal = sheets[sheet]?.rows[rowIdx]?.[col];
    const oldVal = originalSheets[sheet]?.rows[rowIdx]?.[col];
    if (newVal !== oldVal) diffs.push({ sheet, rowIdx, col, oldVal, newVal });
  });
  const bySheet = {};
  diffs.forEach((d) => { if (!bySheet[d.sheet]) bySheet[d.sheet] = []; bySheet[d.sheet].push(d); });

  return (
    <div style={{ position: "fixed", inset: 0, background: "#00000099", zIndex: 1000,
      display: "flex", alignItems: "center", justifyContent: "center" }} onClick={onClose}>
      <div style={{ background: theme.cardBg, border: `1px solid ${theme.cardBorder}`, borderRadius: 12,
        padding: "24px", maxWidth: 720, width: "90vw", maxHeight: "80vh",
        overflowY: "auto", display: "flex", flexDirection: "column", gap: 16, fontFamily: theme.font,
      }} onClick={(e) => e.stopPropagation()}>
        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between" }}>
          <span style={{ fontSize: 14, fontWeight: 700, color: theme.pageText }}>
            What Changed — {diffs.length} edit{diffs.length !== 1 ? "s" : ""}
          </span>
          <button onClick={onClose} style={{ ...S.btn(), fontSize: 13, padding: "3px 10px" }}>✕ Close</button>
        </div>
        {diffs.length === 0 ? (
          <div style={{ color: theme.mutedText, fontSize: 13, fontStyle: "italic" }}>No changes yet.</div>
        ) : Object.entries(bySheet).map(([sheetName, items]) => (
          <div key={sheetName} style={{ display: "flex", flexDirection: "column", gap: 6 }}>
            <div style={{ fontSize: 11, fontWeight: 700, color: theme.mutedText, textTransform: "uppercase", letterSpacing: 0.5 }}>
              Sheet: {sheetName}
            </div>
            <div style={{ border: `1px solid ${theme.cardBorder}`, borderRadius: 8, overflow: "hidden" }}>
              <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
                <thead>
                  <tr style={{ background: theme.headerBg }}>
                    {["Row","Column","Before","After"].map((h, i) => (
                      <th key={h} style={{ padding: "7px 12px", fontWeight: 600, fontSize: 11, textAlign: "left",
                        color: i === 2 ? theme.invalidColor : i === 3 ? "#22c55e" : theme.mutedText }}>{h}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {items.map((d, i) => (
                    <tr key={i} style={{ borderTop: `1px solid ${theme.rowBorder}` }}>
                      <td style={{ padding: "6px 12px", color: theme.mutedText }}>{d.rowIdx + 1}</td>
                      <td style={{ padding: "6px 12px", color: theme.accentText, fontWeight: 600 }}>{d.col}</td>
                      <td style={{ padding: "6px 12px", color: theme.invalidColor, fontStyle: d.oldVal == null ? "italic" : "normal" }}>
                        {d.oldVal != null ? String(d.oldVal) : <span style={{ color: theme.dimText }}>—</span>}
                      </td>
                      <td style={{ padding: "6px 12px", color: "#22c55e", fontStyle: d.newVal == null ? "italic" : "normal" }}>
                        {d.newVal != null ? String(d.newVal) : <span style={{ color: theme.dimText }}>—</span>}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        ))}
      </div>
    </div>
  );
}

// ── Audit log modal ───────────────────────────────────────────────────────────
function AuditModal({ auditLog, xlsx, fileName, onClose }) {
  const { theme } = useTheme();
  const S = useMemo(() => makeStyles(theme), [theme]);
  return (
    <div style={{ position: "fixed", inset: 0, background: "#00000099", zIndex: 1000,
      display: "flex", alignItems: "center", justifyContent: "center" }} onClick={onClose}>
      <div style={{ background: theme.cardBg, border: `1px solid ${theme.cardBorder}`, borderRadius: 12,
        padding: "24px", maxWidth: 760, width: "90vw", maxHeight: "80vh",
        overflowY: "auto", display: "flex", flexDirection: "column", gap: 16, fontFamily: theme.font,
      }} onClick={(e) => e.stopPropagation()}>
        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between" }}>
          <span style={{ fontSize: 14, fontWeight: 700, color: theme.pageText }}>
            Audit Log — {auditLog.length} event{auditLog.length !== 1 ? "s" : ""}
          </span>
          <div style={{ display: "flex", gap: 6 }}>
            {xlsx && auditLog.length > 0 && (
              <button style={S.btn("success")} onClick={() => exportAuditXlsx(xlsx, auditLog, fileName)}>
                ↓ Export XLSX
              </button>
            )}
            <button onClick={onClose} style={{ ...S.btn(), fontSize: 13, padding: "3px 10px" }}>✕</button>
          </div>
        </div>
        {auditLog.length === 0 ? (
          <div style={{ color: theme.mutedText, fontStyle: "italic" }}>No edits recorded yet.</div>
        ) : (
          <div style={{ border: `1px solid ${theme.cardBorder}`, borderRadius: 8, overflow: "hidden" }}>
            <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
              <thead>
                <tr style={{ background: theme.headerBg }}>
                  {["#","Time","Sheet","Row","Column","Old Value","New Value"].map((h) => (
                    <th key={h} style={{ padding: "7px 12px", color: theme.headerText, fontWeight: 700,
                      fontSize: 11, textAlign: "left", borderBottom: `2px solid ${theme.headerBorder}`,
                      textTransform: "uppercase", letterSpacing: 0.5, whiteSpace: "nowrap" }}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {[...auditLog].reverse().map((e, i) => (
                  <tr key={i} style={{ borderTop: `1px solid ${theme.rowBorder}`,
                    background: i % 2 === 0 ? theme.rowEven : theme.rowOdd }}>
                    <td style={{ padding: "5px 12px", color: theme.dimText }}>{auditLog.length - i}</td>
                    <td style={{ padding: "5px 12px", color: theme.mutedText, whiteSpace: "nowrap" }}>{e.ts}</td>
                    <td style={{ padding: "5px 12px", color: theme.accentText }}>{e.sheet}</td>
                    <td style={{ padding: "5px 12px", color: theme.cellText }}>{e.rowIdx + 1}</td>
                    <td style={{ padding: "5px 12px", color: theme.accentText, fontWeight: 600 }}>{e.col}</td>
                    <td style={{ padding: "5px 12px", color: theme.invalidColor, fontStyle: e.oldVal == null ? "italic" : "normal" }}>
                      {e.oldVal != null ? String(e.oldVal) : <span style={{ color: theme.dimText }}>—</span>}
                    </td>
                    <td style={{ padding: "5px 12px", color: "#22c55e", fontStyle: e.newVal == null ? "italic" : "normal" }}>
                      {e.newVal != null ? String(e.newVal) : <span style={{ color: theme.dimText }}>—</span>}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </div>
    </div>
  );
}

// ── Schema validation panel ───────────────────────────────────────────────────
function SchemaPanel({ sheets, activeSheet, onClose }) {
  const { theme } = useTheme();
  const S = useMemo(() => makeStyles(theme), [theme]);
  const [schemaText, setSchemaText] = useState("");
  const [result,     setResult]     = useState(null);

  function runCheck() {
    const lines = schemaText.split("\n").map((l) => l.trim()).filter(Boolean);
    const expected = new Set(lines);
    const actual   = new Set(sheets[activeSheet]?.headers || []);
    const missing  = [...expected].filter((c) => !actual.has(c));
    const extra    = [...actual].filter((c) => !expected.has(c));
    const matched  = [...expected].filter((c) => actual.has(c));
    setResult({ missing, extra, matched });
  }

  return (
    <div style={{ position: "fixed", inset: 0, background: "#00000099", zIndex: 1000,
      display: "flex", alignItems: "center", justifyContent: "center" }} onClick={onClose}>
      <div style={{ background: theme.cardBg, border: `1px solid ${theme.cardBorder}`, borderRadius: 12,
        padding: "24px", maxWidth: 560, width: "90vw", maxHeight: "80vh",
        overflowY: "auto", display: "flex", flexDirection: "column", gap: 16, fontFamily: theme.font,
      }} onClick={(e) => e.stopPropagation()}>
        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between" }}>
          <span style={{ fontSize: 14, fontWeight: 700, color: theme.pageText }}>Schema Validation — {activeSheet}</span>
          <button onClick={onClose} style={{ ...S.btn(), fontSize: 13, padding: "3px 10px" }}>✕</button>
        </div>
        <div style={{ fontSize: 11, color: theme.mutedText }}>
          Enter the expected column names (one per line). The checker will flag missing and extra columns.
        </div>
        <textarea
          value={schemaText}
          onChange={(e) => setSchemaText(e.target.value)}
          placeholder={"source_system\ncol_name\ndata_type\nmandatory\n…"}
          style={{ minHeight: 120, background: theme.inputBg, border: `1px solid ${theme.cardBorder}`,
            borderRadius: 6, color: theme.pageText, fontSize: 12, padding: "8px 10px",
            fontFamily: "monospace", resize: "vertical" }}
        />
        <button style={S.btn("primary")} onClick={runCheck}>Check Schema</button>
        {result && (
          <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
            <ResultGroup label="Missing columns" items={result.missing} color={theme.invalidColor} icon="✗" />
            <ResultGroup label="Extra columns (not in schema)" items={result.extra} color={theme.dirtyColor} icon="+" />
            <ResultGroup label="Matched" items={result.matched} color="#22c55e" icon="✓" />
          </div>
        )}
      </div>
    </div>
  );
}
function ResultGroup({ label, items, color, icon }) {
  const { theme } = useTheme();
  if (!items.length) return null;
  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 4 }}>
      <span style={{ fontSize: 11, fontWeight: 700, color, textTransform: "uppercase", letterSpacing: 0.5 }}>{label} ({items.length})</span>
      <div style={{ display: "flex", gap: 4, flexWrap: "wrap" }}>
        {items.map((c) => (
          <span key={c} style={{ padding: "2px 8px", borderRadius: 4, fontSize: 11,
            background: `${color}18`, border: `1px solid ${color}44`, color, fontFamily: "monospace" }}>
            {icon} {c}
          </span>
        ))}
      </div>
    </div>
  );
}

// ── Sheet table ───────────────────────────────────────────────────────────────
function SheetTable({ sheetName, sheet, dirtySet, validationMap, onCellChange, onAddRow, onDeleteRow,
                      customRules, columnAssignments, searchQuery, onRenameCol, onDeleteCol, onAddCol,
                      comments, onSetComment, showProfile, profiles }) {
  const { theme } = useTheme();
  const S = useMemo(() => makeStyles(theme), [theme]);
  const { headers, rows } = sheet;

  const headerMeta = headers.map((h) => {
    const explicitRule = customRules.find((r) => columnAssignments[r.id]?.has(h));
    if (explicitRule) return { rule: explicitRule, source: "explicit" };
    const patternRule = customRules.find((r) => r.pattern.test(h));
    if (patternRule) return { rule: patternRule, source: "pattern" };
    return { rule: null, source: "none" };
  });

  const PAGE_SIZE_OPTIONS = [10, 15, 25, 50];
  const [pageSize, setPageSize] = useState(15);
  const [page,     setPage]     = useState(1);

  const q = searchQuery.trim().toLowerCase();
  const filteredRows = q
    ? rows.map((row, i) => ({ row, origIdx: i })).filter(({ row }) =>
        headers.some((h) => row[h] != null && String(row[h]).toLowerCase().includes(q))
      )
    : rows.map((row, i) => ({ row, origIdx: i }));

  const totalPages  = Math.max(1, Math.ceil(filteredRows.length / pageSize));
  const safePage    = Math.min(page, totalPages);
  const visibleRows = filteredRows.slice((safePage - 1) * pageSize, safePage * pageSize);

  // Reset to page 1 when filter or sheet changes
  useEffect(() => { setPage(1); }, [q, sheetName]);

  function goTo(p) { setPage(Math.max(1, Math.min(p, totalPages))); }

  // Build compact page numbers: always show first, last, current ±1, with ellipsis gaps
  function pageNumbers() {
    const pages = [];
    const add = (p) => { if (!pages.includes(p) && p >= 1 && p <= totalPages) pages.push(p); };
    add(1); add(totalPages);
    for (let d = -1; d <= 1; d++) add(safePage + d);
    pages.sort((a, b) => a - b);
    const result = [];
    pages.forEach((p, i) => {
      if (i > 0 && p - pages[i - 1] > 1) result.push("…");
      result.push(p);
    });
    return result;
  }

  const pgBtn = (label, onClick, active = false, disabled = false) => (
    <button key={label} onClick={onClick} disabled={disabled}
      style={{
        minWidth: 30, height: 28, padding: "0 8px",
        background: active ? theme.accent : "none",
        border: `1px solid ${active ? theme.accent : theme.cardBorder}`,
        borderRadius: 5, cursor: disabled ? "not-allowed" : "pointer",
        color: active ? "#fff" : disabled ? theme.dimText : theme.pageSubText,
        fontSize: 11, fontWeight: active ? 700 : 400, fontFamily: theme.font,
        opacity: disabled ? 0.4 : 1, transition: "background 0.15s",
      }}
    >{label}</button>
  );

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
      <div style={S.tableWrap}>
        <table style={S.table}>
          <thead>
            <tr>
              {headers.map((h, i) => (
                <ColHeader key={i} name={h} rule={headerMeta[i].rule} source={headerMeta[i].source}
                  onRename={onRenameCol} onDeleteCol={onDeleteCol}
                  profile={profiles?.[h]} showProfile={showProfile} />
              ))}
              <th style={{ ...S.thDelete, width: 40 }}>
                <button onClick={onAddCol} title="Add column"
                  style={{ background: "none", border: `1px solid ${theme.cardBorder}`, borderRadius: 4,
                    cursor: "pointer", color: theme.dimText, fontSize: 14, lineHeight: 1, padding: "2px 6px" }}
                  onMouseEnter={(e) => { e.currentTarget.style.color = theme.accentText; e.currentTarget.style.borderColor = `${theme.accent}55`; }}
                  onMouseLeave={(e) => { e.currentTarget.style.color = theme.dimText; e.currentTarget.style.borderColor = theme.cardBorder; }}
                >+</button>
              </th>
              <th style={S.thDelete} />
            </tr>
          </thead>
          <tbody>
            {visibleRows.length === 0 ? (
              <tr>
                <td colSpan={headers.length + 2}
                  style={{ padding: "24px", textAlign: "center", color: theme.dimText, fontSize: 13, fontStyle: "italic" }}>
                  {q ? `No rows match "${searchQuery}"` : `No rows — click "+ Add Row" to start`}
                </td>
              </tr>
            ) : visibleRows.map(({ row, origIdx: ri }) => (
              <tr key={ri}>
                {headers.map((h, hi) => {
                  const key = `${sheetName}:${ri}:${h}`;
                  const msg = validationMap.get(key) || null;
                  return (
                    <EditableCell key={h} value={row[h]} rowIdx={ri} colKey={h}
                      dirty={dirtySet.has(key)} rule={headerMeta[hi].rule}
                      invalid={!!msg} validationMsg={msg} onChange={onCellChange}
                      comment={comments?.[key]} onSetComment={onSetComment} />
                  );
                })}
                <td style={{ ...S.td(ri, false, false), textAlign: "center", padding: "0 4px" }} />
                <td style={{ ...S.td(ri, false, false), textAlign: "center", padding: "0 4px" }}>
                  <button onClick={() => onDeleteRow(ri)}
                    style={{ background: "none", border: "none", cursor: "pointer",
                      color: `${theme.invalidColor}66`, fontSize: 14, lineHeight: 1, padding: "4px 6px" }}
                    title="Delete row">✕</button>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      {/* ── Bottom bar: Add Row + pagination ── */}
      <div style={{ display: "flex", alignItems: "center", gap: 8, flexWrap: "wrap" }}>
        <button style={S.btn("success")} onClick={onAddRow}>+ Add Row</button>

        {/* row count info */}
        <span style={{ fontSize: 11, color: theme.mutedText, marginRight: 4 }}>
          {filteredRows.length === 0 ? "0 rows" : (
            <>
              {(safePage - 1) * pageSize + 1}–{Math.min(safePage * pageSize, filteredRows.length)}
              {" of "}{filteredRows.length}{q ? ` (filtered from ${rows.length})` : ""} rows
            </>
          )}
        </span>

        {/* pagination controls — only shown when needed */}
        {totalPages > 1 && (
          <div style={{ display: "flex", gap: 4, alignItems: "center", marginLeft: "auto" }}>
            {pgBtn("«", () => goTo(1),         false, safePage === 1)}
            {pgBtn("‹", () => goTo(safePage - 1), false, safePage === 1)}
            {pageNumbers().map((p, i) =>
              p === "…"
                ? <span key={`ellipsis-${i}`} style={{ fontSize: 11, color: theme.dimText, padding: "0 2px" }}>…</span>
                : pgBtn(p, () => goTo(p), p === safePage)
            )}
            {pgBtn("›", () => goTo(safePage + 1), false, safePage === totalPages)}
            {pgBtn("»", () => goTo(totalPages),    false, safePage === totalPages)}
          </div>
        )}

        {/* rows-per-page selector */}
        <select value={pageSize} onChange={(e) => { setPageSize(Number(e.target.value)); setPage(1); }}
          style={{
            background: theme.inputBg, border: `1px solid ${theme.cardBorder}`, borderRadius: 5,
            color: theme.pageSubText, fontSize: 11, padding: "3px 6px", cursor: "pointer",
            fontFamily: theme.font, marginLeft: totalPages > 1 ? 4 : "auto",
          }}>
          {PAGE_SIZE_OPTIONS.map((n) => (
            <option key={n} value={n}>{n} / page</option>
          ))}
        </select>

        {validationMap.size > 0 && (
          <span style={S.invalidBadge}>⚠ {validationMap.size} error{validationMap.size !== 1 ? "s" : ""}</span>
        )}
      </div>
    </div>
  );
}

// ── Column chip ───────────────────────────────────────────────────────────────
function ColChip({ name, color, onRemove, isAuto, theme }) {
  return (
    <span style={{
      display: "inline-flex", alignItems: "center", gap: 4,
      padding: "2px 8px", borderRadius: 4, fontSize: 10, fontWeight: 600,
      background: isAuto ? theme.surfaceAlt : `${color}22`,
      border: `1px solid ${isAuto ? theme.btnGhost : color + "55"}`,
      color: isAuto ? theme.mutedText : color,
    }}>
      {name}
      {isAuto && <span style={{ fontSize: 9, color: theme.dimText }}>auto</span>}
      {!isAuto && (
        <button onClick={onRemove} style={{
          background: "none", border: "none", cursor: "pointer",
          color, fontSize: 11, lineHeight: 1, padding: 0, marginLeft: 2,
        }} title="Remove assignment">×</button>
      )}
    </span>
  );
}

// ── Rules panel ───────────────────────────────────────────────────────────────
function RulesPanel({ customRules, columnAssignments, allColumns, onUpdateRule, onResetRules, onAssignColumn, onUnassignColumn }) {
  const { theme } = useTheme();
  const S = useMemo(() => makeStyles(theme), [theme]);
  const [expanded,    setExpanded]    = useState(false);
  const [editingId,   setEditingId]   = useState(null);
  const [editOptions, setEditOptions] = useState("");
  const [addingFor,   setAddingFor]   = useState(null);
  const [colSearch,   setColSearch]   = useState("");

  function saveOptions(id) {
    const opts = editOptions.split(",").map((s) => s.trim()).filter(Boolean);
    onUpdateRule(id, { options: opts });
    setEditingId(null);
  }

  function ruleColumns(rule) {
    const explicit = [...(columnAssignments[rule.id] || new Set())];
    const auto     = allColumns.filter((c) =>
      !explicit.includes(c) &&
      rule.pattern.test(c) &&
      !customRules.some((r) => r.id !== rule.id && columnAssignments[r.id]?.has(c))
    );
    return { explicit, auto };
  }

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

  const ruleAccent = theme.accent === "#2563eb" ? "#7c3aed" : theme.accent;

  return (
    <div style={{ background: theme.rulesHeaderBg, border: `1px solid ${theme.cardBorder}`, borderRadius: 8 }}>
      <button onClick={() => setExpanded((v) => !v)} style={{
        width: "100%", background: "none", border: "none", cursor: "pointer",
        padding: "10px 14px", display: "flex", alignItems: "center", gap: 8,
        color: theme.pageSubText, fontSize: 12, fontWeight: 600, textAlign: "left",
        fontFamily: theme.font,
      }}>
        <span style={{ color: ruleAccent, fontSize: 14 }}>⚙</span>
        Column Validation Rules
        <span style={{ marginLeft: "auto", fontSize: 10 }}>{expanded ? "▲" : "▼"}</span>
        <span style={S.badge(ruleAccent)}>{customRules.length} rules</span>
        {allColumns.length > 0 && (
          <span style={S.badge(theme.mutedText)}>
            {allColumns.filter((c) => customRules.some((r) => {
              const { explicit, auto } = ruleColumns(r);
              return explicit.includes(c) || auto.includes(c);
            })).length} / {allColumns.length} columns covered
          </span>
        )}
      </button>

      {expanded && (
        <div style={{ padding: "0 14px 14px", display: "flex", flexDirection: "column", gap: 10 }}>
          <div style={{ fontSize: 11, color: theme.dimText }}>
            <strong style={{ color: theme.mutedText }}>auto</strong> — matched by name pattern.&nbsp;
            <strong style={{ color: theme.accentText }}>explicit</strong> — manually pinned. Explicit assignments override patterns.
            {allColumns.length === 0 && <span style={{ color: `${theme.invalidColor}66`, marginLeft: 8 }}>Upload a file to see available columns.</span>}
          </div>

          {customRules.map((rule) => {
            const { explicit, auto } = ruleColumns(rule);
            const isAddingHere = addingFor === rule.id;
            const avail = isAddingHere ? availableColumns(rule.id).filter((c) => !explicit.includes(c)) : [];

            return (
              <div key={rule.id} style={{
                background: theme.ruleItemBg, border: `1px solid ${rule.color}33`,
                borderLeft: `3px solid ${rule.color}`, borderRadius: 6, padding: "10px 12px",
                display: "flex", flexDirection: "column", gap: 8,
              }}>
                <div style={{ display: "flex", alignItems: "center", gap: 8, flexWrap: "wrap" }}>
                  <span style={{ fontSize: 12, fontWeight: 700, color: rule.color, minWidth: 130 }}>{rule.label}</span>
                  <span style={S.badge(theme.mutedText)}>{rule.type}</span>
                  <span style={{ fontSize: 10, color: theme.dimText, flex: 1 }}>
                    matches: <code style={{ color: theme.mutedText }}>{rule.pattern.toString()}</code>
                  </span>
                  {rule.type === "enum" && (
                    editingId === rule.id ? (
                      <div style={{ display: "flex", gap: 6, alignItems: "center", width: "100%" }}>
                        <input value={editOptions} onChange={(e) => setEditOptions(e.target.value)}
                          placeholder="Comma-separated values…"
                          style={{ flex: 1, background: theme.inputBg, border: `1px solid ${theme.accentBorder}`,
                            borderRadius: 5, padding: "4px 8px", color: theme.pageText, fontSize: 11, fontFamily: theme.font }} />
                        <button style={S.btn("primary")} onClick={() => saveOptions(rule.id)}>Save</button>
                        <button style={S.btn()} onClick={() => setEditingId(null)}>Cancel</button>
                      </div>
                    ) : (
                      <>
                        <span style={{ fontSize: 10, color: theme.mutedText }}>
                          {rule.options.slice(0,5).join(", ")}{rule.options.length>5 ? ` +${rule.options.length-5}` : ""}
                        </span>
                        <button style={{ ...S.btn(), fontSize: 10, padding: "2px 8px" }}
                          onClick={() => { setEditingId(rule.id); setEditOptions(rule.options.join(", ")); }}>
                          Edit options
                        </button>
                      </>
                    )
                  )}
                  {rule.type === "integer" && <span style={{ fontSize: 10, color: theme.mutedText }}>Whole numbers only</span>}
                  {rule.type === "regex"   && <span style={{ fontSize: 10, color: theme.mutedText }}>Must be a valid regex</span>}
                </div>

                <div style={{ display: "flex", alignItems: "flex-start", gap: 8, flexWrap: "wrap" }}>
                  <span style={{ fontSize: 10, color: theme.mutedText, fontWeight: 600, paddingTop: 2, whiteSpace: "nowrap" }}>
                    Applied to:
                  </span>
                  <div style={{ display: "flex", gap: 4, flexWrap: "wrap", flex: 1 }}>
                    {auto.map((c) => (
                      <ColChip key={c} name={c} color={rule.color} isAuto theme={theme} />
                    ))}
                    {explicit.map((c) => (
                      <ColChip key={c} name={c} color={rule.color} isAuto={false} theme={theme}
                        onRemove={() => onUnassignColumn(rule.id, c)} />
                    ))}
                    {auto.length === 0 && explicit.length === 0 && (
                      <span style={{ fontSize: 10, color: theme.dimText, fontStyle: "italic" }}>No columns matched yet</span>
                    )}
                  </div>

                  {allColumns.length > 0 && (
                    isAddingHere ? (
                      <div style={{ display: "flex", flexDirection: "column", gap: 4, minWidth: 200 }}>
                        <input autoFocus value={colSearch} onChange={(e) => setColSearch(e.target.value)}
                          placeholder="Search columns…"
                          style={{ background: theme.inputBg, border: `1px solid ${rule.color}55`,
                            borderRadius: 5, padding: "4px 8px", color: theme.pageText, fontSize: 11, fontFamily: theme.font }}
                        />
                        <div style={{ background: theme.inputBg, border: `1px solid ${rule.color}33`,
                          borderRadius: 5, maxHeight: 140, overflowY: "auto" }}>
                          {avail.length === 0
                            ? <div style={{ padding: "8px 10px", fontSize: 11, color: theme.dimText, fontStyle: "italic" }}>No columns available</div>
                            : avail.map((c) => (
                              <button key={c} onClick={() => { onAssignColumn(rule.id, c); setColSearch(""); }}
                                style={{ display: "block", width: "100%", textAlign: "left",
                                  background: "none", border: "none", cursor: "pointer",
                                  padding: "5px 10px", color: theme.pageText, fontSize: 11,
                                  borderBottom: `1px solid ${theme.cardBorder}`, fontFamily: theme.font }}
                                onMouseEnter={(e) => e.currentTarget.style.background = theme.surfaceAlt}
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

// ── Main component ────────────────────────────────────────────────────────────
export default function MetadataTab() {
  const { theme } = useTheme();
  const S = useMemo(() => makeStyles(theme), [theme]);
  const xlsx = useXLSX();

  const { state: sheets, set: setSheets, undo, redo, reset: resetSheets,
          canUndo, canRedo } = useHistory(null);

  const [originalSheets,    setOriginalSheets]    = useState(null);
  const [activeSheet,       setActiveSheet]        = useState(null);
  const [fileName,          setFileName]           = useState(null);
  const [fileType,          setFileType]           = useState(null);
  const [dragOver,          setDragOver]           = useState(false);
  const [error,             setError]              = useState(null);
  const [dirtySet,          setDirtySet]           = useState(new Set());
  const [savedFlash,        setSavedFlash]         = useState(false);
  const [validationMap,     setValidationMap]      = useState(new Map());
  const [customRules,       setCustomRules]        = useState(() =>
    DEFAULT_RULES.map((r) => ({ ...r, options: r.options ? [...r.options] : undefined }))
  );
  const [columnAssignments, setColumnAssignments]  = useState({});
  const [searchQuery,       setSearchQuery]        = useState("");
  const [showDiff,          setShowDiff]           = useState(false);
  const [showAudit,         setShowAudit]          = useState(false);
  const [showProfile,       setShowProfile]        = useState(false);
  const [showProfileInline, setShowProfileInline]  = useState(false);
  const [showSchema,        setShowSchema]         = useState(false);
  const [showCmp,           setShowCmp]            = useState(false);
  // auditLog: [{ts, sheet, rowIdx, col, oldVal, newVal}]
  const [auditLog,          setAuditLog]           = useState([]);
  // comments: { "sheet:row:col": "text" }
  const [comments,          setComments]           = useState({});
  const inputRef = useRef();

  // Keyboard shortcuts: Ctrl/Cmd+Z / Ctrl/Cmd+Y
  useEffect(() => {
    function onKey(e) {
      const mod = e.metaKey || e.ctrlKey;
      if (!mod) return;
      if (e.key === "z" && !e.shiftKey) { e.preventDefault(); undo(); }
      if (e.key === "y" || (e.key === "z" && e.shiftKey)) { e.preventDefault(); redo(); }
    }
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, [undo, redo]);

  const allColumns = useMemo(() => {
    if (!sheets) return [];
    const cols = new Set();
    Object.values(sheets).forEach(({ headers }) => headers.forEach((h) => cols.add(h)));
    return [...cols].sort();
  }, [sheets]);

  // Per-column profiles for the active sheet
  const activeProfiles = useMemo(() => {
    if (!sheets || !activeSheet) return {};
    const sheet = sheets[activeSheet];
    const p = {};
    sheet.headers.forEach((h) => { p[h] = profileColumn(sheet.rows, h); });
    return p;
  }, [sheets, activeSheet]);

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
    resetSheets(parsed);
    setOriginalSheets(JSON.parse(JSON.stringify(parsed)));
    setActiveSheet(Object.keys(parsed)[0] || null);
    setFileName(name); setFileType(type); setError(null);
    setDirtySet(new Set()); setColumnAssignments({});
    setSearchQuery(""); setAuditLog([]); setComments({});
  }, [resetSheets]);

  const parseFile = useCallback((file) => {
    if (!file) return;
    setError(null);
    const type = isJson(file.name) ? "json" : isXlsx(file.name) ? "xlsx" : null;
    if (!type) { setError("Unsupported file type. Upload XLSX, XLS, ODS, CSV, or JSON."); return; }

    parseViaApi(file)
      .then((parsed) => {
        if (!Object.keys(parsed).length) throw new Error("No sheets found.");
        applyParsed(parsed, file.name, type);
      })
      .catch(() => {
        if (type === "json") {
          const r = new FileReader();
          r.onload = (e) => {
            try {
              const p = parseJsonMapping(e.target.result);
              if (!Object.keys(p).length) throw new Error("No sheets found.");
              applyParsed(p, file.name, "json");
            } catch (err) { setError("JSON parse error: " + err.message); }
          };
          r.readAsText(file);
        } else {
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
          r.readAsArrayBuffer(file);
        }
      });
  }, [xlsx, applyParsed]);

  const onDrop        = useCallback((e) => { e.preventDefault(); setDragOver(false); parseFile(e.dataTransfer.files[0]); }, [parseFile]);
  const onInputChange = useCallback((e) => { parseFile(e.target.files[0]); e.target.value = ""; }, [parseFile]);
  const clear         = () => {
    resetSheets(null); setOriginalSheets(null);
    setActiveSheet(null); setFileName(null); setFileType(null);
    setError(null); setDirtySet(new Set()); setValidationMap(new Map());
    setColumnAssignments({}); setSearchQuery(""); setAuditLog([]); setComments({});
  };

  const onCellChange = useCallback((rowIdx, colKey, newVal) => {
    setSheets((prev) => {
      const sheet = prev[activeSheet];
      const oldVal = sheet.rows[rowIdx]?.[colKey];
      // Append to audit log
      setAuditLog((log) => [...log, {
        ts: new Date().toLocaleTimeString(),
        sheet: activeSheet, rowIdx, col: colKey, oldVal, newVal,
      }]);
      return { ...prev, [activeSheet]: { ...sheet, rows: sheet.rows.map((r, i) => i === rowIdx ? { ...r, [colKey]: newVal } : r) } };
    });
    setDirtySet((prev) => { const n = new Set(prev); n.add(`${activeSheet}:${rowIdx}:${colKey}`); return n; });
  }, [activeSheet, setSheets]);

  const onAddRow = useCallback(() => {
    setSheets((prev) => {
      const sheet = prev[activeSheet];
      const empty = Object.fromEntries(sheet.headers.map((h) => [h, null]));
      return { ...prev, [activeSheet]: { ...sheet, rows: [...sheet.rows, empty] } };
    });
  }, [activeSheet, setSheets]);

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
  }, [activeSheet, setSheets]);

  // ── Column operations ─────────────────────────────────────────────────────
  const onRenameCol = useCallback((oldName, newName) => {
    setSheets((prev) => {
      const sheet = prev[activeSheet];
      const headers = sheet.headers.map((h) => h === oldName ? newName : h);
      const rows    = sheet.rows.map((r) => {
        const next = { ...r };
        if (oldName in next) { next[newName] = next[oldName]; delete next[oldName]; }
        return next;
      });
      return { ...prev, [activeSheet]: { ...sheet, headers, rows } };
    });
    setDirtySet((prev) => {
      const next = new Set();
      prev.forEach((key) => {
        const parts = key.split(":");
        const sn = parts[0], ri = parts[1], col = parts.slice(2).join(":");
        next.add(sn === activeSheet && col === oldName ? `${sn}:${ri}:${newName}` : key);
      });
      return next;
    });
  }, [activeSheet, setSheets]);

  const onDeleteCol = useCallback((colName) => {
    if (!window.confirm(`Delete column "${colName}" from all rows in "${activeSheet}"?`)) return;
    setSheets((prev) => {
      const sheet   = prev[activeSheet];
      const headers = sheet.headers.filter((h) => h !== colName);
      const rows    = sheet.rows.map((r) => { const n = { ...r }; delete n[colName]; return n; });
      return { ...prev, [activeSheet]: { ...sheet, headers, rows } };
    });
    setDirtySet((prev) => {
      const next = new Set();
      prev.forEach((key) => {
        const parts = key.split(":");
        const col = parts.slice(2).join(":");
        if (!(parts[0] === activeSheet && col === colName)) next.add(key);
      });
      return next;
    });
  }, [activeSheet, setSheets]);

  const onAddCol = useCallback(() => {
    const name = window.prompt("New column name:");
    if (!name?.trim()) return;
    const colName = name.trim();
    setSheets((prev) => {
      const sheet = prev[activeSheet];
      if (sheet.headers.includes(colName)) { alert(`Column "${colName}" already exists.`); return prev; }
      return { ...prev, [activeSheet]: {
        ...sheet,
        headers: [...sheet.headers, colName],
        rows: sheet.rows.map((r) => ({ ...r, [colName]: null })),
      }};
    });
  }, [activeSheet, setSheets]);

  // ── Comments ──────────────────────────────────────────────────────────────
  const onSetComment = useCallback((rowIdx, colKey, text) => {
    const key = `${activeSheet}:${rowIdx}:${colKey}`;
    setComments((prev) => {
      const next = { ...prev };
      if (text) next[key] = text; else delete next[key];
      return next;
    });
  }, [activeSheet]);

  const onSave = useCallback(() => {
    setDirtySet(new Set()); setSavedFlash(true);
    setOriginalSheets(JSON.parse(JSON.stringify(sheets)));
    setTimeout(() => setSavedFlash(false), 2000);
    if (fileType === "json") exportJson(sheets, fileName);
    else if (xlsx)           exportXlsx(xlsx, sheets, fileName);
  }, [sheets, fileName, fileType, xlsx]);

  const onUpdateRule     = useCallback((id, patch) => setCustomRules((prev) => prev.map((r) => r.id === id ? { ...r, ...patch } : r)), []);
  const onResetRules     = useCallback(() => {
    setCustomRules(DEFAULT_RULES.map((r) => ({ ...r, options: r.options ? [...r.options] : undefined })));
    setColumnAssignments({});
  }, []);
  const onAssignColumn   = useCallback((ruleId, colName) => {
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
  const commentCount = Object.keys(comments).length;

  return (
    <div style={S.container}>
      {showCmp && (
        <FileCmpModal xlsxLib={xlsx} onClose={() => setShowCmp(false)} />
      )}
      {showDiff    && sheets && originalSheets && (
        <DiffModal sheets={sheets} originalSheets={originalSheets} dirtySet={dirtySet} onClose={() => setShowDiff(false)} />
      )}
      {showAudit   && (
        <AuditModal auditLog={auditLog} xlsx={xlsx} fileName={fileName} onClose={() => setShowAudit(false)} />
      )}
      {showProfile && sheets && (
        <ProfileModal sheets={sheets} activeSheet={activeSheet} onClose={() => setShowProfile(false)} />
      )}
      {showSchema  && sheets && (
        <SchemaPanel sheets={sheets} activeSheet={activeSheet} onClose={() => setShowSchema(false)} />
      )}

      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between" }}>
        <div style={{ fontSize: 14, fontWeight: 700, color: theme.pageText }}>Data Mapping Specification</div>
        <button style={{ ...S.btn(), fontSize: 11 }} onClick={() => setShowCmp(true)} title="Compare two files side by side">
          ⇄ Compare Files
        </button>
      </div>

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
          {/* ── File info bar ─────────────────────────────────────────────── */}
          <div style={S.fileInfo}>
            <span style={{ fontSize: 16 }}>{fileType === "json" ? "{ }" : "📄"}</span>
            <span style={S.fileName}>{fileName}</span>
            {formatLabel && <span style={S.badge(formatColor)}>{formatLabel}</span>}
            <span style={{ fontSize: 11, color: theme.mutedText }}>{sheetNames.length} sheet{sheetNames.length !== 1 ? "s" : ""}</span>
            {totalErrors > 0 && <span style={S.invalidBadge}>⚠ {totalErrors} error{totalErrors !== 1 ? "s" : ""}</span>}

            {/* Undo / Redo */}
            <button style={{ ...S.btn(), opacity: canUndo ? 1 : 0.3 }}
              onClick={undo} disabled={!canUndo} title="Undo (Ctrl+Z)">↩ Undo</button>
            <button style={{ ...S.btn(), opacity: canRedo ? 1 : 0.3 }}
              onClick={redo} disabled={!canRedo} title="Redo (Ctrl+Y)">↪ Redo</button>

            {/* Tools */}
            <button style={S.btn()} onClick={() => setShowCmp(true)} title="Compare two files side by side">⇄ Compare</button>
            <button style={S.btn()} onClick={() => setShowProfile(true)} title="Data profile — null %, unique values, min/max">
              ∑ Profile
            </button>
            <button style={S.btn()} onClick={() => setShowSchema(true)} title="Validate columns against expected schema">
              ✓ Schema
            </button>
            <button
              style={{ ...S.btn(), color: showProfileInline ? theme.accentText : undefined }}
              onClick={() => setShowProfileInline((v) => !v)}
              title="Toggle null % badges on column headers"
            >∅ {showProfileInline ? "Hide" : "Show"} Nulls</button>
            <button
              style={{ ...S.btn(), color: auditLog.length > 0 ? theme.dirtyColor : undefined }}
              onClick={() => setShowAudit(true)}
              title="View audit log"
            >📋 Log{auditLog.length > 0 ? ` (${auditLog.length})` : ""}</button>
            {commentCount > 0 && (
              <span style={{ ...S.badge("#f59e0b"), fontSize: 10 }}>💬 {commentCount} note{commentCount !== 1 ? "s" : ""}</span>
            )}

            {hasDirty && (
              <>
                <button onClick={() => setShowDiff(true)} style={S.btn()} title="Review all changes">
                  ⊡ Diff ({dirtySet.size})
                </button>
                <button onClick={onSave} style={{
                  background: `linear-gradient(135deg, ${theme.btnSuccessBg}, ${theme.btnSuccessBg})`,
                  border: `1px solid ${theme.btnSuccessBorder}`,
                  borderRadius: 7, padding: "5px 14px", cursor: "pointer", color: theme.btnSuccessText,
                  fontSize: 12, fontWeight: 700, display: "flex", alignItems: "center", gap: 6,
                  fontFamily: theme.font,
                }} title="Commit edits and download">✓ Save Changes</button>
              </>
            )}
            {savedFlash && !hasDirty && <span style={{ fontSize: 11, color: theme.btnSuccessText, fontWeight: 600 }}>✓ Saved</span>}
            {hasDirty && <span style={S.dirtyBadge}>✎ {dirtySet.size} edit{dirtySet.size !== 1 ? "s" : ""}</span>}
            <button style={S.btn("success")} onClick={() => exportJson(sheets, fileName)}>↓ JSON</button>
            {xlsx && <button style={S.btn("primary")} onClick={() => exportXlsx(xlsx, sheets, fileName)}>↓ XLSX</button>}
            <button style={S.btn()} onClick={clear}>Clear</button>
          </div>

          {/* ── Hint bar ──────────────────────────────────────────────────── */}
          <div style={{ fontSize: 11, color: theme.dimText, fontStyle: "italic" }}>
            Click cell to edit · Double-click column header to rename · Hover cell → 💬 to comment · Ctrl+Z/Y to undo/redo
            {hasDirty && <span style={{ color: `${theme.dirtyColor}88`, marginLeft: 8 }}>
              · <strong style={{ color: theme.dirtyColor }}>Save Changes</strong> to commit &amp; download
            </span>}
          </div>

          {/* ── Search bar ────────────────────────────────────────────────── */}
          <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
            <div style={{ position: "relative", flex: 1, maxWidth: 360 }}>
              <span style={{ position: "absolute", left: 10, top: "50%", transform: "translateY(-50%)",
                color: theme.dimText, fontSize: 13, pointerEvents: "none" }}>🔍</span>
              <input
                value={searchQuery}
                onChange={(e) => {
                  const q = e.target.value;
                  setSearchQuery(q);
                  // auto-switch to first tab that has a match
                  if (q.trim()) {
                    const ql = q.trim().toLowerCase();
                    const firstMatch = sheetNames.find((name) => {
                      const { headers, rows } = sheets[name];
                      return rows.some((row) =>
                        headers.some((h) => row[h] != null && String(row[h]).toLowerCase().includes(ql))
                      );
                    });
                    if (firstMatch && firstMatch !== activeSheet) setActiveSheet(firstMatch);
                  }
                }}
                placeholder="Search across all tabs…"
                style={{ width: "100%", boxSizing: "border-box",
                  background: theme.cardBg, border: `1px solid ${theme.cardBorder}`, borderRadius: 7,
                  color: theme.pageText, fontSize: 12, padding: "7px 10px 7px 32px",
                  outline: "none", fontFamily: theme.font }}
                onFocus={(e) => e.target.style.borderColor = theme.accent}
                onBlur={(e) => e.target.style.borderColor = theme.cardBorder}
              />
              {searchQuery && (
                <button onClick={() => setSearchQuery("")}
                  style={{ position: "absolute", right: 8, top: "50%", transform: "translateY(-50%)",
                    background: "none", border: "none", cursor: "pointer", color: theme.mutedText, fontSize: 13 }}>✕</button>
              )}
            </div>
            {searchQuery.trim() && (() => {
              const ql = searchQuery.trim().toLowerCase();
              const total = sheetNames.reduce((acc, name) => {
                const { headers, rows } = sheets[name];
                return acc + rows.filter((row) =>
                  headers.some((h) => row[h] != null && String(row[h]).toLowerCase().includes(ql))
                ).length;
              }, 0);
              return (
                <span style={{ fontSize: 11, color: theme.mutedText, whiteSpace: "nowrap" }}>
                  {total} match{total !== 1 ? "es" : ""} across {sheetNames.length} tab{sheetNames.length !== 1 ? "s" : ""}
                </span>
              );
            })()}
          </div>

          {/* ── Rules panel ───────────────────────────────────────────────── */}
          <RulesPanel
            customRules={customRules}
            columnAssignments={columnAssignments}
            allColumns={allColumns}
            onUpdateRule={onUpdateRule}
            onResetRules={onResetRules}
            onAssignColumn={onAssignColumn}
            onUnassignColumn={onUnassignColumn}
          />

          {/* ── Sheet tabs ────────────────────────────────────────────────── */}
          <div style={S.sheetNav}>
            {sheetNames.map((name) => {
              const sheetDirty  = [...dirtySet].some((k) => k.startsWith(name + ":"));
              const sheetErrs   = [...validationMap.keys()].filter((k) => k.startsWith(name + ":")).length;
              const ql          = searchQuery.trim().toLowerCase();
              const matchCount  = ql ? (() => {
                const { headers, rows } = sheets[name];
                return rows.filter((row) =>
                  headers.some((h) => row[h] != null && String(row[h]).toLowerCase().includes(ql))
                ).length;
              })() : 0;
              const noMatch     = ql && matchCount === 0;
              return (
                <button key={name} style={{ ...S.sheetBtn(activeSheet === name), opacity: noMatch ? 0.4 : 1 }}
                  onClick={() => setActiveSheet(name)}>
                  {name}
                  {sheetDirty && <span style={{ marginLeft: 5, color: theme.dirtyColor, fontSize: 10 }}>✎</span>}
                  {sheetErrs > 0 && <span style={{ marginLeft: 4, color: theme.invalidColor, fontSize: 10 }}>⚠{sheetErrs}</span>}
                  {ql && matchCount > 0 && (
                    <span style={{ marginLeft: 5, fontSize: 9, fontWeight: 700,
                      background: `${theme.accent}33`, border: `1px solid ${theme.accent}55`,
                      color: theme.accentText, borderRadius: 3, padding: "1px 5px" }}>
                      {matchCount}
                    </span>
                  )}
                </button>
              );
            })}
          </div>

          {/* ── Editable table ────────────────────────────────────────────── */}
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
              searchQuery={searchQuery}
              onRenameCol={onRenameCol}
              onDeleteCol={onDeleteCol}
              onAddCol={onAddCol}
              comments={comments}
              onSetComment={onSetComment}
              showProfile={showProfileInline}
              profiles={activeProfiles}
            />
          )}
        </>
      )}
    </div>
  );
}
