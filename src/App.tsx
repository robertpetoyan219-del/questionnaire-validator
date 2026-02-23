import { useState, useCallback, useRef } from "react";
import * as XLSX from "xlsx";
import Papa from "papaparse";
import mammoth from "mammoth";

const ISSUE_TYPES = {
  SKIP_VIOLATION: { label: "Skip Pattern Violation", color: "#ef4444", bg: "#fef2f2" },
  OUT_OF_RANGE:   { label: "Out of Range",            color: "#f97316", bg: "#fff7ed" },
  MISMATCHED_CODE:{ label: "Mismatched Code",         color: "#eab308", bg: "#fefce8" },
  MISSING_DATA:   { label: "Missing Data",            color: "#8b5cf6", bg: "#f5f3ff" },
} as const;

type IssueType = keyof typeof ISSUE_TYPES;

// ── SAV binary parser: extract variable names + Armenian/Russian labels ────────
function parseSavBinary(raw: string) {
  const vars: { name: string; label: string }[] = [];
  const seen = new Set<string>();

  function cleanLabel(s: string) {
    return s
      .replace(/[\x00-\x08\x0b\x0c\x0e-\x1f\x7f-\x9f]/g, " ")
      .replace(/[ÿþ]+/g, " ")
      .replace(/ð\?/g, " ")
      .replace(/@+/g, " ")
      .replace(/\s+/g, " ")
      .trim();
  }

  const chunks = raw.split(/(?=\b[A-Za-z_][A-Za-z0-9_.]{1,39}\b)/);
  for (const chunk of chunks) {
    const nm = chunk.match(/^([A-Za-z_][A-Za-z0-9_.]{1,39})/);
    if (!nm) continue;
    const name = nm[1];
    if (seen.has(name)) continue;
    if (/^(the|and|or|in|of|to|is|for|with|from|not|ÿ)$/i.test(name)) continue;

    const rest = cleanLabel(chunk.slice(name.length, name.length + 400));
    const parts = rest.match(/[\u0531-\u058F\u0400-\u04FF\w]{2,}(?:\s+[\u0531-\u058F\u0400-\u04FF\w]{2,})*/g) || [];
    const label = parts.join(" ").trim();

    seen.add(name);
    vars.push({ name, label: label && label !== name ? label : "" });
  }

  const valueLabels: Record<number, string> = {};
  const vlRe = /\b(\d{1,3})\b\s+([\u0531-\u058F\u0400-\u04FF][^\n\r\x00-\x1f]{2,60})/g;
  let m: RegExpExecArray | null;
  while ((m = vlRe.exec(raw)) !== null) {
    const code = parseInt(m[1]);
    if (!valueLabels[code]) valueLabels[code] = m[2].trim().replace(/\s+/g, " ");
  }

  return { vars, valueLabels };
}

// ── Questionnaire (.docx) parser ──────────────────────────────────────────────
interface VarDef {
  validValues: number[];
  range: [number, number] | null;
  skipRules: { condVar: string; op: string; condVal: string; skipTo: string }[];
  required: boolean;
}

function parseQuestionnaire(text: string): Record<string, VarDef> {
  const variables: Record<string, VarDef> = {};
  const lines = text.split("\n").map(l => l.trim()).filter(Boolean);
  let cur: string | null = null;
  const varRe    = /^([A-Za-z_][A-Za-z0-9_.]{0,39})\s*[:\-–]/;
  const rangeRe  = /range[:\s]+(\d+)\s*[-–to]+\s*(\d+)/i;
  const validRe  = /valid\s+(?:values?|codes?)[:\s]+([0-9,\s]+)/i;
  const optRe    = /^\s*(\d+)\s*[=\.\)]\s*.+/;
  const skipRe   = /if\s+(.+?)\s*(=|==|>|<|>=|<=)\s*([^\s,]+)\s*[,;]?\s*skip\s+to\s+([A-Za-z_][A-Za-z0-9_.]*)/gi;
  const reqRe    = /required|mandatory/i;

  for (const line of lines) {
    const vm = line.match(varRe);
    if (vm) { cur = vm[1]; variables[cur] = { validValues: [], range: null, skipRules: [], required: false }; }
    if (!cur) continue;
    const rm = line.match(rangeRe);
    if (rm) variables[cur].range = [+rm[1], +rm[2]];
    const vlm = line.match(validRe);
    if (vlm) variables[cur].validValues = vlm[1].split(",").map(v => +v.trim());
    const om = line.match(optRe);
    if (om && !variables[cur].validValues.includes(+om[1])) variables[cur].validValues.push(+om[1]);
    if (reqRe.test(line)) variables[cur].required = true;
    let sm: RegExpExecArray | null; skipRe.lastIndex = 0;
    while ((sm = skipRe.exec(line)) !== null)
      variables[cur].skipRules.push({ condVar: sm[1].trim(), op: sm[2], condVal: sm[3], skipTo: sm[4] });
  }
  return variables;
}

// ── Data validation ────────────────────────────────────────────────────────────
function evalCond(rowVal: unknown, op: string, condVal: string) {
  const a = isNaN(rowVal as number) ? rowVal : +(rowVal as number);
  const b = isNaN(+condVal) ? condVal : +condVal;
  switch (op) {
    case "=": case "==": return a == b;
    case ">": return (a as number) > (b as number);
    case "<": return (a as number) < (b as number);
    case ">=": return (a as number) >= (b as number);
    case "<=": return (a as number) <= (b as number);
    default: return false;
  }
}

interface Issue {
  id: string | number;
  variable: string;
  type: IssueType;
  value: unknown;
  detail: string;
}

function validateData(data: Record<string, unknown>[], variables: Record<string, VarDef>): Issue[] {
  const issues: Issue[] = [];
  const varKeys = Object.keys(variables);
  data.forEach((row, ri) => {
    const id = (row["ID"] || row["id"] || row["RespondentID"] || `Row ${ri + 1}`) as string | number;
    varKeys.forEach(vn => {
      if (!(vn in row)) return;
      const def = variables[vn];
      const val = row[vn];
      const empty = val === "" || val == null || val === "." || String(val).trim() === "";
      if (def.required && empty) { issues.push({ id, variable: vn, type: "MISSING_DATA", value: val, detail: "Required field is empty" }); return; }
      if (empty) return;
      const n = parseFloat(String(val));
      if (def.range && !isNaN(n) && (n < def.range[0] || n > def.range[1]))
        issues.push({ id, variable: vn, type: "OUT_OF_RANGE", value: val, detail: `${val} outside range [${def.range[0]}, ${def.range[1]}]` });
      if (def.validValues.length && !isNaN(n) && !def.validValues.includes(n))
        issues.push({ id, variable: vn, type: "MISMATCHED_CODE", value: val, detail: `${val} not in valid codes: ${def.validValues.join(", ")}` });
      def.skipRules.forEach(rule => {
        const cv = row[rule.condVar];
        if (cv === undefined) return;
        if (evalCond(cv, rule.op, rule.condVal)) {
          const vi = varKeys.indexOf(vn), si = varKeys.indexOf(rule.skipTo);
          if (si > vi) for (let i = vi; i < si; i++) {
            const sv = varKeys[i], sval = row[sv];
            const sempty = sval === "" || sval == null || String(sval).trim() === "";
            if (!sempty) issues.push({ id, variable: sv, type: "SKIP_VIOLATION", value: sval, detail: `Should be skipped (${rule.condVar} ${rule.op} ${rule.condVal}) but has value ${sval}` });
          }
        }
      });
    });
  });
  return issues;
}

// ── UI helpers ─────────────────────────────────────────────────────────────────
function FileZone({ label, sub, accept, onLoad, loaded }: {
  label: string; sub: string; accept: string;
  onLoad: (f: File) => void; loaded: string;
}) {
  const ref = useRef<HTMLInputElement>(null);
  const [drag, setDrag] = useState(false);
  return (
    <div onClick={() => ref.current?.click()}
      onDragOver={e => { e.preventDefault(); setDrag(true); }}
      onDragLeave={() => setDrag(false)}
      onDrop={e => { e.preventDefault(); setDrag(false); onLoad(e.dataTransfer.files[0]); }}
      style={{ border: `2px dashed ${drag ? "#2563eb" : loaded ? "#22c55e" : "#cbd5e1"}`, borderRadius: 12, padding: "18px 14px", textAlign: "center", cursor: "pointer", background: loaded ? "#f0fdf4" : drag ? "#eff6ff" : "#f8fafc", transition: "all .2s" }}>
      <input ref={ref} type="file" accept={accept} style={{ display: "none" }} onChange={e => e.target.files && onLoad(e.target.files[0])} />
      <div style={{ fontSize: 26, marginBottom: 5 }}>{loaded ? "✅" : "📂"}</div>
      <div style={{ fontSize: 13, fontWeight: 600, color: loaded ? "#16a34a" : "#475569" }}>{loaded || label}</div>
      <div style={{ fontSize: 11, color: "#94a3b8", marginTop: 3 }}>{sub}</div>
    </div>
  );
}

function Badge({ type }: { type: IssueType }) {
  const m = ISSUE_TYPES[type];
  return <span style={{ background: m.bg, color: m.color, padding: "2px 8px", borderRadius: 4, fontSize: 11, fontWeight: 700 }}>{m.label}</span>;
}

function Th({ children }: { children: React.ReactNode }) {
  return <th style={{ padding: "9px 13px", textAlign: "left", fontWeight: 600, color: "#475569", borderBottom: "1px solid #e2e8f0", whiteSpace: "nowrap", fontSize: 12 }}>{children}</th>;
}

// ── Main app ───────────────────────────────────────────────────────────────────
export default function App() {
  const [spssVars, setSpssVars]   = useState<{ name: string; label: string }[]>([]);
  const [valLabels, setValLabels] = useState<Record<number, string>>({});
  const [data, setData]           = useState<Record<string, unknown>[]>([]);
  const [dataFileName, setDataFileName] = useState("");
  const [savFileName, setSavFileName]   = useState("");
  const [questText, setQuestText] = useState("");
  const [questFileName, setQuestFileName] = useState("");
  const [variables, setVariables] = useState<Record<string, VarDef>>({});
  const [issues, setIssues]       = useState<Issue[]>([]);
  const [analyzed, setAnalyzed]   = useState(false);
  const [tab, setTab]             = useState("issues");
  const [filterType, setFilterType] = useState<IssueType | "ALL">("ALL");
  const [search, setSearch]       = useState("");
  const [loading, setLoading]     = useState(false);
  const [error, setError]         = useState("");

  const loadSpss = useCallback((file: File) => {
    if (!file) return; setError("");
    const ext = file.name.split(".").pop()?.toLowerCase();
    if (ext === "sav" || ext === "por") {
      const reader = new FileReader();
      reader.onload = e => {
        const { vars, valueLabels } = parseSavBinary(e.target!.result as string);
        setSpssVars(vars); setValLabels(valueLabels); setSavFileName(file.name);
      };
      reader.readAsBinaryString(file);
    } else if (ext === "csv") {
      Papa.parse<Record<string, unknown>>(file, {
        header: true, skipEmptyLines: true, dynamicTyping: true,
        complete: r => { setData(r.data); setDataFileName(file.name); },
        error: (e: { message: string }) => setError("CSV error: " + e.message),
      });
    } else {
      const reader = new FileReader();
      reader.onload = e => {
        try {
          const wb = XLSX.read(e.target!.result, { type: "array" });
          const ws = wb.Sheets[wb.SheetNames[0]];
          setData(XLSX.utils.sheet_to_json(ws, { defval: "" }));
          setDataFileName(file.name);
        } catch (err) { setError("Cannot parse file: " + (err as Error).message); }
      };
      reader.readAsArrayBuffer(file);
    }
  }, []);

  const loadDoc = useCallback(async (file: File) => {
    if (!file) return; setError("");
    try {
      const buf = await file.arrayBuffer();
      const r = await mammoth.extractRawText({ arrayBuffer: buf });
      setQuestText(r.value); setQuestFileName(file.name);
    } catch (e) { setError("Word doc error: " + (e as Error).message); }
  }, []);

  const analyze = () => {
    if (!data.length && !spssVars.length) { setError("Please upload a data file."); return; }
    setLoading(true); setError("");
    setTimeout(() => {
      const vars = questText ? parseQuestionnaire(questText) : {};
      const found = data.length ? validateData(data, vars) : [];
      setVariables(vars); setIssues(found); setAnalyzed(true);
      setTab(found.length ? "issues" : spssVars.length ? "savvars" : "data");
      setLoading(false);
    }, 80);
  };

  const typeCounts = (Object.keys(ISSUE_TYPES) as IssueType[]).reduce((a, t) => { a[t] = issues.filter(i => i.type === t).length; return a; }, {} as Record<IssueType, number>);
  const issueSet: Record<string, IssueType> = {};
  issues.forEach(i => { issueSet[`${i.id}__${i.variable}`] = i.type; });

  const filteredIssues = issues.filter(i =>
    (filterType === "ALL" || i.type === filterType) &&
    (!search || i.variable.toLowerCase().includes(search.toLowerCase()) || String(i.id).toLowerCase().includes(search.toLowerCase()))
  );
  const filteredSavVars = spssVars.filter(v =>
    !search || v.name.toLowerCase().includes(search.toLowerCase()) || v.label.toLowerCase().includes(search.toLowerCase())
  );
  const dataColumns = data.length ? Object.keys(data[0]) : [];

  const dl = (rows: object[], name: string) => {
    const csv = Papa.unparse(rows);
    const a = document.createElement("a"); a.href = URL.createObjectURL(new Blob([csv], { type: "text/csv" })); a.download = name; a.click();
  };

  const tabs = [
    issues.length > 0        && ["issues",  `🚩 Issues (${issues.length})`],
    data.length > 0          && ["data",    `📊 Data Table`],
    Object.keys(variables).length > 0 && ["qvars", `📋 Questionnaire Vars (${Object.keys(variables).length})`],
    spssVars.length > 0      && ["savvars", `🗂 SPSS Variables (${spssVars.length})`],
    Object.keys(valLabels).length > 0 && ["vallabels", `🏷 Value Labels (${Object.keys(valLabels).length})`],
  ].filter(Boolean) as [string, string][];

  return (
    <div style={{ fontFamily: "Inter, system-ui, sans-serif", minHeight: "100vh", background: "#f1f5f9" }}>

      {/* Header */}
      <div style={{ background: "linear-gradient(135deg,#1e3a5f,#2563eb)", padding: "18px 28px", color: "#fff" }}>
        <div style={{ maxWidth: 1120, margin: "0 auto", display: "flex", alignItems: "center", gap: 16 }}>
          <div style={{ flex: 1 }}>
            <h1 style={{ margin: 0, fontSize: 20, fontWeight: 700 }}>🔍 Questionnaire Logic Validator</h1>
            <p style={{ margin: "3px 0 0", fontSize: 12, opacity: .75 }}>Skip patterns · Out-of-range · Mismatched codes · Missing data — Armenian & English</p>
          </div>
          <div style={{ fontSize: 12, opacity: .65, textAlign: "right", lineHeight: 1.6 }}>
            {savFileName   && <div>🗃 {savFileName} — {spssVars.length} vars extracted</div>}
            {dataFileName  && <div>📋 {dataFileName} — {data.length} rows</div>}
            {questFileName && <div>📄 {questFileName}</div>}
          </div>
        </div>
      </div>

      <div style={{ maxWidth: 1120, margin: "0 auto", padding: "20px 16px" }}>

        {/* Upload row */}
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 12, marginBottom: 14 }}>
          <div>
            <div style={{ fontSize: 11, fontWeight: 700, color: "#64748b", marginBottom: 6, textTransform: "uppercase", letterSpacing: .5 }}>1 — SPSS / Data file</div>
            <FileZone label="Drop SPSS, CSV or Excel" sub=".sav · .csv · .xlsx" onLoad={loadSpss} loaded={savFileName || dataFileName} accept=".sav,.por,.csv,.xlsx,.xls" />
          </div>
          <div>
            <div style={{ fontSize: 11, fontWeight: 700, color: "#64748b", marginBottom: 6, textTransform: "uppercase", letterSpacing: .5 }}>2 — Questionnaire (optional)</div>
            <FileZone label="Drop Word questionnaire" sub=".docx — Armenian or English" onLoad={loadDoc} loaded={questFileName} accept=".docx" />
          </div>
          <div style={{ display: "flex", flexDirection: "column", justifyContent: "flex-end" }}>
            <button onClick={analyze} disabled={loading || (!data.length && !spssVars.length)}
              style={{ padding: "14px", background: loading ? "#94a3b8" : (!data.length && !spssVars.length) ? "#cbd5e1" : "#2563eb", color: "#fff", border: "none", borderRadius: 10, fontSize: 14, fontWeight: 700, cursor: "pointer", width: "100%", transition: "background .2s" }}>
              {loading ? "⏳ Analyzing…" : "🚀 Run Validation"}
            </button>
            {spssVars.length > 0 && !data.length && (
              <div style={{ marginTop: 10, background: "#fffbeb", border: "1px solid #fde68a", borderRadius: 8, padding: "8px 12px", fontSize: 12, color: "#92400e" }}>
                ⚠️ <strong>.sav detected</strong> — variable list extracted. For row-level validation, also upload a <strong>CSV export</strong> from SPSS.
              </div>
            )}
          </div>
        </div>

        {error && <div style={{ background: "#fef2f2", border: "1px solid #fca5a5", borderRadius: 8, padding: "10px 14px", color: "#dc2626", fontSize: 13, marginBottom: 12 }}>⚠️ {error}</div>}

        {analyzed && (
          <>
            {/* Summary cards */}
            <div style={{ display: "grid", gridTemplateColumns: "repeat(4,1fr)", gap: 10, marginBottom: 18 }}>
              {(Object.entries(ISSUE_TYPES) as [IssueType, typeof ISSUE_TYPES[IssueType]][]).map(([type, meta]) => (
                <div key={type} onClick={() => setFilterType(filterType === type ? "ALL" : type)}
                  style={{ background: "#fff", border: `2px solid ${filterType === type ? meta.color : "#e2e8f0"}`, borderRadius: 10, padding: "12px 14px", cursor: "pointer", transition: "border .15s", boxShadow: filterType === type ? `0 0 0 3px ${meta.bg}` : "none" }}>
                  <div style={{ fontSize: 24, fontWeight: 800, color: typeCounts[type] > 0 ? meta.color : "#cbd5e1" }}>{typeCounts[type]}</div>
                  <div style={{ fontSize: 11, color: "#64748b", marginTop: 2 }}>{meta.label}</div>
                </div>
              ))}
            </div>

            {/* Tabs + actions */}
            <div style={{ display: "flex", gap: 2, borderBottom: "2px solid #e2e8f0", marginBottom: 14, flexWrap: "wrap" }}>
              {tabs.map(([id, label]) => (
                <button key={id} onClick={() => setTab(id)}
                  style={{ padding: "8px 14px", border: "none", background: "none", borderBottom: tab === id ? "2px solid #2563eb" : "2px solid transparent", color: tab === id ? "#2563eb" : "#64748b", fontWeight: tab === id ? 700 : 400, cursor: "pointer", fontSize: 13, marginBottom: -2 }}>
                  {label}
                </button>
              ))}
              <div style={{ flex: 1 }} />
              {issues.length > 0 && (
                <button onClick={() => dl(issues.map(i => ({ ID: i.id, Variable: i.variable, Type: ISSUE_TYPES[i.type].label, Value: i.value, Detail: i.detail })), "issues_report.csv")}
                  style={{ fontSize: 12, padding: "6px 10px", background: "#f1f5f9", border: "1px solid #e2e8f0", borderRadius: 6, cursor: "pointer", color: "#475569" }}>⬇ Issues CSV</button>
              )}
              {data.length > 0 && (
                <button onClick={() => {
                  const marked = data.map(row => {
                    const id = row["ID"] || row["id"] || row["RespondentID"];
                    const extra: Record<string, string> = {};
                    Object.keys(row).forEach(col => { const k = `${id}__${col}`; if (issueSet[k]) extra[`${col}_FLAG`] = ISSUE_TYPES[issueSet[k]].label; });
                    return { ...row, ...extra };
                  });
                  dl(marked, "flagged_data.csv");
                }} style={{ fontSize: 12, padding: "6px 10px", background: "#f1f5f9", border: "1px solid #e2e8f0", borderRadius: 6, cursor: "pointer", color: "#475569", marginLeft: 4 }}>⬇ Flagged Data CSV</button>
              )}
              {spssVars.length > 0 && (
                <button onClick={() => dl(spssVars.map(v => ({ Variable: v.name, Label: v.label })), "variable_list.csv")}
                  style={{ fontSize: 12, padding: "6px 10px", background: "#f1f5f9", border: "1px solid #e2e8f0", borderRadius: 6, cursor: "pointer", color: "#475569", marginLeft: 4 }}>⬇ Variables CSV</button>
              )}
            </div>

            {/* Search */}
            <input placeholder="Search by variable name, label, or respondent ID…"
              value={search} onChange={e => setSearch(e.target.value)}
              style={{ width: "100%", padding: "8px 12px", border: "1px solid #e2e8f0", borderRadius: 8, fontSize: 13, marginBottom: 12, boxSizing: "border-box" }} />

            {/* ── Issues tab ── */}
            {tab === "issues" && (
              filteredIssues.length === 0
                ? <div style={{ background: "#fff", borderRadius: 10, padding: 48, textAlign: "center", color: "#94a3b8" }}>
                    <div style={{ fontSize: 44, marginBottom: 10 }}>{issues.length === 0 ? "✅" : "🔍"}</div>
                    <div style={{ fontWeight: 600, fontSize: 15 }}>{issues.length === 0 ? "No logic issues found — data looks clean!" : "No issues match your current filter."}</div>
                  </div>
                : <div style={{ background: "#fff", borderRadius: 10, border: "1px solid #e2e8f0", overflow: "auto" }}>
                    <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
                      <thead><tr style={{ background: "#f8fafc" }}>
                        <Th>Respondent ID</Th><Th>Variable</Th><Th>Issue Type</Th><Th>Value</Th><Th>Details</Th>
                      </tr></thead>
                      <tbody>
                        {filteredIssues.map((iss, i) => (
                          <tr key={i} style={{ background: i%2===0?"#fff":"#fafafa", borderLeft: `3px solid ${ISSUE_TYPES[iss.type].color}` }}>
                            <td style={{ padding: "8px 13px", fontWeight: 500 }}>{iss.id}</td>
                            <td style={{ padding: "8px 13px", fontFamily: "monospace", fontSize: 12 }}>{iss.variable}</td>
                            <td style={{ padding: "8px 13px" }}><Badge type={iss.type} /></td>
                            <td style={{ padding: "8px 13px", fontFamily: "monospace", color: "#7c3aed" }}>{String(iss.value ?? "")}</td>
                            <td style={{ padding: "8px 13px", color: "#475569" }}>{iss.detail}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                    <div style={{ padding: "7px 13px", fontSize: 12, color: "#94a3b8", borderTop: "1px solid #e2e8f0" }}>Showing {filteredIssues.length} of {issues.length} issues</div>
                  </div>
            )}

            {/* ── Data table tab ── */}
            {tab === "data" && (
              <div style={{ background: "#fff", borderRadius: 10, border: "1px solid #e2e8f0", overflow: "auto", maxHeight: 520 }}>
                <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
                  <thead style={{ position: "sticky", top: 0, background: "#f8fafc", zIndex: 1 }}>
                    <tr>{dataColumns.map(col => <Th key={col}>{col}</Th>)}</tr>
                  </thead>
                  <tbody>
                    {data.slice(0, 200).map((row, ri) => {
                      const id = row["ID"] || row["id"] || row["RespondentID"] || `Row ${ri+1}`;
                      return (
                        <tr key={ri} style={{ background: ri%2===0?"#fff":"#fafafa" }}>
                          {dataColumns.map(col => {
                            const it = issueSet[`${id}__${col}`];
                            const meta = it ? ISSUE_TYPES[it] : null;
                            return (
                              <td key={col} title={meta ? meta.label : ""}
                                style={{ padding: "6px 11px", background: meta ? meta.bg : "transparent", color: meta ? meta.color : "#334155", fontFamily: "monospace", borderBottom: "1px solid #f1f5f9", borderLeft: meta ? `2px solid ${meta.color}` : "none", whiteSpace: "nowrap" }}>
                                {String(row[col] ?? "")}
                              </td>
                            );
                          })}
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
                {data.length > 200 && <div style={{ padding: "7px 13px", fontSize: 12, color: "#94a3b8", borderTop: "1px solid #e2e8f0" }}>Showing first 200 of {data.length} rows</div>}
              </div>
            )}

            {/* ── Questionnaire vars tab ── */}
            {tab === "qvars" && (
              <div style={{ background: "#fff", borderRadius: 10, border: "1px solid #e2e8f0", overflow: "hidden" }}>
                {Object.keys(variables).length === 0
                  ? <div style={{ padding: 32, textAlign: "center", color: "#94a3b8" }}>No variables parsed — check that variable names appear at line starts followed by ":" or "-".</div>
                  : <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
                      <thead><tr style={{ background: "#f8fafc" }}>
                        <Th>Variable</Th><Th>Valid Codes</Th><Th>Range</Th><Th>Required</Th><Th>Skip Rules</Th>
                      </tr></thead>
                      <tbody>
                        {Object.entries(variables).map(([name, def], i) => (
                          <tr key={name} style={{ background: i%2===0?"#fff":"#fafafa" }}>
                            <td style={{ padding: "8px 13px", fontFamily: "monospace", fontWeight: 700 }}>{name}</td>
                            <td style={{ padding: "8px 13px", color: "#475569" }}>{def.validValues.length ? def.validValues.join(", ") : "—"}</td>
                            <td style={{ padding: "8px 13px", color: "#475569" }}>{def.range ? `${def.range[0]}–${def.range[1]}` : "—"}</td>
                            <td style={{ padding: "8px 13px" }}>{def.required ? <span style={{ color: "#ef4444", fontWeight: 700 }}>Yes</span> : "No"}</td>
                            <td style={{ padding: "8px 13px", fontSize: 12, color: "#475569" }}>{def.skipRules.length ? def.skipRules.map((r,j) => <div key={j}>If {r.condVar} {r.op} {r.condVal} → skip to {r.skipTo}</div>) : "—"}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                }
              </div>
            )}

            {/* ── SPSS variables tab ── */}
            {tab === "savvars" && (
              <div style={{ background: "#fff", borderRadius: 10, border: "1px solid #e2e8f0", overflow: "auto", maxHeight: 520 }}>
                <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
                  <thead style={{ position: "sticky", top: 0, background: "#f8fafc", zIndex: 1 }}>
                    <tr><Th>#</Th><Th>Variable Name</Th><Th>Label (extracted from .sav)</Th></tr>
                  </thead>
                  <tbody>
                    {filteredSavVars.slice(0, 400).map((v, i) => (
                      <tr key={v.name} style={{ background: i%2===0?"#fff":"#fafafa" }}>
                        <td style={{ padding: "7px 13px", color: "#94a3b8", fontSize: 11 }}>{i+1}</td>
                        <td style={{ padding: "7px 13px", fontFamily: "monospace", fontWeight: 700, color: "#1e293b", fontSize: 12 }}>{v.name}</td>
                        <td style={{ padding: "7px 13px", color: v.label ? "#334155" : "#cbd5e1", fontStyle: v.label ? "normal" : "italic" }}>
                          {v.label || "no label extracted"}
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
                {filteredSavVars.length > 400 && <div style={{ padding: "7px 13px", fontSize: 12, color: "#94a3b8", borderTop: "1px solid #e2e8f0" }}>Showing 400 of {filteredSavVars.length}</div>}
              </div>
            )}

            {/* ── Value labels tab ── */}
            {tab === "vallabels" && (
              <div style={{ background: "#fff", borderRadius: 10, border: "1px solid #e2e8f0", overflow: "hidden" }}>
                <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
                  <thead><tr style={{ background: "#f8fafc" }}><Th>Code</Th><Th>Label</Th></tr></thead>
                  <tbody>
                    {Object.entries(valLabels).sort((a,b)=>+a[0]-+b[0]).map(([code, label], i) => (
                      <tr key={code} style={{ background: i%2===0?"#fff":"#fafafa" }}>
                        <td style={{ padding: "7px 13px", fontFamily: "monospace", fontWeight: 700, color: "#2563eb", width: 80 }}>{code}</td>
                        <td style={{ padding: "7px 13px", color: "#334155" }}>{label}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )}
          </>
        )}

        {!analyzed && (
          <div style={{ textAlign: "center", padding: "56px 20px", color: "#94a3b8" }}>
            <div style={{ fontSize: 52, marginBottom: 14 }}>📋</div>
            <div style={{ fontSize: 16, fontWeight: 600, color: "#64748b" }}>Upload your files to begin</div>
            <div style={{ fontSize: 13, marginTop: 6, maxWidth: 440, margin: "8px auto 0" }}>
              Drop a <strong>.sav</strong> file — variable names and Armenian labels are extracted automatically, no garbled text shown.<br />
              Add a <strong>CSV export</strong> for full row-by-row logic validation.
            </div>
          </div>
        )}
      </div>
    </div>
  );
}
