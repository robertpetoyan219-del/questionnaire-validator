import { useState, useCallback, useRef } from "react";
import type { ReactNode } from "react";
import * as XLSX from "xlsx";
import Papa from "papaparse";
import mammoth from "mammoth";

// ─────────────────────────────────────────────────────────────────────────────
// TYPES
// ─────────────────────────────────────────────────────────────────────────────

const ISSUE_TYPES = {
  SKIP_VIOLATION:  { label: "Skip/Routing Violation", color: "#ef4444", bg: "#fef2f2" },
  OUT_OF_RANGE:    { label: "Out of Range",            color: "#f97316", bg: "#fff7ed" },
  MISMATCHED_CODE: { label: "Mismatched Code",         color: "#eab308", bg: "#fefce8" },
  MISSING_DATA:    { label: "Missing Data",            color: "#8b5cf6", bg: "#f5f3ff" },
  DATA_QUALITY:    { label: "Data Quality",            color: "#06b6d4", bg: "#ecfeff" },
  OPEN_TEXT_ISSUE: { label: "Open Text Issue",         color: "#10b981", bg: "#f0fdf4" },
} as const;
type IssueType = keyof typeof ISSUE_TYPES;

interface Issue {
  id: string | number;
  variable: string;
  type: IssueType;
  value: unknown;
  detail: string;
  explanation: string;
}

interface DatasetWarning {
  type: string;
  variable: string;
  detail: string;
  explanation: string;
}

// SAV: per-variable metadata parsed from binary
interface SavVariable {
  name: string;
  label: string;               // variable label (English/Armenian question text)
  valueLabels: Record<number, string>; // code → label
  validCodes: number[];        // derived from valueLabels keys (excluding system missing)
  type: "numeric" | "string";
  missingValues: number[];     // explicitly declared missing values
}

// Docx: a single parsed routing rule
interface RoutingRule {
  condVar: string;   // e.g. "S4"
  condOp: string;    // "=", "!=", "<", ">", "<=", ">="
  condVals: number[];// e.g. [1]  or [9,10]
  targets: string[]; // variables that should be asked
  skipTargets: string[]; // variables that should be SKIPPED
  rawText: string;
}

// Docx: question metadata
interface DocxQuestion {
  code: string;          // variable name/code e.g. "S4", "E14"
  label: string;         // question text
  validCodes: number[];  // answer codes parsed from docx
  codeLabels: Record<number, string>;
  section: string;       // section name
}

type RowData = Record<string, unknown>;

// ─────────────────────────────────────────────────────────────────────────────
// SAV BINARY PARSER
// Parses SPSS .sav (System file) format:
// Record type 2 = variable records, 3/4 = value label records, 7 = info records
// ─────────────────────────────────────────────────────────────────────────────

function parseSavFile(buffer: ArrayBuffer): {
  variables: SavVariable[];
  varMap: Record<string, SavVariable>;
} {
  const bytes = new Uint8Array(buffer);
  const latin1 = new TextDecoder("latin1");

  const readI32 = (o: number) =>
    (bytes[o] | (bytes[o+1]<<8) | (bytes[o+2]<<16) | (bytes[o+3]<<24)) >>> 0;
  const readI32s = (o: number) => // signed
    bytes[o] | (bytes[o+1]<<8) | (bytes[o+2]<<16) | (bytes[o+3]<<24);
  const readF64 = (o: number) => new DataView(buffer, o, 8).getFloat64(0, true);
  const readStr = (o: number, n: number) => latin1.decode(bytes.slice(o, o+n)).trimEnd();
  const pad4 = (n: number) => Math.ceil(n / 4) * 4;

  // Verify SAV magic "$FL2" or "$FL3"
  const magic = readStr(0, 4);
  if (!magic.startsWith("$FL")) {
    // fallback: legacy ASCII scan
    return legacySavScan(buffer);
  }

  const variables: SavVariable[] = [];
  const varBySlot: (SavVariable | null)[] = [];
  let slot = 0;
  let offset = 176; // after 176-byte header
  const maxOff = buffer.byteLength - 4;

  while (offset < maxOff) {
    const recType = readI32s(offset); offset += 4;

    if (recType === 999) break; // end of dictionary

    if (recType === 2) {
      // Variable record: 28 fixed bytes + optional label + optional missing values
      if (offset + 28 > buffer.byteLength) break;
      const typeCode  = readI32s(offset);     // 0=numeric, >0=string width
      const hasLabel  = readI32s(offset + 4);
      const nMissing  = readI32s(offset + 8); // 0, ±1, ±2, ±3
      const rawName   = readStr(offset + 20, 8).trim();
      offset += 28;

      let varLabel = "";
      if (hasLabel !== 0) {
        if (offset + 4 > buffer.byteLength) break;
        const lblLen = readI32s(offset); offset += 4;
        varLabel = readStr(offset, lblLen);
        offset += pad4(lblLen);
      }

      const absMissing = Math.abs(nMissing);
      const missingVals: number[] = [];
      for (let i = 0; i < absMissing; i++) {
        missingVals.push(readF64(offset));
        offset += 8;
      }

      // Register variable (skip continuation slots — they have empty names)
      if (rawName !== "") {
        const sv: SavVariable = {
          name: rawName,
          label: varLabel,
          valueLabels: {},
          validCodes: [],
          type: typeCode === 0 ? "numeric" : "string",
          missingValues: missingVals,
        };
        variables.push(sv);
        varBySlot[slot] = sv;
      } else {
        varBySlot[slot] = null; // continuation slot
      }
      slot++;

      // Long strings span extra slots
      if (typeCode > 8) {
        const extra = Math.ceil(typeCode / 8) - 1;
        for (let i = 0; i < extra; i++) { varBySlot[slot++] = null; }
      }

    } else if (recType === 3) {
      // Value labels record
      if (offset + 4 > buffer.byteLength) break;
      const nLabels = readI32s(offset); offset += 4;
      const tempLabels: Record<number, string> = {};

      for (let i = 0; i < nLabels; i++) {
        if (offset + 9 > buffer.byteLength) break;
        const code = readF64(offset); offset += 8;
        const lblLen = bytes[offset]; offset++;
        const lbl = readStr(offset, lblLen);
        offset += Math.ceil((lblLen + 1) / 8) * 8 - 1;
        tempLabels[code] = lbl;
      }

      // Followed immediately by record type 4 (variable index list)
      if (offset + 4 > buffer.byteLength) break;
      const next = readI32s(offset); offset += 4;
      if (next === 4) {
        if (offset + 4 > buffer.byteLength) break;
        const nVars = readI32s(offset); offset += 4;
        for (let i = 0; i < nVars; i++) {
          if (offset + 4 > buffer.byteLength) break;
          const idx = readI32s(offset) - 1; offset += 4; // 1-based
          const sv = varBySlot[idx];
          if (sv) {
            Object.assign(sv.valueLabels, tempLabels);
          }
        }
      }

    } else if (recType === 6) {
      // Document record — skip
      if (offset + 4 > buffer.byteLength) break;
      const nLines = readI32s(offset); offset += 4;
      offset += nLines * 80;

    } else if (recType === 7) {
      // Info/extension record — skip
      if (offset + 8 > buffer.byteLength) break;
      const elemSize = readI32s(offset + 4); offset += 8;
      if (offset + 4 > buffer.byteLength) break;
      const nElems = readI32s(offset); offset += 4;
      offset += elemSize * nElems;

    } else {
      break; // Unknown record type
    }
  }

  // Build validCodes from valueLabels (exclude system-missing 1.7976931348623157e+308)
  const SYSMIS = 1.7976931348623157e+308;
  for (const sv of variables) {
    sv.validCodes = Object.keys(sv.valueLabels)
      .map(Number)
      .filter(n => !isNaN(n) && n !== SYSMIS && !sv.missingValues.includes(n))
      .sort((a, b) => a - b);
  }

  const varMap: Record<string, SavVariable> = {};
  for (const sv of variables) varMap[sv.name] = sv;

  if (variables.length === 0) return legacySavScan(buffer);
  return { variables, varMap };
}

function legacySavScan(buffer: ArrayBuffer): { variables: SavVariable[]; varMap: Record<string, SavVariable> } {
  const bytes = new Uint8Array(buffer);
  const SKIP = new Set(["the","and","or","in","of","to","is","for","with","from","not","are","at","by",
    "this","that","be","as","it","on","if","do","so","SPSS","DATA","LIST","SAVE","GET","COMPUTE",
    "RECODE","EXECUTE","VALUE","VARIABLE","LABELS","TYPE","STRING","NUMERIC","MISSING","SYSMIS",
    "ALL","EQ","NE","LT","GT","LE","GE","SUM","MEAN","BEGIN","END","FILE","NAME","FORMAT",
    "MEASURE","ROLE","NOMINAL","SCALE","ORDINAL","INPUT","OUTPUT","YES","NO"]);
  let ascii = "";
  for (let i = 0; i < Math.min(bytes.length, 300000); i++)
    ascii += bytes[i] < 128 ? String.fromCharCode(bytes[i]) : " ";
  const seen = new Set<string>(), variables: SavVariable[] = [];
  const re = /\b([A-Za-z][A-Za-z0-9_.]{1,39})\b/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(ascii)) !== null) {
    const nm = m[1];
    if (!seen.has(nm) && !SKIP.has(nm) && !SKIP.has(nm.toLowerCase())) {
      seen.add(nm);
      variables.push({ name: nm, label: "", valueLabels: {}, validCodes: [], type: "numeric", missingValues: [] });
    }
  }
  const varMap: Record<string, SavVariable> = {};
  for (const sv of variables) varMap[sv.name] = sv;
  return { variables, varMap };
}

// ─────────────────────────────────────────────────────────────────────────────
// DOCX PARSER (mammoth)
// Extracts: question labels, valid codes per question, routing rules
// ─────────────────────────────────────────────────────────────────────────────

async function parseDocxFile(buffer: ArrayBuffer): Promise<{
  questions: DocxQuestion[];
  questionMap: Record<string, DocxQuestion>;
  routingRules: RoutingRule[];
  rawText: string;
  currentSection: string;
}> {
  const result = await mammoth.extractRawText({ arrayBuffer: buffer });
  const rawText = result.value;
  const lines = rawText.split(/\r?\n/).map(l => l.trim()).filter(Boolean);

  const questions: DocxQuestion[] = [];
  const questionMap: Record<string, DocxQuestion> = {};
  const routingRules: RoutingRule[] = [];

  // Section detection
  const sectionKeywords = [
    "SCREENING", "SECTION", "DEMOGRAPH", "INFORMATION SOURCE",
    "AWARENESS", "ATTITUDE", "EU IMAGE", "EU DELEGATION",
    "INTERNATIONAL ORGAN", "FUNDED PROJECT",
  ];
  let currentSection = "General";

  // Question code pattern: starts with uppercase letter, then digits/underscore, followed by a separator
  const qCodeRe = /^([A-Z][A-Za-z0-9_.]{0,19})\s*[-–—.:)]\s*(.*)/;

  // Routing patterns — expanded to cover diverse questionnaire styles
  const routingPatterns: Array<{
    re: RegExp;
    parse: (m: RegExpMatchArray) => Partial<RoutingRule> | null;
  }> = [
    // "ASK IF X=1" / "ASK IF X=1-4" / "ASK IF X=1,2,3"
    {
      re: /ASK\s+IF\s+([A-Z][A-Za-z0-9_.]{0,19})\s*([=!<>]{1,2})\s*([\d,\-–]+)/i,
      parse: m => ({ condVar: m[1], condOp: m[2], condVals: parseCondVals(m[3]) }),
    },
    // "IF X=1 ASK Y" / "IF X=1-4 ASK Y,Z"
    {
      re: /IF\s+([A-Z][A-Za-z0-9_.]{0,19})\s*([=!<>]{1,2})\s*([\d,\-–]+)\s+(?:ASK|GO TO|SHOW|DISPLAY)\s+(.*)/i,
      parse: m => ({ condVar: m[1], condOp: m[2], condVals: parseCondVals(m[3]), targets: extractVarNames(m[4]) }),
    },
    // "IF X=1 SKIP TO Y" or "SKIP X IF Y≠1"
    {
      re: /SKIP\s+(?:TO\s+)?([A-Z][A-Za-z0-9_.]{0,19})\s+IF\s+([A-Z][A-Za-z0-9_.]{0,19})\s*([=!<>]{1,2})\s*([\d,\-–]+)/i,
      parse: m => ({ condVar: m[2], condOp: m[3], condVals: parseCondVals(m[4]), skipTargets: [m[1]] }),
    },
    // "IF X=1 END INTERVIEW" / "TERMINATE IF X=1"
    {
      re: /(?:IF\s+([A-Z][A-Za-z0-9_.]{0,19})\s*([=!<>]{1,2})\s*([\d,\-–]+)\s+(?:END|TERMINATE|CLOSE|STOP))|(?:(?:END|TERMINATE)\s+IF\s+([A-Z][A-Za-z0-9_.]{0,19})\s*([=!<>]{1,2})\s*([\d,\-–]+))/i,
      parse: m => m[1]
        ? { condVar: m[1], condOp: m[2], condVals: parseCondVals(m[3]), targets: ["TERMINATE"] }
        : { condVar: m[4], condOp: m[5], condVals: parseCondVals(m[6]), targets: ["TERMINATE"] },
    },
    // Arrow: "X=1 → Y" / "X=1-4 → Y, Z"
    {
      re: /([A-Z][A-Za-z0-9_.]{0,19})\s*([=!<>]{1,2})\s*([\d,\-–]+)\s*[→➔>]+\s*(.*)/,
      parse: m => ({ condVar: m[1], condOp: m[2], condVals: parseCondVals(m[3]), targets: extractVarNames(m[4]) }),
    },
    // "Control: X=1-4" (condition-only, used as filter marker)
    {
      re: /[Cc]ontrol\s*[:\-]\s*([A-Z][A-Za-z0-9_.]{0,19})\s*([=!<>]{1,2})\s*([\d,\-–]+)/,
      parse: m => ({ condVar: m[1], condOp: m[2], condVals: parseCondVals(m[3]) }),
    },
  ];

  // Parse condition values like "1-4", "1,2,3", "9-10", "3-4"
  function parseCondVals(raw: string): number[] {
    const vals = new Set<number>();
    for (const part of raw.split(",")) {
      const rangeMatch = part.trim().match(/^(\d+)\s*[-–]\s*(\d+)$/);
      if (rangeMatch) {
        const lo = parseInt(rangeMatch[1]), hi = parseInt(rangeMatch[2]);
        for (let v = lo; v <= hi; v++) vals.add(v);
      } else {
        const n = parseInt(part.trim());
        if (!isNaN(n)) vals.add(n);
      }
    }
    return [...vals];
  }

  function extractVarNames(text: string): string[] {
    const out: string[] = [];
    const re2 = /\b([A-Z][A-Za-z0-9_.]{0,19})\b/g;
    let m2: RegExpExecArray | null;
    while ((m2 = re2.exec(text)) !== null) {
      const nm = m2[1];
      if (nm.length >= 2 && nm !== "ASK" && nm !== "GO" && nm !== "TO" && nm !== "IF" && nm !== "AND" && nm !== "OR") {
        out.push(nm);
      }
    }
    return out;
  }

  // Code list pattern: lines like "1 = Yes" / "1=Yes" / "1. Yes"
  const codeLabelRe = /^\s*(\d+)\s*[=.\-:)]\s*(.+)$/;

  let currentQ: DocxQuestion | null = null;

  for (const line of lines) {
    // ── Section detection ────────────────────────────────────────────────
    const upperLine = line.toUpperCase();
    if (sectionKeywords.some(kw => upperLine.includes(kw))) {
      currentSection = line.slice(0, 80);
    }

    // ── Routing rule extraction ──────────────────────────────────────────
    for (const { re, parse } of routingPatterns) {
      const m = line.match(re);
      if (m) {
        const partial = parse(m);
        if (partial && partial.condVar) {
          routingRules.push({
            condVar: partial.condVar!,
            condOp: partial.condOp ?? "=",
            condVals: partial.condVals ?? [],
            targets: partial.targets ?? [],
            skipTargets: partial.skipTargets ?? [],
            rawText: line.slice(0, 200),
          });
        }
        break;
      }
    }

    // ── Question code detection ──────────────────────────────────────────
    const qm = line.match(qCodeRe);
    if (qm && qm[1].length >= 1 && qm[1].length <= 12) {
      // Plausibility: code must look like a survey variable (not a sentence)
      const code = qm[1];
      const label = qm[2].trim();
      if (/^[A-Z][A-Za-z0-9_.]*$/.test(code)) {
        currentQ = {
          code,
          label: label.slice(0, 300),
          validCodes: [],
          codeLabels: {},
          section: currentSection,
        };
        questions.push(currentQ);
        questionMap[code] = currentQ;
      }
    }

    // ── Code/answer option extraction (under current question) ───────────
    if (currentQ) {
      const cm = line.match(codeLabelRe);
      if (cm) {
        const code = parseInt(cm[1]);
        const lbl = cm[2].trim();
        if (!isNaN(code) && code >= 0 && code <= 99999 && lbl.length > 0) {
          currentQ.codeLabels[code] = lbl;
          if (!currentQ.validCodes.includes(code)) currentQ.validCodes.push(code);
        }
      }
    }
  }

  return { questions, questionMap, routingRules, rawText, currentSection };
}

// ─────────────────────────────────────────────────────────────────────────────
// DYNAMIC VALIDATION ENGINE
// No hardcoded survey knowledge. All rules derive from SAV + docx metadata.
// ─────────────────────────────────────────────────────────────────────────────

function buildValidCodes(
  varName: string,
  savVarMap: Record<string, SavVariable>,
  docxQMap: Record<string, DocxQuestion>,
): number[] | null {
  // Priority 1: SAV value labels (most authoritative)
  const sv = savVarMap[varName];
  if (sv && sv.validCodes.length > 0) return sv.validCodes;

  // Priority 2: docx code list
  const dq = docxQMap[varName];
  if (dq && dq.validCodes.length > 0) return dq.validCodes;

  return null; // unknown
}

function getVarLabel(
  varName: string,
  savVarMap: Record<string, SavVariable>,
  docxQMap: Record<string, DocxQuestion>,
): string {
  const sv = savVarMap[varName];
  if (sv?.label) return sv.label;
  const dq = docxQMap[varName];
  if (dq?.label) return dq.label;
  return varName;
}

function getCodeLabel(
  varName: string,
  code: number,
  savVarMap: Record<string, SavVariable>,
  docxQMap: Record<string, DocxQuestion>,
): string {
  const sv = savVarMap[varName];
  if (sv?.valueLabels[code]) return sv.valueLabels[code];
  const dq = docxQMap[varName];
  if (dq?.codeLabels[code]) return dq.codeLabels[code];
  return "";
}

// Build a map: condVar → list of routing rules, so we can look up "what does X=1 imply"
function buildRoutingIndex(rules: RoutingRule[]): Record<string, RoutingRule[]> {
  const idx: Record<string, RoutingRule[]> = {};
  for (const r of rules) {
    (idx[r.condVar] ??= []).push(r);
  }
  return idx;
}

// Given a variable, find which conditions gate it (i.e. rules whose targets include it)
function findGatingRules(
  varName: string,
  rules: RoutingRule[],
): RoutingRule[] {
  return rules.filter(r => r.targets.includes(varName) || r.skipTargets.includes(varName));
}

function runDynamicValidation(
  data: RowData[],
  savVars: SavVariable[],
  savVarMap: Record<string, SavVariable>,
  docxQMap: Record<string, DocxQuestion>,
  routingRules: RoutingRule[],
): { issues: Issue[]; datasetWarnings: DatasetWarning[]; columnSummary: ColumnSummary[] } {

  const issues: Issue[] = [];
  const datasetWarnings: DatasetWarning[] = [];
  const n = data.length;
  if (n === 0) return { issues, datasetWarnings, columnSummary: [] };

  const dataColumns = new Set(Object.keys(data[0]));
  const routingIndex = buildRoutingIndex(routingRules);

  // ── Helpers ─────────────────────────────────────────────────────────────

  const getNum = (row: RowData, col: string): number | null => {
    const raw = row[col];
    if (raw === null || raw === undefined || raw === "" || raw === ".") return null;
    const n = Number(String(raw).trim());
    return isNaN(n) ? null : n;
  };
  const getStr = (row: RowData, col: string): string | null => {
    const raw = row[col];
    if (raw == null) return null;
    const s = String(raw).trim();
    return s === "" || s === "." ? null : s;
  };
  const hasVal = (row: RowData, col: string) => getNum(row, col) !== null || getStr(row, col) !== null;
  const emptyVal = (row: RowData, col: string) => !hasVal(row, col);
  const eqVal = (row: RowData, col: string, v: number) => getNum(row, col) === v;
  const inVals = (row: RowData, col: string, list: number[]) => {
    const x = getNum(row, col); return x !== null && list.includes(x);
  };

  function evalCondition(row: RowData, rule: RoutingRule): boolean {
    const x = getNum(row, rule.condVar);
    if (x === null) return false;
    switch (rule.condOp) {
      case "=":  return rule.condVals.includes(x);
      case "!=": return !rule.condVals.includes(x);
      case "<":  return x < (rule.condVals[0] ?? 0);
      case ">":  return x > (rule.condVals[0] ?? 0);
      case "<=": return x <= (rule.condVals[0] ?? 0);
      case ">=": return x >= (rule.condVals[0] ?? 0);
      default:   return false;
    }
  }

  const flag = (id: unknown, variable: string, type: IssueType, value: unknown, detail: string, explanation: string) => {
    issues.push({ id: id as string | number, variable, type, value, detail, explanation });
  };

  // ── Dataset-level: SAV vars not in CSV ──────────────────────────────────
  const criticalSavVars = savVars.filter(sv =>
    sv.type === "numeric" &&
    sv.validCodes.length > 0 &&
    !dataColumns.has(sv.name)
  );
  // Report only first 10 to avoid noise
  for (const sv of criticalSavVars.slice(0, 10)) {
    datasetWarnings.push({
      type: "STRUCTURAL",
      variable: sv.name,
      detail: `Column "${sv.name}" is defined in the SAV file but MISSING from the data export`,
      explanation: `The SAV file defines variable "${sv.name}" (${sv.label || "no label"}) with ${sv.validCodes.length} valid codes, but this column does not exist in the uploaded CSV/Excel data. This indicates an incomplete data export from SPSS. Re-export with all variables included.`,
    });
  }
  if (criticalSavVars.length > 10) {
    datasetWarnings.push({
      type: "STRUCTURAL",
      variable: "(multiple)",
      detail: `${criticalSavVars.length - 10} more SAV-defined columns are missing from the data export`,
      explanation: `SAV defines ${criticalSavVars.length} variables total that are absent from the CSV. Only the first 10 are listed individually. Please re-export from SPSS with all columns.`,
    });
  }

  // ── Duplicate ID check ───────────────────────────────────────────────────
  const idKey = data[0] ? (["id","ID","Id","RespondentID","respondent_id"].find(k => k in data[0])) : undefined;
  if (idKey) {
    const seen = new Set<unknown>(), dupes = new Set<unknown>();
    data.forEach(r => {
      const v = r[idKey];
      if (v != null) { if (seen.has(v)) dupes.add(v); seen.add(v); }
    });
    if (dupes.size > 0) {
      datasetWarnings.push({
        type: "STRUCTURAL",
        variable: idKey,
        detail: `Duplicate respondent IDs: ${[...dupes].slice(0,20).join(", ")}`,
        explanation: "Each respondent should have a unique ID. Duplicates indicate a merge error or repeated import.",
      });
    }
  }

  // ── Per-column summary (for column-centric view) ─────────────────────────
  // We compute this during row iteration
  const colIssueCount: Record<string, number> = {};
  const colIssueTypes: Record<string, Set<IssueType>> = {};

  const flagColTracked = (id: unknown, variable: string, type: IssueType, value: unknown, detail: string, explanation: string) => {
    flag(id, variable, type, value, detail, explanation);
    colIssueCount[variable] = (colIssueCount[variable] ?? 0) + 1;
    (colIssueTypes[variable] ??= new Set()).add(type);
  };

  // ── Row-level validation ─────────────────────────────────────────────────
  for (const row of data) {
    const id = (idKey ? row[idKey] : null) ?? row.id ?? row.ID ?? "?";

    // 1. Valid code checks (for every SAV-defined numeric variable)
    for (const sv of savVars) {
      if (!dataColumns.has(sv.name)) continue;
      if (sv.type !== "numeric") continue;
      if (sv.validCodes.length === 0) continue;
      if (emptyVal(row, sv.name)) continue; // blanks handled separately

      const val = getNum(row, sv.name);
      if (val === null) continue;

      // Skip if value is a declared system-missing or missing value
      if (sv.missingValues.includes(val)) continue;

      if (!sv.validCodes.includes(val)) {
        const lbl = getVarLabel(sv.name, savVarMap, docxQMap);
        flagColTracked(id, sv.name, "MISMATCHED_CODE", val,
          `${sv.name}=${val} is not a valid code (expected: ${sv.validCodes.slice(0,8).join(", ")}${sv.validCodes.length > 8 ? "…" : ""})`,
          `Variable: ${lbl}\nSAV-defined valid codes: ${sv.validCodes.map(c => `${c}=${sv.valueLabels[c] ?? c}`).slice(0,15).join(", ")}${sv.validCodes.length > 15 ? "…" : ""}\nObserved value ${val} is not among them.`
        );
      }
    }

    // 2. Docx-sourced valid code checks (variables not in SAV but in docx)
    for (const dq of Object.values(docxQMap)) {
      if (!dataColumns.has(dq.code)) continue;
      if (savVarMap[dq.code]?.validCodes.length > 0) continue; // already handled by SAV
      if (emptyVal(row, dq.code)) continue;
      const val = getNum(row, dq.code);
      if (val === null) continue;
      if (dq.validCodes.length > 0 && !dq.validCodes.includes(val)) {
        flagColTracked(id, dq.code, "MISMATCHED_CODE", val,
          `${dq.code}=${val} not in docx-defined codes (${dq.validCodes.slice(0,6).join(", ")}…)`,
          `Question: ${dq.label}\nDocx-defined valid codes: ${dq.validCodes.map(c => `${c}=${dq.codeLabels[c] ?? c}`).slice(0,12).join(", ")}\nObserved value ${val} is not among them.`
        );
      }
    }

    // 3. Routing rule violations
    for (const rule of routingRules) {
      if (!dataColumns.has(rule.condVar)) continue;
      const condMet = hasVal(row, rule.condVar) && evalCondition(row, rule);
      const condNotMet = hasVal(row, rule.condVar) && !evalCondition(row, rule);

      const condDesc = `${rule.condVar}${rule.condOp}${rule.condVals.join(",")}`;
      const condVarLbl = getVarLabel(rule.condVar, savVarMap, docxQMap);
      const condValDescs = rule.condVals.map(v => {
        const lbl = getCodeLabel(rule.condVar, v, savVarMap, docxQMap);
        return lbl ? `${v}=${lbl}` : String(v);
      }).join(", ");

      // If condition IS met → target variables should be present
      if (condMet) {
        for (const target of rule.targets) {
          if (target === "TERMINATE") continue;
          if (!dataColumns.has(target)) continue;
          if (emptyVal(row, target)) {
            const tLbl = getVarLabel(target, savVarMap, docxQMap);
            flagColTracked(id, target, "MISSING_DATA", null,
              `${target} missing — required when ${condDesc}`,
              `Routing rule: "${rule.rawText}"\n${rule.condVar} (${condVarLbl}) = ${condValDescs}, so ${target} (${tLbl}) must be answered.\nCurrent value of ${rule.condVar}: ${getNum(row, rule.condVar)}`
            );
          }
        }
        // If condition IS met → skipTargets should be empty
        for (const skip of rule.skipTargets) {
          if (!dataColumns.has(skip)) continue;
          if (hasVal(row, skip)) {
            const sLbl = getVarLabel(skip, savVarMap, docxQMap);
            flagColTracked(id, skip, "SKIP_VIOLATION", getNum(row, skip),
              `${skip} filled but should be skipped when ${condDesc}`,
              `Routing rule: "${rule.rawText}"\nWhen ${condDesc} (${condVarLbl}=${condValDescs}), ${skip} (${sLbl}) should be empty.`
            );
          }
        }
      }

      // If condition is NOT met → target variables should be empty (skip violation)
      if (condNotMet) {
        for (const target of rule.targets) {
          if (target === "TERMINATE") continue;
          if (!dataColumns.has(target)) continue;
          const val = getNum(row, target);
          if (hasVal(row, target)) {
            // Only flag if it's not a "refusal" code (99, 999)
            if (val !== 99 && val !== 999 && val !== 9999) {
              const tLbl = getVarLabel(target, savVarMap, docxQMap);
              const cVar = getNum(row, rule.condVar);
              const cLbl = cVar !== null ? getCodeLabel(rule.condVar, cVar, savVarMap, docxQMap) : "";
              flagColTracked(id, target, "SKIP_VIOLATION", val,
                `${target} filled but ${rule.condVar}=${cVar}${cLbl ? ` (${cLbl})` : ""} (routing condition ${condDesc} not met)`,
                `Routing rule: "${rule.rawText}"\n${target} (${tLbl}) should only be asked when ${condDesc}. Current ${rule.condVar}=${cVar}${cLbl ? ` (${cLbl})` : ""}, so this field should be empty.`
              );
            }
          }
        }
      }
    }

    // 4. Open text quality check (garbled text detection)
    for (const col of Object.keys(row)) {
      const sv = savVarMap[col];
      if (sv && sv.type !== "string") continue; // only string/open-text vars
      if (sv && sv.type === "string") {
        const txt = getStr(row, col);
        if (!txt || txt.length < 3) continue;
        const meaningful = (txt.match(/[\u0531-\u058Fa-zA-Z]{2,}/g) ?? []).join("").length;
        const total = txt.replace(/\s/g, "").length;
        if (total > 6 && meaningful / total < 0.20) {
          const lbl = getVarLabel(col, savVarMap, docxQMap);
          flagColTracked(id, col, "DATA_QUALITY", txt.slice(0, 60),
            `${col}: open text appears garbled/random characters`,
            `Variable: ${lbl}\nThe text "${txt.slice(0,80)}…" contains less than 20% recognizable Armenian/Latin characters. This is consistent with interviewers entering random characters to bypass mandatory open-text fields.`
          );
        }
      }
    }
  } // end row loop

  // ── Build per-column summary ─────────────────────────────────────────────
  const columnSummary: ColumnSummary[] = [];
  const allVarNames = [
    ...savVars.map(sv => sv.name),
    ...Object.keys(docxQMap).filter(k => !savVarMap[k]),
    ...Object.keys(data[0] ?? {}).filter(k => !savVarMap[k] && !docxQMap[k]),
  ].filter((v, i, a) => a.indexOf(v) === i); // unique

  for (const varName of allVarNames) {
    if (!dataColumns.has(varName)) continue;

    const sv = savVarMap[varName];
    const dq = docxQMap[varName];
    const label = getVarLabel(varName, savVarMap, docxQMap);
    const validCodes = buildValidCodes(varName, savVarMap, docxQMap);
    const gatingRules = findGatingRules(varName, routingRules);
    const condRules = routingIndex[varName] ?? [];

    // Compute value frequency in data
    const freqMap: Record<string, number> = {};
    let nFilled = 0, nEmpty = 0;
    for (const row of data) {
      const v = getStr(row, varName);
      if (v === null) { nEmpty++; continue; }
      nFilled++;
      freqMap[v] = (freqMap[v] ?? 0) + 1;
    }
    const topValues = Object.entries(freqMap)
      .sort((a, b) => b[1] - a[1])
      .slice(0, 10)
      .map(([val, cnt]) => {
        const numVal = Number(val);
        const lbl = !isNaN(numVal) ? getCodeLabel(varName, numVal, savVarMap, docxQMap) : "";
        return { val, cnt, label: lbl };
      });

    columnSummary.push({
      varName,
      label,
      section: dq?.section ?? sv?.label ?? "",
      type: sv?.type ?? "unknown",
      validCodes: validCodes ?? [],
      valueLabels: sv?.valueLabels ?? dq?.codeLabels ?? {},
      nFilled,
      nEmpty,
      nIssues: colIssueCount[varName] ?? 0,
      issueTypes: [...(colIssueTypes[varName] ?? [])],
      topValues,
      gatingRules,
      condRules,
      hasInDocx: !!dq,
      hasInSav: !!sv,
    });
  }

  return { issues, datasetWarnings, columnSummary };
}

// ─────────────────────────────────────────────────────────────────────────────
// COLUMN SUMMARY TYPE
// ─────────────────────────────────────────────────────────────────────────────

interface ColumnSummary {
  varName: string;
  label: string;
  section: string;
  type: string;
  validCodes: number[];
  valueLabels: Record<number, string>;
  nFilled: number;
  nEmpty: number;
  nIssues: number;
  issueTypes: IssueType[];
  topValues: { val: string; cnt: number; label: string }[];
  gatingRules: RoutingRule[];
  condRules: RoutingRule[];
  hasInDocx: boolean;
  hasInSav: boolean;
}

// ─────────────────────────────────────────────────────────────────────────────
// UI COMPONENTS
// ─────────────────────────────────────────────────────────────────────────────

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
      onDrop={e => { e.preventDefault(); setDrag(false); if (e.dataTransfer.files[0]) onLoad(e.dataTransfer.files[0]); }}
      style={{ border: `2px dashed ${drag ? "#2563eb" : loaded ? "#22c55e" : "#cbd5e1"}`, borderRadius: 12, padding: "16px 12px", textAlign: "center", cursor: "pointer", background: loaded ? "#f0fdf4" : drag ? "#eff6ff" : "#f8fafc", transition: "all .2s", minHeight: 90 }}>
      <input ref={ref} type="file" accept={accept} style={{ display: "none" }}
        onChange={e => { if (e.target.files?.[0]) onLoad(e.target.files[0]); }} />
      <div style={{ fontSize: 24, marginBottom: 4 }}>{loaded ? "✅" : "📂"}</div>
      <div style={{ fontSize: 12, fontWeight: 600, color: loaded ? "#16a34a" : "#475569" }}>{loaded || label}</div>
      <div style={{ fontSize: 10, color: "#94a3b8", marginTop: 2 }}>{sub}</div>
    </div>
  );
}

function Badge({ type }: { type: IssueType }) {
  const m = ISSUE_TYPES[type];
  return <span style={{ background: m.bg, color: m.color, padding: "2px 7px", borderRadius: 4, fontSize: 10, fontWeight: 700, whiteSpace: "nowrap" }}>{m.label}</span>;
}

function Th({ children, style = {} }: { children: ReactNode; style?: React.CSSProperties }) {
  return <th style={{ padding: "8px 12px", textAlign: "left", fontWeight: 600, color: "#475569", borderBottom: "1px solid #e2e8f0", whiteSpace: "nowrap", fontSize: 11, ...style }}>{children}</th>;
}

function Td({ children, style = {} }: { children: ReactNode; style?: React.CSSProperties }) {
  return <td style={{ padding: "7px 12px", fontSize: 12, color: "#374151", borderBottom: "1px solid #f1f5f9", ...style }}>{children}</td>;
}

function ExplanationBox({ explanation }: { explanation: string }) {
  const [open, setOpen] = useState(false);
  const parts = explanation.split(/(📌[^\n]*|📋[^\n]*|Routing rule:[^\n]*)/g);
  return (
    <div>
      <button onClick={e => { e.stopPropagation(); setOpen(o => !o); }}
        style={{ fontSize: 10, color: "#6366f1", background: "none", border: "1px solid #e0e7ff", borderRadius: 4, padding: "2px 7px", cursor: "pointer" }}>
        {open ? "▲ Hide" : "▼ Why?"}
      </button>
      {open && (
        <div style={{ marginTop: 5, background: "#f5f3ff", border: "1px solid #ddd6fe", borderRadius: 6, padding: "7px 9px", fontSize: 11, color: "#4c1d95", lineHeight: 1.6, maxWidth: 480 }}>
          {parts.map((p, i) => {
            if (p.startsWith("📌") || p.startsWith("Routing rule:")) {
              return <div key={i} style={{ marginTop: 4, background: "#eff6ff", border: "1px solid #bfdbfe", borderRadius: 3, padding: "3px 7px", color: "#1e40af", fontSize: 11 }}>{p}</div>;
            }
            return <span key={i} style={{ whiteSpace: "pre-wrap" }}>{p}</span>;
          })}
        </div>
      )}
    </div>
  );
}

// Column detail panel
function ColumnPanel({ col, onClose }: { col: ColumnSummary; onClose: () => void }) {
  return (
    <div style={{ position: "fixed", top: 0, right: 0, width: 420, height: "100vh", background: "#fff", boxShadow: "-4px 0 20px rgba(0,0,0,.12)", overflowY: "auto", zIndex: 100, padding: "20px 18px" }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 16 }}>
        <div>
          <div style={{ fontFamily: "monospace", fontSize: 18, fontWeight: 800, color: "#1e293b" }}>{col.varName}</div>
          <div style={{ fontSize: 12, color: "#64748b", marginTop: 2, maxWidth: 340 }}>{col.label || "—"}</div>
        </div>
        <button onClick={onClose} style={{ background: "none", border: "none", fontSize: 18, cursor: "pointer", color: "#94a3b8" }}>✕</button>
      </div>

      {/* Source badges */}
      <div style={{ display: "flex", gap: 6, marginBottom: 14 }}>
        {col.hasInSav && <span style={{ background: "#fef9c3", border: "1px solid #fde047", borderRadius: 4, fontSize: 10, padding: "2px 7px", color: "#854d0e" }}>🗃 SAV-defined</span>}
        {col.hasInDocx && <span style={{ background: "#f0fdf4", border: "1px solid #86efac", borderRadius: 4, fontSize: 10, padding: "2px 7px", color: "#166534" }}>📝 Docx-defined</span>}
        <span style={{ background: "#f1f5f9", border: "1px solid #e2e8f0", borderRadius: 4, fontSize: 10, padding: "2px 7px", color: "#475569" }}>{col.type}</span>
      </div>

      {/* Fill stats */}
      <div style={{ background: "#f8fafc", borderRadius: 8, padding: "10px 12px", marginBottom: 14, display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 8 }}>
        <div style={{ textAlign: "center" }}>
          <div style={{ fontSize: 20, fontWeight: 800, color: "#1e293b" }}>{col.nFilled + col.nEmpty}</div>
          <div style={{ fontSize: 10, color: "#64748b" }}>Total rows</div>
        </div>
        <div style={{ textAlign: "center" }}>
          <div style={{ fontSize: 20, fontWeight: 800, color: "#16a34a" }}>{col.nFilled}</div>
          <div style={{ fontSize: 10, color: "#64748b" }}>Filled</div>
        </div>
        <div style={{ textAlign: "center" }}>
          <div style={{ fontSize: 20, fontWeight: 800, color: col.nIssues > 0 ? "#ef4444" : "#22c55e" }}>{col.nIssues}</div>
          <div style={{ fontSize: 10, color: "#64748b" }}>Issues</div>
        </div>
      </div>

      {/* Issue type badges */}
      {col.issueTypes.length > 0 && (
        <div style={{ marginBottom: 14 }}>
          <div style={{ fontSize: 11, fontWeight: 700, color: "#64748b", marginBottom: 6 }}>ISSUE TYPES</div>
          <div style={{ display: "flex", gap: 5, flexWrap: "wrap" }}>
            {col.issueTypes.map(t => <Badge key={t} type={t} />)}
          </div>
        </div>
      )}

      {/* Valid codes */}
      {col.validCodes.length > 0 && (
        <div style={{ marginBottom: 14 }}>
          <div style={{ fontSize: 11, fontWeight: 700, color: "#64748b", marginBottom: 6 }}>VALID CODES</div>
          <div style={{ display: "flex", flexWrap: "wrap", gap: 4 }}>
            {col.validCodes.map(c => (
              <span key={c} style={{ background: "#f0f9ff", border: "1px solid #bae6fd", borderRadius: 3, padding: "2px 6px", fontSize: 10, color: "#0369a1" }}>
                {c}{col.valueLabels[c] ? `=${col.valueLabels[c]}` : ""}
              </span>
            ))}
          </div>
        </div>
      )}

      {/* Value distribution */}
      {col.topValues.length > 0 && (
        <div style={{ marginBottom: 14 }}>
          <div style={{ fontSize: 11, fontWeight: 700, color: "#64748b", marginBottom: 6 }}>VALUE DISTRIBUTION</div>
          <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 11 }}>
            <thead><tr><Th>Value</Th><Th>Label</Th><Th>Count</Th><Th>%</Th></tr></thead>
            <tbody>
              {col.topValues.map(({ val, cnt, label }) => (
                <tr key={val}>
                  <Td><code>{val}</code></Td>
                  <Td style={{ color: "#64748b" }}>{label || "—"}</Td>
                  <Td>{cnt}</Td>
                  <Td style={{ color: "#94a3b8" }}>{((cnt / (col.nFilled || 1)) * 100).toFixed(1)}%</Td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}

      {/* Routing rules that gate this variable */}
      {col.gatingRules.length > 0 && (
        <div style={{ marginBottom: 14 }}>
          <div style={{ fontSize: 11, fontWeight: 700, color: "#64748b", marginBottom: 6 }}>ASKED WHEN (routing rules)</div>
          {col.gatingRules.map((r, i) => (
            <div key={i} style={{ background: "#fdf4ff", border: "1px solid #e9d5ff", borderRadius: 5, padding: "6px 9px", marginBottom: 5, fontSize: 11, color: "#6b21a8" }}>
              <strong>{r.condVar}{r.condOp}{r.condVals.join(",")}</strong><br />
              <span style={{ color: "#9333ea", fontSize: 10 }}>{r.rawText.slice(0, 150)}</span>
            </div>
          ))}
        </div>
      )}

      {/* Routing rules triggered by this variable */}
      {col.condRules.length > 0 && (
        <div style={{ marginBottom: 14 }}>
          <div style={{ fontSize: 11, fontWeight: 700, color: "#64748b", marginBottom: 6 }}>TRIGGERS ROUTING</div>
          {col.condRules.map((r, i) => (
            <div key={i} style={{ background: "#eff6ff", border: "1px solid #bfdbfe", borderRadius: 5, padding: "6px 9px", marginBottom: 5, fontSize: 11, color: "#1e40af" }}>
              <strong>{r.condVar}{r.condOp}{r.condVals.join(",")}</strong> → {r.targets.join(", ") || r.skipTargets.join(", ") || "skip"}<br />
              <span style={{ color: "#3b82f6", fontSize: 10 }}>{r.rawText.slice(0, 150)}</span>
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

// ─────────────────────────────────────────────────────────────────────────────
// MAIN APP
// ─────────────────────────────────────────────────────────────────────────────

export default function App() {
  // Data
  const [data, setData] = useState<RowData[]>([]);
  const [dataFileName, setDataFileName] = useState("");

  // SAV
  const [savVars, setSavVars] = useState<SavVariable[]>([]);
  const [savVarMap, setSavVarMap] = useState<Record<string, SavVariable>>({});
  const [savFileName, setSavFileName] = useState("");

  // Docx
  const [docxQMap, setDocxQMap] = useState<Record<string, DocxQuestion>>({});
  const [routingRules, setRoutingRules] = useState<RoutingRule[]>([]);
  const [docxRawText, setDocxRawText] = useState("");
  const [docxFileName, setDocxFileName] = useState("");
  const [docxLoading, setDocxLoading] = useState(false);

  // Results
  const [issues, setIssues] = useState<Issue[]>([]);
  const [dsWarnings, setDsWarnings] = useState<DatasetWarning[]>([]);
  const [columnSummary, setColumnSummary] = useState<ColumnSummary[]>([]);
  const [analyzed, setAnalyzed] = useState(false);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");

  // UI state
  const [tab, setTab] = useState("columns");
  const [filterType, setFilterType] = useState<IssueType | "ALL">("ALL");
  const [search, setSearch] = useState("");
  const [selectedCol, setSelectedCol] = useState<ColumnSummary | null>(null);
  const [colFilter, setColFilter] = useState<"all" | "issues" | "sav" | "docx">("all");

  // ── File loaders ───────────────────────────────────────────────────────

  const loadDataFile = useCallback((file: File) => {
    setError("");
    const ext = file.name.split(".").pop()?.toLowerCase() ?? "";
    if (ext === "csv") {
      Papa.parse<RowData>(file, {
        header: true, skipEmptyLines: true, dynamicTyping: true,
        complete: r => { setData(r.data); setDataFileName(file.name); },
        error: (e: Error) => setError("CSV parse error: " + e.message),
      });
    } else if (ext === "xlsx" || ext === "xls") {
      const reader = new FileReader();
      reader.onload = e => {
        try {
          const wb = XLSX.read(e.target?.result as ArrayBuffer, { type: "array" });
          setData(XLSX.utils.sheet_to_json<RowData>(wb.Sheets[wb.SheetNames[0]], { defval: "" }));
          setDataFileName(file.name);
        } catch (err) { setError("Excel parse error: " + (err as Error).message); }
      };
      reader.readAsArrayBuffer(file);
    } else {
      setError(`Unsupported file: .${ext}`);
    }
  }, []);

  const loadSavFile = useCallback((file: File) => {
    setError("");
    const reader = new FileReader();
    reader.onload = e => {
      const { variables, varMap } = parseSavFile(e.target?.result as ArrayBuffer);
      setSavVars(variables);
      setSavVarMap(varMap);
      setSavFileName(file.name);
    };
    reader.readAsArrayBuffer(file);
  }, []);

  const loadDocxFile = useCallback(async (file: File) => {
    setError(""); setDocxLoading(true);
    try {
      const { questions, questionMap, routingRules: rules, rawText } = await parseDocxFile(await file.arrayBuffer());
      setDocxQMap(questionMap);
      setRoutingRules(rules);
      setDocxRawText(rawText);
      setDocxFileName(file.name);
      void questions; // stored via questionMap
    } catch (err) {
      setError("Docx parse error: " + (err as Error).message);
    } finally { setDocxLoading(false); }
  }, []);

  // ── Analyze ────────────────────────────────────────────────────────────

  const analyze = () => {
    if (!data.length) { setError("Upload a CSV or Excel data file to run validation."); return; }
    setLoading(true); setError("");
    setTimeout(() => {
      const { issues: found, datasetWarnings: warnings, columnSummary: colSummary } =
        runDynamicValidation(data, savVars, savVarMap, docxQMap, routingRules);
      setIssues(found);
      setDsWarnings(warnings);
      setColumnSummary(colSummary);
      setAnalyzed(true);
      setTab("columns");
      setLoading(false);
    }, 60);
  };

  // ── Derived counts ─────────────────────────────────────────────────────

  const typeCounts = (Object.keys(ISSUE_TYPES) as IssueType[]).reduce((acc, t) => {
    acc[t] = issues.filter(i => i.type === t).length; return acc;
  }, {} as Record<IssueType, number>);

  const issueMap: Record<string, IssueType> = {};
  issues.forEach(i => { issueMap[`${i.id}__${i.variable}`] = i.type; });

  const filteredIssues = issues.filter(i => {
    const q = search.toLowerCase();
    return (filterType === "ALL" || i.type === filterType) &&
      (!q || i.variable.toLowerCase().includes(q) || String(i.id).toLowerCase().includes(q) || i.detail.toLowerCase().includes(q));
  });

  const filteredCols = columnSummary.filter(col => {
    const q = search.toLowerCase();
    const matchSearch = !q || col.varName.toLowerCase().includes(q) || col.label.toLowerCase().includes(q);
    const matchFilter =
      colFilter === "all" ? true :
      colFilter === "issues" ? col.nIssues > 0 :
      colFilter === "sav" ? col.hasInSav :
      colFilter === "docx" ? col.hasInDocx : true;
    return matchSearch && matchFilter;
  });

  const dataColumns = data.length ? Object.keys(data[0]) : [];
  const totalIssues = issues.length + dsWarnings.length;

  const downloadCSV = (rows: object[], name: string) => {
    const csv = Papa.unparse(rows);
    const a = document.createElement("a");
    a.href = URL.createObjectURL(new Blob([csv], { type: "text/csv" }));
    a.download = name; a.click();
  };

  const nSavLabels = savVars.filter(sv => sv.validCodes.length > 0).length;
  const nDocxRouting = routingRules.length;
  const nDocxQ = Object.keys(docxQMap).length;

  const tabs = [
    analyzed && ["columns", `📊 Columns (${columnSummary.length})`],
    analyzed && totalIssues > 0 && ["issues", `🚩 Issues (${totalIssues})`],
    analyzed && data.length > 0 && ["data", `📋 Data Table`],
    savVars.length > 0 && ["savvars", `🗃 SAV (${savVars.length})`],
    nDocxRouting > 0 && ["routing", `📋 Routing (${nDocxRouting})`],
    docxRawText && ["docxtext", `📄 Questionnaire`],
  ].filter(Boolean) as [string, string][];

  return (
    <div style={{ fontFamily: "Inter, system-ui, sans-serif", minHeight: "100vh", background: "#f1f5f9" }}>

      {/* Header */}
      <div style={{ background: "linear-gradient(135deg,#1e3a5f,#2563eb)", padding: "16px 24px", color: "#fff" }}>
        <div style={{ maxWidth: 1300, margin: "0 auto", display: "flex", alignItems: "center", gap: 16 }}>
          <div style={{ flex: 1 }}>
            <h1 style={{ margin: 0, fontSize: 18, fontWeight: 700 }}>🔍 Survey Data Validator</h1>
            <p style={{ margin: "2px 0 0", fontSize: 11, opacity: .75 }}>
              SAV-first dynamic validation · Skip logic · Code checks · Column-centric analysis
            </p>
          </div>
          <div style={{ fontSize: 11, opacity: .7, textAlign: "right", lineHeight: 1.8 }}>
            {savFileName && <div>🗃 {savFileName} · {savVars.length} vars, {nSavLabels} with codes</div>}
            {docxFileName && <div>📝 {docxFileName} · {nDocxQ} questions · {nDocxRouting} routing rules</div>}
            {dataFileName && <div>📋 {dataFileName} · {data.length} rows · {dataColumns.length} cols</div>}
          </div>
        </div>
      </div>

      <div style={{ maxWidth: 1300, margin: "0 auto", padding: "18px 16px" }}>

        {/* Upload zones */}
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr 1fr", gap: 10, marginBottom: 12 }}>

          <div>
            <div style={{ fontSize: 10, fontWeight: 700, color: "#64748b", marginBottom: 5, textTransform: "uppercase", letterSpacing: .5 }}>
              1 · SPSS .sav <span style={{ color: "#2563eb" }}>(primary)</span>
            </div>
            <FileZone label="Drop .sav file" sub="Variable names, valid codes, value labels" onLoad={loadSavFile} loaded={savFileName} accept=".sav,.por" />
            {nSavLabels > 0 && <div style={{ marginTop: 3, fontSize: 10, color: "#16a34a", background: "#f0fdf4", border: "1px solid #bbf7d0", borderRadius: 4, padding: "3px 7px" }}>✅ {nSavLabels} vars with valid codes</div>}
          </div>

          <div>
            <div style={{ fontSize: 10, fontWeight: 700, color: "#64748b", marginBottom: 5, textTransform: "uppercase", letterSpacing: .5 }}>
              2 · Word .docx <span style={{ color: "#7c3aed" }}>(routing + labels)</span>
            </div>
            <FileZone label={docxLoading ? "⏳ Parsing…" : "Drop questionnaire .docx"} sub="Routing rules, question text, answer codes" onLoad={loadDocxFile} loaded={docxFileName} accept=".docx,.doc" />
            {nDocxRouting > 0 && <div style={{ marginTop: 3, fontSize: 10, color: "#7c3aed", background: "#f5f3ff", border: "1px solid #ddd6fe", borderRadius: 4, padding: "3px 7px" }}>✅ {nDocxRouting} routing rules · {nDocxQ} questions</div>}
          </div>

          <div>
            <div style={{ fontSize: 10, fontWeight: 700, color: "#64748b", marginBottom: 5, textTransform: "uppercase", letterSpacing: .5 }}>
              3 · Data file <span style={{ color: "#ef4444" }}>*required</span>
            </div>
            <FileZone label="Drop CSV or Excel" sub=".csv · .xlsx — row-level data" onLoad={loadDataFile} loaded={dataFileName} accept=".csv,.xlsx,.xls" />
          </div>

          <div style={{ display: "flex", flexDirection: "column", justifyContent: "flex-end" }}>
            <button onClick={analyze} disabled={loading || !data.length}
              style={{ padding: "13px", background: loading ? "#94a3b8" : !data.length ? "#cbd5e1" : "#2563eb", color: "#fff", border: "none", borderRadius: 10, fontSize: 13, fontWeight: 700, cursor: data.length ? "pointer" : "not-allowed", width: "100%", marginBottom: 6 }}>
              {loading ? "⏳ Analyzing…" : "🚀 Run Validation"}
            </button>
            <div style={{ background: "#fffbeb", border: "1px solid #fde68a", borderRadius: 6, padding: "6px 10px", fontSize: 10, color: "#92400e" }}>
              💡 SAV defines what's valid. Docx defines routing. CSV is the data to check.<br />
              All rules are derived from your files — nothing is hardcoded.
            </div>
          </div>
        </div>

        {error && <div style={{ background: "#fef2f2", border: "1px solid #fca5a5", borderRadius: 8, padding: "9px 13px", color: "#dc2626", fontSize: 12, marginBottom: 10 }}>⚠️ {error}</div>}

        {/* Enrichment banner */}
        {(nSavLabels > 0 || nDocxRouting > 0) && !analyzed && (
          <div style={{ background: "#eff6ff", border: "1px solid #bfdbfe", borderRadius: 8, padding: "8px 13px", fontSize: 11, color: "#1e40af", marginBottom: 10, display: "flex", gap: 14, flexWrap: "wrap" }}>
            <span>🔗 <strong>Ready:</strong></span>
            {nSavLabels > 0 && <span>🗃 {nSavLabels} SAV variables with valid codes loaded</span>}
            {nDocxRouting > 0 && <span>📋 {nDocxRouting} routing rules from docx</span>}
            {nDocxQ > 0 && <span>📝 {nDocxQ} question labels from docx</span>}
            <span style={{ opacity: .7 }}>→ Upload data file and click Run Validation</span>
          </div>
        )}

        {analyzed && (
          <>
            {/* Type summary cards */}
            <div style={{ display: "grid", gridTemplateColumns: "repeat(6,1fr)", gap: 7, marginBottom: 14 }}>
              {(Object.entries(ISSUE_TYPES) as [IssueType, { label: string; color: string; bg: string }][]).map(([type, meta]) => (
                <div key={type} onClick={() => setFilterType(filterType === type ? "ALL" : type)}
                  style={{ background: "#fff", border: `2px solid ${filterType === type ? meta.color : "#e2e8f0"}`, borderRadius: 9, padding: "9px 11px", cursor: "pointer", transition: "all .15s" }}>
                  <div style={{ fontSize: 20, fontWeight: 800, color: typeCounts[type] > 0 ? meta.color : "#cbd5e1" }}>{typeCounts[type]}</div>
                  <div style={{ fontSize: 9, color: "#64748b", marginTop: 2, lineHeight: 1.3 }}>{meta.label}</div>
                </div>
              ))}
            </div>

            {/* Tabs */}
            <div style={{ display: "flex", gap: 1, borderBottom: "2px solid #e2e8f0", marginBottom: 12, flexWrap: "wrap", alignItems: "center" }}>
              {tabs.map(([id, label]) => (
                <button key={id} onClick={() => setTab(id)}
                  style={{ padding: "7px 13px", border: "none", background: "none", borderBottom: tab === id ? "2px solid #2563eb" : "2px solid transparent", color: tab === id ? "#2563eb" : "#64748b", fontWeight: tab === id ? 700 : 400, cursor: "pointer", fontSize: 12, marginBottom: -2 }}>
                  {label}
                </button>
              ))}
              <div style={{ flex: 1 }} />
              {issues.length > 0 && (
                <button onClick={() => downloadCSV(issues.map(i => ({ ID: i.id, Variable: i.variable, Type: ISSUE_TYPES[i.type].label, Value: String(i.value ?? ""), Detail: i.detail, Explanation: i.explanation })), "issues_report.csv")}
                  style={{ fontSize: 11, padding: "4px 9px", background: "#f1f5f9", border: "1px solid #e2e8f0", borderRadius: 5, cursor: "pointer", color: "#475569" }}>
                  ⬇ Export Issues CSV
                </button>
              )}
            </div>

            {/* Search */}
            <input placeholder="Search variable name, ID, or issue description…" value={search} onChange={e => setSearch(e.target.value)}
              style={{ width: "100%", padding: "7px 11px", border: "1px solid #e2e8f0", borderRadius: 7, fontSize: 12, marginBottom: 10, boxSizing: "border-box" }} />

            {/* ── Columns tab (primary view) ── */}
            {tab === "columns" && (
              <div>
                {/* Column filter pills */}
                <div style={{ display: "flex", gap: 6, marginBottom: 10, flexWrap: "wrap" }}>
                  {(["all","issues","sav","docx"] as const).map(f => (
                    <button key={f} onClick={() => setColFilter(f)}
                      style={{ padding: "4px 11px", borderRadius: 20, border: `1px solid ${colFilter === f ? "#2563eb" : "#e2e8f0"}`, background: colFilter === f ? "#eff6ff" : "#fff", color: colFilter === f ? "#2563eb" : "#64748b", fontSize: 11, cursor: "pointer", fontWeight: colFilter === f ? 700 : 400 }}>
                      {f === "all" ? `All (${columnSummary.length})` :
                       f === "issues" ? `Has issues (${columnSummary.filter(c => c.nIssues > 0).length})` :
                       f === "sav" ? `SAV-defined (${columnSummary.filter(c => c.hasInSav).length})` :
                       `Docx-defined (${columnSummary.filter(c => c.hasInDocx).length})`}
                    </button>
                  ))}
                </div>

                <div style={{ background: "#fff", borderRadius: 10, border: "1px solid #e2e8f0", overflow: "hidden" }}>
                  <div style={{ overflowX: "auto", maxHeight: 600 }}>
                    <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
                      <thead style={{ position: "sticky", top: 0, background: "#f8fafc", zIndex: 1 }}>
                        <tr>
                          <Th>Variable</Th>
                          <Th>Question / Label</Th>
                          <Th>Valid Codes (SAV)</Th>
                          <Th style={{ textAlign: "center" }}>Filled</Th>
                          <Th style={{ textAlign: "center" }}>Empty</Th>
                          <Th style={{ textAlign: "center" }}>Issues</Th>
                          <Th>Issue Types</Th>
                          <Th>Source</Th>
                        </tr>
                      </thead>
                      <tbody>
                        {filteredCols.map((col, i) => (
                          <tr key={col.varName}
                            onClick={() => setSelectedCol(selectedCol?.varName === col.varName ? null : col)}
                            style={{ background: selectedCol?.varName === col.varName ? "#eff6ff" : i % 2 === 0 ? "#fff" : "#fafafa", cursor: "pointer", borderLeft: col.nIssues > 0 ? `3px solid #ef4444` : "3px solid transparent" }}>
                            <Td><code style={{ fontWeight: 700, color: "#7c3aed" }}>{col.varName}</code></Td>
                            <Td style={{ maxWidth: 280, color: "#374151" }}>
                              <div style={{ overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", maxWidth: 280 }} title={col.label}>
                                {col.label || <span style={{ color: "#cbd5e1", fontStyle: "italic" }}>—</span>}
                              </div>
                            </Td>
                            <Td>
                              {col.validCodes.length > 0 ? (
                                <div style={{ display: "flex", flexWrap: "wrap", gap: 2 }}>
                                  {col.validCodes.slice(0, 8).map(c => (
                                    <span key={c} style={{ background: "#f0f9ff", border: "1px solid #bae6fd", borderRadius: 2, padding: "1px 4px", fontSize: 9, color: "#0369a1" }}>
                                      {c}{col.valueLabels[c] ? `=${col.valueLabels[c].slice(0,10)}` : ""}
                                    </span>
                                  ))}
                                  {col.validCodes.length > 8 && <span style={{ fontSize: 9, color: "#94a3b8" }}>+{col.validCodes.length - 8}</span>}
                                </div>
                              ) : <span style={{ color: "#cbd5e1", fontSize: 10, fontStyle: "italic" }}>—</span>}
                            </Td>
                            <Td style={{ textAlign: "center", color: "#16a34a", fontWeight: 600 }}>{col.nFilled}</Td>
                            <Td style={{ textAlign: "center", color: col.nEmpty > 0 ? "#f97316" : "#94a3b8" }}>{col.nEmpty}</Td>
                            <Td style={{ textAlign: "center" }}>
                              {col.nIssues > 0
                                ? <span style={{ background: "#fef2f2", color: "#ef4444", borderRadius: 12, padding: "2px 8px", fontWeight: 700, fontSize: 11 }}>{col.nIssues}</span>
                                : <span style={{ color: "#22c55e", fontSize: 11 }}>✓</span>}
                            </Td>
                            <Td>
                              <div style={{ display: "flex", gap: 3, flexWrap: "wrap" }}>
                                {col.issueTypes.map(t => <Badge key={t} type={t} />)}
                              </div>
                            </Td>
                            <Td>
                              <div style={{ display: "flex", gap: 3 }}>
                                {col.hasInSav && <span style={{ fontSize: 9, background: "#fef9c3", border: "1px solid #fde047", borderRadius: 3, padding: "1px 4px", color: "#854d0e" }}>SAV</span>}
                                {col.hasInDocx && <span style={{ fontSize: 9, background: "#f0fdf4", border: "1px solid #86efac", borderRadius: 3, padding: "1px 4px", color: "#166534" }}>DOCX</span>}
                              </div>
                            </Td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                  <div style={{ padding: "6px 12px", fontSize: 11, color: "#94a3b8", borderTop: "1px solid #e2e8f0" }}>
                    {filteredCols.length} of {columnSummary.length} columns shown · Click a row to open column details
                  </div>
                </div>
              </div>
            )}

            {/* ── Issues tab ── */}
            {tab === "issues" && (
              <div>
                {dsWarnings.length > 0 && (
                  <div style={{ marginBottom: 12 }}>
                    <div style={{ fontSize: 11, fontWeight: 700, color: "#64748b", marginBottom: 5, textTransform: "uppercase" }}>⚠️ Structural / Dataset Issues ({dsWarnings.length})</div>
                    {dsWarnings.map((w, i) => (
                      <div key={i} style={{ background: "#fff7ed", border: "1px solid #fed7aa", borderRadius: 8, padding: "10px 13px", marginBottom: 7, borderLeft: "4px solid #f97316" }}>
                        <div style={{ fontWeight: 700, fontSize: 12, color: "#9a3412", marginBottom: 3 }}>
                          <code style={{ background: "#fef3c7", padding: "1px 5px", borderRadius: 2 }}>{w.variable}</code> {w.detail}
                        </div>
                        <div style={{ fontSize: 11, color: "#92400e", lineHeight: 1.6 }}>{w.explanation}</div>
                      </div>
                    ))}
                  </div>
                )}

                {filteredIssues.length === 0 ? (
                  <div style={{ background: "#fff", borderRadius: 10, padding: 40, textAlign: "center", color: "#94a3b8" }}>
                    <div style={{ fontSize: 40, marginBottom: 8 }}>{issues.length === 0 ? "✅" : "🔍"}</div>
                    <div style={{ fontWeight: 600 }}>{issues.length === 0 ? "No row-level issues found!" : "No issues match your filter."}</div>
                  </div>
                ) : (
                  <div style={{ background: "#fff", borderRadius: 10, border: "1px solid #e2e8f0", overflow: "hidden" }}>
                    <div style={{ overflowX: "auto", maxHeight: 560 }}>
                      <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
                        <thead style={{ position: "sticky", top: 0, background: "#f8fafc", zIndex: 1 }}>
                          <tr><Th>ID</Th><Th>Variable</Th><Th>Type</Th><Th>Value</Th><Th>Description</Th><Th>Explanation</Th></tr>
                        </thead>
                        <tbody>
                          {filteredIssues.map((iss, i) => (
                            <tr key={i} style={{ background: i % 2 === 0 ? "#fff" : "#fafafa", borderLeft: `3px solid ${ISSUE_TYPES[iss.type].color}` }}>
                              <Td style={{ fontWeight: 700, whiteSpace: "nowrap" }}>{String(iss.id)}</Td>
                              <Td><code style={{ color: "#7c3aed", fontSize: 11 }}>{iss.variable}</code></Td>
                              <Td><Badge type={iss.type} /></Td>
                              <Td><code style={{ color: "#dc2626", fontSize: 11 }}>{String(iss.value ?? "—").slice(0, 50)}</code></Td>
                              <Td style={{ maxWidth: 260 }}>{iss.detail}</Td>
                              <Td><ExplanationBox explanation={iss.explanation} /></Td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                    <div style={{ padding: "6px 12px", fontSize: 11, color: "#94a3b8", borderTop: "1px solid #e2e8f0" }}>
                      {filteredIssues.length} of {issues.length} row-level issues
                    </div>
                  </div>
                )}
              </div>
            )}

            {/* ── Data table tab ── */}
            {tab === "data" && (
              <div style={{ background: "#fff", borderRadius: 10, border: "1px solid #e2e8f0", overflow: "auto", maxHeight: 560 }}>
                <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 10 }}>
                  <thead style={{ position: "sticky", top: 0, background: "#f8fafc", zIndex: 1 }}>
                    <tr>{dataColumns.map(col => <Th key={col}>{col}</Th>)}</tr>
                  </thead>
                  <tbody>
                    {data.slice(0, 200).map((row, ri) => {
                      const rid = (["id","ID","Id","RespondentID"].find(k => row[k] != null) ? row[["id","ID","Id","RespondentID"].find(k => row[k] != null)!] : `Row${ri+1}`);
                      return (
                        <tr key={ri} style={{ background: ri % 2 === 0 ? "#fff" : "#fafafa" }}>
                          {dataColumns.map(col => {
                            const it = issueMap[`${rid}__${col}`];
                            const meta = it ? ISSUE_TYPES[it] : null;
                            return (
                              <td key={col} title={meta ? `${meta.label}: ${issues.find(iss => iss.id == rid && iss.variable === col)?.detail}` : ""}
                                style={{ padding: "4px 8px", background: meta ? meta.bg : "transparent", color: meta ? meta.color : "#334155", fontFamily: "monospace", fontSize: 10, borderBottom: "1px solid #f1f5f9", whiteSpace: "nowrap" }}>
                                {String(row[col] ?? "").slice(0, 35)}
                              </td>
                            );
                          })}
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            )}

            {/* ── SAV vars tab ── */}
            {tab === "savvars" && (
              <div style={{ background: "#fff", borderRadius: 10, border: "1px solid #e2e8f0", overflow: "auto", maxHeight: 540 }}>
                <div style={{ padding: "9px 13px", background: "#fffbeb", borderBottom: "1px solid #fde68a", fontSize: 11, color: "#92400e" }}>
                  {savVars.length} variables from SAV · {nSavLabels} have value labels
                </div>
                <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
                  <thead style={{ position: "sticky", top: 0, background: "#f8fafc", zIndex: 1 }}>
                    <tr><Th>#</Th><Th>Name</Th><Th>Label</Th><Th>Value Labels</Th></tr>
                  </thead>
                  <tbody>
                    {savVars.filter(sv => !search || sv.name.toLowerCase().includes(search.toLowerCase()) || sv.label.toLowerCase().includes(search.toLowerCase())).slice(0, 500).map((sv, i) => (
                      <tr key={sv.name} style={{ background: i % 2 === 0 ? "#fff" : "#fafafa" }}>
                        <Td style={{ color: "#94a3b8", fontSize: 10 }}>{i + 1}</Td>
                        <Td><code style={{ fontWeight: 700 }}>{sv.name}</code></Td>
                        <Td style={{ color: sv.label ? "#334155" : "#cbd5e1", fontStyle: sv.label ? "normal" : "italic" }}>{sv.label || "—"}</Td>
                        <Td>
                          {Object.keys(sv.valueLabels).length > 0 ? (
                            <div style={{ display: "flex", flexWrap: "wrap", gap: 2 }}>
                              {Object.entries(sv.valueLabels).slice(0, 15).map(([code, lbl]) => (
                                <span key={code} style={{ background: "#f0f9ff", border: "1px solid #bae6fd", borderRadius: 2, padding: "1px 4px", fontSize: 9, color: "#0369a1", whiteSpace: "nowrap" }}>
                                  {code}={lbl}
                                </span>
                              ))}
                            </div>
                          ) : <span style={{ color: "#cbd5e1", fontSize: 10 }}>—</span>}
                        </Td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )}

            {/* ── Routing tab ── */}
            {tab === "routing" && (
              <div style={{ background: "#fff", borderRadius: 10, border: "1px solid #e2e8f0", overflow: "auto", maxHeight: 540 }}>
                <div style={{ padding: "9px 13px", background: "#f5f3ff", borderBottom: "1px solid #ddd6fe", fontSize: 11, color: "#4c1d95" }}>
                  {routingRules.length} routing/skip rules extracted from the Word questionnaire
                </div>
                <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
                  <thead style={{ position: "sticky", top: 0, background: "#f8fafc", zIndex: 1 }}>
                    <tr><Th>#</Th><Th>Condition</Th><Th>Ask (targets)</Th><Th>Skip</Th><Th>Source line</Th></tr>
                  </thead>
                  <tbody>
                    {routingRules.filter(r => !search || r.condVar.toLowerCase().includes(search.toLowerCase()) || r.rawText.toLowerCase().includes(search.toLowerCase())).map((r, i) => (
                      <tr key={i} style={{ background: i % 2 === 0 ? "#fff" : "#fafafa" }}>
                        <Td style={{ color: "#94a3b8", fontSize: 10 }}>{i + 1}</Td>
                        <Td><code style={{ fontWeight: 700, color: "#7c3aed" }}>{r.condVar}{r.condOp}{r.condVals.join(",")}</code></Td>
                        <Td>
                          <div style={{ display: "flex", gap: 3, flexWrap: "wrap" }}>
                            {r.targets.map(t => <span key={t} style={{ background: "#eff6ff", border: "1px solid #bfdbfe", borderRadius: 3, padding: "1px 5px", fontSize: 10, color: "#1d4ed8" }}>{t}</span>)}
                          </div>
                        </Td>
                        <Td>
                          <div style={{ display: "flex", gap: 3, flexWrap: "wrap" }}>
                            {r.skipTargets.map(t => <span key={t} style={{ background: "#fef2f2", border: "1px solid #fca5a5", borderRadius: 3, padding: "1px 5px", fontSize: 10, color: "#dc2626" }}>{t}</span>)}
                          </div>
                        </Td>
                        <Td style={{ fontSize: 10, color: "#64748b", maxWidth: 360, wordBreak: "break-word" }}>{r.rawText.slice(0, 180)}</Td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )}

            {/* ── Docx text tab ── */}
            {tab === "docxtext" && (
              <div style={{ background: "#fff", borderRadius: 10, border: "1px solid #e2e8f0", overflow: "auto", maxHeight: 540 }}>
                <div style={{ padding: "9px 13px", background: "#f5f3ff", borderBottom: "1px solid #ddd6fe", fontSize: 11, color: "#4c1d95" }}>
                  {docxFileName} · {docxRawText.length.toLocaleString()} chars · {nDocxQ} questions detected
                </div>
                <pre style={{ padding: "12px 16px", fontSize: 11, color: "#334155", lineHeight: 1.8, whiteSpace: "pre-wrap", wordBreak: "break-word", margin: 0 }}>
                  {docxRawText.slice(0, 60000)}
                  {docxRawText.length > 60000 && `\n… (showing first 60,000 of ${docxRawText.length.toLocaleString()} chars)`}
                </pre>
              </div>
            )}
          </>
        )}

        {/* Empty state */}
        {!analyzed && (
          <div style={{ textAlign: "center", padding: "50px 20px", color: "#94a3b8" }}>
            <div style={{ fontSize: 50, marginBottom: 14 }}>📋</div>
            <div style={{ fontSize: 15, fontWeight: 600, color: "#64748b" }}>Upload your files to begin validation</div>
            <div style={{ fontSize: 12, marginTop: 8, maxWidth: 500, margin: "10px auto 0", lineHeight: 1.8, color: "#64748b" }}>
              <strong>This validator is fully dynamic — it learns from your files:</strong><br />
              🗃 <strong>.sav</strong> → defines which codes are valid for each variable (primary authority)<br />
              📝 <strong>.docx</strong> → defines routing logic, skip conditions, question labels<br />
              📋 <strong>.csv/.xlsx</strong> → the actual data to validate against the above<br /><br />
              No survey-specific rules are hardcoded. Upload any questionnaire dataset.
            </div>
          </div>
        )}
      </div>

      {/* Column detail side panel */}
      {selectedCol && <ColumnPanel col={selectedCol} onClose={() => setSelectedCol(null)} />}
    </div>
  );
}
