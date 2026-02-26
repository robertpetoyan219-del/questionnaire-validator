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
// UNIVERSAL SPECIAL CODES
// These appear across ALL your surveys. Never flag them as invalid answer codes.
// ─────────────────────────────────────────────────────────────────────────────
const UNIVERSAL_SPECIAL_CODES = new Set([
  99, 999, 9999,     // Refuse / DK / no answer
  98, 998,           // None / DK / "none of them"
  97, 997,           // Other (specify)
  0,                 // NOT SELECTED on binary multi-select dummies
  666666,            // LimeSurvey internal "other" code
]);

// ─────────────────────────────────────────────────────────────────────────────
// SAV BINARY PARSER
// Parses SPSS .sav (System file) format:
// Record type 2 = variable records, 3/4 = value label records, 7 = info records
// ─────────────────────────────────────────────────────────────────────────────

// Try to decode bytes as UTF-8; if it fails, try Windows-1251 (Cyrillic/Armenian SPSS files),
// then Windows-1252, then fall back to latin1.
function smartDecode(bytes: Uint8Array, start: number, len: number): string {
  const slice = bytes.slice(start, start + len);
  // 1. Try UTF-8 (preferred)
  try {
    const utf8 = new TextDecoder("utf-8", { fatal: true }).decode(slice);
    return utf8.trimEnd();
  } catch { /* not valid UTF-8 */ }
  // 2. Try Windows-1251 (common for Armenian/Cyrillic SPSS exports)
  try {
    const win1251 = new TextDecoder("windows-1251", { fatal: false }).decode(slice);
    // Check if result contains meaningful Cyrillic or Armenian characters
    const hasCyrillic = /[\u0400-\u04FF\u0531-\u058F]/.test(win1251);
    if (hasCyrillic) return win1251.trimEnd();
  } catch { /* decoder not available */ }
  // 3. Try Windows-1252 (Western European)
  try {
    const win1252 = new TextDecoder("windows-1252", { fatal: false }).decode(slice);
    return win1252.trimEnd();
  } catch { /* decoder not available */ }
  // 4. Final fallback: latin1
  return new TextDecoder("latin1").decode(slice).trimEnd();
}

function parseSavFile(buffer: ArrayBuffer): {
  variables: SavVariable[];
  varMap: Record<string, SavVariable>;
} {
  const bytes = new Uint8Array(buffer);

  const readI32s = (o: number) => // signed
    bytes[o] | (bytes[o+1]<<8) | (bytes[o+2]<<16) | (bytes[o+3]<<24);
  const readF64 = (o: number) => new DataView(buffer, o, 8).getFloat64(0, true);
  const readStr = (o: number, n: number) => smartDecode(bytes, o, n);
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
  // Only scan for ASCII variable names; non-ASCII is labels/text, not variable names
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
  let currentSection = "General";
  let currentQ: DocxQuestion | null = null;

  // ── Helpers ────────────────────────────────────────────────────────────────

  // Normalize Unicode operators → ASCII
  function normalizeOp(op: string): string {
    return op === "≠" ? "!=" : op === "≤" ? "<=" : op === "≥" ? ">=" : op;
  }

  // Parse "1-4", "1,2,3", "9,10", "1-3,5" into [1,2,3,4] / [1,2,3] / [9,10] / [1,2,3,5]
  function parseCondVals(raw: string): number[] {
    const vals = new Set<number>();
    for (const part of raw.split(",")) {
      const r = part.trim().match(/^(\d+)\s*[-–]\s*(\d+)$/);
      if (r) { for (let v = +r[1]; v <= +r[2]; v++) vals.add(v); }
      else { const n = parseInt(part.trim()); if (!isNaN(n)) vals.add(n); }
    }
    return [...vals];
  }

  // Detect scale/range patterns in a line of text.
  // The phrasing in Armenian surveys is typically:
  //   "...1-7 բdelays delays delays, delays 1-delays delays X, 7-delays delays Y..."
  //   Meaning: "...answer on a 1-7 point scale, where 1 means X, 7 means Y..."
  // The numbers are always ASCII digits. We match the N-M range pattern in context.
  //
  // Patterns matched:
  //   1) "N-M" followed by Armenian script word(s) — the most common Armenian format
  //   2) "N-M" preceded by Armenian script word(s) — alternate word order
  //   3) "SCALE N-M" / "SCALE: N-M" — English instruction format
  //   4) "scale of N to M" / "N to M scale/point" — English prose
  //   5) "шdelays N-M" / "N-M шdelays" — Russian format
  //   6) "where N means..." / "delays N-delays..." pattern confirming scale endpoints
  function extractScaleRange(text: string): [number, number] | null {
    // Pattern A: N-M followed or preceded by Armenian/Cyrillic text (the word for "scale", "point", etc.)
    // We don't hardcode the Armenian words — we detect: digits-dash-digits near Armenian script
    const rangeMatch = text.match(/(\d{1,2})\s*[-–—]\s*(\d{1,2})/);
    if (rangeMatch) {
      const lo = parseInt(rangeMatch[1]);
      const hi = parseInt(rangeMatch[2]);
      if (!isNaN(lo) && !isNaN(hi) && hi > lo && (hi - lo) >= 2 && (hi - lo) <= 20) {
        // Check context: is this range near Armenian/Cyrillic words that suggest a scale?
        // Armenian Unicode: \u0531-\u058F, Cyrillic: \u0400-\u04FF
        const hasArmenianContext = /[\u0531-\u058F]{2,}/.test(text);
        const hasCyrillicContext = /[\u0400-\u04FF]{2,}/.test(text);

        // English scale keywords
        const hasEnglishScaleWord = /\b(scale|point|score|rate|rating|grading)\b/i.test(text);

        // "SCALE" as standalone instruction
        const isScaleInstruction = /^SCALE\b/i.test(text.trim());

        // The range appears in text that also has "where N means" / "delays N-" endpoint explanation
        // This is the "delays 1-delays X, 7-delays Y" pattern
        const hasEndpointExplanation = new RegExp(
          `\\b${lo}\\b[^\\d]*\\b${hi}\\b`
        ).test(text) && text.length > 20;

        if (hasArmenianContext || hasCyrillicContext || hasEnglishScaleWord || isScaleInstruction || hasEndpointExplanation) {
          return [lo, hi];
        }

        // Also match if the line is short and looks like a scale instruction
        // e.g., "1-7" alone or "1-10 point" or "0-10"
        if (text.trim().length < 30 && /^\s*\d{1,2}\s*[-–—]\s*\d{1,2}\s*\S*\s*$/.test(text.trim())) {
          return [lo, hi];
        }
      }
    }

    // Pattern B: "scale of N to M" / "from N to M"
    const toMatch = text.match(/(?:scale|шкал|rating)\s*(?:of|от|из)?\s*(\d{1,2})\s*(?:to|до|[-–—])\s*(\d{1,2})/i);
    if (toMatch) {
      const lo = parseInt(toMatch[1]);
      const hi = parseInt(toMatch[2]);
      if (!isNaN(lo) && !isNaN(hi) && hi > lo && (hi - lo) >= 2 && (hi - lo) <= 20) {
        return [lo, hi];
      }
    }

    return null;
  }

  // ── Line classification helpers ────────────────────────────────────────────

  // Lines that are purely interviewer instructions — never a question or answer code
  const INSTRUCTION_RE = /^(ԿԱՐԴԱԼ|ՉԿԱՐԴԱԼ|Read\s+if|Do\s+not\s+read|INT\s*:|Rotate|Bring\s+brands|Control\s+(list|the)|If\s+in\s+.*\bdo\s+not|Automatically\s+code|Show\s+\d|REGISTER|INTEGRATE|MULTIPLE|SINGLE|OPEN\s+(RESPONSE|ANSWER)|Multiple\s+response|Single\s+response|Single\s+answer|Multiple\s+answer|Open\s+(response|answer)|ՄԻ\s+ՔԱՆԻ|ՄԵԿ\s+ՊԱՏاHԱՆ|ROTATE|SCALE|BRING|CORON|Recall|Remind|Read\s+the|Record|Please\s+be|Your\s+Tel|This\s+is|This\s+call|Text\s+in|NPS|MAX\s+\d|Recoding\s+key)/i;

  // Section header patterns:
  //   =ԲԱԺԻՆ 1. TITLE= (Armenian)
  //   = Мас 1. TITLE   (Russian transliteration)
  //   Section1: Title  (English)
  //   =бажин N.        (lowercase Armenian)
  const SECTION_RE = /^(?:=+\s*(?:ԲԱԺԻՆ|бажин|Мас|МАС|SECTION|Section)\s*\d|Section\s*\d\s*:|=\s*Мас\s+\d)/i;

  // Routing trigger lines — the whole line is a routing instruction
  // Must start with Ask/ASK or End/Terminate, and contain a condition
  const ROUTING_LINE_RE = /^(?:Ask\s+(?:if|all|this\s+section|the\s+section|questions)|ASK\s+(?:IF|ALL|THIS\s+SECTION|THE\s+SECTION|QUESTIONS)|End\s+if|Terminate\s+if)/i;

  // Answer option line: starts with a number then ":" or "։" (Armenian colon) or ")"
  // e.g. "1: Yes", "2։ No (Terminate)", "97: Other"
  const ANSWER_RE = /^(\d{1,5})\s*[：:։)]\s*(.+)$/;

  // Question line: VAR_CODE. Question text
  // Code must be ALL-CAPS start, contain at least one digit OR be a known alpha code (S0, B1, etc.)
  // Separators: . ․ (Armenian middle dot) : – — )
  // The code must NOT be a pure number
  const QUESTION_RE = /^([A-Z][A-Za-z0-9_.]{0,19})\s*[.․:–—)]\s*(.{2,})/;

  // Words/phrases that look like question codes but are actually instructions
  const FAKE_CODE_WORDS = new Set([
    "CATI","CAPI","NPS","INT","END","ASK","ALL","MULTIPLE","SINGLE","OPEN",
    "ROTATE","BRING","CONTROL","SCALE","READ","DO","ID","CODE","AGE","SEX",
    "YES","NO","MAX","MIN","TOM","PROM","TEXT","Note","URL","PC","TV","SMS",
  ]);

  // ── Routing condition parser ───────────────────────────────────────────────
  // Handles:
  //   "Ask if S0=1"
  //   "Ask if S5=1 and (S7≠1 AND S7≠2) and S9≠1"   → multiple conditions (AND)
  //   "ASK QUESTIONS G2-G6 IF G1=1-3"
  //   "ASK THIS SECTION IF BR1=2 or 3"              → "or 3" = additional value for same var
  //   "End if S4=98 or 999"
  //   "Terminate if S5=2 and S9=2"  (from inline answer option text)
  //   "ASK IF E16 ≠ 97"
  //   "ASK IF E3<5"
  //   "ASK IF BR1=1-3"

  function parseRoutingLine(line: string): RoutingRule[] {
    const rules: RoutingRule[] = [];

    // Strip leading verb: "Ask if", "ASK IF", "ASK QUESTIONS X-Y IF", "ASK THIS SECTION IF",
    //                     "End if", "Terminate if"
    let body = line
      .replace(/^Ask\s+(?:if|questions\s+[A-Z0-9_,\-–]+\s+if|this\s+section\s+if|the\s+section\s+if)/i, "")
      .replace(/^ASK\s+(?:IF|QUESTIONS\s+[A-Z0-9_,\-–]+\s+IF|THIS\s+SECTION\s+IF|THE\s+SECTION\s+IF)/i, "")
      .replace(/^(?:End|Terminate)\s+if/i, "")
      .trim();

    // Split on AND/and/&&, process each condition sub-clause
    // Each sub-clause looks like: VAR OP VALUE  or  VAR OP VALUE1,VALUE2  or  (VAR OP VALUE)
    const clauses = body.split(/\s+AND\s+|\s*&&\s*/i);

    for (const clause of clauses) {
      const clean = clause.replace(/[()]/g, "").trim();

      // Single condition: VAR OP VALS
      // Also handles "or N" continuation: "S4=98 or 999" → vals=[98,999]
      const m = clean.match(
        /^([A-Z][A-Za-z0-9_.]{0,19})\s*([=!<>≠]{1,2})\s*([\d,\-–]+(?:\s+or\s+[\d,\-–]+)*)/i
      );
      if (m) {
        const condVar = m[1];
        const condOp = normalizeOp(m[2]);
        // Handle "98 or 999" → "98,999"
        const valStr = m[3].replace(/\s+or\s+/gi, ",");
        const condVals = parseCondVals(valStr);
        if (condVar && condVals.length > 0) {
          rules.push({ condVar, condOp, condVals, targets: [], skipTargets: [], rawText: line.slice(0, 200) });
        }
      }
    }

    return rules;
  }

  // ── Per-line processing ───────────────────────────────────────────────────

  for (const line of lines) {

    // 1. Section headers
    if (SECTION_RE.test(line)) {
      currentSection = line.replace(/^=+|=+$/g, "").trim().slice(0, 80);
      currentQ = null;
      continue;
    }

    // 2. Pure instruction lines — but first extract scale hints before skipping
    if (INSTRUCTION_RE.test(line)) {
      // Before skipping, check if this instruction mentions a scale range
      if (currentQ) {
        const scaleRange = extractScaleRange(line);
        if (scaleRange) {
          for (let v = scaleRange[0]; v <= scaleRange[1]; v++) {
            if (!currentQ.validCodes.includes(v)) currentQ.validCodes.push(v);
          }
        }
      }
      continue;
    }

    // 3. Routing lines (standalone "Ask if …" / "End if …")
    if (ROUTING_LINE_RE.test(line)) {
      const parsed = parseRoutingLine(line);
      routingRules.push(...parsed);
      // Don't set currentQ — routing line is a filter, not a question
      continue;
    }

    // 4. Answer option lines — "N: text" or "N։ text"
    const am = line.match(ANSWER_RE);
    if (am) {
      const code = parseInt(am[1]);
      const label = am[2].trim();
      if (!isNaN(code) && label.length > 0) {
        // Attach to current question's code labels
        if (currentQ && !UNIVERSAL_SPECIAL_CODES.has(code)) {
          // Strip inline instructions from label: "(Terminate)", "TERMINATE", "(go to X)", "ЧKАРДАЛ", "ЧКАРDАЛ"
          const cleanLabel = label
            .replace(/\s*\(Terminate[^)]*\)/gi, "")
            .replace(/\s*TERMINATE\b/gi, "")
            .replace(/\s*\(go\s+to\s+\S+\)/gi, "")
            .replace(/\s*→\s*Go\s+to\s+\S+.*/i, "")
            .replace(/\s*CONTINUE\b/gi, "")
            .replace(/ՉԿԱՐԴԱԼchinese|ЧКАРDАЛ|ЧKАРДАЛ/g, "")
            .trim();
          if (cleanLabel) {
            currentQ.codeLabels[code] = cleanLabel;
            if (!currentQ.validCodes.includes(code)) currentQ.validCodes.push(code);
          }
        }

        // Check for inline routing: "(Terminate if X=N)" / "(go to QN)"
        const terminateInline = label.match(/(?:Terminate\s+if|terminate\s+if)\s+([A-Z][A-Za-z0-9_.]{0,19})\s*([=!<>≠]{1,2})\s*([\d,\-–]+(?:\s+and\s+[A-Z][A-Za-z0-9_.]{0,19}\s*[=!<>≠]{1,2}\s*[\d,\-–]+)*)/i);
        if (terminateInline) {
          const parsed = parseRoutingLine("Terminate if " + terminateInline[1] + terminateInline[2] + terminateInline[3]);
          routingRules.push(...parsed);
        }
        // "(go to QN)" — note as skip but without condVar we can't enforce it
        // Arrow "→ Go to QN" in answer option
      }
      continue;
    }

    // 5. Question lines — "VARCODE. Question text" or "VARCODE․ text"
    const qm = line.match(QUESTION_RE);
    if (qm) {
      const code = qm[1];
      const label = qm[2].trim();

      // Reject fake codes
      if (FAKE_CODE_WORDS.has(code)) continue;
      // Code must look like a survey variable: starts uppercase, reasonable length
      if (code.length < 1 || code.length > 15) continue;
      // Must contain a digit OR be a well-known single-letter+digit pattern (S0, B1, G9, etc.)
      // Reject pure alphabetic codes that are too long (likely English words/instructions)
      const looksLikeVar = /\d/.test(code) || /^[A-Z]\d/.test(code);
      if (!looksLikeVar) continue;
      // Reject if label starts with a lowercase letter (likely a sentence continuation, not a question)
      // Exception: Armenian characters always start uppercase
      if (/^[a-z]/.test(label)) continue;

      currentQ = {
        code,
        label: label.slice(0, 300),
        validCodes: [],
        codeLabels: {},
        section: currentSection,
      };

      // Check if the question label itself contains a scale range hint
      // e.g., "...1-7 բdelays..., delays 1 means X, 7 means Y..."
      const scaleFromLabel = extractScaleRange(label);
      if (scaleFromLabel) {
        for (let v = scaleFromLabel[0]; v <= scaleFromLabel[1]; v++) {
          if (!currentQ.validCodes.includes(v)) currentQ.validCodes.push(v);
        }
      }

      // Only register the first occurrence of each code
      if (!questionMap[code]) {
        questions.push(currentQ);
        questionMap[code] = currentQ;
      } else {
        // Update section/label if we see it again with more context
        currentQ = questionMap[code];
      }
      continue;
    }

    // 6. Everything else — but check for scale hints in continuation lines
    // Lines following a question like "Կrequire 1-7 բdelays" or
    // "Answer on a 1 to 7 scale, where 1 means definitely not, 7 means definitely yes"
    if (currentQ) {
      const contScale = extractScaleRange(line);
      if (contScale) {
        for (let v = contScale[0]; v <= contScale[1]; v++) {
          if (!currentQ.validCodes.includes(v)) currentQ.validCodes.push(v);
        }
      }
    }
  }

  return { questions, questionMap, routingRules, rawText, currentSection };
}

// ─────────────────────────────────────────────────────────────────────────────
// SURVEY CONVENTION KNOWLEDGE
// Derived from analysis of 4 real SAV+DOCX pairs across multiple projects.
// ─────────────────────────────────────────────────────────────────────────────

// Variable name patterns that are NEVER content-validated (only structural):
// – open-text / verbatim variables: _T, _1T, _97T, _977T, _other, T suffix after digit
// – coding companion variables: _coding, _coding1, _coding2
// – specify companion variables (Barerar pattern): S_ prefix
// – platform admin variables
// – TOM (Top of Mind) positional mention variables: *otherSpont*, *_TOM
// – iteration tracking (Barerar CAPI): I_N_* pattern
// – derived/recoded variables with English names (Region, AGE, CARS, etc.)
function isSkipValidationVar(varName: string, sv?: SavVariable): boolean {
  const n = varName;
  // Open-text suffixes
  if (/(_1T|_T|_97T|_977T|_other|_OTHER)$/i.test(n)) return true;
  // Coding companions
  if (/(_coding\d*$)/i.test(n)) return true;
  // Specify prefix (Barerar: S_QN_N)
  if (/^S_/.test(n)) return true;
  // Spontaneous mention / TOM positional
  if (/otherSpont/i.test(n) || /_TOM$/i.test(n) || /TOM_/.test(n)) return true;
  // Iteration tracking (Barerar: I_N_*)
  if (/^I_\d+_/.test(n)) return true;
  // Platform admin fields
  if (/^(submitdate|lastpage|startlanguage|seed|startdate|datestamp|IVDate|IVDur|ContactID|UserID|UserName|UserLgIn|Latitude|Longitude|SbjNum)$/i.test(n)) return true;
  // SAV string type — open text
  if (sv?.type === "string") return true;
  return false;
}

// Detect if a variable is a multi-select binary dummy (valid codes = {0,1} only)
function isBinaryDummy(sv?: SavVariable): boolean {
  if (!sv) return false;
  const codes = sv.validCodes;
  if (codes.length === 0) return false;
  // Exclusively 0 and/or 1
  return codes.every(c => c === 0 || c === 1);
}

// Detect if a variable looks like a scale (e.g., 1-7, 1-10, 0-10).
// SAV files often only define endpoint labels like "1=Definitely not" and "7=Definitely yes",
// but all integers in between are valid answers.
//
// Detection strategies:
// 1. Question label text contains a range pattern (N-M) near Armenian/Cyrillic/English text.
//    Armenian surveys phrase it as: "Answer on a 1-7 scale, where 1 means X and 7 means Y"
//    We don't hardcode Armenian words — we detect N-M near Armenian/Cyrillic script.
// 2. Docx parser already expanded scale codes during parsing (from question/instruction text).
// 3. SAV defines only 2-4 codes with a gap suggesting scale endpoints (e.g., {1,7} with labels).
function detectScaleRange(
  sv?: SavVariable,
  dq?: DocxQuestion,
): number[] | null {
  // Strategy 1: Check question label text for N-M range pattern in context
  const labelText = (sv?.label ?? "") + " " + (dq?.label ?? "");
  if (labelText.trim()) {
    const rangeMatch = labelText.match(/(\d{1,2})\s*[-–—]\s*(\d{1,2})/);
    if (rangeMatch) {
      const lo = parseInt(rangeMatch[1]);
      const hi = parseInt(rangeMatch[2]);
      if (!isNaN(lo) && !isNaN(hi) && hi > lo && (hi - lo) >= 2 && (hi - lo) <= 20) {
        // Check if text context suggests this is a scale:
        // Armenian script nearby (Ա-֏) — common in Armenian surveys
        // Cyrillic script nearby (Ѐ-ӿ) — Russian surveys
        // English scale keywords
        const hasArmenian = /[Ա-֏]{2,}/.test(labelText);
        const hasCyrillic = /[Ѐ-ӿ]{2,}/.test(labelText);
        const hasEnglishScale = /(scale|point|score|rate|rating)/i.test(labelText);

        if (hasArmenian || hasCyrillic || hasEnglishScale) {
          const range: number[] = [];
          for (let v = lo; v <= hi; v++) range.push(v);
          return range;
        }

        // Also confirm if SAV/docx codes look like endpoints of this range
        const codes = sv?.validCodes ?? dq?.validCodes ?? [];
        const codesInRange = codes.filter(c => c >= lo && c <= hi);
        if (codesInRange.length >= 1 && codesInRange.length < (hi - lo + 1)) {
          const range: number[] = [];
          for (let v = lo; v <= hi; v++) range.push(v);
          return range;
        }
      }
    }
  }

  // Strategy 2: Docx already expanded — if docx validCodes look like a full consecutive range,
  // it means the docx parser detected a scale. Confirm by checking consecutiveness.
  if (dq && dq.validCodes.length >= 3) {
    const sorted = [...dq.validCodes].sort((a, b) => a - b);
    const lo = sorted[0];
    const hi = sorted[sorted.length - 1];
    const isConsecutive = sorted.length === (hi - lo + 1) && sorted.every((v, i) => v === lo + i);
    if (isConsecutive && (hi - lo) >= 2 && (hi - lo) <= 20) {
      const range: number[] = [];
      for (let v = lo; v <= hi; v++) range.push(v);
      return range;
    }
  }

  // Strategy 3: SAV heuristic — if SAV defines exactly 2-4 codes that look like
  // scale endpoints (e.g., {1,7} or {1,5,10} or {0,10}), expand the range.
  if (sv && sv.validCodes.length >= 2 && sv.validCodes.length <= 4) {
    const codes = [...sv.validCodes].sort((a, b) => a - b);
    const lo = codes[0];
    const hi = codes[codes.length - 1];
    if (hi > lo && (hi - lo) >= 3 && (hi - lo) <= 20) {
      const expectedCount = hi - lo + 1;
      const definedInRange = codes.filter(c => c >= lo && c <= hi).length;
      // If less than half the range has defined labels -> scale with only endpoint labels
      if (definedInRange < expectedCount * 0.6) {
        const hasEndpointLabels = sv.valueLabels[lo] && sv.valueLabels[hi];
        if (hasEndpointLabels) {
          const range: number[] = [];
          for (let v = lo; v <= hi; v++) range.push(v);
          return range;
        }
      }
    }
  }

  return null;
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

  function evalCondition(row: RowData, rule: RoutingRule): boolean {
    const x = getNum(row, rule.condVar);
    if (x === null) return false;
    if (rule.condVals.length === 0) return false; // no values to evaluate
    switch (rule.condOp) {
      case "=":  return rule.condVals.includes(x);
      case "!=": return !rule.condVals.includes(x);
      case "<":  return x < rule.condVals[0];
      case ">":  return x > rule.condVals[rule.condVals.length - 1]; // compare against max of range
      case "<=": return x <= rule.condVals[rule.condVals.length - 1];
      case ">=": return x >= rule.condVals[0];
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

  // ── Scale detection pre-computation ──────────────────────────────────────
  // For each variable, check if it's a scale and compute expanded valid codes
  const scaleRanges: Record<string, number[]> = {};
  const effectiveValidCodes: Record<string, number[]> = {};
  for (const sv of savVars) {
    const dq = docxQMap[sv.name];
    const scaleRange = detectScaleRange(sv, dq);
    if (scaleRange) {
      scaleRanges[sv.name] = scaleRange;
      // Merge scale range with any explicitly defined codes + universal special codes
      const merged = new Set([...scaleRange, ...sv.validCodes]);
      effectiveValidCodes[sv.name] = [...merged].sort((a, b) => a - b);
    } else {
      effectiveValidCodes[sv.name] = sv.validCodes;
    }
  }
  // Also check docx-only variables for scales
  for (const dq of Object.values(docxQMap)) {
    if (savVarMap[dq.code]) continue; // already handled above
    const scaleRange = detectScaleRange(undefined, dq);
    if (scaleRange) {
      scaleRanges[dq.code] = scaleRange;
      const merged = new Set([...scaleRange, ...dq.validCodes]);
      effectiveValidCodes[dq.code] = [...merged].sort((a, b) => a - b);
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
      if (emptyVal(row, sv.name)) continue;

      // Skip open-text, coding companions, admin, iteration, TOM, binary dummies
      if (isSkipValidationVar(sv.name, sv)) continue;
      // Binary dummy vars (0/1): only values 0 and 1 are meaningful — skip deep validation
      if (isBinaryDummy(sv)) continue;

      const val = getNum(row, sv.name);
      if (val === null) continue;

      // Never flag declared missing values or universal special codes
      if (sv.missingValues.includes(val)) continue;
      if (UNIVERSAL_SPECIAL_CODES.has(val)) continue;

      // Use scale-expanded codes if detected, otherwise original SAV codes
      const validForCheck = effectiveValidCodes[sv.name] ?? sv.validCodes;
      const isScale = !!scaleRanges[sv.name];

      if (!validForCheck.includes(val)) {
        const lbl = getVarLabel(sv.name, savVarMap, docxQMap);
        // Build a human-readable code list: "1=Yes, 2=No, 3=DK"
        const codeDesc = sv.validCodes
          .map(c => sv.valueLabels[c] ? `${c}=${sv.valueLabels[c]}` : String(c))
          .slice(0, 15).join(", ") + (sv.validCodes.length > 15 ? "…" : "");
        const scaleNote = isScale
          ? `\nNote: This variable was detected as a scale (${scaleRanges[sv.name][0]}-${scaleRanges[sv.name][scaleRanges[sv.name].length - 1]}). All values within the scale range are considered valid.`
          : "";
        flagColTracked(id, sv.name, isScale ? "OUT_OF_RANGE" : "MISMATCHED_CODE", val,
          `${sv.name}=${val} — ${isScale ? "out of scale range" : "not a valid answer code"}`,
          `Question: ${lbl}\nValid codes from SAV: ${codeDesc}${isScale ? ` (scale: ${scaleRanges[sv.name][0]}–${scaleRanges[sv.name][scaleRanges[sv.name].length - 1]})` : ""}\nObserved value: ${val}${scaleNote}\nNote: codes 97/98/99/999 are always allowed as DK/Refuse/Other.`
        );
      }
    }

    // 2. Docx-sourced valid code checks (variables not in SAV but in docx)
    for (const dq of Object.values(docxQMap)) {
      if (!dataColumns.has(dq.code)) continue;
      if (savVarMap[dq.code]?.validCodes.length > 0) continue; // already handled by SAV
      if (emptyVal(row, dq.code)) continue;
      if (isSkipValidationVar(dq.code, savVarMap[dq.code])) continue;
      const val = getNum(row, dq.code);
      if (val === null) continue;
      if (UNIVERSAL_SPECIAL_CODES.has(val)) continue;
      // Use scale-expanded codes if detected
      const docxEffective = effectiveValidCodes[dq.code] ?? dq.validCodes;
      const isScale = !!scaleRanges[dq.code];
      if (docxEffective.length > 0 && !docxEffective.includes(val)) {
        const codeDesc = dq.validCodes
          .map(c => dq.codeLabels[c] ? `${c}=${dq.codeLabels[c]}` : String(c))
          .slice(0, 12).join(", ");
        flagColTracked(id, dq.code, isScale ? "OUT_OF_RANGE" : "MISMATCHED_CODE", val,
          `${dq.code}=${val} — ${isScale ? "out of scale range" : "not a valid answer code"} (from questionnaire)`,
          `Question: ${dq.label}\nValid codes from questionnaire: ${codeDesc}${isScale ? ` (scale: ${scaleRanges[dq.code][0]}–${scaleRanges[dq.code][scaleRanges[dq.code].length - 1]})` : ""}\nObserved value: ${val} is not among them.`
        );
      }
    }

    // 3. Routing rule violations
    // Strategy:
    //   • condMet + target empty     → MISSING_DATA (required but absent)      ✓ always check
    //   • condMet + skipTarget filled → SKIP_VIOLATION (should have been skipped) ✓ always check
    //   • condNotMet + target filled  → only check for explicit SKIP rules (skipTargets list),
    //                                   NOT for plain "ASK IF / IF … ASK" rules.
    //     Reason: most questions appear in many routing rules as targets; flagging them whenever
    //     *any one* condition is unmet causes massive false positives.  Only explicit "SKIP TO X
    //     IF Y≠1" rules reliably tell us the variable must be empty.
    for (const rule of routingRules) {
      if (!dataColumns.has(rule.condVar)) continue;
      const condVarVal = getNum(row, rule.condVar);
      if (condVarVal === null) continue; // condVar itself is blank — can't evaluate

      const condMet = evalCondition(row, rule);

      const condDesc = `${rule.condVar}${rule.condOp}${rule.condVals.join(",")}`;
      const condVarLbl = getVarLabel(rule.condVar, savVarMap, docxQMap);
      const condValDescs = rule.condVals.map(v => {
        const lbl = getCodeLabel(rule.condVar, v, savVarMap, docxQMap);
        return lbl ? `${v}=${lbl}` : String(v);
      }).join(", ");

      if (condMet) {
        // Condition met → named targets must be present
        for (const target of rule.targets) {
          if (target === "TERMINATE") continue;
          if (!dataColumns.has(target)) continue;
          if (emptyVal(row, target)) {
            const tLbl = getVarLabel(target, savVarMap, docxQMap);
            flagColTracked(id, target, "MISSING_DATA", null,
              `${target} missing — required when ${condDesc}`,
              `Routing rule: "${rule.rawText}"\n${rule.condVar} (${condVarLbl}) = ${condValDescs}, so ${target} (${tLbl}) must be answered.\nCurrent value of ${rule.condVar}: ${condVarVal}`
            );
          }
        }
        // Condition met → explicit skipTargets must be empty
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
      } else {
        // Condition NOT met → only enforce explicit skip/terminate targets
        // (i.e., variables that the docx literally says to skip)
        for (const skip of rule.skipTargets) {
          if (!dataColumns.has(skip)) continue;
          const val = getNum(row, skip);
          if (hasVal(row, skip) && (val === null || !UNIVERSAL_SPECIAL_CODES.has(val))) {
            const sLbl = getVarLabel(skip, savVarMap, docxQMap);
            const cLbl = getCodeLabel(rule.condVar, condVarVal, savVarMap, docxQMap);
            flagColTracked(id, skip, "SKIP_VIOLATION", val,
              `${skip} filled but should be skipped — ${rule.condVar}=${condVarVal}${cLbl ? ` (${cLbl})` : ""} (condition ${condDesc} not met)`,
              `Routing rule: "${rule.rawText}"\nThe questionnaire explicitly says to skip ${skip} (${sLbl}) when condition ${condDesc} is not met.\nCurrent ${rule.condVar}=${condVarVal}${cLbl ? ` (${cLbl})` : ""}.`
            );
          }
        }
      }
    }

    // 4. Open text quality check (garbled text detection)
    // Applies to: SAV string vars, _T/_1T/_97T/_other suffix vars, S_ prefix vars (Barerar)
    for (const col of Object.keys(row)) {
      const sv = savVarMap[col];
      const isOpenText =
        (sv?.type === "string") ||
        /(_1T|_T|_97T|_977T|_other|_OTHER)$/i.test(col) ||
        /^S_/.test(col);
      if (!isOpenText) continue;
      // Skip coding companions even if they look like open text
      if (/(_coding\d*$)/i.test(col)) continue;

      const txt = getStr(row, col);
      if (!txt || txt.length < 3) continue;

      const total = txt.replace(/\s/g, "").length;
      if (total < 3) continue;

      // Check 1: Encoding artifact patterns — mojibake from double-encoding or wrong codepage
      // Common patterns: Ô (Armenian in latin1), Ã (UTF-8 misread as latin1),
      // Â (BOM artifacts), repeated Ð/Ñ (Cyrillic as latin1)
      const encodingArtifacts = (txt.match(/[ÃÂÔÕÖàáâãäåèéêëìíîïðñòóôõöøùúûüý]{2,}/g) ?? []).join("").length;
      const hasEncodingGarble = encodingArtifacts / total > 0.30;

      // Check 2: Meaningful character ratio — Armenian, Cyrillic, Latin, digits, punctuation
      // Armenian: \u0531-\u058F (uppercase + lowercase)
      // Cyrillic: \u0400-\u04FF
      // Latin: a-zA-Z
      // Also count digits and common punctuation as meaningful
      const meaningfulChars = (txt.match(/[\u0531-\u058F\u0400-\u04FFa-zA-Z0-9.,;:!?'"()\-–—/]/g) ?? []).length;
      const meaningfulRatio = meaningfulChars / total;
      const hasLowMeaningful = total > 5 && meaningfulRatio < 0.30;

      // Check 3: Repetitive character patterns (keyboard mashing: "asdfasdf", "ааааа", "11111")
      const hasRepetition = /(.)\1{4,}/.test(txt) || // same char 5+ times
        /(.{2,4})\1{2,}/.test(txt); // same 2-4 char pattern 3+ times

      // Check 4: Control characters or unusual Unicode blocks that shouldn't appear in survey text
      const controlChars = (txt.match(/[\x00-\x08\x0E-\x1F\x7F-\x9F\uFFFD]/g) ?? []).length;
      const hasControlGarble = controlChars > 0 && controlChars / total > 0.10;

      // Check 5: Mixed script gibberish — text that mixes unrelated scripts unnaturally
      // (e.g., Armenian + CJK + Latin control chars in the same short string)
      const hasArmenian = /[\u0531-\u058F]/.test(txt);
      const hasCyrillic = /[\u0400-\u04FF]/.test(txt);
      const hasLatin = /[a-zA-Z]/.test(txt);
      const hasCJK = /[\u4E00-\u9FFF\u3040-\u309F\u30A0-\u30FF]/.test(txt);
      const scriptCount = [hasArmenian, hasCyrillic, hasLatin, hasCJK].filter(Boolean).length;
      const suspiciousMixedScript = scriptCount >= 3 && total < 50;

      // Determine issue type and message
      let garbleReason = "";
      if (hasEncodingGarble) {
        garbleReason = "encoding artifacts detected (possible double-encoding or wrong codepage)";
      } else if (hasControlGarble) {
        garbleReason = "contains control characters or replacement characters";
      } else if (hasLowMeaningful) {
        garbleReason = `only ${Math.round(meaningfulRatio * 100)}% recognizable characters`;
      } else if (hasRepetition && total < 30) {
        garbleReason = "repetitive character pattern (possible keyboard mashing)";
      } else if (suspiciousMixedScript) {
        garbleReason = "unnatural mix of unrelated writing systems";
      }

      if (garbleReason) {
        const lbl = getVarLabel(col, savVarMap, docxQMap);
        flagColTracked(id, col, "DATA_QUALITY", txt.slice(0, 60),
          `${col}: open text appears garbled — ${garbleReason}`,
          `Variable: ${lbl || col}\nText recorded: "${txt.slice(0, 150)}"\n${garbleReason.charAt(0).toUpperCase() + garbleReason.slice(1)}.\n${meaningfulRatio < 1 ? `Meaningful character ratio: ${Math.round(meaningfulRatio * 100)}%\n` : ""}This may indicate encoding corruption, interviewer typing random characters, or data import issues.`
        );
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

    const isScale = !!scaleRanges[varName];
    const scaleArr = scaleRanges[varName];

    columnSummary.push({
      varName,
      label,
      section: dq?.section ?? sv?.label ?? "",
      type: sv?.type ?? "unknown",
      validCodes: effectiveValidCodes[varName] ?? validCodes ?? [],
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
      isScale,
      scaleRange: isScale && scaleArr ? [scaleArr[0], scaleArr[scaleArr.length - 1]] : undefined,
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
  isScale: boolean;
  scaleRange?: [number, number];
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
function ColumnPanel({ col, colIssues, onClose }: { col: ColumnSummary; colIssues: Issue[]; onClose: () => void }) {
  return (
    <div style={{ position: "fixed", top: 0, right: 0, width: 460, height: "100vh", background: "#fff", boxShadow: "-4px 0 20px rgba(0,0,0,.12)", overflowY: "auto", zIndex: 100, padding: "20px 18px" }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 16 }}>
        <div>
          <div style={{ fontFamily: "monospace", fontSize: 18, fontWeight: 800, color: "#1e293b" }}>{col.varName}</div>
          <div style={{ fontSize: 12, color: "#64748b", marginTop: 2, maxWidth: 380 }}>{col.label || "—"}</div>
        </div>
        <button onClick={onClose} style={{ background: "none", border: "none", fontSize: 18, cursor: "pointer", color: "#94a3b8" }}>✕</button>
      </div>

      {/* Source badges */}
      <div style={{ display: "flex", gap: 6, marginBottom: 14, flexWrap: "wrap" }}>
        {col.hasInSav && <span style={{ background: "#fef9c3", border: "1px solid #fde047", borderRadius: 4, fontSize: 10, padding: "2px 7px", color: "#854d0e" }}>🗃 SAV-defined</span>}
        {col.hasInDocx && <span style={{ background: "#f0fdf4", border: "1px solid #86efac", borderRadius: 4, fontSize: 10, padding: "2px 7px", color: "#166534" }}>📝 Docx-defined</span>}
        <span style={{ background: "#f1f5f9", border: "1px solid #e2e8f0", borderRadius: 4, fontSize: 10, padding: "2px 7px", color: "#475569" }}>{col.type}</span>
        {col.isScale && col.scaleRange && <span style={{ background: "#fdf4ff", border: "1px solid #e9d5ff", borderRadius: 4, fontSize: 10, padding: "2px 7px", color: "#7c3aed" }}>📏 Scale {col.scaleRange[0]}–{col.scaleRange[1]}</span>}
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

      {/* Issues for this column — with "Why?" explanations inline */}
      {colIssues.length > 0 && (
        <div style={{ marginBottom: 14 }}>
          <div style={{ fontSize: 11, fontWeight: 700, color: "#ef4444", marginBottom: 6 }}>
            ISSUES IN THIS COLUMN ({colIssues.length})
          </div>
          {colIssues.slice(0, 50).map((iss, i) => (
            <div key={i} style={{ background: ISSUE_TYPES[iss.type].bg, border: `1px solid ${ISSUE_TYPES[iss.type].color}40`, borderRadius: 6, padding: "8px 10px", marginBottom: 6, borderLeft: `3px solid ${ISSUE_TYPES[iss.type].color}` }}>
              <div style={{ display: "flex", gap: 6, alignItems: "center", marginBottom: 4 }}>
                <Badge type={iss.type} />
                <span style={{ fontSize: 11, fontWeight: 700, color: "#475569" }}>ID: {String(iss.id)}</span>
                <span style={{ fontSize: 11, color: "#dc2626", fontFamily: "monospace" }}>
                  {String(iss.value ?? "—").slice(0, 30)}
                </span>
              </div>
              <div style={{ fontSize: 11, color: "#374151", marginBottom: 4 }}>{iss.detail}</div>
              <ExplanationBox explanation={iss.explanation} />
            </div>
          ))}
          {colIssues.length > 50 && (
            <div style={{ fontSize: 10, color: "#94a3b8", textAlign: "center", padding: "4px 0" }}>
              Showing first 50 of {colIssues.length} issues
            </div>
          )}
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
              💡 <strong>SAV is the primary authority</strong> — it defines valid codes, variable types, and value labels.<br />
              Docx adds routing/skip logic. CSV/Excel is the data to validate. All rules are derived from your files.
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
                              <div style={{ display: "flex", gap: 3, flexWrap: "wrap" }}>
                                {col.hasInSav && <span style={{ fontSize: 9, background: "#fef9c3", border: "1px solid #fde047", borderRadius: 3, padding: "1px 4px", color: "#854d0e" }}>SAV</span>}
                                {col.hasInDocx && <span style={{ fontSize: 9, background: "#f0fdf4", border: "1px solid #86efac", borderRadius: 3, padding: "1px 4px", color: "#166534" }}>DOCX</span>}
                                {col.isScale && col.scaleRange && <span style={{ fontSize: 9, background: "#fdf4ff", border: "1px solid #e9d5ff", borderRadius: 3, padding: "1px 4px", color: "#7c3aed" }}>{col.scaleRange[0]}–{col.scaleRange[1]}</span>}
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
      {selectedCol && (
        <ColumnPanel
          col={selectedCol}
          colIssues={issues.filter(iss => iss.variable === selectedCol.varName)}
          onClose={() => setSelectedCol(null)}
        />
      )}
    </div>
  );
}
