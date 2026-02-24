import { useState, useCallback, useRef } from "react";
import type { ReactNode } from "react";
import * as XLSX from "xlsx";
import Papa from "papaparse";
import mammoth from "mammoth";

// ── Issue type definitions ─────────────────────────────────────────────────────
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

// ── Routing rule parsed from docx ──────────────────────────────────────────────
interface RoutingRule {
  condition: string;   // e.g. "B1=9-10"
  targets: string[];   // e.g. ["B2"]
  rawText: string;     // original line from docx
}

// ── Parsed SAV value labels ────────────────────────────────────────────────────
// Map: varName → { code: label }
type SavValueLabels = Record<string, Record<number, string>>;

// ── Armenian question labels from docx ────────────────────────────────────────
// Map: varName → Armenian question text
type ArmenianLabels = Record<string, string>;

// ── Survey-specific valid codes ────────────────────────────────────────────────
// Based on: ACBA Bank Employer Branding questionnaire (imr_Acba EB_Q-re_vF.docx)
// Dataset:  New_link_till_22.02.sav (61 respondents, 232 variables)

const VALID_CODES: Record<string, number[]> = {
  S0:   [1, 2],
  S2:   [1, 2],
  S31:  [1, 2, 3, 4, 5, 6, 7, 99],
  S14:  [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 99],
  S4:   [1, 2],
  S5:   [1, 2, 3, 4],
  S6:   [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 666666, 99],
  S9:   [1, 2, 3, 4, 5, 99],
  S8:   [1, 2],
  D3:   [1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,97,666666,99],
  D31:  [1, 2, 3, 4, 5, 7, 8, 9, 10, 11, 12, 13, 14, 15],
  S12:  [1, 2, 99],
  S7:   [1, 2, 3, 4, 5, 6],
  B1:   [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 99],
  N1:   [1, 2, 3, 99],
  D1:   [1, 2, 3, 4, 5, 6, 99],
  D4:   [1, 2, 3, 4, 99],
  D5:   [1, 2, 3, 4, 5, 6, 99],
  D6:   [1,2,3,4,5,6,7,8,10,11,12,13,14,15,16,17,18,98,99],
};

// Human-readable variable descriptions (shown in SAV vars tab + explanations)
const VAR_DESCRIPTIONS: Record<string, string> = {
  S0: "Consent to participate (1=Yes, 2=No)",
  S2: "Gender observed (1=Male, 2=Female)",
  S3: "Age in years (numeric open)",
  S31: "Age group: 1=<18 TERM, 2=18-20, 3=21-25, 4=26-35, 5=36-40, 6=41-55, 7=56+ TERM, 99=Refusal TERM",
  S14: "Region of Armenia: 1=Yerevan, 2-11=marzes, 99=Refusal",
  S4: "Currently a student? (1=Yes, 2=No)",
  S5: "Year of study (1=1-2yr, 2=3-4yr, 3=Masters, 4=Doctoral) — asked if S4=1",
  S6: "University (1-11=named, 12=Other specify, 666666=Other open) — asked if S4=1",
  S6_other: "University other specification (open text) — required if S6=12 or 666666",
  S9: "Employment status (1=Employed, 2=Self-employed, 3=Homemaker, 4=Unemployed, 5=Pensioner, 99=Refusal)",
  S8: "Worked in past 1 year? (1=Yes, 2=No) — asked if S9=2,3,4,5",
  S10: "Years at current company (numeric) — asked if S9=1",
  D3: "Work sector (1=Finance/Banking, …, 97=Other, 666666=Other open, 99=Refusal) — asked if S9=1",
  D31: "Employer bank: 1=Acba, 2=Ararat, 3=Ameria, 4=ID, 5=Ardshin, 7=Evoca, 8=Ineco, 9=Converse, 10=Hayeconom, 11=AMIO, 12=Uni, 13=VTB, 14=Other, 15=Refusal — asked if D3=1",
  S12: "Currently seeking new job? (1=Yes, 2=No, 99=Refusal)",
  S7: "Segmentation dummy: 1=Competitor bank empl, 2=ACBA empl, 3=Other sector, 4=Working student, 5=Job seeker, 6=Out-of-segment (TERMINATE)",
  B1: "eNPS: Likelihood to recommend employer (0-10, 99=DK/Refusal) — asked if S7=1,2,3,4",
  B2: "Why recommend? (open text) — REQUIRED if B1=9 or 10",
  B3: "What should improve? (open text) — REQUIRED if B1=0-8",
  N2_1: "Benefit: Day-offs (0/1)",
  N2_2: "Benefit: Social package (0/1)",
  N2_3: "Benefit: Health insurance (0/1)",
  N2_4: "Benefit: Gym membership (0/1)",
  N2_5: "Benefit: Lunch breaks (0/1)",
  N2_6: "Benefit: Coffee/tea (0/1)",
  N2_7: "Benefit: Discounts (0/1)",
  N2_98: "Benefit: None / DK (0/1)",
  B4_1: "B4 loyalty statement 1: Proud to be part of company (1-5) — ONLY S7=1",
  B4_2: "B4 statement 2: Would recommend as best place to work (1-5) — ONLY S7=1",
  B4_3: "B4 statement 3: Company promotes diversity/inclusion (1-5) — ONLY S7=1",
  A4TOM: "Most attractive sector TOM - first mention (spontaneous, 1-24, 97=Other, 99=Refusal)",
  A101: "Why consider ACBA Bank as employer? (open) — required if A9_1=4-5 AND D31≠1",
  A102: "Why consider Ameriabank as employer? (open) — required if A9_2=4-5 AND D31≠3",
  A103: "Why consider Inekobank as employer? (open) — required if A9_8=4-5 AND D31≠8",
  A104: "Why consider Ardshinbank as employer? (open) — required if A9_15=4-5 AND D31≠5",
  D4: "Job level (1=Unskilled, 2=Skilled, 3=Middle mgmt, 4=Senior mgmt, 99=Refusal) — ONLY S7=1",
  D5: "Salary range (1=<100K, 2=101-150K, 3=151-300K, 4=301-500K, 5=501-800K, 6=801K+, 99=Refusal) — ONLY S7=1",
  D6: "Main bank (1=Acba, 2=Ararat, 3=Ameria, 4=ID, 5=Ardshin, 6=Armswiss, 7=Artsakh, 8=Byblos, 10=Evoca, 11=Ineco, 12=Converse, 13=Hayeconom, 14=AMIO, 15=Mellat, 16=Unibank, 17=VTB, 18=Fast, 98=None, 99=Refusal)",
  N1: "Preferred work mode (1=Office, 2=Remote, 3=Hybrid, 99=Refusal)",
};

// ── Rule engine (survey-specific) ─────────────────────────────────────────────
type RowData = Record<string, unknown>;

function runSurveyRules(
  data: RowData[],
  savValueLabels: SavValueLabels,
  docxRoutingRules: RoutingRule[],
  armenianLabels: ArmenianLabels,
): Issue[] {
  const issues: Issue[] = [];

  // Merge hardcoded valid codes with SAV value labels
  const effectiveValidCodes: Record<string, number[]> = { ...VALID_CODES };
  for (const [varName, labelMap] of Object.entries(savValueLabels)) {
    const codes = Object.keys(labelMap).map(Number).filter(n => !isNaN(n));
    if (codes.length > 0) {
      // SAV enriches (union) — hardcoded takes precedence only if SAV is empty
      if (!effectiveValidCodes[varName]) {
        effectiveValidCodes[varName] = codes;
      } else {
        // Merge: add any SAV codes not already in hardcoded list
        const hardSet = new Set(effectiveValidCodes[varName]);
        codes.forEach(c => hardSet.add(c));
        effectiveValidCodes[varName] = [...hardSet].sort((a, b) => a - b);
      }
    }
  }

  // Helper: get Armenian label + docx routing rule description for a variable
  const armLabel = (col: string): string => armenianLabels[col] ?? "";
  const armPrefix = (col: string): string => {
    const arm = armLabel(col);
    return arm ? `\n📌 Armenian question text: «${arm}»` : "";
  };

  // Helper: describe SAV value label for a code
  const codeLabel = (varName: string, code: number | null): string => {
    if (code === null) return "";
    const labels = savValueLabels[varName];
    if (!labels) return "";
    return labels[code] ? ` (${labels[code]})` : "";
  };

  // Helper: find docx routing rule for a condition like "S4=1"
  const findRoutingRule = (conditionSubstr: string): string => {
    const match = docxRoutingRules.find(r =>
      r.condition.toLowerCase().includes(conditionSubstr.toLowerCase())
    );
    return match ? `\n📋 Questionnaire routing: "${match.rawText}"` : "";
  };

  function flag(id: unknown, variable: string, type: IssueType, value: unknown, detail: string, explanation: string) {
    issues.push({ id: id as string | number, variable, type, value, detail, explanation });
  }

  data.forEach(row => {
    const id = row.id ?? row.ID ?? row.RespondentID ?? "?";

    // Helper: get numeric value or null if empty/missing
    const v = (col: string): number | null => {
      const raw = row[col];
      if (raw === null || raw === undefined || raw === "" || raw === ".") return null;
      const s = String(raw).trim();
      if (s === "") return null;
      const n = Number(s);
      return isNaN(n) ? null : n;
    };
    // Helper: get string value or null
    const vStr = (col: string): string | null => {
      const raw = row[col];
      if (raw === null || raw === undefined) return null;
      const s = String(raw).trim();
      return s === "" ? null : s;
    };
    const has = (col: string) => v(col) !== null || (vStr(col) !== null && !["", "."].includes(vStr(col) ?? ""));
    const empty = (col: string) => !has(col);
    const eq = (col: string, val: number) => v(col) === val;
    const inList = (col: string, list: number[]) => { const x = v(col); return x !== null && list.includes(x); };

    // ── Screening ─────────────────────────────────────────────────────────────

    // S4=1 (student) → S5, S6 required
    if (eq("S4", 1)) {
      if (empty("S5"))
        flag(id, "S5", "MISSING_DATA", null, "S5 (year of study) missing for student (S4=1)",
          `S4=1 means respondent is currently a student. S5 (which year/course of study) must be answered. The survey routing requires S5 immediately after S4=1.${findRoutingRule("S4=1")}${armPrefix("S5")}`);
      if (empty("S6"))
        flag(id, "S6", "MISSING_DATA", null, "S6 (university name) missing for student (S4=1)",
          `S4=1 means respondent is a student. S6 (name of educational institution) must be answered. Without S6, we cannot segment by institution type.${findRoutingRule("S4=1")}${armPrefix("S6")}`);
    }

    // S4≠1 → S5, S6 should NOT be filled
    if (!eq("S4", 1) && has("S4")) {
      if (has("S5") && !eq("S5", 99))
        flag(id, "S5", "SKIP_VIOLATION", v("S5"), `S5 filled but S4=${v("S4")} (not a student)`,
          `S4=${v("S4")}${codeLabel("S4", v("S4"))} — respondent is NOT a student. S5 (year of study) should only be asked when S4=1. This response should be empty — the CAPI routing should have skipped S5.${findRoutingRule("S4")}${armPrefix("S5")}`);
      if (has("S6") && !eq("S6", 99))
        flag(id, "S6", "SKIP_VIOLATION", v("S6"), `S6 filled but S4=${v("S4")} (not a student)`,
          `S4=${v("S4")}${codeLabel("S4", v("S4"))} — respondent is not a student, so S6 (university) should be empty. Routing skipped incorrectly.${armPrefix("S6")}`);
    }

    // S6 = "Other" (12 or 666666) → S6_other must be filled
    if (eq("S4", 1) && (eq("S6", 12) || eq("S6", 666666))) {
      const txt = vStr("S6_other");
      if (!txt || txt.length < 2)
        flag(id, "S6_other", "MISSING_DATA", txt,
          "S6_other empty — respondent chose 'Other' university but didn't specify",
          `S6=12 or S6=666666 means the respondent's university is not in the standard list and they selected 'Other'. The open text field S6_other (where they specify the institution) is empty. This is a confirmed CAPI programming bug documented in the data management log (Question S6): when 'Other' is selected, the follow-up text field is not being saved.${armPrefix("S6_other")}`);
    }

    // S9=1 (employed) → S10, D3 required
    if (eq("S9", 1)) {
      if (empty("S10"))
        flag(id, "S10", "MISSING_DATA", null, "S10 (years at company) missing for employed respondent (S9=1)",
          `S9=1${codeLabel("S9", 1)} means currently employed under contract. S10 (how many years at current company) is a mandatory follow-up for employed respondents and is absent.${findRoutingRule("S9=1")}${armPrefix("S10")}`);
      if (empty("D3"))
        flag(id, "D3", "MISSING_DATA", null, "D3 (work sector) missing for employed respondent (S9=1)",
          `S9=1 means employed. D3 (which sector do you work in?) is required for all currently employed respondents. This determines the segmentation variable S7.${findRoutingRule("S9=1")}${armPrefix("D3")}`);
    }

    // S9≠1 → S10, D3, D31 should be empty
    if (has("S9") && !eq("S9", 1) && !eq("S9", 99)) {
      if (has("S10"))
        flag(id, "S10", "SKIP_VIOLATION", v("S10"), `S10 filled but S9=${v("S9")} (not employed under contract)`,
          `S9=${v("S9")}${codeLabel("S9", v("S9"))} — respondent is not currently employed under contract. S10 (years at company) is only asked when S9=1. This field should be empty.${armPrefix("S10")}`);
      if (has("D3") && !eq("D3", 99))
        flag(id, "D3", "SKIP_VIOLATION", v("D3"), `D3 filled but S9=${v("S9")} (not employed)`,
          `S9=${v("S9")}${codeLabel("S9", v("S9"))} — respondent is not employed. D3 (work sector) should only be asked for S9=1.${armPrefix("D3")}`);
    }

    // D3=1 (finance sector) → D31 required
    if (eq("D3", 1)) {
      if (empty("D31"))
        flag(id, "D31", "MISSING_DATA", null, "D31 (employer bank) missing for finance sector employee (D3=1)",
          `D3=1${codeLabel("D3", 1)} means respondent works in the Finance/Banking sector. D31 (name of employer bank) is required — it is the key variable for determining S7 segment (ACBA employee vs. competitor bank employee vs. other finance).${findRoutingRule("D3=1")}${armPrefix("D31")}`);
    }

    // D3≠1 → D31 should be empty
    if (has("D3") && !eq("D3", 1) && !eq("D3", 99) && !eq("D3", 97) && !eq("D3", 666666)) {
      if (has("D31"))
        flag(id, "D31", "SKIP_VIOLATION", v("D31"), `D31 filled but D3=${v("D3")} (not finance sector)`,
          `D3=${v("D3")}${codeLabel("D3", v("D3"))} — respondent does not work in Finance/Banking. D31 (employer bank name) is only asked when D3=1. This field should be empty.${armPrefix("D31")}`);
    }

    // S9=2,3,4,5 → S8 required
    if (inList("S9", [2, 3, 4, 5])) {
      if (empty("S8"))
        flag(id, "S8", "MISSING_DATA", null, `S8 (worked in past year?) missing — S9=${v("S9")}`,
          `S9=${v("S9")}${codeLabel("S9", v("S9"))} means respondent is not currently employed under contract. S8 (did you work at any company in the past 1 year?) is required as a follow-up for all non-employed respondents.${findRoutingRule("S9=")}${armPrefix("S8")}`);
    }

    // S9=1 → S8 should NOT be filled (employed, no need to ask about past year)
    if (eq("S9", 1) && has("S8"))
      flag(id, "S8", "SKIP_VIOLATION", v("S8"), "S8 filled but S9=1 (currently employed — S8 should be skipped)",
        `S8 (worked in past 1 year?) is only asked if S9=2,3,4,5 (not currently employed). S9=1 means currently employed — the routing should have skipped S8 entirely.${armPrefix("S8")}`);

    // S7 valid range
    if (has("S7") && !inList("S7", [1, 2, 3, 4, 5, 6]))
      flag(id, "S7", "MISMATCHED_CODE", v("S7"), `S7=${v("S7")} is not a valid segment code (1-6)`,
        `S7 is the computed segmentation variable. Valid values: 1=Competitor bank employees, 2=ACBA employees, 3=Other sector employees, 4=Working students, 5=Job seekers, 6=Out-of-segment (TERMINATE).${armPrefix("S7")}`);

    // ── B-section (Loyalty / NPS) ─────────────────────────────────────────────
    const s7 = v("S7");
    const b1 = v("B1");
    const inB1Scope = s7 !== null && [1, 2, 3, 4].includes(s7);

    // B1 out of range
    if (has("B1") && !inList("B1", [0,1,2,3,4,5,6,7,8,9,10,99]))
      flag(id, "B1", "OUT_OF_RANGE", b1, `B1=${b1} outside valid range 0-10 (99=DK/Refusal)`,
        `B1 is the eNPS question: 'How likely are you to recommend your current employer to friends or family?' Valid range: 0 (not at all likely) to 10 (extremely likely), plus 99 for DK/Refusal.${armPrefix("B1")}`);

    // B1 missing for in-scope respondents
    if (inB1Scope && empty("B1"))
      flag(id, "B1", "MISSING_DATA", null, `B1 (eNPS) missing for S7=${s7} respondent`,
        `S7=${s7}${codeLabel("S7", s7)} places this respondent in a segment that should answer B1 (segments 1-4 all receive the eNPS question). The answer is absent.${findRoutingRule("S7=")}${armPrefix("B1")}`);

    // B1=9-10 → B2 required (promoter open text)
    if (inList("B1", [9, 10])) {
      const b2 = vStr("B2");
      if (!b2 || b2.length < 2)
        flag(id, "B2", "MISSING_DATA", b2, `B2 (why recommend) empty — B1=${b1} (Promoter score)`,
          `B1=${b1} means this respondent is a PROMOTER (NPS score 9-10). The follow-up question B2 ('For what specific advantage are you ready to recommend your workplace?') is mandatory for all promoters and must contain a meaningful open-text response. This exact issue was flagged in the data management log for 10 respondent IDs (453, 526, 638, 351, 687, 682, 166, 637, 659, 249).${findRoutingRule("B1=9")}${armPrefix("B2")}`);
    }

    // B1=0-8 → B3 required (improvement open text)
    if (b1 !== null && b1 >= 0 && b1 <= 8) {
      const b3 = vStr("B3");
      if (!b3 || b3.length < 2)
        flag(id, "B3", "MISSING_DATA", b3, `B3 (what to improve) empty — B1=${b1} (Detractor/Passive)`,
          `B1=${b1} means this respondent is a DETRACTOR (0-6) or PASSIVE (7-8). The follow-up question B3 ('What should the employer improve?') is mandatory for all non-promoters.${findRoutingRule("B1=0")}${armPrefix("B3")}`);
    }

    // B2 filled when B1 is NOT 9-10 → skip violation
    const b2txt = vStr("B2");
    if (b2txt && b2txt.length > 1 && (b1 === null || b1 < 9 || b1 === 99)) {
      flag(id, "B2", "SKIP_VIOLATION", b2txt.slice(0, 50), `B2 filled but B1=${b1} (should only be asked if B1=9-10)`,
        `B2 (why recommend?) must only be asked when B1=9 or 10 (Promoter). B1=${b1} here, so B2 should be empty. This indicates a routing error or data entry in the wrong field.${armPrefix("B2")}`);
    }

    // B3 filled when B1 is NOT 0-8 → skip violation
    const b3txt = vStr("B3");
    if (b3txt && b3txt.length > 1 && b1 !== null && b1 >= 9 && b1 !== 99) {
      flag(id, "B3", "SKIP_VIOLATION", b3txt.slice(0, 50), `B3 filled but B1=${b1} (should only be asked if B1=0-8)`,
        `B3 (what to improve?) must only be asked when B1=0-8. B1=${b1} here (Promoter), so B3 should be empty.${armPrefix("B3")}`);
    }

    // Open text quality checks (garbled text detection)
    const checkOpenTextQuality = (col: string, context: string) => {
      const txt = vStr(col);
      if (!txt || txt.length < 3) return;
      const meaningful = (txt.match(/[\u0531-\u058Fa-zA-Z]{2,}/g) || []).join("").length;
      const total = txt.replace(/\s/g, "").length;
      if (total > 4 && meaningful / total < 0.25)
        flag(id, col, "DATA_QUALITY", txt.slice(0, 60),
          `${col} open text appears to be garbled/random characters`,
          `The ${col} field (${context}) contains text where less than 25% of characters form meaningful words in Armenian or Latin script. This pattern is consistent with interviewers typing random characters to bypass a mandatory open-text field. This was explicitly documented in the data management log for question B3: "Most of the answers are just garbled/meaningless characters written in."${armPrefix(col)}`);
    };
    checkOpenTextQuality("B2", "why recommend employer?");
    checkOpenTextQuality("B3", "what should employer improve?");
    checkOpenTextQuality("A101", "why consider ACBA Bank?");
    checkOpenTextQuality("A102", "why consider Ameriabank?");
    checkOpenTextQuality("A103", "why consider Inekobank?");
    checkOpenTextQuality("A104", "why consider Ardshinbank?");

    // B4 battery: ONLY for S7=1 (competitor bank employees)
    for (let i = 1; i <= 12; i++) {
      const col = `B4_${i}`;
      if (has(col) && s7 !== 1)
        flag(id, col, "SKIP_VIOLATION", v(col), `${col} filled but S7=${s7} — B4 only for S7=1`,
          `The B4 battery (12 loyalty/engagement statements about employer relationship) is ONLY asked to respondents in Segment S7=1 (competitor bank employees). This respondent is S7=${s7} and should have B4 completely empty.${armPrefix(col)}`);
      if (eq("S7", 1) && has(col) && !inList(col, [1,2,3,4,5,99]))
        flag(id, col, "OUT_OF_RANGE", v(col), `${col}=${v(col)} outside valid range 1-5 (99=DK)`,
          `B4 loyalty statements use a 1-5 agreement scale (1=Completely disagree, 5=Completely agree, 99=DK/Refusal).${armPrefix(col)}`);
    }

    // ── A1 Importance / A2 Satisfaction ratings ───────────────────────────────

    for (let i = 1; i <= 16; i++) {
      const col = `A1_${i}`;
      if (has(col) && !inList(col, [1,2,3,4,5,99]))
        flag(id, col, "OUT_OF_RANGE", v(col), `${col}=${v(col)} outside valid range 1-5 (99=DK)`,
          `A1 rates importance of employer selection factors on a 1-5 scale (1=Not important at all, 5=Extremely important, 99=DK/Refusal). All 16 items asked to all in-scope respondents.${armPrefix(col)}`);
    }

    // A2: only for S7=1 or S7=2
    for (let i = 1; i <= 15; i++) {
      const col = `A2_${i}`;
      if (has(col) && !inList("S7", [1, 2]))
        flag(id, col, "SKIP_VIOLATION", v(col), `${col} filled but S7=${s7} — A2 only for S7=1 or 2`,
          `A2 (satisfaction with current employer factors) is only shown to S7=1 (competitor bank employees) and S7=2 (ACBA employees). This respondent is S7=${s7} and should have A2 empty.${armPrefix(col)}`);
      if (inList("S7", [1, 2]) && has(col) && !inList(col, [1,2,3,4,5,6,99]))
        flag(id, col, "OUT_OF_RANGE", v(col), `${col}=${v(col)} outside valid range 1-5 (6=DK, 99=Refusal)`,
          `A2 satisfaction items use 1-5 scale. Code 6 is used as DK/Not applicable in this dataset (confirmed from value labels).${armPrefix(col)}`);
    }

    // ── A3 Productivity detractors (binary 0/1) ───────────────────────────────
    for (let i = 1; i <= 11; i++) {
      const col = `A3_${i}`;
      if (has(col) && !inList(col, [0, 1]))
        flag(id, col, "MISMATCHED_CODE", v(col), `${col}=${v(col)} — binary variable should be 0 or 1`,
          `A3 is a multiple-response question stored as binary dummies (0=not selected, 1=selected). Values other than 0 or 1 indicate a data entry error.${armPrefix(col)}`);
    }
    if (has("A3_98") && !inList("A3_98", [0, 1]))
      flag(id, "A3_98", "MISMATCHED_CODE", v("A3_98"), "A3_98 (DK/Refusal) should be 0 or 1",
        "A3_98 is the DK/Refusal flag for A3. It should be 0 (not selected) or 1 (selected).");

    // ── A4 Sector attractiveness (TOM + checkboxes) ───────────────────────────
    const a4tom = v("A4TOM");

    // A4TOM=99 (refusal) → no A4_X checkbox should be ticked
    if (eq("A4TOM", 99)) {
      for (let i = 1; i <= 24; i++) {
        const col = `A4_${i}`;
        if (eq(col, 1))
          flag(id, col, "SKIP_VIOLATION", 1, `${col}=1 but A4TOM=99 (refusal on TOM question)`,
            `A4TOM=99 means the respondent refused to name any attractive sector spontaneously. Therefore NONE of the A4_X sector checkboxes should be selected. This is the exact skip logic violation documented in the data management log: 'ID 441 — refusal is marked in TOM, but a sector is selected in A4_1. If there is a refusal in TOM, the spontaneous question should not be asked.'${findRoutingRule("A4TOM")}`);
      }
    }

    // A4TOM filled with valid sector → that sector should NOT also appear in A4_X (double-count)
    if (a4tom !== null && a4tom !== 99 && a4tom !== 666666 && a4tom >= 1 && a4tom <= 24) {
      if (eq(`A4_${a4tom}`, 1))
        flag(id, `A4_${a4tom}`, "DATA_QUALITY", 1,
          `A4TOM=${a4tom} and A4_${a4tom}=1 — sector ${a4tom} counted as both TOM and secondary mention`,
          `A4TOM=${a4tom} means sector ${a4tom} was the respondent's first/spontaneous mention. A4_${a4tom}=1 means it was also ticked in the secondary checkbox grid. The same sector is being counted twice. TOM should be stored only in A4TOM; the secondary checkboxes (A4_1 to A4_24) should capture different sectors (up to 2 more).`);
    }

    // Max 3 sectors: 1 TOM + max 2 checkboxes
    if (a4tom !== null && a4tom !== 99) {
      const nChecked = Array.from({length: 24}, (_, i) => `A4_${i+1}`).filter(c => eq(c, 1)).length;
      if (nChecked > 2)
        flag(id, "A4", "OUT_OF_RANGE", nChecked,
          `${nChecked} sectors checked in A4 grid + 1 TOM = ${nChecked + 1} total (max is 3)`,
          "The A4 question allows a MAXIMUM of 3 sector selections total: 1 captured in A4TOM (first spontaneous mention) and up to 2 more in the A4_X checkbox grid. This respondent has exceeded the limit.");
    }

    // ── A5 Sector attractiveness ratings ─────────────────────────────────────
    for (let i = 1; i <= 4; i++) {
      const col = `A5_${i}`;
      if (has(col) && !inList(col, [1,2,3,4,5,99]))
        flag(id, col, "OUT_OF_RANGE", v(col), `${col}=${v(col)} outside valid range 1-5 (99=DK)`,
          `A5 rates attractiveness of 4 sectors (Finance/Banking, IT, Manufacturing, Services) on a 1-5 scale.${armPrefix(col)}`);
    }

    // ── A8/A9 Brand awareness + employer consideration ────────────────────────
    const brandCodes = ["1","2","4","5","7","8","10","12","14","15","16","17","18","19"];

    brandCodes.forEach(brand => {
      const a8 = `A8_${brand}`;
      const a9 = `A9_${brand}`;

      // A8 should be binary (0/1) if filled
      if (has(a8) && !inList(a8, [0, 1]))
        flag(id, a8, "MISMATCHED_CODE", v(a8), `${a8}=${v(a8)} — should be 0 (unaware) or 1 (aware)`,
          "A8 brand awareness is stored as binary dummies: 1=Yes, respondent knows this company; 0=No. Other values indicate a data entry error.");

      // A8=0 or unaware → A9 should NOT be filled
      if (eq(a8, 0) && has(a9))
        flag(id, a9, "SKIP_VIOLATION", v(a9), `${a9} filled but ${a8}=0 (not aware of brand ${brand})`,
          `${a8}=0 means respondent does NOT know company ${brand}. A9 (employer consideration rating) should only be asked for companies the respondent is aware of (A8=1). The routing should have skipped ${a9}.${findRoutingRule("A8")}`);

      // A8=1 (aware) → A9 required
      if (eq(a8, 1) && empty(a9))
        flag(id, a9, "MISSING_DATA", null, `${a9} missing but ${a8}=1 (aware of brand ${brand})`,
          `${a8}=1 means respondent knows company ${brand} by name. The follow-up ${a9} (on a 1-5 scale, would you consider this company as your employer?) is mandatory for all brands the respondent is aware of.`);

      // A9 range check
      if (has(a9) && !inList(a9, [1,2,3,4,5,99]))
        flag(id, a9, "OUT_OF_RANGE", v(a9), `${a9}=${v(a9)} outside valid range 1-5 (99=DK)`,
          "A9 rates employer consideration on a 1-5 scale (1=Would not consider at all, 5=Would definitely consider, 99=DK/Refusal).");
    });

    // ACBA Bank (A8_1) should NOT be filled for ACBA employees (D3=1 AND D31=1)
    if (eq("D3", 1) && eq("D31", 1) && has("A8_1"))
      flag(id, "A8_1", "SKIP_VIOLATION", v("A8_1"), "A8_1 (ACBA awareness) filled for ACBA employee (D31=1)",
        "Per the questionnaire instructions, ACBA Bank is SKIPPED in the A8 awareness question for respondents who are ACBA employees (D3=1 AND D31=1), since they obviously know their own employer. A8_1 should be empty for ACBA employees.");

    // ── A10/A101-A104 Why consider bank ──────────────────────────────────────
    const bankChecks = [
      { a9col: "A9_1",  excl: 1,  openCol: "A101", bankName: "ACBA Bank" },
      { a9col: "A9_2",  excl: 3,  openCol: "A102", bankName: "Ameriabank" },
      { a9col: "A9_8",  excl: 8,  openCol: "A103", bankName: "Inekobank" },
      { a9col: "A9_15", excl: 5,  openCol: "A104", bankName: "Ardshinbank" },
    ];
    bankChecks.forEach(({ a9col, excl, openCol, bankName }) => {
      const a9val = v(a9col);
      const d31val = v("D31");
      // Required if A9=4-5 AND D31≠excl (not an employee of that bank)
      if (a9val !== null && [4, 5].includes(a9val) && d31val !== excl) {
        const txt = vStr(openCol);
        if (!txt || txt.length < 2)
          flag(id, openCol, "MISSING_DATA", txt,
            `${openCol} (why consider ${bankName}?) missing — ${a9col}=${a9val}, D31=${d31val}`,
            `${a9col}=${a9val} means respondent rates ${bankName} as 4-5 ('would consider' or 'would definitely consider' as employer), and D31=${d31val}≠${excl} (not a current ${bankName} employee). Per questionnaire logic, the open-text question '${openCol}: Why would you consider ${bankName} as your future employer?' is mandatory in this case. ${openCol === "A101" ? "This specific issue was confirmed in the data management log: ID 453 was missing A101." : ""}${findRoutingRule(a9col)}${armPrefix(openCol)}`);
      }
      // A10x filled when not applicable
      if (has(openCol) && (a9val === null || ![4, 5].includes(a9val)))
        flag(id, openCol, "SKIP_VIOLATION", vStr(openCol)?.slice(0, 50),
          `${openCol} filled but ${a9col}=${a9val} (should only be asked if ${a9col}=4-5)`,
          `${openCol} should only be asked when ${a9col}=4 or 5 (high employer consideration of ${bankName}). ${a9col}=${a9val} here, so ${openCol} should be empty.${armPrefix(openCol)}`);
    });

    // ── D4, D5: ONLY for S7=1 ────────────────────────────────────────────────
    if (has("D4") && s7 !== 1)
      flag(id, "D4", "SKIP_VIOLATION", v("D4"), `D4 (job level) filled but S7=${s7} — only for S7=1`,
        `D4 (current job position level) is ONLY asked to respondents in Segment S7=1 (competitor bank employees). This allows competitive benchmarking of job levels. All other segments should have D4 empty.${armPrefix("D4")}`);
    if (has("D5") && s7 !== 1)
      flag(id, "D5", "SKIP_VIOLATION", v("D5"), `D5 (salary range) filled but S7=${s7} — only for S7=1`,
        `D5 (monthly salary range) is ONLY asked to S7=1 respondents (competitor bank employees). Salary data is deliberately excluded from ACBA employees (S7=2) and other segments for ethical/sensitivity reasons.${armPrefix("D5")}`);
    if (eq("S7", 1)) {
      if (has("D4") && !inList("D4", [1,2,3,4,99]))
        flag(id, "D4", "OUT_OF_RANGE", v("D4"), `D4=${v("D4")} outside valid codes 1-4 (99=Refusal)`,
          `D4 job level: 1=Unskilled worker, 2=Skilled specialist (non-manager), 3=Middle management, 4=Senior management, 99=Refusal.${armPrefix("D4")}`);
      if (has("D5") && !inList("D5", [1,2,3,4,5,6,99]))
        flag(id, "D5", "OUT_OF_RANGE", v("D5"), `D5=${v("D5")} outside valid codes 1-6 (99=Refusal)`,
          `D5 salary: 1=<100K AMD, 2=101-150K, 3=151-300K, 4=301-500K, 5=501-800K, 6=801K+, 99=Refusal.${armPrefix("D5")}`);
    }

    // ── D6 Main bank ─────────────────────────────────────────────────────────
    if (has("D6") && !effectiveValidCodes.D6.includes(v("D6")!))
      flag(id, "D6", "MISMATCHED_CODE", v("D6"), `D6=${v("D6")} is not a valid bank code`,
        `D6 (main bank used by respondent). Valid codes: 1=Acba, 2=Ararat, 3=Ameria, 4=IDBank, 5=Ardshin, 6=Armswiss, 7=Artsakh, 8=Byblos, 10=Evoca, 11=Ineco, 12=Converse, 13=Hayeconom, 14=AMIO, 15=Mellat, 16=Unibank, 17=VTB, 18=Fast Bank, 98=None (do not read), 99=Refusal.${armPrefix("D6")}`);

    // ── S7 internal consistency ───────────────────────────────────────────────
    if (eq("S7", 1)) {
      if (!eq("S9", 1))
        flag(id, "S7", "DATA_QUALITY", 1, `S7=1 but S9=${v("S9")} (should be 1=employed)`,
          "S7=1 (competitor bank employee) requires S9=1 (employed under contract). The S9 value is inconsistent with the assigned segment.");
      if (!eq("D3", 1))
        flag(id, "S7", "DATA_QUALITY", 1, `S7=1 but D3=${v("D3")} (should be 1=finance sector)`,
          "S7=1 (competitor bank employee) requires D3=1 (finance/banking sector). The sector value is inconsistent.");
      if (eq("D31", 1))
        flag(id, "S7", "DATA_QUALITY", 1, `S7=1 but D31=1 (ACBA Bank) — should be S7=2`,
          "S7=1 is for COMPETITOR bank employees. D31=1 means this respondent works at ACBA Bank and should be in segment S7=2, not S7=1. This indicates a segmentation computation error.");
    }
    if (eq("S7", 2) && (has("D3") || has("D31"))) {
      if (!eq("D3", 1) || !eq("D31", 1))
        flag(id, "S7", "DATA_QUALITY", 2, `S7=2 but D3=${v("D3")}, D31=${v("D31")} — expected D3=1 AND D31=1`,
          "S7=2 (ACBA employee) requires D3=1 (finance sector) AND D31=1 (ACBA Bank). The sector/employer values don't match the assigned segment.");
    }

    // ── Misc valid code checks ────────────────────────────────────────────────
    (["S14", "S9", "N1", "D1"] as const).forEach(col => {
      if (has(col) && effectiveValidCodes[col] && !effectiveValidCodes[col].includes(v(col)!))
        flag(id, col, "MISMATCHED_CODE", v(col), `${col}=${v(col)} is not in the valid code list`,
          `${VAR_DESCRIPTIONS[col] || col}${armPrefix(col)}`);
    });

    // ── A12/A13 control list: selections must come from brands where A8=1 ─────
    brandCodes.forEach(brand => {
      const a8col = `A8_${brand}`;
      const a12col = `A12_${brand}`;
      const a13col = `A13_${brand}`;
      if (eq(a12col, 1) && eq(a8col, 0))
        flag(id, a12col, "SKIP_VIOLATION", 1, `A12_${brand} selected but A8_${brand}=0 (not aware of brand ${brand})`,
          `A12 ('Which company best fits: prospective employer?') should only include companies the respondent is aware of (A8=1). Brand ${brand} was not known to this respondent but was selected in A12. The control list was not properly applied.`);
      if (eq(a13col, 1) && eq(a8col, 0))
        flag(id, a13col, "SKIP_VIOLATION", 1, `A13_${brand} selected but A8_${brand}=0 (not aware of brand ${brand})`,
          `A13 ('Which company gives young people career growth?') should only include brands where A8=1. Brand ${brand} was unknown to this respondent but appeared in A13. Control list filtering failed.`);
    });

  }); // end forEach row

  return issues;
}

// ── Dataset-level structural checks ───────────────────────────────────────────
function runDatasetChecks(data: RowData[]): DatasetWarning[] {
  const warnings: DatasetWarning[] = [];
  if (!data.length) return warnings;
  const cols = new Set(Object.keys(data[0]));

  // A3_12 missing from the database entirely
  if (!cols.has("A3_12"))
    warnings.push({
      type: "STRUCTURAL",
      variable: "A3_12",
      detail: "Column A3_12 ('Colleague relations' productivity detractor) is MISSING from the dataset",
      explanation: "The questionnaire has 12 productivity detractor options in question A3 (codes 1-12). Code 12 is 'Relationships with colleagues' (Գործընկերների հետ հարաբերությունները). However, the dataset only contains A3_1 through A3_11 — A3_12 was never created as a column. This is a confirmed structural error documented in the data management log (Question A3, row 6): 'The option 12: Relationships with colleagues is missing from the database.' Any respondent who selected this option cannot be analyzed. ACTION REQUIRED: Add A3_12 to the data structure and re-export from CAPI system."
    });

  // Check for duplicate IDs
  const ids = data.map(r => r.id ?? r.ID ?? r.RespondentID).filter(Boolean);
  const seen = new Set<unknown>(), dupes = new Set<unknown>();
  ids.forEach(id => { if (seen.has(id)) dupes.add(id); seen.add(id); });
  if (dupes.size > 0)
    warnings.push({
      type: "STRUCTURAL",
      variable: "id",
      detail: `Duplicate respondent IDs detected: ${[...dupes].join(", ")}`,
      explanation: "Each respondent should have a unique ID. Duplicate IDs indicate a data merge error, repeated import of records, or a CAPI system issue. Duplicate records will inflate all counts and percentages in analysis."
    });

  return warnings;
}

// ── SAV binary parser: variable names + value labels ──────────────────────────
// Parses SPSS .sav binary format (IEEE 754, little-endian)
// Returns { varNames, valueLabels }
function parseSavFile(buffer: ArrayBuffer): {
  varNames: string[];
  valueLabels: SavValueLabels;
} {
  const bytes = new Uint8Array(buffer);
  const decoder = new TextDecoder("latin1");

  // Read a 4-byte little-endian int32
  const readInt32 = (offset: number): number => {
    return bytes[offset] | (bytes[offset+1] << 8) | (bytes[offset+2] << 16) | (bytes[offset+3] << 24);
  };

  // Read a 8-byte little-endian float64
  const readFloat64 = (offset: number): number => {
    const dv = new DataView(buffer, offset, 8);
    return dv.getFloat64(0, true); // little-endian
  };

  // Read fixed-length string, trimming trailing spaces
  const readStr = (offset: number, len: number): string => {
    return decoder.decode(bytes.slice(offset, offset + len)).trimEnd();
  };

  // Verify SAV magic header "$FL2" or "$FL3"
  const magic = readStr(0, 4);
  if (!magic.startsWith("$FL")) {
    // Fallback: just extract variable names from binary like before
    return { varNames: parseSavVariableNamesLegacy(buffer), valueLabels: {} };
  }

  // Header record starts at offset 0
  // Prod name: bytes 4-64 (60 bytes)
  // layout_code at 64 (int32)
  // case_size at 68 (int32) — # of 8-byte slots per case
  // compression at 72 (int32)
  // weight_index at 76 (int32)
  // ncases at 80 (int32)
  // bias at 84 (float64)
  // creation_date at 92 (9 bytes)
  // creation_time at 101 (8 bytes)
  // file_label at 109 (64 bytes)
  // padding at 173 (3 bytes)
  // First record after header starts at byte 176

  const varNames: string[] = [];
  const varIndex: string[] = []; // variable at each 8-byte slot (for value labels)
  const valueLabels: SavValueLabels = {};

  let offset = 176; // start of records
  const MAX_OFFSET = buffer.byteLength - 4;

  // We track which variable index (slot) corresponds to which variable name
  // for value labels association
  let slotIndex = 0;

  while (offset < MAX_OFFSET) {
    if (offset + 4 > buffer.byteLength) break;
    const recType = readInt32(offset);
    offset += 4;

    if (recType === 999) {
      // End-of-dictionary record
      break;
    }

    if (recType === 2) {
      // Variable record
      if (offset + 28 > buffer.byteLength) break;
      const typeCode = readInt32(offset);       // 0=numeric, 1-255=string width
      // const hasVarLabel = readInt32(offset + 4);
      const hasLabel = readInt32(offset + 4);
      const nMissing = readInt32(offset + 8);   // 0, ±1, ±2, ±3
      // format: offset+12 (int32)
      // format2: offset+16 (int32) — not used here
      const rawName = readStr(offset + 20, 8).trim();
      offset += 28;

      // Variable label (if hasLabel != 0)
      if (hasLabel !== 0) {
        if (offset + 4 > buffer.byteLength) break;
        const labelLen = readInt32(offset);
        offset += 4;
        // label is padded to multiple of 4
        const paddedLen = Math.ceil(labelLen / 4) * 4;
        offset += paddedLen;
      }

      // Missing values
      const absMissing = Math.abs(nMissing);
      offset += absMissing * 8;

      // Register this variable
      if (typeCode === 0 || typeCode > 0) {
        // Only register real variables (not continuation slots for long strings)
        if (rawName !== "" && !rawName.startsWith("@")) {
          varNames.push(rawName);
        }
        varIndex[slotIndex] = rawName;
        slotIndex++;

        // Long strings occupy multiple 8-byte slots (continuation records)
        if (typeCode > 0) {
          const extraSlots = Math.ceil(typeCode / 8) - 1;
          for (let s = 0; s < extraSlots; s++) {
            varIndex[slotIndex++] = ""; // continuation, no name
          }
        }
      }
    } else if (recType === 3) {
      // Value labels record
      if (offset + 4 > buffer.byteLength) break;
      const nLabels = readInt32(offset);
      offset += 4;

      const tempLabels: Record<number, string> = {};
      for (let i = 0; i < nLabels; i++) {
        if (offset + 9 > buffer.byteLength) break;
        const code = readFloat64(offset);
        offset += 8;
        const labelLen = bytes[offset];
        offset += 1;
        const paddedLen = Math.ceil((labelLen + 1) / 8) * 8 - 1;
        const label = readStr(offset, labelLen);
        offset += paddedLen;
        tempLabels[code] = label;
      }

      // Next record MUST be record type 4 (variable index for these labels)
      if (offset + 4 > buffer.byteLength) break;
      const nextRec = readInt32(offset);
      offset += 4;
      if (nextRec === 4) {
        if (offset + 4 > buffer.byteLength) break;
        const nVars = readInt32(offset);
        offset += 4;
        for (let i = 0; i < nVars; i++) {
          if (offset + 4 > buffer.byteLength) break;
          const varSlot = readInt32(offset) - 1; // 1-based → 0-based
          offset += 4;
          const varName = varIndex[varSlot] ?? "";
          if (varName && varName !== "") {
            valueLabels[varName] = { ...(valueLabels[varName] ?? {}), ...tempLabels };
          }
        }
      }
      // (if nextRec ≠ 4 we just consumed its type bytes — skip gracefully)
    } else if (recType === 6) {
      // Document record
      if (offset + 4 > buffer.byteLength) break;
      const nLines = readInt32(offset);
      offset += 4;
      offset += nLines * 80; // each line is 80 bytes
    } else if (recType === 7) {
      // Extension (informational) record — skip
      if (offset + 8 > buffer.byteLength) break;
      const subType = readInt32(offset);       // not used
      void subType;
      const elemSize = readInt32(offset + 4);
      offset += 8;
      if (offset + 4 > buffer.byteLength) break;
      const nElems = readInt32(offset);
      offset += 4;
      offset += elemSize * nElems;
    } else {
      // Unknown record type — we can't safely continue
      break;
    }
  }

  // If parser returned no var names (malformed SAV), fall back to legacy
  if (varNames.length === 0) {
    return { varNames: parseSavVariableNamesLegacy(buffer), valueLabels: {} };
  }

  return { varNames, valueLabels };
}

// Legacy ASCII-scan fallback (original approach)
function parseSavVariableNamesLegacy(buffer: ArrayBuffer): string[] {
  const bytes = new Uint8Array(buffer);
  const SKIP_WORDS = new Set([
    "the","and","or","in","of","to","is","for","with","from","not","are","at","by",
    "this","that","be","as","it","on","if","do","so","SPSS","DATA","LIST","SAVE",
    "GET","COMPUTE","RECODE","EXECUTE","VALUE","VARIABLE","LABELS","TYPE","STRING",
    "NUMERIC","MISSING","SYSMIS","ALL","EQ","NE","LT","GT","LE","GE","SUM","MEAN",
    "BEGIN","END","FILE","NAME","FORMAT","MEASURE","ROLE","NOMINAL","SCALE","ORDINAL",
    "INPUT","OUTPUT","YES","NO","true","false","null","undefined","var","let","const",
    "function","return","class","import","export","default","new","super","extends",
  ]);
  let ascii = "";
  for (let i = 0; i < Math.min(bytes.length, 300000); i++)
    ascii += bytes[i] < 128 ? String.fromCharCode(bytes[i]) : " ";

  const seen = new Set<string>(), names: string[] = [];
  const re = /\b([A-Za-z][A-Za-z0-9_]{1,39})\b/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(ascii)) !== null) {
    const nm = m[1];
    if (!seen.has(nm) && !SKIP_WORDS.has(nm) && !SKIP_WORDS.has(nm.toLowerCase())) {
      seen.add(nm);
      names.push(nm);
    }
  }
  return names;
}

// ── Docx parser: extract Armenian labels + routing rules ──────────────────────
async function parseDocxFile(buffer: ArrayBuffer): Promise<{
  armenianLabels: ArmenianLabels;
  routingRules: RoutingRule[];
  rawText: string;
}> {
  const result = await mammoth.extractRawText({ arrayBuffer: buffer });
  const rawText = result.value;
  const lines = rawText.split(/\r?\n/).map(l => l.trim()).filter(l => l.length > 0);

  const armenianLabels: ArmenianLabels = {};
  const routingRules: RoutingRule[] = [];

  // Armenian Unicode block: \u0531-\u058F (Armenian letters)
  const hasArmenian = (s: string) => /[\u0531-\u058F]/.test(s);

  // Variable name pattern e.g. S4, B1, D3, A9_1, N2_98, A4TOM, etc.
  const varPattern = /\b([A-Za-z][A-Za-z0-9_]{0,19})\b/g;

  // Routing patterns (case-insensitive):
  // "if B1=9-10 ask B2", "if S4=1 ask S5 and S6", "B1=9-10→B2", etc.
  // Also Armenian patterns using "ask" / "→" / "→" / "-ի դեպքում"
  const routingPatterns = [
    // English-style: "if VAR=X ask VAR" or "if VAR=X-Y ask VAR"
    /if\s+([A-Za-z][A-Za-z0-9_]{0,19})\s*=\s*([\d\-,]+)\s+(?:ask|go to|skip to)\s+(.*)/i,
    // Arrow style: "VAR=X → VAR" or "VAR=X-Y → VAR, VAR"
    /([A-Za-z][A-Za-z0-9_]{0,19})\s*[=:]\s*([\d\-,]+)\s*[→➔>]+\s*(.*)/,
    // Armenian routing cue: "VAR=X դեպքում VAR"
    /([A-Za-z][A-Za-z0-9_]{0,19})\s*[=:]\s*([\d\-,]+)[^А-яA-Za-z0-9_]*(?:դեպքում|հարցեք|հարց)\s*(.*)/i,
    // SPSS-style: "FILTER BY VAR=X" or routing notes
    /filter\s+(?:by\s+)?([A-Za-z][A-Za-z0-9_]{0,19})\s*=\s*([\d\-,]+)/i,
  ];

  // Heuristic to extract targets (variable names) from the "ask XXX" part
  const extractTargets = (text: string): string[] => {
    const targets: string[] = [];
    const re = /\b([A-Za-z][A-Za-z0-9_]{0,19})\b/g;
    let m: RegExpExecArray | null;
    while ((m = re.exec(text)) !== null) {
      const name = m[1];
      // Only accept reasonable variable names (not common English words)
      if (name.length >= 2 && /^[A-Z][A-Za-z0-9_]+$/.test(name)) {
        targets.push(name);
      }
    }
    return targets;
  };

  // Track: what was the last variable name we saw on a line?
  let lastVarSeen = "";

  for (const line of lines) {
    // ── Extract routing rules ──────────────────────────────────────────────
    for (const pattern of routingPatterns) {
      const m = line.match(pattern);
      if (m) {
        const condVar = m[1];
        const condVal = m[2];
        const targetText = m[3] ?? "";
        const targets = extractTargets(targetText);
        routingRules.push({
          condition: `${condVar}=${condVal}`,
          targets,
          rawText: line.slice(0, 200),
        });
        break;
      }
    }

    // ── Extract Armenian question labels ───────────────────────────────────
    // Strategy: look for lines that contain a variable name AND Armenian text
    // OR lines that follow a variable name line and contain Armenian text

    // Check if line has a variable name pattern at start
    const varMatch = line.match(/^([A-Z][A-Za-z0-9_]{0,19})[\.\s\:\-]/);
    if (varMatch) {
      lastVarSeen = varMatch[1];
    }

    // Also check: line starts with a question code pattern like "B1.", "S4.", "A9_1."
    const qCodeMatch = line.match(/^([A-Z][A-Za-z0-9_]{0,19})[\.:\)]/);
    if (qCodeMatch && qCodeMatch[1].length <= 10) {
      const potentialVar = qCodeMatch[1];
      lastVarSeen = potentialVar;

      // If the rest of the line has Armenian, store it
      const rest = line.slice(qCodeMatch[0].length).trim();
      if (hasArmenian(rest) && rest.length > 5) {
        armenianLabels[potentialVar] = rest.slice(0, 300);
      }
    }

    // Standalone Armenian line → associate with last seen variable
    if (hasArmenian(line) && lastVarSeen && !armenianLabels[lastVarSeen]) {
      // Make sure the line looks like a question (has some substance)
      if (line.length > 8 && !line.match(/^\d+[\.\)]/)) {
        armenianLabels[lastVarSeen] = line.slice(0, 300);
      }
    }

    // Scan all variable names in the line and update lastVarSeen
    let vm: RegExpExecArray | null;
    varPattern.lastIndex = 0;
    while ((vm = varPattern.exec(line)) !== null) {
      const nm = vm[1];
      if (nm.length >= 2 && /^[A-Z]/.test(nm)) {
        lastVarSeen = nm;
      }
    }
  }

  return { armenianLabels, routingRules, rawText };
}

// ── UI Components ──────────────────────────────────────────────────────────────
function FileZone({ label, sub, accept, onLoad, loaded }: {
  label: string; sub: string; accept: string;
  onLoad: (f: File) => void; loaded: string;
}) {
  const ref = useRef<HTMLInputElement>(null);
  const [drag, setDrag] = useState(false);
  return (
    <div
      onClick={() => ref.current?.click()}
      onDragOver={e => { e.preventDefault(); setDrag(true); }}
      onDragLeave={() => setDrag(false)}
      onDrop={e => { e.preventDefault(); setDrag(false); if (e.dataTransfer.files[0]) onLoad(e.dataTransfer.files[0]); }}
      style={{ border: `2px dashed ${drag ? "#2563eb" : loaded ? "#22c55e" : "#cbd5e1"}`, borderRadius: 12, padding: "18px 14px", textAlign: "center", cursor: "pointer", background: loaded ? "#f0fdf4" : drag ? "#eff6ff" : "#f8fafc", transition: "all .2s" }}>
      <input ref={ref} type="file" accept={accept} style={{ display: "none" }}
        onChange={e => { if (e.target.files?.[0]) onLoad(e.target.files[0]); }} />
      <div style={{ fontSize: 26, marginBottom: 5 }}>{loaded ? "✅" : "📂"}</div>
      <div style={{ fontSize: 13, fontWeight: 600, color: loaded ? "#16a34a" : "#475569" }}>{loaded || label}</div>
      <div style={{ fontSize: 11, color: "#94a3b8", marginTop: 3 }}>{sub}</div>
    </div>
  );
}

function Badge({ type }: { type: IssueType }) {
  const m = ISSUE_TYPES[type];
  return <span style={{ background: m.bg, color: m.color, padding: "2px 8px", borderRadius: 4, fontSize: 11, fontWeight: 700, whiteSpace: "nowrap" }}>{m.label}</span>;
}

function Th({ children, style = {} }: { children: ReactNode; style?: React.CSSProperties }) {
  return <th style={{ padding: "9px 13px", textAlign: "left", fontWeight: 600, color: "#475569", borderBottom: "1px solid #e2e8f0", whiteSpace: "nowrap", fontSize: 12, ...style }}>{children}</th>;
}

function ExplanationBox({ explanation }: { explanation: string }) {
  const [open, setOpen] = useState(false);

  // Split explanation into segments: plain text and Armenian text (📌 prefix)
  const renderExplanation = (text: string) => {
    const parts = text.split(/(📌[^\n]*|📋[^\n]*)/g);
    return parts.map((part, i) => {
      if (part.startsWith("📌")) {
        return (
          <div key={i} style={{ marginTop: 6, background: "#fdf4ff", border: "1px solid #e9d5ff", borderRadius: 4, padding: "5px 8px", fontSize: 12, color: "#6b21a8", fontStyle: "italic" }}>
            {part}
          </div>
        );
      }
      if (part.startsWith("📋")) {
        return (
          <div key={i} style={{ marginTop: 6, background: "#eff6ff", border: "1px solid #bfdbfe", borderRadius: 4, padding: "5px 8px", fontSize: 12, color: "#1e40af" }}>
            {part}
          </div>
        );
      }
      return <span key={i}>{part}</span>;
    });
  };

  return (
    <div>
      <button
        onClick={e => { e.stopPropagation(); setOpen(o => !o); }}
        style={{ fontSize: 11, color: "#6366f1", background: "none", border: "1px solid #e0e7ff", borderRadius: 4, padding: "2px 8px", cursor: "pointer" }}>
        {open ? "▲ Hide" : "▼ Why?"}
      </button>
      {open && (
        <div style={{ marginTop: 6, background: "#f5f3ff", border: "1px solid #ddd6fe", borderRadius: 6, padding: "8px 10px", fontSize: 12, color: "#4c1d95", lineHeight: 1.6, maxWidth: 500 }}>
          {renderExplanation(explanation)}
        </div>
      )}
    </div>
  );
}

// ── Main App ───────────────────────────────────────────────────────────────────
export default function App() {
  const [data, setData]             = useState<RowData[]>([]);
  const [dataFileName, setDataFileName] = useState("");

  // SAV state
  const [savVarNames, setSavVarNames]       = useState<string[]>([]);
  const [savValueLabels, setSavValueLabels] = useState<SavValueLabels>({});
  const [savFileName, setSavFileName]       = useState("");

  // Docx state
  const [docxFileName, setDocxFileName]         = useState("");
  const [docxRouting, setDocxRouting]           = useState<RoutingRule[]>([]);
  const [armenianLabels, setArmenianLabels]     = useState<ArmenianLabels>({});
  const [docxRawText, setDocxRawText]           = useState("");
  const [docxLoading, setDocxLoading]           = useState(false);

  const [issues, setIssues]         = useState<Issue[]>([]);
  const [dsWarnings, setDsWarnings] = useState<DatasetWarning[]>([]);
  const [analyzed, setAnalyzed]     = useState(false);
  const [tab, setTab]               = useState("issues");
  const [filterType, setFilterType] = useState<IssueType | "ALL">("ALL");
  const [search, setSearch]         = useState("");
  const [loading, setLoading]       = useState(false);
  const [error, setError]           = useState("");

  // ── File loaders ─────────────────────────────────────────────────────────

  const loadDataFile = useCallback((file: File) => {
    if (!file) return;
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
          const ws = wb.Sheets[wb.SheetNames[0]];
          setData(XLSX.utils.sheet_to_json<RowData>(ws, { defval: "" }));
          setDataFileName(file.name);
        } catch (err) { setError("Excel parse error: " + (err as Error).message); }
      };
      reader.readAsArrayBuffer(file);
    } else {
      setError(`Unsupported data file: .${ext}. Upload .csv or .xlsx`);
    }
  }, []);

  const loadSavFile = useCallback((file: File) => {
    if (!file) return;
    setError("");
    const ext = file.name.split(".").pop()?.toLowerCase() ?? "";
    if (ext !== "sav" && ext !== "por") {
      setError(`Expected a .sav file, got .${ext}`);
      return;
    }
    const reader = new FileReader();
    reader.onload = e => {
      const buf = e.target?.result as ArrayBuffer;
      const { varNames, valueLabels } = parseSavFile(buf);
      setSavVarNames(varNames);
      setSavValueLabels(valueLabels);
      setSavFileName(file.name);
    };
    reader.readAsArrayBuffer(file);
  }, []);

  const loadDocxFile = useCallback(async (file: File) => {
    if (!file) return;
    setError("");
    const ext = file.name.split(".").pop()?.toLowerCase() ?? "";
    if (ext !== "docx" && ext !== "doc") {
      setError(`Expected a .docx file, got .${ext}`);
      return;
    }
    setDocxLoading(true);
    try {
      const buf = await file.arrayBuffer();
      const { armenianLabels: labels, routingRules: rules, rawText } = await parseDocxFile(buf);
      setArmenianLabels(labels);
      setDocxRouting(rules);
      setDocxRawText(rawText);
      setDocxFileName(file.name);
    } catch (err) {
      setError("docx parse error: " + (err as Error).message);
    } finally {
      setDocxLoading(false);
    }
  }, []);

  const analyze = () => {
    if (!data.length) { setError("Upload a CSV or Excel data file to run validation."); return; }
    setLoading(true); setError("");
    setTimeout(() => {
      const found = runSurveyRules(data, savValueLabels, docxRouting, armenianLabels);
      const warnings = runDatasetChecks(data);
      setIssues(found);
      setDsWarnings(warnings);
      setAnalyzed(true);
      setTab(found.length || warnings.length ? "issues" : "data");
      setLoading(false);
    }, 60);
  };

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

  const dataColumns = data.length ? Object.keys(data[0]) : [];
  const totalIssues = issues.length + dsWarnings.length;

  const downloadCSV = (rows: object[], name: string) => {
    const csv = Papa.unparse(rows);
    const a = document.createElement("a");
    a.href = URL.createObjectURL(new Blob([csv], { type: "text/csv" }));
    a.download = name; a.click();
  };

  // Summary counts for enrichments
  const nSavLabels = Object.keys(savValueLabels).length;
  const nDocxRouting = docxRouting.length;
  const nArmenianLabels = Object.keys(armenianLabels).length;

  const tabs = [
    totalIssues > 0 && ["issues", `🚩 Issues (${totalIssues})`],
    data.length > 0 && ["data", `📊 Data Table`],
    savVarNames.length > 0 && ["savvars", `🗂 SAV Variables (${savVarNames.length})`],
    docxRouting.length > 0 && ["routing", `📋 Routing Rules (${nDocxRouting})`],
    docxRawText.length > 0 && ["docxtext", `📄 Questionnaire Text`],
  ].filter(Boolean) as [string, string][];

  return (
    <div style={{ fontFamily: "Inter, system-ui, sans-serif", minHeight: "100vh", background: "#f1f5f9" }}>
      {/* Header */}
      <div style={{ background: "linear-gradient(135deg,#1e3a5f,#2563eb)", padding: "18px 28px", color: "#fff" }}>
        <div style={{ maxWidth: 1200, margin: "0 auto", display: "flex", alignItems: "center", gap: 16 }}>
          <div style={{ flex: 1 }}>
            <h1 style={{ margin: 0, fontSize: 20, fontWeight: 700 }}>🔍 ACBA Bank Survey Validator</h1>
            <p style={{ margin: "3px 0 0", fontSize: 12, opacity: .75 }}>
              Employer Branding Survey · Skip logic · Missing data · Range checks · Data quality · Armenian survey rules
            </p>
          </div>
          <div style={{ fontSize: 12, opacity: .7, textAlign: "right", lineHeight: 1.7 }}>
            {savFileName   && <div>🗃 {savFileName}{nSavLabels > 0 ? ` · ${nSavLabels} vars with value labels` : ""}</div>}
            {docxFileName  && <div>📝 {docxFileName}{nDocxRouting > 0 ? ` · ${nDocxRouting} routing rules` : ""}{nArmenianLabels > 0 ? ` · ${nArmenianLabels} Armenian labels` : ""}</div>}
            {dataFileName  && <div>📋 {dataFileName} — {data.length} rows, {dataColumns.length} cols</div>}
          </div>
        </div>
      </div>

      <div style={{ maxWidth: 1200, margin: "0 auto", padding: "20px 16px" }}>
        {/* Upload zones — 2×2 grid */}
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr 1fr", gap: 12, marginBottom: 14 }}>

          {/* Zone 1: Data file */}
          <div>
            <div style={{ fontSize: 11, fontWeight: 700, color: "#64748b", marginBottom: 6, textTransform: "uppercase", letterSpacing: .5 }}>
              1 — Data file <span style={{ color: "#ef4444" }}>*</span>
            </div>
            <FileZone
              label="Drop CSV or Excel data"
              sub=".csv · .xlsx — main data for row-level validation"
              onLoad={loadDataFile}
              loaded={dataFileName}
              accept=".csv,.xlsx,.xls"
            />
          </div>

          {/* Zone 2: SAV file */}
          <div>
            <div style={{ fontSize: 11, fontWeight: 700, color: "#64748b", marginBottom: 6, textTransform: "uppercase", letterSpacing: .5 }}>
              2 — SPSS .sav <span style={{ color: "#94a3b8", fontWeight: 400 }}>(optional)</span>
            </div>
            <FileZone
              label="Drop SPSS .sav file"
              sub=".sav — variable names + value labels extracted"
              onLoad={loadSavFile}
              loaded={savFileName}
              accept=".sav,.por"
            />
            {nSavLabels > 0 && (
              <div style={{ marginTop: 4, fontSize: 11, color: "#16a34a", background: "#f0fdf4", border: "1px solid #bbf7d0", borderRadius: 6, padding: "4px 8px" }}>
                ✅ {nSavLabels} vars with value labels parsed
              </div>
            )}
          </div>

          {/* Zone 3: Docx questionnaire */}
          <div>
            <div style={{ fontSize: 11, fontWeight: 700, color: "#64748b", marginBottom: 6, textTransform: "uppercase", letterSpacing: .5 }}>
              3 — Questionnaire .docx <span style={{ color: "#94a3b8", fontWeight: 400 }}>(optional)</span>
            </div>
            <FileZone
              label={docxLoading ? "⏳ Parsing docx…" : "Drop Word questionnaire"}
              sub=".docx — Armenian labels + routing rules extracted"
              onLoad={loadDocxFile}
              loaded={docxFileName}
              accept=".docx,.doc"
            />
            {(nDocxRouting > 0 || nArmenianLabels > 0) && (
              <div style={{ marginTop: 4, fontSize: 11, color: "#7c3aed", background: "#f5f3ff", border: "1px solid #ddd6fe", borderRadius: 6, padding: "4px 8px" }}>
                ✅ {nDocxRouting} routing rules · {nArmenianLabels} Armenian labels
              </div>
            )}
          </div>

          {/* Zone 4: Run button */}
          <div style={{ display: "flex", flexDirection: "column", justifyContent: "flex-end" }}>
            <button onClick={analyze} disabled={loading || !data.length}
              style={{ padding: "14px", background: loading ? "#94a3b8" : !data.length ? "#cbd5e1" : "#2563eb", color: "#fff", border: "none", borderRadius: 10, fontSize: 14, fontWeight: 700, cursor: data.length ? "pointer" : "not-allowed", width: "100%" }}>
              {loading ? "⏳ Analyzing…" : "🚀 Run Validation"}
            </button>
            <div style={{ marginTop: 8, background: "#fffbeb", border: "1px solid #fde68a", borderRadius: 8, padding: "8px 12px", fontSize: 11, color: "#92400e" }}>
              💡 SPSS: <strong>File → Export → CSV Data</strong> then upload .csv here.<br/>
              The .sav + .docx files <em>enrich</em> the built-in rules — they don't replace them.
            </div>
          </div>
        </div>

        {error && (
          <div style={{ background: "#fef2f2", border: "1px solid #fca5a5", borderRadius: 8, padding: "10px 14px", color: "#dc2626", fontSize: 13, marginBottom: 12 }}>
            ⚠️ {error}
          </div>
        )}

        {/* Enrichment banner — shown when SAV or docx is loaded */}
        {(nSavLabels > 0 || nDocxRouting > 0) && (
          <div style={{ background: "#eff6ff", border: "1px solid #bfdbfe", borderRadius: 8, padding: "8px 14px", fontSize: 12, color: "#1e40af", marginBottom: 12, display: "flex", gap: 16, flexWrap: "wrap" }}>
            <span>🔗 <strong>Enrichment active:</strong></span>
            {nSavLabels > 0 && <span>🗃 SAV value labels for {nSavLabels} variables</span>}
            {nDocxRouting > 0 && <span>📋 {nDocxRouting} docx routing rules loaded</span>}
            {nArmenianLabels > 0 && <span>🇦🇲 {nArmenianLabels} Armenian question labels — shown in ▼ Why? explanations</span>}
          </div>
        )}

        {analyzed && (
          <>
            {/* Summary cards */}
            <div style={{ display: "grid", gridTemplateColumns: "repeat(6,1fr)", gap: 8, marginBottom: 18 }}>
              {(Object.entries(ISSUE_TYPES) as [IssueType, {label:string;color:string;bg:string}][]).map(([type, meta]) => (
                <div key={type} onClick={() => setFilterType(filterType === type ? "ALL" : type)}
                  style={{ background: "#fff", border: `2px solid ${filterType === type ? meta.color : "#e2e8f0"}`, borderRadius: 10, padding: "10px 12px", cursor: "pointer", boxShadow: filterType === type ? `0 0 0 3px ${meta.bg}` : "none", transition: "all .15s" }}>
                  <div style={{ fontSize: 22, fontWeight: 800, color: typeCounts[type] > 0 ? meta.color : "#cbd5e1" }}>{typeCounts[type]}</div>
                  <div style={{ fontSize: 10, color: "#64748b", marginTop: 2, lineHeight: 1.3 }}>{meta.label}</div>
                </div>
              ))}
            </div>

            {/* Tabs + actions */}
            <div style={{ display: "flex", gap: 2, borderBottom: "2px solid #e2e8f0", marginBottom: 14, flexWrap: "wrap", alignItems: "center" }}>
              {tabs.map(([id, label]) => (
                <button key={id} onClick={() => setTab(id)}
                  style={{ padding: "8px 14px", border: "none", background: "none", borderBottom: tab === id ? "2px solid #2563eb" : "2px solid transparent", color: tab === id ? "#2563eb" : "#64748b", fontWeight: tab === id ? 700 : 400, cursor: "pointer", fontSize: 13, marginBottom: -2 }}>
                  {label}
                </button>
              ))}
              <div style={{ flex: 1 }} />
              {filterType !== "ALL" && (
                <button onClick={() => setFilterType("ALL")}
                  style={{ fontSize: 12, padding: "5px 10px", background: "#e0e7ff", border: "none", borderRadius: 6, cursor: "pointer", color: "#3730a3", marginRight: 4 }}>
                  ✕ Clear filter
                </button>
              )}
              {issues.length > 0 && (
                <button onClick={() => downloadCSV(issues.map(i => ({ ID: i.id, Variable: i.variable, Type: ISSUE_TYPES[i.type].label, Value: String(i.value ?? ""), Detail: i.detail, Explanation: i.explanation })), "issues_report.csv")}
                  style={{ fontSize: 12, padding: "5px 10px", background: "#f1f5f9", border: "1px solid #e2e8f0", borderRadius: 6, cursor: "pointer", color: "#475569" }}>
                  ⬇ Export Issues CSV
                </button>
              )}
            </div>

            {/* Search */}
            <input
              placeholder="Search by variable name, respondent ID, or issue description…"
              value={search} onChange={e => setSearch(e.target.value)}
              style={{ width: "100%", padding: "8px 12px", border: "1px solid #e2e8f0", borderRadius: 8, fontSize: 13, marginBottom: 12, boxSizing: "border-box" }}
            />

            {/* ── Issues tab ── */}
            {tab === "issues" && (
              <div>
                {/* Structural / dataset-level warnings */}
                {dsWarnings.length > 0 && (
                  <div style={{ marginBottom: 14 }}>
                    <div style={{ fontSize: 12, fontWeight: 700, color: "#64748b", marginBottom: 6, textTransform: "uppercase", letterSpacing: .5 }}>⚠️ Dataset-level structural issues ({dsWarnings.length})</div>
                    {dsWarnings.map((w, i) => (
                      <div key={i} style={{ background: "#fff7ed", border: "1px solid #fed7aa", borderRadius: 8, padding: "12px 14px", marginBottom: 8, borderLeft: "4px solid #f97316" }}>
                        <div style={{ fontWeight: 700, fontSize: 13, color: "#9a3412", marginBottom: 4 }}>
                          <code style={{ background: "#fef3c7", padding: "1px 5px", borderRadius: 3 }}>{w.variable}</code>
                          {" "}{w.detail}
                        </div>
                        <div style={{ fontSize: 12, color: "#92400e", lineHeight: 1.6, background: "#fef9f0", border: "1px solid #fde68a", borderRadius: 4, padding: "7px 10px" }}>
                          📋 {w.explanation}
                        </div>
                      </div>
                    ))}
                  </div>
                )}

                {filteredIssues.length === 0 ? (
                  <div style={{ background: "#fff", borderRadius: 10, padding: 48, textAlign: "center", color: "#94a3b8" }}>
                    <div style={{ fontSize: 44, marginBottom: 10 }}>{issues.length === 0 ? "✅" : "🔍"}</div>
                    <div style={{ fontWeight: 600, fontSize: 15 }}>{issues.length === 0 ? "No row-level issues found!" : "No issues match your current filter."}</div>
                  </div>
                ) : (
                  <div style={{ background: "#fff", borderRadius: 10, border: "1px solid #e2e8f0", overflow: "hidden" }}>
                    <div style={{ overflowX: "auto" }}>
                      <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
                        <thead>
                          <tr style={{ background: "#f8fafc" }}>
                            <Th>ID</Th>
                            <Th>Variable</Th>
                            <Th>Issue Type</Th>
                            <Th>Value</Th>
                            <Th>Description</Th>
                            <Th>Explanation</Th>
                          </tr>
                        </thead>
                        <tbody>
                          {filteredIssues.map((iss, i) => (
                            <tr key={i} style={{ background: i % 2 === 0 ? "#fff" : "#fafafa", borderLeft: `3px solid ${ISSUE_TYPES[iss.type].color}` }}>
                              <td style={{ padding: "8px 13px", fontWeight: 700, color: "#1e293b", whiteSpace: "nowrap" }}>{String(iss.id)}</td>
                              <td style={{ padding: "8px 13px", fontFamily: "monospace", fontSize: 12, color: "#7c3aed", whiteSpace: "nowrap" }}>{iss.variable}</td>
                              <td style={{ padding: "8px 13px", whiteSpace: "nowrap" }}><Badge type={iss.type} /></td>
                              <td style={{ padding: "8px 13px", fontFamily: "monospace", color: "#dc2626", maxWidth: 120, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }} title={String(iss.value ?? "")}>
                                {String(iss.value ?? "—").slice(0, 60)}
                              </td>
                              <td style={{ padding: "8px 13px", color: "#374151" }}>{iss.detail}</td>
                              <td style={{ padding: "8px 13px", minWidth: 100 }}>
                                <ExplanationBox explanation={iss.explanation} />
                              </td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                    <div style={{ padding: "7px 13px", fontSize: 12, color: "#94a3b8", borderTop: "1px solid #e2e8f0", display: "flex", justifyContent: "space-between" }}>
                      <span>Showing {filteredIssues.length} of {issues.length} row-level issues</span>
                      <span>Click <strong>▼ Why?</strong> in any row for the full validation rule explanation</span>
                    </div>
                  </div>
                )}
              </div>
            )}

            {/* ── Data table tab ── */}
            {tab === "data" && (
              <div style={{ background: "#fff", borderRadius: 10, border: "1px solid #e2e8f0", overflow: "auto", maxHeight: 560 }}>
                <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 11 }}>
                  <thead style={{ position: "sticky", top: 0, background: "#f8fafc", zIndex: 1 }}>
                    <tr>{dataColumns.map(col => <Th key={col}>{col}</Th>)}</tr>
                  </thead>
                  <tbody>
                    {data.slice(0, 200).map((row, ri) => {
                      const rid = row.id ?? row.ID ?? row.RespondentID ?? `Row ${ri + 1}`;
                      return (
                        <tr key={ri} style={{ background: ri % 2 === 0 ? "#fff" : "#fafafa" }}>
                          {dataColumns.map(col => {
                            const it = issueMap[`${rid}__${col}`];
                            const meta = it ? ISSUE_TYPES[it] : null;
                            const tooltip = meta ? `${meta.label}: ${issues.find(iss => iss.id == rid && iss.variable === col)?.detail || ""}` : "";
                            return (
                              <td key={col} title={tooltip}
                                style={{ padding: "5px 9px", background: meta ? meta.bg : "transparent", color: meta ? meta.color : "#334155", fontFamily: "monospace", fontSize: 11, borderBottom: "1px solid #f1f5f9", borderLeft: meta ? `2px solid ${meta.color}` : "none", whiteSpace: "nowrap" }}>
                                {String(row[col] ?? "").slice(0, 40)}
                              </td>
                            );
                          })}
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
                {data.length > 200 && <div style={{ padding: "7px 13px", fontSize: 12, color: "#94a3b8", borderTop: "1px solid #e2e8f0" }}>Showing first 200 of {data.length} rows. Highlighted cells have detected issues — hover for details.</div>}
              </div>
            )}

            {/* ── SAV vars tab ── */}
            {tab === "savvars" && (
              <div style={{ background: "#fff", borderRadius: 10, border: "1px solid #e2e8f0", overflow: "auto", maxHeight: 520 }}>
                <div style={{ padding: "10px 14px", background: "#fffbeb", borderBottom: "1px solid #fde68a", fontSize: 12, color: "#92400e" }}>
                  Variable names + value labels extracted from .sav binary.
                  {nSavLabels > 0 && <> · <strong>{nSavLabels} variables</strong> have value labels that enrich validation.</>}
                </div>
                <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
                  <thead style={{ position: "sticky", top: 0, background: "#f8fafc", zIndex: 1 }}>
                    <tr><Th>#</Th><Th>Variable Name</Th><Th>Description</Th><Th>SAV Value Labels</Th></tr>
                  </thead>
                  <tbody>
                    {savVarNames.filter(nm => !search || nm.toLowerCase().includes(search.toLowerCase())).slice(0, 400).map((nm, i) => {
                      const vlabels = savValueLabels[nm];
                      return (
                        <tr key={nm} style={{ background: i % 2 === 0 ? "#fff" : "#fafafa" }}>
                          <td style={{ padding: "7px 13px", color: "#94a3b8", fontSize: 11 }}>{i + 1}</td>
                          <td style={{ padding: "7px 13px", fontFamily: "monospace", fontWeight: 700, color: "#1e293b" }}>{nm}</td>
                          <td style={{ padding: "7px 13px", color: VAR_DESCRIPTIONS[nm] ? "#334155" : "#cbd5e1", fontStyle: VAR_DESCRIPTIONS[nm] ? "normal" : "italic", fontSize: 12 }}>
                            {VAR_DESCRIPTIONS[nm] || "—"}
                          </td>
                          <td style={{ padding: "7px 13px", fontSize: 11, color: "#475569" }}>
                            {vlabels ? (
                              <div style={{ display: "flex", flexWrap: "wrap", gap: 3 }}>
                                {Object.entries(vlabels).slice(0, 20).map(([code, label]) => (
                                  <span key={code} style={{ background: "#f0f9ff", border: "1px solid #bae6fd", borderRadius: 3, padding: "1px 5px", fontSize: 10, color: "#0369a1", whiteSpace: "nowrap" }}>
                                    {code}={label}
                                  </span>
                                ))}
                                {Object.keys(vlabels).length > 20 && <span style={{ color: "#94a3b8", fontSize: 10 }}>+{Object.keys(vlabels).length - 20} more</span>}
                              </div>
                            ) : <span style={{ color: "#cbd5e1", fontStyle: "italic" }}>—</span>}
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            )}

            {/* ── Routing rules tab ── */}
            {tab === "routing" && (
              <div style={{ background: "#fff", borderRadius: 10, border: "1px solid #e2e8f0", overflow: "auto", maxHeight: 520 }}>
                <div style={{ padding: "10px 14px", background: "#f5f3ff", borderBottom: "1px solid #ddd6fe", fontSize: 12, color: "#4c1d95" }}>
                  Skip/routing rules extracted from the Word questionnaire (.docx). These rules enrich the built-in validation logic.
                </div>
                <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
                  <thead style={{ position: "sticky", top: 0, background: "#f8fafc", zIndex: 1 }}>
                    <tr><Th>#</Th><Th>Condition</Th><Th>Target Variables</Th><Th>Source Text</Th></tr>
                  </thead>
                  <tbody>
                    {docxRouting.filter(r => !search || r.condition.toLowerCase().includes(search.toLowerCase()) || r.rawText.toLowerCase().includes(search.toLowerCase())).map((rule, i) => (
                      <tr key={i} style={{ background: i % 2 === 0 ? "#fff" : "#fafafa" }}>
                        <td style={{ padding: "7px 13px", color: "#94a3b8", fontSize: 11 }}>{i + 1}</td>
                        <td style={{ padding: "7px 13px", fontFamily: "monospace", fontWeight: 700, color: "#7c3aed" }}>{rule.condition}</td>
                        <td style={{ padding: "7px 13px" }}>
                          <div style={{ display: "flex", gap: 4, flexWrap: "wrap" }}>
                            {rule.targets.map(t => (
                              <span key={t} style={{ background: "#eff6ff", border: "1px solid #bfdbfe", borderRadius: 3, padding: "1px 6px", fontSize: 11, color: "#1d4ed8" }}>{t}</span>
                            ))}
                          </div>
                        </td>
                        <td style={{ padding: "7px 13px", fontSize: 11, color: "#475569", maxWidth: 400, wordBreak: "break-word" }}>
                          {rule.rawText.slice(0, 200)}
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )}

            {/* ── Docx raw text tab ── */}
            {tab === "docxtext" && (
              <div style={{ background: "#fff", borderRadius: 10, border: "1px solid #e2e8f0", overflow: "auto", maxHeight: 520 }}>
                <div style={{ padding: "10px 14px", background: "#f5f3ff", borderBottom: "1px solid #ddd6fe", fontSize: 12, color: "#4c1d95", display: "flex", justifyContent: "space-between" }}>
                  <span>Raw text extracted from <strong>{docxFileName}</strong> — {docxRawText.length.toLocaleString()} characters</span>
                  <span style={{ color: "#7c3aed" }}>🇦🇲 {nArmenianLabels} Armenian question labels detected</span>
                </div>
                <div style={{ padding: "12px 16px", fontSize: 12, color: "#334155", lineHeight: 1.8, fontFamily: "monospace", whiteSpace: "pre-wrap", wordBreak: "break-word" }}>
                  {docxRawText.slice(0, 50000)}
                  {docxRawText.length > 50000 && (
                    <div style={{ color: "#94a3b8", marginTop: 12, fontFamily: "sans-serif" }}>
                      … (showing first 50,000 of {docxRawText.length.toLocaleString()} characters)
                    </div>
                  )}
                </div>
              </div>
            )}
          </>
        )}

        {!analyzed && (
          <div style={{ textAlign: "center", padding: "60px 20px", color: "#94a3b8" }}>
            <div style={{ fontSize: 52, marginBottom: 16 }}>📋</div>
            <div style={{ fontSize: 16, fontWeight: 600, color: "#64748b" }}>Upload your data to begin validation</div>
            <div style={{ fontSize: 13, marginTop: 8, maxWidth: 540, margin: "10px auto 0", lineHeight: 1.8, color: "#64748b" }}>
              This validator has the <strong>ACBA Bank Employer Branding questionnaire rules built in</strong>.<br />
              The <strong>.sav</strong> and <strong>.docx</strong> files enrich those rules with real value labels and Armenian question text.
            </div>
            <div style={{ marginTop: 24, display: "inline-block", textAlign: "left", background: "#fff", border: "1px solid #e2e8f0", borderRadius: 10, padding: "16px 20px", fontSize: 12, color: "#475569" }}>
              <div style={{ fontWeight: 700, marginBottom: 10, color: "#1e293b", fontSize: 13 }}>What this tool checks:</div>
              {[
                ["🔴 Skip/Routing Violations", "A4_1=1 but A4TOM=99 (refusal) · B2 filled but B1<9 · D4 filled but S7≠1 · A9 filled but A8=0"],
                ["🟠 Out of Range", "B1 not in 0-10 · A1/A2 ratings not in 1-5 · D5 salary code not in 1-6"],
                ["🟡 Mismatched Code", "D6 bank code invalid · A8 value not 0/1 · S7 segment code invalid"],
                ["🟣 Missing Data", "B2 empty when B1=9-10 · B3 empty when B1≤8 · A101 missing when A9_1=4-5 · D31 missing when D3=1"],
                ["🔵 Data Quality", "Garbled open-text (random characters) · A4TOM double-counted in checkboxes"],
                ["🟢 Open Text", "Mandatory open fields too short or empty"],
                ["⚠️ Structural", "A3_12 (colleague relations) column MISSING from entire dataset · Duplicate IDs"],
              ].map(([type, ex]) => (
                <div key={type} style={{ marginBottom: 8 }}>
                  <strong style={{ display: "block" }}>{type}</strong>
                  <span style={{ color: "#94a3b8", fontSize: 11 }}>{ex}</span>
                </div>
              ))}
              <div style={{ marginTop: 12, padding: "8px 12px", background: "#f5f3ff", borderRadius: 6, border: "1px solid #ddd6fe", fontSize: 11, color: "#6d28d9" }}>
                🗃 Upload <strong>.sav</strong> → value labels enrich code validation<br/>
                📝 Upload <strong>.docx</strong> → Armenian question text appears in ▼ Why? explanations + routing rules tab
              </div>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}
