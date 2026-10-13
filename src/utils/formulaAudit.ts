/**
 * Formula auditor: error values, broken references, hard-coded numbers inside formulas, and formulas
 * that differ from their neighbours down a column. Pure functions (no Office) – unit tested.
 */
import { a1 } from "./a1";

export type AuditKind = "error" | "brokenRef" | "inconsistent" | "valueInFormulaColumn" | "hardcoded";

export interface AuditIssue {
  /** 0-based sheet row / column. */
  row: number;
  col: number;
  address: string;
  kind: AuditKind;
  severity: "high" | "medium" | "low";
  detail: string;
  formula: string;
}

export interface AuditBlock {
  rowIndex: number;
  columnIndex: number;
  formulas: unknown[][];
  formulasR1C1: unknown[][];
  values: unknown[][];
  valueTypes: string[][];
}

const ERROR_RE = /^#(DIV\/0!|N\/A|NAME\?|NULL!|NUM!|REF!|VALUE!|SPILL!|CALC!|GETTING_DATA|FIELD!|BUSY!|BLOCKED!|CONNECT!|UNKNOWN!)/i;

const ERROR_HELP: Record<string, string> = {
  "#DIV/0!": "Division by zero or an empty cell – wrap with IFERROR or check the divisor.",
  "#N/A": "A lookup found no match – check the lookup value or use IFNA.",
  "#NAME?": "Unknown function or name – check spelling or a missing named range.",
  "#REF!": "Points to a deleted cell or sheet.",
  "#VALUE!": "Wrong type of value (text where a number is expected).",
  "#NUM!": "Invalid number (e.g. negative square root, no convergence).",
  "#SPILL!": "The result can't spill – clear the cells in its way.",
  "#NULL!": "Ranges that don't intersect (space instead of comma?).",
};

/** Numbers that are normal inside formulas (0, 1, percentages, unit-free counters). */
const OK_NUMBERS = new Set(["0", "1", "2", "10", "12", "24", "60", "100", "1000", "365", "-1"]);

/** Numeric literals in a formula, ignoring cell references, strings, sheet names and function names. */
export function numericLiterals(formula: string): string[] {
  const body = formula
    .replace(/^=/, "")
    .replace(/"([^"]|"")*"/g, " ") // strings
    .replace(/'([^']|'')*'!/g, " ") // quoted sheet names
    .replace(/\b[A-Za-z_][A-Za-z0-9_.]*!/g, " ") // sheet names
    .replace(/\$?\b[A-Za-z]{1,3}\$?\d+(:\$?[A-Za-z]{1,3}\$?\d+)?\b/g, " ") // A1 / A1:B2
    .replace(/\$?\b[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}\b/g, " ") // A:B
    .replace(/\$?\b\d+:\$?\d+\b/g, " ") // 1:5
    .replace(/\b[A-Za-z_][A-Za-z0-9_.]*\s*\(/g, "(") // functions
    .replace(/\b[A-Za-z_][A-Za-z0-9_.]*\b/g, " "); // names
  return numbersIn(body).filter((n) => !OK_NUMBERS.has(n));
}

/** Standalone numbers in a text (no lookbehind – older Mac Office webviews don't support it). */
function numbersIn(text: string): string[] {
  const out: string[] = [];
  const re = /(^|[^\w.$])(\d+(?:\.\d+)?(?:[eE][-+]?\d+)?)(?![\w.])/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(text))) out.push(m[2]);
  return out;
}

/** Second argument of ROUND/ROUNDUP/ROUNDDOWN/LEFT/RIGHT/MID/INDEX… is a count, not a constant: drop those numbers. */
function countArgs(formula: string): Set<string> {
  const out = new Set<string>();
  const re = /\b(ROUND|ROUNDUP|ROUNDDOWN|LEFT|RIGHT|MID|INDEX|OFFSET|CHOOSE|FIXED|TEXT|WEEKDAY|EDATE|EOMONTH|MATCH|VLOOKUP|HLOOKUP|XLOOKUP|LARGE|SMALL|PERCENTILE|QUARTILE)\s*\(([^()]*)\)/gi;
  let m: RegExpExecArray | null;
  while ((m = re.exec(formula))) for (const n of numbersIn(m[2])) out.add(n);
  return out;
}

const isFormula = (f: unknown) => typeof f === "string" && f.startsWith("=");

export function auditFormulas(b: AuditBlock, maxIssues = 2000): AuditIssue[] {
  const issues: AuditIssue[] = [];
  const push = (i: number, j: number, kind: AuditKind, severity: AuditIssue["severity"], detail: string) => {
    if (issues.length >= maxIssues) return;
    const f = b.formulas[i]?.[j];
    issues.push({ row: b.rowIndex + i, col: b.columnIndex + j, address: a1(b.rowIndex + i, b.columnIndex + j), kind, severity, detail, formula: typeof f === "string" ? f : String(f ?? "") });
  };
  const rows = b.formulas.length;
  const cols = b.formulas[0]?.length ?? 0;

  for (let i = 0; i < rows; i++) {
    for (let j = 0; j < cols; j++) {
      const f = b.formulas[i][j];
      const v = b.values[i]?.[j];
      const isErr = b.valueTypes[i]?.[j] === "Error" || (typeof v === "string" && ERROR_RE.test(v));
      if (isFormula(f) && /#REF!/i.test(f as string)) push(i, j, "brokenRef", "high", "Formula contains #REF! – it points to a deleted cell, row, column or sheet.");
      else if (isErr) {
        const code = String(v).toUpperCase().match(/^#[A-Z/0!?_]+[!?]?/)?.[0] ?? String(v);
        push(i, j, "error", "high", `${code} – ${ERROR_HELP[code] ?? "Error value."}`);
      }
      if (isFormula(f)) {
        const counts = countArgs(f as string);
        const lits = numericLiterals(f as string).filter((n) => !counts.has(n));
        if (lits.length) push(i, j, "hardcoded", "low", `Hard-coded number${lits.length > 1 ? "s" : ""} ${Array.from(new Set(lits)).slice(0, 4).join(", ")} – consider an input cell or named constant so it can be changed in one place.`);
      }
    }
  }

  // Inconsistent formulas: in each column, compare a formula with the R1C1 shape most cells share.
  for (let j = 0; j < cols; j++) {
    let start = 0;
    while (start < rows) {
      // A "run" = consecutive rows that are formulas or numbers (stops at blanks / text such as headers).
      while (start < rows && !isFormula(b.formulas[start][j]) && typeof b.formulas[start][j] !== "number") start++;
      let end = start;
      while (end < rows && (isFormula(b.formulas[end][j]) || typeof b.formulas[end][j] === "number")) end++;
      const run = Array.from({ length: end - start }, (_, k) => start + k);
      const formulaRows = run.filter((i) => isFormula(b.formulas[i][j]));
      if (formulaRows.length >= 3) {
        const counts = new Map<string, number>();
        formulaRows.forEach((i) => counts.set(String(b.formulasR1C1[i][j]), (counts.get(String(b.formulasR1C1[i][j])) ?? 0) + 1));
        const [shape, n] = Array.from(counts.entries()).sort((a, c) => c[1] - a[1])[0];
        if (n >= 3 && n / formulaRows.length >= 0.6 && n < run.length) {
          const typical = b.formulas[formulaRows.find((i) => String(b.formulasR1C1[i][j]) === shape) as number][j];
          for (const i of run) {
            const f = b.formulas[i][j];
            if (isFormula(f) && String(b.formulasR1C1[i][j]) !== shape)
              push(i, j, "inconsistent", "medium", `Differs from the ${n} other formulas in this column (they look like ${String(typical)}).`);
            else if (!isFormula(f) && n / formulaRows.length >= 0.6)
              push(i, j, "valueInFormulaColumn", "medium", `Typed value ${String(f)} in a column of formulas (like ${String(typical)}) – was a formula overwritten?`);
          }
        }
      }
      start = Math.max(end, start + 1);
    }
  }
  const order = { high: 0, medium: 1, low: 2 };
  return issues.sort((a, c) => order[a.severity] - order[c.severity] || a.col - c.col || a.row - c.row);
}

export const AUDIT_LABELS: Record<AuditKind, string> = {
  error: "Error value",
  brokenRef: "Broken reference",
  inconsistent: "Inconsistent formula",
  valueInFormulaColumn: "Value in formula column",
  hardcoded: "Hard-coded number",
};
