/**
 * Pure helpers that turn loosely-typed agent tool arguments into Office.js-ready values.
 * No Office dependency, so they are unit-tested.
 */

/** "A" -> 0, "AB" -> 27; null when not a column letter. */
export function columnIndexFromLetter(s: string): number | null {
  const t = s.trim().toUpperCase().replace(/^\$/, "");
  if (!/^[A-Z]{1,3}$/.test(t)) return null;
  let n = 0;
  for (const ch of t) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n - 1;
}

/**
 * Find a column inside a range: by sheet column letter ("C"), 1-based position within the range ("2"),
 * or header text (case-insensitive, exact first, then "contains").
 * Returns the 0-based index relative to the range, or null.
 */
export function resolveColumn(spec: unknown, headerRow: unknown[], rangeColumnIndex: number): number | null {
  const width = headerRow.length;
  if (typeof spec === "number" && Number.isInteger(spec)) return spec >= 1 && spec <= width ? spec - 1 : null;
  const s = String(spec ?? "").trim();
  if (!s) return null;
  const headers = headerRow.map((h) => String(h ?? "").trim().toLowerCase());
  const exact = headers.indexOf(s.toLowerCase());
  if (exact >= 0) return exact;
  // "C", "AB", "$D" or upper-case "OIL" are column letters; lower-case words like "oil" are header text.
  const looksLikeLetter = /^\$?[A-Z]{1,3}$/.test(s) || /^[a-z]{1,2}$/i.test(s);
  if (looksLikeLetter) {
    const rel = (columnIndexFromLetter(s) as number) - rangeColumnIndex;
    if (rel >= 0 && rel < width) return rel;
    if (s.replace("$", "").length <= 2) return null;
  }
  if (/^\d+$/.test(s)) {
    const n = Number(s);
    return n >= 1 && n <= width ? n - 1 : null;
  }
  if (s.length < 3) return null;
  const partial = headers.findIndex((h) => h !== "" && h.includes(s.toLowerCase()));
  return partial >= 0 ? partial : null;
}

export interface FilterArgs {
  values?: unknown;
  condition?: unknown;
  condition2?: unknown;
  operator?: unknown;
  top?: unknown;
  bottom?: unknown;
  top_percent?: unknown;
  bottom_percent?: unknown;
}

export interface FilterCriteriaLike {
  filterOn: "Values" | "Custom" | "TopItems" | "BottomItems" | "TopPercent" | "BottomPercent";
  values?: string[];
  criterion1?: string;
  criterion2?: string;
  operator?: "And" | "Or";
}

const posInt = (v: unknown) => {
  const n = Number(v);
  return Number.isFinite(n) && n > 0 ? Math.round(n) : null;
};

/** Build AutoFilter criteria from agent arguments. Throws a readable error when nothing usable was given. */
export function buildFilterCriteria(a: FilterArgs): FilterCriteriaLike {
  if (Array.isArray(a.values) && a.values.length) return { filterOn: "Values", values: a.values.map((v) => String(v)) };
  const top = posInt(a.top);
  if (top) return { filterOn: "TopItems", criterion1: String(top) };
  const bottom = posInt(a.bottom);
  if (bottom) return { filterOn: "BottomItems", criterion1: String(bottom) };
  const tp = posInt(a.top_percent);
  if (tp) return { filterOn: "TopPercent", criterion1: String(tp) };
  const bp = posInt(a.bottom_percent);
  if (bp) return { filterOn: "BottomPercent", criterion1: String(bp) };
  const c1 = typeof a.condition === "string" ? a.condition.trim() : "";
  if (c1) {
    const c2 = typeof a.condition2 === "string" ? a.condition2.trim() : "";
    const norm = (c: string) => (/^(=|<>|>=|<=|>|<)/.test(c) ? c : `=${c}`);
    return c2
      ? { filterOn: "Custom", criterion1: norm(c1), criterion2: norm(c2), operator: String(a.operator).toLowerCase() === "or" ? "Or" : "And" }
      : { filterOn: "Custom", criterion1: norm(c1) };
  }
  throw new Error("Give values, a condition like \">100\", or top/bottom N.");
}

export type CellValueOperator = "Between" | "NotBetween" | "EqualTo" | "NotEqualTo" | "GreaterThan" | "LessThan" | "GreaterThanOrEqual" | "LessThanOrEqual";

const OPERATORS: Record<string, CellValueOperator> = {
  between: "Between",
  notbetween: "NotBetween",
  equal: "EqualTo",
  equalto: "EqualTo",
  "=": "EqualTo",
  notequal: "NotEqualTo",
  notequalto: "NotEqualTo",
  "<>": "NotEqualTo",
  greaterthan: "GreaterThan",
  ">": "GreaterThan",
  lessthan: "LessThan",
  "<": "LessThan",
  greaterthanorequal: "GreaterThanOrEqual",
  greaterthanorequalto: "GreaterThanOrEqual",
  ">=": "GreaterThanOrEqual",
  lessthanorequal: "LessThanOrEqual",
  lessthanorequalto: "LessThanOrEqual",
  "<=": "LessThanOrEqual",
};

/** "greater_than", "GreaterThan", ">" … -> Office operator name. */
export function normalizeOperator(op: unknown): CellValueOperator | null {
  const key = String(op ?? "").toLowerCase().replace(/[\s_-]/g, "");
  return OPERATORS[key] ?? null;
}

/** Conditional-format and validation formulas: numbers stay numbers, everything else becomes "=…" text. */
export function cfFormula(v: unknown): string {
  if (typeof v === "number") return String(v);
  const s = String(v ?? "").trim();
  if (!s) return "";
  if (s.startsWith("=")) return s;
  if (Number.isFinite(Number(s))) return s;
  return `="${s.replace(/"/g, '""')}"`;
}

export type Aggregation = "Sum" | "Count" | "Average" | "Max" | "Min" | "Product" | "CountNumbers" | "StandardDeviation" | "Variance";

/** "avg", "mean", "sum", "count" … -> PivotTable summarize-by name. */
export function normalizeAggregation(v: unknown): Aggregation {
  const k = String(v ?? "sum").toLowerCase().replace(/[\s_-]/g, "");
  if (["avg", "average", "mean"].includes(k)) return "Average";
  if (["count", "counta"].includes(k)) return "Count";
  if (["countnumbers", "countnums"].includes(k)) return "CountNumbers";
  if (["max", "maximum"].includes(k)) return "Max";
  if (["min", "minimum"].includes(k)) return "Min";
  if (k === "product") return "Product";
  if (["stdev", "std", "stddev", "standarddeviation"].includes(k)) return "StandardDeviation";
  if (["var", "variance"].includes(k)) return "Variance";
  return "Sum";
}

/** Find a header case-insensitively; returns the header exactly as written in the sheet, or null. */
export function matchHeader(name: unknown, headers: string[]): string | null {
  const s = String(name ?? "").trim().toLowerCase();
  if (!s) return null;
  return headers.find((h) => h.trim().toLowerCase() === s) ?? headers.find((h) => h.trim().toLowerCase().includes(s)) ?? null;
}

/** Excel defined-name rules (simplified): letter/underscore first, no spaces, not a cell reference. */
export function validName(name: string): boolean {
  return /^[A-Za-z_\\][A-Za-z0-9_.\\]{0,254}$/.test(name) && !/^[A-Za-z]{1,3}\d+$/.test(name) && !/^[rRcC]$/.test(name);
}
