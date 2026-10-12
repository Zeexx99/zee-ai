/** Pure data-cleaning helpers used by the "Clean Data" quick action. */

export type CaseMode = "none" | "upper" | "lower" | "proper" | "sentence";
export type DateOrder = "off" | "MDY" | "DMY" | "YMD";

export interface CleanOptions {
  trimWhitespace: boolean;
  removeNonPrintable: boolean;
  caseMode: CaseMode;
  convertNumbers: boolean;
  dateOrder: DateOrder;
}

export const DEFAULT_CLEAN_OPTIONS: CleanOptions = {
  trimWhitespace: true,
  removeNonPrintable: true,
  caseMode: "none",
  convertNumbers: true,
  dateOrder: "off",
};

export interface CleanResult {
  value: unknown;
  /** New number format for the cell, when the type changed (dates, percentages). */
  numberFormat?: string;
  changed: boolean;
}

const NON_PRINTABLE = /[\u0000-\u0008\u000B\u000C\u000E-\u001F\u007F​-‍⁠﻿]/g;
const MONTHS: Record<string, number> = {
  jan: 1, feb: 2, mar: 3, apr: 4, may: 5, jun: 6, jul: 7, aug: 8, sep: 9, sept: 9, oct: 10, nov: 11, dec: 12,
};

function looksLikeEmailOrUrl(s: string): boolean {
  return /@|https?:\/\/|www\./i.test(s);
}

export function toProperCase(s: string): string {
  return s.toLowerCase().replace(/(^|[\s\-(/"])(\p{L})/gu, (_m, p: string, c: string) => p + c.toUpperCase());
}

export function toSentenceCase(s: string): string {
  return s.toLowerCase().replace(/(^\s*\p{L}|[.!?]\s+\p{L})/gu, (m) => m.toUpperCase());
}

/** Excel serial date (1900 system) from a calendar date, or null when invalid. */
export function toExcelSerial(y: number, m: number, d: number): number | null {
  if (y < 100) y += y < 50 ? 2000 : 1900;
  if (m < 1 || m > 12 || d < 1 || d > 31 || y < 1900 || y > 9999) return null;
  const ms = Date.UTC(y, m - 1, d);
  const check = new Date(ms);
  if (check.getUTCFullYear() !== y || check.getUTCMonth() !== m - 1 || check.getUTCDate() !== d) return null;
  return Math.round((ms - Date.UTC(1899, 11, 30)) / 86400000);
}

/** Parse common textual dates. Returns an Excel serial or null. */
export function parseDateText(s: string, order: DateOrder): number | null {
  if (order === "off") return null;
  const t = s.trim();
  let m = /^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})$/.exec(t); // ISO-like always Y-M-D
  if (m) return toExcelSerial(+m[1], +m[2], +m[3]);
  m = /^(\d{1,2})[-/.](\d{1,2})[-/.](\d{2}|\d{4})$/.exec(t);
  if (m) {
    const a = +m[1], b = +m[2], y = +m[3];
    if (order === "MDY") return toExcelSerial(y, a, b);
    if (order === "DMY") return toExcelSerial(y, b, a);
    return null;
  }
  // "5 Jan 2024", "05-Jan-24"
  m = /^(\d{1,2})[\s\-.]+([A-Za-z]{3,9})\.?[\s\-.,]+(\d{2}|\d{4})$/.exec(t);
  if (m) {
    const mon = MONTHS[m[2].slice(0, 4).toLowerCase()] ?? MONTHS[m[2].slice(0, 3).toLowerCase()];
    return mon ? toExcelSerial(+m[3], mon, +m[1]) : null;
  }
  // "Jan 5, 2024", "January 5 2024"
  m = /^([A-Za-z]{3,9})\.?\s+(\d{1,2})(?:st|nd|rd|th)?,?\s+(\d{2}|\d{4})$/.exec(t);
  if (m) {
    const mon = MONTHS[m[1].slice(0, 4).toLowerCase()] ?? MONTHS[m[1].slice(0, 3).toLowerCase()];
    return mon ? toExcelSerial(+m[3], mon, +m[2]) : null;
  }
  return null;
}

/** Converts "1,234.50" / "-12" / "45%" to numbers; leaves IDs with leading zeros or >15 digits alone. */
export function parseNumberText(s: string): { value: number; numberFormat?: string } | null {
  const t = s.trim();
  if (!t) return null;
  const digits = t.replace(/\D/g, "");
  if (digits.length > 15) return null; // Excel would lose precision (IDs, card numbers…)
  if (/^[-+]?0\d/.test(t)) return null; // leading zeros = identifier (zip, product code)
  if (/^[-+]?(\d{1,3}(,\d{3})+|\d+)(\.\d+)?$/.test(t) || /^[-+]?\.\d+$/.test(t)) {
    const v = Number(t.replace(/,/g, ""));
    return Number.isFinite(v) ? { value: v } : null;
  }
  const pct = /^([-+]?\d+(\.\d+)?)\s?%$/.exec(t);
  if (pct) {
    const v = Number(pct[1]) / 100;
    return { value: v, numberFormat: pct[2] ? "0.00%" : "0%" };
  }
  return null;
}

/** Clean a single cell value (as returned by Range.formulas). Formulas are never touched. */
export function cleanCell(v: unknown, o: CleanOptions): CleanResult {
  if (typeof v !== "string" || v.length === 0 || v.startsWith("=")) return { value: v, changed: false };
  let s = v;
  if (o.removeNonPrintable) s = s.replace(NON_PRINTABLE, "").replace(/ /g, " ");
  if (o.trimWhitespace) {
    s = s
      .split(/\r?\n/)
      .map((line) => line.replace(/[ \t]+/g, " ").trim())
      .join("\n")
      .trim();
  }
  if (o.caseMode !== "none" && !looksLikeEmailOrUrl(s)) {
    if (o.caseMode === "upper") s = s.toUpperCase();
    else if (o.caseMode === "lower") s = s.toLowerCase();
    else if (o.caseMode === "proper") s = toProperCase(s);
    else if (o.caseMode === "sentence") s = toSentenceCase(s);
  }
  if (o.dateOrder !== "off") {
    const serial = parseDateText(s, o.dateOrder);
    if (serial !== null) return { value: serial, numberFormat: "yyyy-mm-dd", changed: true };
  }
  if (o.convertNumbers) {
    const n = parseNumberText(s);
    if (n) return { value: n.value, numberFormat: n.numberFormat, changed: true };
  }
  return { value: s, changed: s !== v };
}

/* ---------------------------------------------------------------- empty rows / columns */

const NULL_WORDS = new Set(["null", "none", "nil", "nan", "n/a", "#n/a", "na"]);

/** True for a cell that carries no information: blank, spaces, 0, "0", "0.00" or a null word ("NULL", "N/A", "NaN"…). */
export function isEmptyLike(v: unknown): boolean {
  if (v === null || v === undefined) return true;
  if (typeof v === "number") return v === 0 || Number.isNaN(v);
  if (typeof v === "boolean") return false;
  const s = String(v).replace(/[\s ]+/g, "");
  if (s === "") return true;
  if (NULL_WORDS.has(s.toLowerCase())) return true;
  return /^[-+]?0*(\.0*)?$/.test(s) && /0/.test(s);
}

function isHeaderText(v: unknown): boolean {
  return typeof v === "string" && !isEmptyLike(v) && Number.isNaN(Number(v.trim()));
}

/**
 * Row indexes (0-based, ascending) whose every cell is empty-like.
 * A row with even one real value is kept.
 */
export function findEmptyRows(values: unknown[][]): number[] {
  const out: number[] = [];
  values.forEach((row, i) => {
    if (row.every(isEmptyLike)) out.push(i);
  });
  return out;
}

/**
 * Column indexes (0-based, ascending) whose every cell is empty-like.
 * When the first row looks like a header row (has text), its header cell is ignored,
 * so a column titled "Gas rate" that only holds zeros is still removed.
 */
export function findEmptyColumns(values: unknown[][]): number[] {
  if (values.length === 0) return [];
  const cols = values[0].length;
  const hasHeader = values.length > 1 && values[0].some(isHeaderText);
  const start = hasHeader ? 1 : 0;
  const out: number[] = [];
  for (let j = 0; j < cols; j++) {
    let empty = true;
    for (let i = start; i < values.length && empty; i++) if (!isEmptyLike(values[i][j])) empty = false;
    if (empty) out.push(j);
  }
  return out;
}

/** Group ascending indexes into contiguous runs, returned last-first so they can be deleted safely. */
export function runsDescending(indexes: number[]): { start: number; count: number }[] {
  const runs: { start: number; count: number }[] = [];
  for (const i of indexes) {
    const last = runs[runs.length - 1];
    if (last && last.start + last.count === i) last.count++;
    else runs.push({ start: i, count: 1 });
  }
  return runs.reverse();
}
