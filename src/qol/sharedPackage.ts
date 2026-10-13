/**
 * Hand-over from the Excel add-in to the Word add-in (daily production report).
 * Both add-ins are served from the same site, so they normally share browser storage; the package is
 * also copied to the clipboard as text, so it can be pasted when storage isn't shared.
 */
import { readJson, writeJson } from "./store";

const KEY = "zeeai.shared.dailyReport.v1";

export interface DailyReportPackage {
  zeeDailyReport: 1;
  createdAt: string;
  /** Workbook sheet the report table was written to. */
  sheet: string;
  /** ISO dates. */
  reportDate: string;
  previousDate: string | null;
  /** "13 Oct 2026" */
  dateLabel: string;
  /** Header first, last row = totals. */
  table: Array<Array<string | number>>;
  highlights: string[];
  wells: number;
  missingWells: string[];
  /** PNG of the oil-by-well chart (base64, no data: prefix). */
  chartPng?: string;
}

export function isDailyReportPackage(v: unknown): v is DailyReportPackage {
  const p = v as Partial<DailyReportPackage> | null;
  return !!p && p.zeeDailyReport === 1 && typeof p.reportDate === "string" && Array.isArray(p.table) && Array.isArray(p.highlights);
}

export function saveDailyReportPackage(p: DailyReportPackage): string {
  writeJson(KEY, p);
  return JSON.stringify(p);
}

export function loadDailyReportPackage(): DailyReportPackage | null {
  const p = readJson<unknown>(KEY, null);
  return isDailyReportPackage(p) ? p : null;
}

/** Parses pasted text (the clipboard copy made by the Excel add-in). */
export function parseDailyReportPackage(text: string): DailyReportPackage {
  let v: unknown;
  try {
    v = JSON.parse(text.trim());
  } catch {
    throw new Error("That isn't a Zee AI report package. In Excel, run Petroleum → Daily production report, then paste here.");
  }
  if (!isDailyReportPackage(v)) throw new Error("That isn't a Zee AI report package. In Excel, run Petroleum → Daily production report, then paste here.");
  return v;
}
