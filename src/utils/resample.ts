/**
 * Resample time series to day / week / month / quarter / year (per well when a well column is given).
 * Each value column gets its own aggregation. Optionally inserts empty periods and fills them.
 * Pure – unit tested. (Filling gaps WITHOUT changing the period lives in gaps.ts.)
 */
import { toSerial } from "./dailyReport";

export type Period = "day" | "week" | "month" | "quarter" | "year";
export type Agg = "avg" | "sum" | "last" | "first" | "min" | "max" | "count" | "skip";
export type PeriodFill = "none" | "blank" | "previous" | "linear";

export const PERIOD_LABELS: Record<Period, string> = { day: "Daily", week: "Weekly (Mon–Sun)", month: "Monthly", quarter: "Quarterly", year: "Yearly" };
export const AGG_LABELS: Record<Agg, string> = {
  avg: "Average",
  sum: "Sum",
  last: "Last value",
  first: "First value",
  min: "Minimum",
  max: "Maximum",
  count: "Count",
  skip: "Leave out",
};
export const FILL_LABELS: Record<PeriodFill, string> = {
  none: "Don't add missing periods",
  blank: "Add missing periods, leave empty",
  previous: "Add missing periods, repeat previous",
  linear: "Add missing periods, interpolate",
};

export interface ResampleOptions {
  dateCol: number;
  groupCol?: number;
  period: Period;
  /** Aggregation per column index (date/group ignored). Missing → "avg" for numbers, "last" for text. */
  aggs: Record<number, Agg>;
  fill: PeriodFill;
  /** Add a "Days with data" column. */
  countDays: boolean;
}

export interface ResampleResult {
  headers: string[];
  rows: unknown[][];
  /** Output rows that were inserted for missing periods. */
  added: boolean[];
  stats: { inputRows: number; outputRows: number; groups: number; skippedRows: number; addedPeriods: number };
}

const DAY_MS = 86_400_000;
const EPOCH = Date.UTC(1899, 11, 30);
const toDate = (s: number) => new Date(EPOCH + s * DAY_MS);
const toSer = (y: number, m: number, d: number) => Math.round((Date.UTC(y, m, d) - EPOCH) / DAY_MS);

/** Start of the period that contains serial `s`. */
export function periodStart(s: number, p: Period): number {
  const day = Math.floor(s);
  if (p === "day") return day;
  if (p === "week") {
    const dow = (toDate(day).getUTCDay() + 6) % 7; // Monday = 0
    return day - dow;
  }
  const d = toDate(day);
  const y = d.getUTCFullYear();
  const m = d.getUTCMonth();
  if (p === "month") return toSer(y, m, 1);
  if (p === "quarter") return toSer(y, m - (m % 3), 1);
  return toSer(y, 0, 1);
}

export function nextPeriod(start: number, p: Period): number {
  if (p === "day") return start + 1;
  if (p === "week") return start + 7;
  const d = toDate(start);
  const add = p === "month" ? 1 : p === "quarter" ? 3 : 12;
  return toSer(d.getUTCFullYear(), d.getUTCMonth() + add, 1);
}

const isNum = (v: unknown): v is number => typeof v === "number" && Number.isFinite(v);
const isBlank = (v: unknown) => v === null || v === undefined || (typeof v === "string" && v.trim() === "");

/** Default aggregation from the header and the column's values. */
export function guessAgg(header: string, values: unknown[]): Agg {
  const h = header.toLowerCase();
  const nonBlank = values.filter((v) => !isBlank(v));
  const numeric = nonBlank.length > 0 && nonBlank.filter(isNum).length >= nonBlank.length * 0.8;
  if (!numeric) return "last";
  if (/\bcum|cumulative|total to date|np\b|gp\b|wp\b/.test(h)) return "last";
  if (/hours|hrs|downtime|uptime|volume|vol\b|produced|bbls?\b(?!\/)|\bm3\b(?!\/)|mscf\b(?!\/)|days on/.test(h) && !/\/\s*d|per day|rate|bpd|bopd|bwpd|\/day/.test(h)) return "sum";
  return "avg";
}

function aggregate(vals: unknown[], agg: Agg): unknown {
  const present = vals.filter((v) => !isBlank(v));
  if (agg === "count") return present.length;
  if (!present.length) return "";
  if (agg === "first") return present[0];
  if (agg === "last") return present[present.length - 1];
  const nums = present.filter(isNum);
  if (!nums.length) return present[present.length - 1];
  if (agg === "sum") return round(nums.reduce((a, b) => a + b, 0));
  if (agg === "avg") return round(nums.reduce((a, b) => a + b, 0) / nums.length);
  if (agg === "min") return Math.min(...nums);
  if (agg === "max") return Math.max(...nums);
  return "";
}
const round = (v: number) => Math.round(v * 1e6) / 1e6;

export function resample(headers: string[], rows: unknown[][], o: ResampleOptions): ResampleResult {
  const width = headers.length;
  const valueCols = Array.from({ length: width }, (_, j) => j).filter((j) => j !== o.dateCol && j !== o.groupCol);
  const aggOf = (j: number) => o.aggs[j] ?? guessAgg(headers[j], rows.map((r) => r[j]));
  const cols = valueCols.filter((j) => aggOf(j) !== "skip");
  const colAgg = cols.map(aggOf);

  const groups = new Map<string, Map<number, Array<{ d: number; r: unknown[] }>>>();
  let skipped = 0;
  for (const r of rows) {
    const d = toSerial(r[o.dateCol]);
    if (d === null) {
      if (r.some((v) => !isBlank(v))) skipped++;
      continue;
    }
    const g = o.groupCol !== undefined ? String(r[o.groupCol] ?? "").trim() : "";
    const ps = periodStart(d, o.period);
    const byP = groups.get(g) ?? new Map<number, Array<{ d: number; r: unknown[] }>>();
    const list = byP.get(ps) ?? [];
    list.push({ d, r });
    byP.set(ps, list);
    groups.set(g, byP);
  }

  const outHeaders = [...(o.groupCol !== undefined ? [headers[o.groupCol]] : []), `${PERIOD_LABELS[o.period].replace(/ \(.*\)$/, "")} (start)`, ...cols.map((j, k) => `${headers[j]} (${AGG_LABELS[colAgg[k]].toLowerCase()})`), ...(o.countDays ? ["Days with data"] : [])];
  const out: unknown[][] = [];
  const added: boolean[] = [];
  let addedPeriods = 0;

  const groupNames = Array.from(groups.keys()).sort((a, b) => a.localeCompare(b, undefined, { numeric: true }));
  for (const g of groupNames) {
    const byP = groups.get(g)!;
    const starts = Array.from(byP.keys()).sort((a, b) => a - b);
    const series: Array<{ start: number; values: unknown[]; days: number; added: boolean }> = [];
    for (let k = 0; k < starts.length; k++) {
      const list = byP.get(starts[k])!.sort((a, b) => a.d - b.d);
      series.push({ start: starts[k], values: cols.map((j, c) => aggregate(list.map((x) => x.r[j]), colAgg[c])), days: new Set(list.map((x) => Math.floor(x.d))).size, added: false });
      if (o.fill !== "none" && k + 1 < starts.length) {
        for (let p = nextPeriod(starts[k], o.period), guard = 0; p < starts[k + 1] && guard < 50_000; p = nextPeriod(p, o.period), guard++) {
          series.push({ start: p, values: cols.map(() => ""), days: 0, added: true });
          addedPeriods++;
        }
      }
    }
    if (o.fill === "previous" || o.fill === "linear") {
      cols.forEach((_, c) => {
        if (colAgg[c] === "count") return;
        for (let k = 0; k < series.length; k++) {
          if (!series[k].added) continue;
          let a = k - 1;
          while (a >= 0 && series[a].added) a--;
          if (a < 0) continue;
          const va = series[a].values[c];
          if (o.fill === "previous" || !isNum(va)) {
            series[k].values[c] = va;
            continue;
          }
          let b = k + 1;
          while (b < series.length && series[b].added) b++;
          const vb = b < series.length ? series[b].values[c] : undefined;
          series[k].values[c] = isNum(vb) ? round(va + ((vb - va) * (series[k].start - series[a].start)) / (series[b].start - series[a].start || 1)) : va;
        }
      });
    }
    for (const s of series) {
      out.push([...(o.groupCol !== undefined ? [g] : []), s.start, ...s.values, ...(o.countDays ? [s.days] : [])]);
      added.push(s.added);
    }
  }
  return { headers: outHeaders, rows: out, added, stats: { inputRows: rows.length, outputRows: out.length, groups: groups.size, skippedRows: skipped, addedPeriods } };
}
