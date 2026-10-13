/**
 * Gap filler for time series: adds the missing dates (per well when a well column is given) and fills
 * blank values. Every filled cell is reported so it can be marked. Pure – unit tested.
 */
import { toSerial } from "./dailyReport";

export type GapMethod = "linear" | "previous" | "blank";

export interface GapOptions {
  dateCol: number;
  groupCol?: number;
  method: GapMethod;
  /** Gaps longer than this many steps are left alone (0 = no limit). */
  maxGapSteps: number;
}

export type Step = { kind: "day"; days: number } | { kind: "month" };

const DAY_MS = 86_400_000;
const EPOCH = Date.UTC(1899, 11, 30);

export function addMonths(serial: number, n: number): number {
  const d = new Date(EPOCH + serial * DAY_MS);
  const day = d.getUTCDate();
  const t = Date.UTC(d.getUTCFullYear(), d.getUTCMonth() + n, 1);
  const last = new Date(Date.UTC(d.getUTCFullYear(), d.getUTCMonth() + n + 1, 0)).getUTCDate();
  return Math.round((t + (Math.min(day, last) - 1) * DAY_MS - EPOCH) / DAY_MS);
}

/** Most common spacing between consecutive dates: daily, weekly… or monthly (28–31 days). */
export function detectStep(serials: number[]): Step {
  const s = Array.from(new Set(serials)).sort((a, b) => a - b);
  const counts = new Map<number, number>();
  for (let i = 1; i < s.length; i++) {
    const d = s[i] - s[i - 1];
    counts.set(d, (counts.get(d) ?? 0) + 1);
  }
  if (!counts.size) return { kind: "day", days: 1 };
  const monthly = Array.from(counts.entries()).filter(([d]) => d >= 28 && d <= 31).reduce((n, [, c]) => n + c, 0);
  const [best, bestN] = Array.from(counts.entries()).sort((a, b) => b[1] - a[1] || a[0] - b[0])[0];
  if (monthly > bestN || (best >= 28 && best <= 31)) return { kind: "month" };
  return { kind: "day", days: Math.max(1, best) };
}

const next = (d: number, step: Step) => (step.kind === "month" ? addMonths(d, 1) : d + step.days);

/** Dates strictly between a and b on the step grid. */
export function missingBetween(a: number, b: number, step: Step): number[] {
  const out: number[] = [];
  for (let d = next(a, step); d < b && out.length < 100_000; d = next(d, step)) out.push(d);
  return out;
}

export interface GapResult {
  rows: unknown[][];
  /** Per output row: true when the row was added. */
  added: boolean[];
  /** Per output row: column indexes whose value was filled in. */
  filledCols: number[][];
  stats: { addedRows: number; filledCells: number; groups: number; skippedGaps: number; step: Step };
}

const isNum = (v: unknown) => typeof v === "number" && Number.isFinite(v);
const isBlank = (v: unknown) => v === null || v === undefined || (typeof v === "string" && v.trim() === "");

export function fillGaps(rows: unknown[][], o: GapOptions): GapResult {
  const width = Math.max(0, ...rows.map((r) => r.length));
  const valued = rows.map((r) => ({ r, d: toSerial(r[o.dateCol]) })).filter((x): x is { r: unknown[]; d: number } => x.d !== null);
  // Value columns = mostly numeric (excluding date / group).
  const numericCols = Array.from({ length: width }, (_, j) => j).filter((j) => {
    if (j === o.dateCol || j === o.groupCol) return false;
    const vals = valued.map((x) => x.r[j]).filter((v) => !isBlank(v));
    return vals.length > 0 && vals.filter(isNum).length >= vals.length * 0.8;
  });
  const step = detectStep(valued.map((x) => x.d));
  const groups = new Map<string, Array<{ r: unknown[]; d: number }>>();
  for (const x of valued) {
    const g = o.groupCol !== undefined ? String(x.r[o.groupCol] ?? "").trim() : "";
    groups.set(g, [...(groups.get(g) ?? []), x]);
  }

  const out: unknown[][] = [];
  const added: boolean[] = [];
  const filledCols: number[][] = [];
  let skippedGaps = 0;
  let filledCells = 0;
  let addedRows = 0;

  for (const [g, list] of groups) {
    list.sort((a, b) => a.d - b.d);
    const series: Array<{ row: unknown[]; d: number; added: boolean; filled: number[] }> = [];
    list.forEach((x, k) => {
      series.push({ row: [...x.r, ...Array(width - x.r.length).fill("")], d: x.d, added: false, filled: [] });
      const nxt = list[k + 1];
      if (!nxt || nxt.d === x.d) return;
      const miss = missingBetween(x.d, nxt.d, step);
      if (!miss.length) return;
      if (o.maxGapSteps > 0 && miss.length > o.maxGapSteps) {
        skippedGaps++;
        return;
      }
      for (const d of miss) {
        const row: unknown[] = Array(width).fill("");
        row[o.dateCol] = d;
        if (o.groupCol !== undefined) row[o.groupCol] = g;
        if (o.method === "previous") for (let j = 0; j < width; j++) if (j !== o.dateCol && j !== o.groupCol && !numericCols.includes(j)) row[j] = x.r[j] ?? "";
        series.push({ row, d, added: true, filled: [] });
      }
    });
    if (o.method !== "blank") {
      for (const j of numericCols) {
        for (let k = 0; k < series.length; k++) {
          const s = series[k];
          if (!isBlank(s.row[j])) continue;
          let a = k - 1;
          while (a >= 0 && !isNum(series[a].row[j])) a--;
          if (a < 0) continue;
          let value: number | null = null;
          if (o.method === "previous") value = series[a].row[j] as number;
          else {
            let b = k + 1;
            while (b < series.length && !isNum(series[b].row[j])) b++;
            if (b >= series.length) continue;
            const va = series[a].row[j] as number;
            const vb = series[b].row[j] as number;
            const f = (s.d - series[a].d) / (series[b].d - series[a].d || 1);
            value = Math.round((va + (vb - va) * f) * 1e6) / 1e6;
          }
          s.row[j] = value;
          s.filled.push(j);
          filledCells++;
        }
      }
    }
    for (const s of series) {
      out.push(s.row);
      added.push(s.added);
      filledCols.push(s.filled);
      if (s.added) addedRows++;
    }
  }
  return { rows: out, added, filledCols, stats: { addedRows, filledCells, groups: groups.size, skippedGaps, step } };
}
