/**
 * Merge well data from several tables by Well (+ Date). Pure – unit tested.
 *
 *  - Well names are matched loosely: case, spaces, "-" and "_" are ignored ("W-01" = "w 01" = "W_01").
 *  - Dates can be Excel serials or date texts (day-first unless the data proves month-first).
 *  - "exact": rows join on the same well and the same day.
 *  - "asof":  every row of the FIRST table gets the most recent row of each other table for the same well
 *             on or before its date (e.g. the latest well test or ESP reading for each production day).
 */
import { detectDateOrder, latinDigits, parseLooseDate, type DateOrder } from "./importFix";

export interface MergeTable {
  name: string;
  headers: string[];
  rows: unknown[][];
  wellCol: number;
  /** -1 when the table has no date (e.g. a well master list) – it then joins on well only. */
  dateCol: number;
}

export type MergeHow = "outer" | "left";
export type DateMatch = "exact" | "asof";

export interface MergeOptions {
  how: MergeHow;
  dateMatch: DateMatch;
  /** For "asof": ignore matches older than this many days (0 = no limit). */
  maxAgeDays?: number;
}

export interface MergeResult {
  headers: string[];
  rows: unknown[][];
  /** Index of the date column in the output (-1 when no table has dates). */
  dateCol: number;
  stats: {
    perTable: Array<{ name: string; rows: number; matched: number; duplicateKeys: number; badKeys: number }>;
    outputRows: number;
    wells: number;
  };
}

export function wellKey(v: unknown): string {
  return latinDigits(String(v ?? ""))
    .trim()
    .toUpperCase()
    .replace(/[\s\-_.]+/g, "");
}

function toDate(v: unknown, order: DateOrder): number | null {
  if (typeof v === "number" && Number.isFinite(v)) return v > 1000 && v < 200000 ? Math.floor(v) : null;
  if (typeof v === "string" && v.trim()) return parseLooseDate(v, order);
  return null;
}

interface Keyed {
  well: string;
  wellShown: string;
  date: number | null;
  row: unknown[];
}

function keyRows(t: MergeTable): { keyed: Keyed[]; bad: number } {
  const order = t.dateCol >= 0 ? detectDateOrder(t.rows.map((r) => r[t.dateCol]).filter((v): v is string => typeof v === "string")) : "DMY";
  const keyed: Keyed[] = [];
  let bad = 0;
  for (const row of t.rows) {
    const well = wellKey(row[t.wellCol]);
    const date = t.dateCol >= 0 ? toDate(row[t.dateCol], order) : null;
    if (!well || (t.dateCol >= 0 && date === null)) {
      if (row.some((v) => v !== "" && v !== null && v !== undefined)) bad++;
      continue;
    }
    keyed.push({ well, wellShown: String(row[t.wellCol]).trim(), date, row });
  }
  return { keyed, bad };
}

/** Output column names: data columns of every table, prefixed with the table name when names clash. */
export function mergedHeaders(tables: MergeTable[]): { headers: string[]; cols: number[][] } {
  const cols = tables.map((t) => t.headers.map((_, j) => j).filter((j) => j !== t.wellCol && j !== t.dateCol));
  const count = new Map<string, number>();
  tables.forEach((t, i) => cols[i].forEach((j) => count.set(t.headers[j].toLowerCase(), (count.get(t.headers[j].toLowerCase()) ?? 0) + 1)));
  const headers = tables.flatMap((t, i) => cols[i].map((j) => ((count.get(t.headers[j].toLowerCase()) ?? 0) > 1 ? `${t.name}: ${t.headers[j]}` : t.headers[j])));
  return { headers, cols };
}

export function mergeTables(tables: MergeTable[], o: MergeOptions): MergeResult {
  if (tables.length < 2) throw new Error("Pick at least two tables to merge.");
  const anyDate = tables.some((t) => t.dateCol >= 0);
  if (anyDate && tables[0].dateCol < 0 && (o.how === "left" || o.dateMatch === "asof"))
    throw new Error(`The first table ("${tables[0].name}") needs a date column for this kind of merge – put a dated table (e.g. daily production) first.`);
  const { headers: dataHeaders, cols } = mergedHeaders(tables);
  const keyedAll = tables.map(keyRows);
  const shownName = new Map<string, string>();
  keyedAll.forEach((k) => k.keyed.forEach((r) => !shownName.has(r.well) && shownName.set(r.well, r.wellShown)));

  // Index each table: well → date → row (first wins; later duplicates counted).
  const dupes = tables.map(() => 0);
  const index = keyedAll.map((k, i) => {
    const m = new Map<string, Map<number, unknown[]>>();
    for (const r of k.keyed) {
      const d = r.date ?? -1;
      const byDate = m.get(r.well) ?? new Map<number, unknown[]>();
      if (byDate.has(d)) dupes[i]++;
      else byDate.set(d, r.row);
      m.set(r.well, byDate);
    }
    return m;
  });
  // For as-of lookups: sorted dates per well.
  const sortedDates = index.map((m) => new Map(Array.from(m.entries()).map(([w, byDate]) => [w, Array.from(byDate.keys()).sort((a, b) => a - b)])));

  // Output keys.
  const keys: Array<{ well: string; date: number }> = [];
  const seen = new Set<string>();
  const addKey = (well: string, date: number) => {
    const k = `${well}|${date}`;
    if (seen.has(k)) return;
    seen.add(k);
    keys.push({ well, date });
  };
  const baseTables = o.how === "left" || o.dateMatch === "asof" ? [0] : tables.map((_, i) => i).filter((i) => tables[i].dateCol >= 0 || !anyDate);
  for (const i of baseTables) for (const [well, byDate] of index[i]) for (const d of byDate.keys()) addKey(well, d);
  keys.sort((a, b) => (a.well < b.well ? -1 : a.well > b.well ? 1 : a.date - b.date));

  const matched = tables.map(() => 0);
  const lookup = (i: number, well: string, date: number): unknown[] | undefined => {
    const t = tables[i];
    const byDate = index[i].get(well);
    if (!byDate) return undefined;
    if (t.dateCol < 0) return byDate.get(-1);
    if (o.dateMatch === "exact" || i === 0) return byDate.get(date);
    const ds = sortedDates[i].get(well) ?? [];
    let lo = 0;
    let hi = ds.length - 1;
    let best = -1;
    while (lo <= hi) {
      const mid = (lo + hi) >> 1;
      if (ds[mid] <= date) {
        best = mid;
        lo = mid + 1;
      } else hi = mid - 1;
    }
    if (best < 0) return undefined;
    if (o.maxAgeDays && date - ds[best] > o.maxAgeDays) return undefined;
    return byDate.get(ds[best]);
  };

  const rows: unknown[][] = keys.map(({ well, date }) => {
    const out: unknown[] = [shownName.get(well) ?? well];
    if (anyDate) out.push(date >= 0 ? date : "");
    tables.forEach((_, i) => {
      const r = lookup(i, well, date);
      if (r) matched[i]++;
      for (const j of cols[i]) out.push(r ? r[j] ?? "" : "");
    });
    return out;
  });

  return {
    headers: ["Well", ...(anyDate ? ["Date"] : []), ...dataHeaders],
    rows,
    dateCol: anyDate ? 1 : -1,
    stats: {
      perTable: tables.map((t, i) => ({ name: t.name, rows: t.rows.length, matched: matched[i], duplicateKeys: dupes[i], badKeys: keyedAll[i].bad })),
      outputRows: rows.length,
      wells: new Set(keys.map((k) => k.well)).size,
    },
  };
}

/** Guess the well and date columns from header names. */
export function guessKeys(headers: string[]): { wellCol: number; dateCol: number } {
  const find = (re: RegExp) => headers.findIndex((h) => re.test(h));
  let wellCol = find(/^\s*(well|well name|well ?id|uwi|بئر)\s*$/i);
  if (wellCol < 0) wellCol = find(/well|uwi|بئر/i);
  let dateCol = find(/^\s*(date|day|test date|production date|تاريخ)\s*$/i);
  if (dateCol < 0) dateCol = find(/date|day|time|تاريخ/i);
  return { wellCol, dateCol };
}

/**
 * Header row of a sheet's used range: the first of the top 20 rows that is about as wide as the data and
 * mostly text (skips title lines such as "Daily production – Oct 2026").
 */
export function findHeaderRow(values: unknown[][]): number {
  const filled = (r: unknown[]) => r.filter((v) => v !== "" && v !== null && v !== undefined).length;
  const width = Math.max(0, ...values.slice(0, 50).map(filled));
  for (let i = 0; i < Math.min(values.length, 20); i++) {
    const r = values[i];
    const n = filled(r);
    if (n < Math.max(2, width * 0.6)) continue;
    const text = r.filter((v) => typeof v === "string" && v.trim() !== "" && !/^-?\d+([.,]\d+)?$/.test(v.trim())).length;
    if (text >= n * 0.6) return i;
  }
  return 0;
}
