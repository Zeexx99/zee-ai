/**
 * Duplicate rows, spelling variants ("W-12" / "W12" / "Well 12") and outliers. Pure – unit tested.
 */

const blank = (v: unknown) => v === null || v === undefined || (typeof v === "string" && v.trim() === "");

const cellKey = (v: unknown) => (typeof v === "number" ? String(v) : String(v ?? "").trim().replace(/\s+/g, " ").toLowerCase());

export interface DuplicateGroup {
  /** 0-based data-row index kept (first occurrence). */
  keep: number;
  /** Later rows identical to it. */
  dupes: number[];
}

/** Rows identical in `cols` (all columns by default). Fully blank rows are ignored. */
export function findDuplicateRows(rows: unknown[][], cols?: number[]): DuplicateGroup[] {
  const seen = new Map<string, DuplicateGroup>();
  const out: DuplicateGroup[] = [];
  rows.forEach((r, i) => {
    const use = cols ?? r.map((_, j) => j);
    if (use.every((j) => blank(r[j]))) return;
    const key = use.map((j) => cellKey(r[j])).join("␟");
    const g = seen.get(key);
    if (g) g.dupes.push(i);
    else {
      const ng = { keep: i, dupes: [] as number[] };
      seen.set(key, ng);
      out.push(ng);
    }
  });
  return out.filter((g) => g.dupes.length);
}

/** Loose key for names: case, spaces, dashes, leading zeros and a "Well"/"W" prefix don't matter. */
export function variantKey(s: string): string {
  return s
    .replace(/[٠-٩]/g, (d) => String(d.charCodeAt(0) - 0x0660))
    .toLowerCase()
    .replace(/[\s\-_./#]+/g, "")
    .replace(/^(well|w)(?=\d)/, "w")
    .replace(/(^|\D)0+(\d)/g, "$1$2");
}

export interface VariantGroup {
  /** Most common spelling – the suggested one. */
  canonical: string;
  variants: Array<{ value: string; n: number; rows: number[] }>;
}

/** Text values in a column that differ only in spelling. Groups with a single spelling are not returned. */
export function findSpellingVariants(rows: unknown[][], col: number): VariantGroup[] {
  const groups = new Map<string, Map<string, number[]>>();
  rows.forEach((r, i) => {
    const v = r[col];
    if (typeof v !== "string" || !v.trim()) return;
    const k = variantKey(v.trim());
    if (!k) return;
    const g = groups.get(k) ?? new Map<string, number[]>();
    groups.set(k, g);
    const t = v.trim();
    g.set(t, [...(g.get(t) ?? []), i]);
  });
  const out: VariantGroup[] = [];
  for (const g of groups.values()) {
    if (g.size < 2) continue;
    const variants = Array.from(g.entries())
      .map(([value, rs]) => ({ value, n: rs.length, rows: rs }))
      .sort((a, b) => b.n - a.n || a.rows[0] - b.rows[0]);
    out.push({ canonical: variants[0].value, variants });
  }
  return out.sort((a, b) => b.variants.length - a.variants.length);
}

export interface Outlier {
  row: number;
  value: number;
  /** Robust z-score (median / MAD). */
  score: number;
  direction: "high" | "low";
}

const median = (xs: number[]) => {
  const s = [...xs].sort((a, b) => a - b);
  const m = s.length >> 1;
  return s.length % 2 ? s[m] : (s[m - 1] + s[m]) / 2;
};

/**
 * Outliers in a numeric column by the robust z-score 0.6745·(x − median)/MAD (|z| > `threshold`, 3.5 by
 * default). Needs at least 8 numbers; columns where most values are equal (MAD = 0) are skipped.
 */
export function findOutliers(rows: unknown[][], col: number, threshold = 3.5): Outlier[] {
  const nums: Array<{ row: number; value: number }> = [];
  rows.forEach((r, i) => {
    const v = r[col];
    if (typeof v === "number" && Number.isFinite(v)) nums.push({ row: i, value: v });
  });
  if (nums.length < 8) return [];
  const med = median(nums.map((x) => x.value));
  const mad = median(nums.map((x) => Math.abs(x.value - med)));
  if (!(mad > 0)) return [];
  return nums
    .map((x) => ({ row: x.row, value: x.value, score: (0.6745 * (x.value - med)) / mad }))
    .filter((x) => Math.abs(x.score) > threshold)
    .map((x) => ({ ...x, score: Math.round(x.score * 10) / 10, direction: x.score > 0 ? ("high" as const) : ("low" as const) }));
}
