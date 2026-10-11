/**
 * Pure petroleum-engineering helpers (no Office dependency, unit-tested):
 *  - Arps decline-curve fitting (exponential, harmonic, hyperbolic) + forecast/EUR
 *  - Production data QC
 *  - Oilfield unit conversions
 */

/* ================================================================ Arps DCA */

export type ArpsModel = "exponential" | "harmonic" | "hyperbolic";

export interface ArpsFit {
  model: ArpsModel;
  qi: number; // rate at t = 0
  di: number; // nominal initial decline, per time unit of t
  b: number; // 0 exponential, 1 harmonic, (0,1) hyperbolic (fit allows up to 1.5)
  r2: number;
  rmse: number;
}

/** Arps rate q(t). */
export function arpsRate(f: Pick<ArpsFit, "qi" | "di" | "b">, t: number): number {
  if (f.b < 1e-6) return f.qi * Math.exp(-f.di * t);
  return f.qi / Math.pow(1 + f.b * f.di * t, 1 / f.b);
}

/** Arps cumulative production Np(t) (same time unit as di; rate × time). */
export function arpsCum(f: Pick<ArpsFit, "qi" | "di" | "b">, t: number): number {
  const { qi, di, b } = f;
  if (t <= 0) return 0;
  if (b < 1e-6) return (qi / di) * (1 - Math.exp(-di * t));
  if (Math.abs(b - 1) < 1e-6) return (qi / di) * Math.log(1 + di * t);
  const q = arpsRate(f, t);
  return (Math.pow(qi, b) / ((1 - b) * di)) * (Math.pow(qi, 1 - b) - Math.pow(q, 1 - b));
}

/** Time at which q(t) falls to qLimit (Infinity if never / invalid). */
export function arpsTimeToRate(f: Pick<ArpsFit, "qi" | "di" | "b">, qLimit: number): number {
  const { qi, di, b } = f;
  if (!(qLimit > 0) || qLimit >= qi || di <= 0) return qLimit >= qi ? 0 : Infinity;
  if (b < 1e-6) return Math.log(qi / qLimit) / di;
  return (Math.pow(qi / qLimit, b) - 1) / (b * di);
}

/** Minimal Nelder–Mead minimiser for small problems. */
function nelderMead(fn: (x: number[]) => number, x0: number[], step = 0.5, iters = 400): number[] {
  const n = x0.length;
  let simplex: number[][] = [x0.slice()];
  for (let i = 0; i < n; i++) {
    const p = x0.slice();
    p[i] += step;
    simplex.push(p);
  }
  let vals = simplex.map(fn);
  for (let k = 0; k < iters; k++) {
    const order = vals.map((v, i) => [v, i] as const).sort((a, b) => a[0] - b[0]);
    simplex = order.map(([, i]) => simplex[i]);
    vals = order.map(([v]) => v);
    if (Math.abs(vals[n] - vals[0]) < 1e-12 * (1 + Math.abs(vals[0]))) break;
    const centroid = Array.from({ length: n }, (_, j) => simplex.slice(0, n).reduce((s, p) => s + p[j], 0) / n);
    const worst = simplex[n];
    const at = (c: number) => centroid.map((cj, j) => cj + c * (worst[j] - cj));
    const xr = at(-1);
    const fr = fn(xr);
    if (fr < vals[0]) {
      const xe = at(-2);
      const fe = fn(xe);
      [simplex[n], vals[n]] = fe < fr ? [xe, fe] : [xr, fr];
    } else if (fr < vals[n - 1]) {
      [simplex[n], vals[n]] = [xr, fr];
    } else {
      const xc = at(0.5);
      const fc = fn(xc);
      if (fc < vals[n]) {
        [simplex[n], vals[n]] = [xc, fc];
      } else {
        for (let i = 1; i <= n; i++) {
          simplex[i] = simplex[i].map((v, j) => simplex[0][j] + 0.5 * (v - simplex[0][j]));
          vals[i] = fn(simplex[i]);
        }
      }
    }
  }
  const best = vals.indexOf(Math.min(...vals));
  return simplex[best];
}

function stats(t: number[], q: number[], f: Pick<ArpsFit, "qi" | "di" | "b">) {
  const mean = q.reduce((a, b) => a + b, 0) / q.length;
  let ssr = 0;
  let sst = 0;
  for (let i = 0; i < q.length; i++) {
    const e = q[i] - arpsRate(f, t[i]);
    ssr += e * e;
    sst += (q[i] - mean) ** 2;
  }
  return { r2: sst > 0 ? 1 - ssr / sst : 0, rmse: Math.sqrt(ssr / q.length), ssr };
}

/** Fit one Arps model for a fixed or free b. Data must have t ≥ 0, q > 0. */
function fitWithB(t: number[], q: number[], bMode: "fixed" | "free", bFixed: number): ArpsFit & { ssr: number } {
  // Initial guess from log-linear (exponential) regression.
  const lnq = q.map(Math.log);
  const n = t.length;
  const mt = t.reduce((a, b) => a + b, 0) / n;
  const ml = lnq.reduce((a, b) => a + b, 0) / n;
  let num = 0;
  let den = 0;
  for (let i = 0; i < n; i++) {
    num += (t[i] - mt) * (lnq[i] - ml);
    den += (t[i] - mt) ** 2;
  }
  const slope = den > 0 ? num / den : -0.01;
  const di0 = Math.max(1e-6, -slope || 1e-3);
  const qi0 = Math.exp(ml - slope * mt);

  const decode = (x: number[]) => ({
    qi: Math.exp(x[0]),
    di: Math.exp(x[1]),
    b: bMode === "free" ? 1.5 / (1 + Math.exp(-x[2])) : bFixed,
  });
  const obj = (x: number[]) => {
    const p = decode(x);
    let s = 0;
    for (let i = 0; i < n; i++) {
      const e = q[i] - arpsRate(p, t[i]);
      s += e * e;
    }
    return Number.isFinite(s) ? s : 1e300;
  };
  const x0 = bMode === "free" ? [Math.log(qi0), Math.log(di0), 0] : [Math.log(qi0), Math.log(di0)];
  const x = nelderMead(obj, x0, 0.3, 1500);
  const p = decode(x);
  const s = stats(t, q, p);
  const model: ArpsModel = p.b < 1e-6 ? "exponential" : Math.abs(p.b - 1) < 1e-6 ? "harmonic" : "hyperbolic";
  return { model, ...p, r2: s.r2, rmse: s.rmse, ssr: s.ssr };
}

export interface DcaInput {
  t: number[];
  q: number[];
}

/** Clean input: keep finite t, q > 0; shift t to start at 0. */
export function prepareDca(rawT: unknown[], rawQ: unknown[]): DcaInput & { t0: number; dropped: number } {
  const pts: Array<[number, number]> = [];
  let dropped = 0;
  for (let i = 0; i < Math.min(rawT.length, rawQ.length); i++) {
    const tv = Number(rawT[i]);
    const qv = Number(rawQ[i]);
    if (rawT[i] === "" || rawQ[i] === "" || !Number.isFinite(tv) || !Number.isFinite(qv) || qv <= 0) {
      dropped++;
      continue;
    }
    pts.push([tv, qv]);
  }
  pts.sort((a, b) => a[0] - b[0]);
  const t0 = pts.length ? pts[0][0] : 0;
  return { t: pts.map((p) => p[0] - t0), q: pts.map((p) => p[1]), t0, dropped };
}

/** Fit all three Arps models; returns them sorted best-first by RMSE. */
export function fitArps(input: DcaInput): ArpsFit[] {
  if (input.t.length < 3) throw new Error("Need at least 3 valid points (time, rate > 0) to fit a decline curve.");
  const fits = [fitWithB(input.t, input.q, "fixed", 0), fitWithB(input.t, input.q, "fixed", 1), fitWithB(input.t, input.q, "free", 0.5)];
  fits[0].model = "exponential";
  fits[1].model = "harmonic";
  fits[2].model = "hyperbolic";
  return fits.map(({ ssr: _ssr, ...f }) => f).sort((a, b) => a.rmse - b.rmse);
}

export interface GroupFit {
  group: string;
  points: number;
  dropped: number;
  /** Excel time of the first valid point (t = 0 for the fit). */
  t0: number;
  best: ArpsFit | null;
  error?: string;
}

/** Fit Arps curves separately for every group (e.g. well name). Groups keep first-seen order. */
export function fitArpsByGroup(keys: unknown[], rawT: unknown[], rawQ: unknown[]): GroupFit[] {
  const order: string[] = [];
  const buckets = new Map<string, { t: unknown[]; q: unknown[] }>();
  keys.forEach((k, i) => {
    const key = String(k ?? "").trim();
    if (!key) return;
    if (!buckets.has(key)) {
      buckets.set(key, { t: [], q: [] });
      order.push(key);
    }
    const b = buckets.get(key)!;
    b.t.push(rawT[i]);
    b.q.push(rawQ[i]);
  });
  return order.map((group) => {
    const b = buckets.get(group)!;
    const prep = prepareDca(b.t, b.q);
    try {
      return { group, points: prep.t.length, dropped: prep.dropped, t0: prep.t0, best: fitArps(prep)[0] };
    } catch (e) {
      return { group, points: prep.t.length, dropped: prep.dropped, t0: prep.t0, best: null, error: (e as Error).message };
    }
  });
}

/* ================================================================ QC */

export interface ColumnQc {
  name: string;
  count: number;
  missing: number;
  nonNumeric: number;
  negatives: number;
  zeros: number;
  outliers: number;
  flatRuns: number; // runs of ≥ 5 identical consecutive values
  min: number | null;
  max: number | null;
  mean: number | null;
  issues: string[]; // e.g. "row 14: negative (-3)"
}

export interface QcReport {
  columns: ColumnQc[];
  dateGaps: string[];
  duplicateDates: string[];
}

function quantile(sorted: number[], p: number): number {
  const i = (sorted.length - 1) * p;
  const lo = Math.floor(i);
  const hi = Math.ceil(i);
  return sorted[lo] + (sorted[hi] - sorted[lo]) * (i - lo);
}

/**
 * QC a block: first row = headers. If dateCol >= 0, that column is treated as dates
 * (Excel serials) and checked for gaps/duplicates; other columns are numeric checks.
 * rowOffset = sheet row number of the first DATA row (for messages).
 */
export function qcAnalyze(block: unknown[][], dateCol: number, rowOffset: number, maxIssuesPerColumn = 25): QcReport {
  const headers = (block[0] ?? []).map((h, j) => (String(h ?? "").trim() || `Column ${j + 1}`));
  const data = block.slice(1);
  const columns: ColumnQc[] = [];
  for (let j = 0; j < headers.length; j++) {
    if (j === dateCol) continue;
    const c: ColumnQc = { name: headers[j], count: data.length, missing: 0, nonNumeric: 0, negatives: 0, zeros: 0, outliers: 0, flatRuns: 0, min: null, max: null, mean: null, issues: [] };
    const nums: Array<[number, number]> = [];
    const push = (msg: string) => {
      if (c.issues.length < maxIssuesPerColumn) c.issues.push(msg);
    };
    data.forEach((row, i) => {
      const v = row[j];
      const r = rowOffset + i;
      if (v === "" || v === null || v === undefined) {
        c.missing++;
        push(`row ${r}: missing`);
        return;
      }
      const n = typeof v === "number" ? v : Number(String(v).replace(/,/g, ""));
      if (!Number.isFinite(n)) {
        c.nonNumeric++;
        push(`row ${r}: not a number ("${String(v).slice(0, 20)}")`);
        return;
      }
      if (n < 0) {
        c.negatives++;
        push(`row ${r}: negative (${n})`);
      }
      if (n === 0) c.zeros++;
      nums.push([n, r]);
    });
    if (nums.length) {
      const vals = nums.map((x) => x[0]);
      const sorted = vals.slice().sort((a, b) => a - b);
      c.min = sorted[0];
      c.max = sorted[sorted.length - 1];
      c.mean = vals.reduce((a, b) => a + b, 0) / vals.length;
      if (sorted.length >= 8) {
        const q1 = quantile(sorted, 0.25);
        const q3 = quantile(sorted, 0.75);
        const med = quantile(sorted, 0.5);
        // Fall back to 10% of the median when most values are identical (IQR = 0).
        const iqr = q3 - q1 > 0 ? q3 - q1 : Math.abs(med) * 0.1;
        if (iqr > 0) {
          const lo = q1 - 3 * iqr;
          const hi = q3 + 3 * iqr;
          nums.forEach(([n, r]) => {
            if (n < lo || n > hi) {
              c.outliers++;
              push(`row ${r}: outlier (${n})`);
            }
          });
        }
      }
      let run = 1;
      for (let k = 1; k <= nums.length; k++) {
        if (k < nums.length && nums[k][0] === nums[k - 1][0] && nums[k][0] !== 0) run++;
        else {
          if (run >= 5) {
            c.flatRuns++;
            push(`rows ${nums[k - run][1]}–${nums[k - 1][1]}: ${run} identical values (${nums[k - 1][0]}) – stuck gauge?`);
          }
          run = 1;
        }
      }
    }
    columns.push(c);
  }

  const dateGaps: string[] = [];
  const duplicateDates: string[] = [];
  if (dateCol >= 0) {
    const dates = data
      .map((row, i) => ({ d: Number(row[dateCol]), r: rowOffset + i }))
      .filter((x) => Number.isFinite(x.d) && x.d > 0);
    const steps: number[] = [];
    for (let k = 1; k < dates.length; k++) steps.push(Math.round(dates[k].d - dates[k - 1].d));
    const positive = steps.filter((s) => s > 0).sort((a, b) => a - b);
    const typical = positive.length ? positive[Math.floor(positive.length / 2)] : 1;
    for (let k = 1; k < dates.length; k++) {
      const s = steps[k - 1];
      if (s === 0 && duplicateDates.length < 50) duplicateDates.push(`row ${dates[k].r}: same date as row ${dates[k - 1].r}`);
      else if (s < 0 && duplicateDates.length < 50) duplicateDates.push(`row ${dates[k].r}: date goes backwards`);
      else if (s > typical * 1.5 && dateGaps.length < 50) dateGaps.push(`rows ${dates[k - 1].r}→${dates[k].r}: gap of ${s} days (typical step ${typical})`);
    }
  }
  return { columns, dateGaps, duplicateDates };
}

/* ================================================================ Units */

export interface UnitConversion {
  id: string;
  label: string;
  group: string;
  convert: (v: number) => number;
  suffix: string;
}

const lin = (k: number) => (v: number) => v * k;

export const UNIT_CONVERSIONS: UnitConversion[] = [
  { id: "bbl_m3", label: "bbl → m³ (also bbl/d → m³/d)", group: "Liquid volume / rate", convert: lin(0.158987294928), suffix: "m³" },
  { id: "m3_bbl", label: "m³ → bbl (also m³/d → bbl/d)", group: "Liquid volume / rate", convert: lin(1 / 0.158987294928), suffix: "bbl" },
  { id: "scf_m3", label: "scf → m³ (std)", group: "Gas volume / rate", convert: lin(0.028316846592), suffix: "m³" },
  { id: "mscf_m3", label: "Mscf/d → m³/d", group: "Gas volume / rate", convert: lin(28.316846592), suffix: "m³/d" },
  { id: "mmscf_mm3", label: "MMscf/d → 10³ m³/d", group: "Gas volume / rate", convert: lin(28.316846592), suffix: "10³ m³/d" },
  { id: "m3_scf", label: "m³ → scf", group: "Gas volume / rate", convert: lin(1 / 0.028316846592), suffix: "scf" },
  { id: "psi_bar", label: "psi → bar", group: "Pressure", convert: lin(0.0689475729), suffix: "bar" },
  { id: "bar_psi", label: "bar → psi", group: "Pressure", convert: lin(1 / 0.0689475729), suffix: "psi" },
  { id: "psi_kpa", label: "psi → kPa", group: "Pressure", convert: lin(6.89475729), suffix: "kPa" },
  { id: "kpa_psi", label: "kPa → psi", group: "Pressure", convert: lin(1 / 6.89475729), suffix: "psi" },
  { id: "f_c", label: "°F → °C", group: "Temperature", convert: (v) => ((v - 32) * 5) / 9, suffix: "°C" },
  { id: "c_f", label: "°C → °F", group: "Temperature", convert: (v) => (v * 9) / 5 + 32, suffix: "°F" },
  { id: "ft_m", label: "ft → m", group: "Length / depth", convert: lin(0.3048), suffix: "m" },
  { id: "m_ft", label: "m → ft", group: "Length / depth", convert: lin(1 / 0.3048), suffix: "ft" },
  { id: "api_sg", label: "°API → specific gravity", group: "Density", convert: (v) => 141.5 / (v + 131.5), suffix: "SG" },
  { id: "sg_api", label: "specific gravity → °API", group: "Density", convert: (v) => 141.5 / v - 131.5, suffix: "°API" },
  { id: "psift_sg", label: "pressure gradient psi/ft → SG", group: "Density", convert: lin(1 / 0.433), suffix: "SG" },
  { id: "hp_kw", label: "hp → kW", group: "Power", convert: lin(0.745699872), suffix: "kW" },
  { id: "kw_hp", label: "kW → hp", group: "Power", convert: lin(1 / 0.745699872), suffix: "hp" },
];

export function convertValue(v: unknown, conv: UnitConversion): number | "" {
  if (v === "" || v === null || v === undefined) return "";
  const n = typeof v === "number" ? v : Number(String(v).replace(/,/g, ""));
  if (!Number.isFinite(n)) return "";
  const r = conv.convert(n);
  return Number.isFinite(r) ? Math.round(r * 1e6) / 1e6 : "";
}
