/**
 * Deterministic ESP performance analysis (no AI): TDH, hydraulic HP, motor load, affinity-law
 * correction to the base frequency, head vs catalog curve, operating range and flags.
 * Pure functions – unit-tested.
 */
import { headOnCurve, hydraulicHp, liquidSg, pumpTdh, type CurvePoint } from "./petroFormulas";

export interface EspOptions {
  /** Base (catalog) frequency of the pump curve, Hz. */
  baseHz: number;
  /** Constant fluid SG, used when there is no water-cut column (or no API). */
  sg: number;
  /** Oil API – with a water-cut column gives a per-row mixture SG. */
  api?: number;
  /** Number of stages; with a per-stage curve converts to total head. */
  stages?: number;
  /** Catalog curve at baseHz: rate (bbl/d) vs head per stage (ft) – or total head if stages is empty. */
  curve?: CurvePoint[];
  /** Recommended operating range at baseHz (bbl/d). Defaults to the curve's end points when a curve is given. */
  minRate?: number;
  maxRate?: number;
  /** Motor nameplate current (A) for motor load %. */
  motorAmps?: number;
  /** Flag rows where intake pressure is below this (e.g. bubble point), psi. */
  minPip?: number;
  /** Head shortfall vs curve that counts as degradation, %. */
  degradationPct: number;
}

export interface EspRowInput {
  rate: unknown;
  pip: unknown;
  pd: unknown;
  hz: unknown;
  amps?: unknown;
  waterCut?: unknown;
}

export interface EspRow {
  rate: number | null;
  hz: number | null;
  pip: number | null;
  pd: number | null;
  dp: number | null;
  sg: number | null;
  tdh: number | null;
  rateBase: number | null; // rate scaled to baseHz
  headBase: number | null; // actual head per stage (or total) scaled to baseHz
  headCurve: number | null; // catalog head at rateBase
  headDeviationPct: number | null; // (actual − curve) / curve × 100
  hhp: number | null;
  motorLoadPct: number | null;
  range: "" | "In range" | "Downthrust" | "Upthrust";
  flags: string[];
}

export interface EspSummary {
  rows: number;
  validRows: number;
  downthrust: number;
  upthrust: number;
  inRange: number;
  avgHeadDeviationPct: number | null;
  latestHeadDeviationPct: number | null;
  avgMotorLoadPct: number | null;
  flaggedRows: number;
  flagCounts: Record<string, number>;
  status: string;
}

const num = (v: unknown): number | null => {
  if (v === "" || v === null || v === undefined) return null;
  const n = typeof v === "number" ? v : Number(String(v).replace(/,/g, "").replace(/%$/, ""));
  return Number.isFinite(n) ? n : null;
};
const round = (v: number | null, d = 2) => (v === null ? null : Math.round(v * 10 ** d) / 10 ** d);
const mean = (xs: number[]) => (xs.length ? xs.reduce((a, b) => a + b, 0) / xs.length : null);

export const FLAG = {
  noLift: "No lift (ΔP ≤ 0) – pump off, broken shaft or reversed rotation?",
  lowPip: "Intake pressure below minimum – gas interference risk",
  overload: "Motor overload (> 100 % of nameplate amps)",
  underload: "Motor underload (< 50 %) – gas, pump-off or broken shaft?",
  down: "Below recommended range (downthrust)",
  up: "Above recommended range (upthrust)",
  headLoss: "Head below catalog curve – wear, gas or plugging",
  offCurve: "Rate outside the catalog curve",
} as const;

export function analyzeEspRow(r: EspRowInput, o: EspOptions): EspRow {
  const rate = num(r.rate);
  const pip = num(r.pip);
  const pd = num(r.pd);
  const hz = num(r.hz);
  const amps = num(r.amps);
  let wc = num(r.waterCut);
  if (wc !== null && wc > 1) wc = wc / 100; // percent → fraction
  const flags: string[] = [];

  const sg = wc !== null && o.api !== undefined && wc >= 0 && wc <= 1 ? liquidSg(wc, o.api) : o.sg;
  const dp = pd !== null && pip !== null ? pd - pip : null;
  const tdh = dp !== null && sg > 0 ? pumpTdh(pd as number, pip as number, sg) : null;
  const hhp = rate !== null && dp !== null ? hydraulicHp(rate, dp) : null;
  const ratio = hz !== null && hz > 0 ? o.baseHz / hz : null;
  const rateBase = rate !== null && ratio !== null ? rate * ratio : null;
  const stages = o.stages && o.stages > 0 ? o.stages : 1;
  const headBase = tdh !== null && ratio !== null ? (tdh / stages) * ratio * ratio : null;

  let headCurve: number | null = null;
  let headDeviationPct: number | null = null;
  if (o.curve && o.curve.length >= 2 && rateBase !== null) {
    headCurve = headOnCurve(o.curve, rateBase);
    if (headCurve === null) flags.push(FLAG.offCurve);
    else if (headBase !== null && headCurve > 0) {
      headDeviationPct = ((headBase - headCurve) / headCurve) * 100;
      if (headDeviationPct < -Math.abs(o.degradationPct)) flags.push(FLAG.headLoss);
    }
  }

  const curveQs = o.curve?.map((p) => p.q).filter((q) => Number.isFinite(q)) ?? [];
  const minRate = o.minRate ?? (curveQs.length ? Math.min(...curveQs) : undefined);
  const maxRate = o.maxRate ?? (curveQs.length ? Math.max(...curveQs) : undefined);
  let range: EspRow["range"] = "";
  if (rateBase !== null && minRate !== undefined && maxRate !== undefined) {
    range = rateBase < minRate ? "Downthrust" : rateBase > maxRate ? "Upthrust" : "In range";
    if (range === "Downthrust") flags.push(FLAG.down);
    if (range === "Upthrust") flags.push(FLAG.up);
  }

  if (dp !== null && dp <= 0) flags.push(FLAG.noLift);
  if (pip !== null && o.minPip !== undefined && pip < o.minPip) flags.push(FLAG.lowPip);
  const motorLoadPct = amps !== null && o.motorAmps ? (amps / o.motorAmps) * 100 : null;
  if (motorLoadPct !== null && motorLoadPct > 100) flags.push(FLAG.overload);
  if (motorLoadPct !== null && motorLoadPct > 0 && motorLoadPct < 50) flags.push(FLAG.underload);

  return {
    rate,
    hz,
    pip,
    pd,
    dp: round(dp, 1),
    sg: round(sg, 3),
    tdh: round(tdh, 0),
    rateBase: round(rateBase, 0),
    headBase: round(headBase, 2),
    headCurve: round(headCurve, 2),
    headDeviationPct: round(headDeviationPct, 1),
    hhp: round(hhp, 1),
    motorLoadPct: round(motorLoadPct, 0),
    range,
    flags,
  };
}

export function analyzeEsp(rows: EspRowInput[], o: EspOptions): { rows: EspRow[]; summary: EspSummary } {
  const out = rows.map((r) => analyzeEspRow(r, o));
  const valid = out.filter((r) => r.tdh !== null && r.rate !== null);
  const devs = out.map((r) => r.headDeviationPct).filter((v): v is number => v !== null);
  const loads = out.map((r) => r.motorLoadPct).filter((v): v is number => v !== null);
  const flagCounts: Record<string, number> = {};
  out.forEach((r) => r.flags.forEach((f) => (flagCounts[f] = (flagCounts[f] ?? 0) + 1)));
  const latest = [...out].reverse().find((r) => r.headDeviationPct !== null)?.headDeviationPct ?? null;
  const avgDev = mean(devs);
  const down = out.filter((r) => r.range === "Downthrust").length;
  const up = out.filter((r) => r.range === "Upthrust").length;

  const issues: string[] = [];
  if (latest !== null && latest < -Math.abs(o.degradationPct)) issues.push(`head ${Math.abs(latest).toFixed(0)} % below curve`);
  if (valid.length && down / valid.length > 0.3) issues.push("mostly in downthrust");
  if (valid.length && up / valid.length > 0.3) issues.push("mostly in upthrust");
  if (flagCounts[FLAG.lowPip]) issues.push("low intake pressure");
  if (flagCounts[FLAG.overload]) issues.push("motor overload");
  if (flagCounts[FLAG.underload]) issues.push("motor underload");
  const status = !valid.length ? "No valid rows – check the column mapping" : issues.length ? `Attention: ${issues.join(", ")}` : "OK – no major issues detected";

  return {
    rows: out,
    summary: {
      rows: out.length,
      validRows: valid.length,
      downthrust: down,
      upthrust: up,
      inRange: out.filter((r) => r.range === "In range").length,
      avgHeadDeviationPct: round(avgDev, 1),
      latestHeadDeviationPct: latest,
      avgMotorLoadPct: round(mean(loads), 0),
      flaggedRows: out.filter((r) => r.flags.length).length,
      flagCounts,
      status,
    },
  };
}

/** Read a 2-column pump curve (rate, head) from raw cell values; skips header/blank rows. */
export function parseCurve(values: unknown[][]): CurvePoint[] {
  return values
    .map((r) => ({ q: num(r[0]), h: num(r[1]) }))
    .filter((p): p is CurvePoint => p.q !== null && p.h !== null)
    .sort((a, b) => a.q - b.q);
}
