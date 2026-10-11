/**
 * Pure petroleum-engineering formulas behind the =ZEE.* worksheet functions and the ESP report.
 * Oilfield units throughout: psia/psi, °F, bbl/d, scf/STB, ft. No Office dependency (unit-tested).
 */
import { arpsCum, arpsRate, UNIT_CONVERSIONS } from "./petro";

const finite = (...v: number[]) => v.every((x) => typeof x === "number" && Number.isFinite(x));

export class FormulaInputError extends Error {}
function need(ok: boolean, msg: string) {
  if (!ok) throw new FormulaInputError(msg);
}

/* ---------------------------------------------------------------- inflow (Vogel) */

/** Vogel AOF (q at Pwf = 0) from one test point. */
export function vogelQmax(pr: number, qTest: number, pwfTest: number): number {
  need(finite(pr, qTest, pwfTest) && pr > 0 && qTest >= 0 && pwfTest >= 0 && pwfTest < pr, "Need Pr > Pwf_test ≥ 0 and q_test ≥ 0.");
  const x = pwfTest / pr;
  return qTest / (1 - 0.2 * x - 0.8 * x * x);
}

/** Vogel rate at Pwf, from reservoir pressure and one test point. */
export function vogelRate(pr: number, pwf: number, qTest: number, pwfTest: number): number {
  need(finite(pwf) && pwf >= 0 && pwf <= pr, "Pwf must be between 0 and Pr.");
  const qmax = vogelQmax(pr, qTest, pwfTest);
  const x = pwf / pr;
  return qmax * (1 - 0.2 * x - 0.8 * x * x);
}

/** Straight-line productivity index J = q / (Pr − Pwf). */
export function productivityIndex(q: number, pr: number, pwf: number): number {
  need(finite(q, pr, pwf) && pr > pwf, "Need Pr > Pwf.");
  return q / (pr - pwf);
}

/* ---------------------------------------------------------------- PVT (Standing, 1947) */

export const apiToSg = (api: number) => 141.5 / (api + 131.5);

/** Solution GOR Rs (scf/STB) at pressure p (psia) ≤ bubble point. */
export function standingRs(p: number, tF: number, api: number, gasSg: number): number {
  need(finite(p, tF, api, gasSg) && p > 0 && gasSg > 0, "Need p > 0 and gas SG > 0.");
  const x = 0.0125 * api - 0.00091 * tF;
  return gasSg * Math.pow((p / 18.2 + 1.4) * Math.pow(10, x), 1.2048);
}

/** Bubble-point pressure (psia) from Rs. */
export function standingPb(rs: number, tF: number, api: number, gasSg: number): number {
  need(finite(rs, tF, api, gasSg) && rs >= 0 && gasSg > 0, "Need Rs ≥ 0 and gas SG > 0.");
  const a = 0.00091 * tF - 0.0125 * api;
  return 18.2 * (Math.pow(rs / gasSg, 0.83) * Math.pow(10, a) - 1.4);
}

/** Oil formation volume factor (rb/STB) from Rs. */
export function standingBo(rs: number, tF: number, api: number, gasSg: number): number {
  need(finite(rs, tF, api, gasSg) && rs >= 0 && gasSg > 0, "Need Rs ≥ 0 and gas SG > 0.");
  const go = apiToSg(api);
  return 0.9759 + 0.00012 * Math.pow(rs * Math.sqrt(gasSg / go) + 1.25 * tF, 1.2);
}

/* ---------------------------------------------------------------- decline */

export function arpsRateAt(qi: number, di: number, b: number, t: number): number {
  need(finite(qi, di, b, t) && qi > 0 && di >= 0 && b >= 0 && t >= 0, "Need qi > 0, Di ≥ 0, b ≥ 0, t ≥ 0.");
  return arpsRate({ qi, di, b }, t);
}

export function arpsCumAt(qi: number, di: number, b: number, t: number): number {
  need(finite(qi, di, b, t) && qi > 0 && di > 0 && b >= 0 && t >= 0, "Need qi > 0, Di > 0, b ≥ 0, t ≥ 0.");
  return arpsCum({ qi, di, b }, t);
}

/* ---------------------------------------------------------------- ESP / pumps */

/** Pressure gradient of fresh water, psi/ft. */
export const WATER_GRADIENT = 0.433;

/** Total dynamic head (ft) developed by the pump. */
export function pumpTdh(pDischarge: number, pIntake: number, sg: number): number {
  need(finite(pDischarge, pIntake, sg) && sg > 0, "Need SG > 0.");
  return (pDischarge - pIntake) / (WATER_GRADIENT * sg);
}

/** Hydraulic horsepower: q (bbl/d) × ΔP (psi) / 58,771. */
export function hydraulicHp(q: number, dpPsi: number): number {
  need(finite(q, dpPsi), "Need numbers.");
  return (q * dpPsi) / 58771;
}

export type AffinityKind = "Q" | "H" | "P";

/** Affinity laws: rate ∝ f, head ∝ f², power ∝ f³. */
export function affinity(value: number, fFrom: number, fTo: number, kind: AffinityKind): number {
  need(finite(value, fFrom, fTo) && fFrom > 0 && fTo > 0, "Need positive frequencies.");
  const r = fTo / fFrom;
  const k = String(kind).trim().toUpperCase();
  if (k === "Q") return value * r;
  if (k === "H") return value * r * r;
  if (k === "P" || k === "BHP" || k === "HP") return value * r * r * r;
  throw new FormulaInputError('kind must be "Q", "H" or "P".');
}

/** Mixture specific gravity from water cut (0–1) and oil API (gas ignored). */
export function liquidSg(waterCut: number, api: number, waterSg = 1.0): number {
  need(finite(waterCut, api, waterSg) && waterCut >= 0 && waterCut <= 1, "Water cut must be 0–1.");
  return waterCut * waterSg + (1 - waterCut) * apiToSg(api);
}

/* ---------------------------------------------------------------- pump curve lookup */

export interface CurvePoint {
  q: number; // rate at base frequency
  h: number; // head per stage (or total) at base frequency
}

/** Linear interpolation of head on a pump curve; null outside the curve. */
export function headOnCurve(curve: CurvePoint[], q: number): number | null {
  const pts = curve.filter((p) => finite(p.q, p.h)).sort((a, b) => a.q - b.q);
  if (pts.length < 2 || q < pts[0].q || q > pts[pts.length - 1].q) return null;
  for (let i = 1; i < pts.length; i++) {
    if (q <= pts[i].q) {
      const a = pts[i - 1];
      const b = pts[i];
      return b.q === a.q ? a.h : a.h + ((q - a.q) * (b.h - a.h)) / (b.q - a.q);
    }
  }
  return pts[pts.length - 1].h;
}

/* ---------------------------------------------------------------- units */

export function convertUnits(value: number, id: string): number {
  need(finite(value), "Value must be a number.");
  const conv = UNIT_CONVERSIONS.find((u) => u.id.toLowerCase() === String(id).trim().toLowerCase());
  need(!!conv, `Unknown conversion "${id}". Use one of: ${UNIT_CONVERSIONS.map((u) => u.id).join(", ")}`);
  return conv!.convert(value);
}
