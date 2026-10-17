/**
 * Gas-well toolkit behind =ZEE.Z, ZEE.BG, ZEE.GAS_VISC, ZEE.DELIV and ZEE.DELIV_Q, also used by nodal analysis.
 * Oilfield units: psia, °F, scf. Pure (no Office) – unit tested.
 *
 *  - Pseudo-critical properties: Sutton (1985) from gas SG, Wichert–Aziz correction for CO2 / H2S.
 *  - Z-factor: Hall–Yarborough (1973), solved with Newton–Raphson.
 *  - Viscosity: Lee–Gonzalez–Eakin (1966).
 *  - Deliverability: Rawlins–Schellhardt back-pressure equation q = C (Pr² − Pwf²)^n.
 */
import { FormulaInputError } from "./petroFormulas";

const finite = (...v: number[]) => v.every((x) => typeof x === "number" && Number.isFinite(x));
function need(ok: boolean, msg: string) {
  if (!ok) throw new FormulaInputError(msg);
}

export interface PseudoCritical {
  /** °R */
  tpc: number;
  /** psia */
  ppc: number;
}

/** Sutton pseudo-critical properties, Wichert–Aziz corrected when CO2 / H2S mole fractions are given. */
export function pseudoCritical(gasSg: number, co2 = 0, h2s = 0): PseudoCritical {
  need(finite(gasSg, co2, h2s) && gasSg > 0.5 && gasSg < 3, "Gas SG must be between 0.5 and 3 (air = 1).");
  need(co2 >= 0 && h2s >= 0 && co2 + h2s < 1, "CO2 and H2S are mole fractions (0–1).");
  let tpc = 169.2 + 349.5 * gasSg - 74.0 * gasSg * gasSg;
  let ppc = 756.8 - 131.0 * gasSg - 3.6 * gasSg * gasSg;
  const a = co2 + h2s;
  if (a > 0) {
    const eps = 120 * (Math.pow(a, 0.9) - Math.pow(a, 1.6)) + 15 * (Math.sqrt(h2s) - Math.pow(h2s, 4));
    const tNew = tpc - eps;
    ppc = (ppc * tNew) / (tpc + h2s * (1 - h2s) * eps);
    tpc = tNew;
  }
  return { tpc, ppc };
}

/** Hall–Yarborough Z from pseudo-reduced pressure and temperature. */
export function hallYarborough(ppr: number, tpr: number): number {
  need(finite(ppr, tpr) && ppr >= 0 && tpr > 0, "Need Ppr ≥ 0 and Tpr > 0.");
  need(tpr >= 1.0 && tpr <= 3.0, `Tpr = ${tpr.toFixed(2)} is outside Hall–Yarborough's range (1.0–3.0).`);
  need(ppr <= 30, `Ppr = ${ppr.toFixed(1)} is outside Hall–Yarborough's range (0–30).`);
  if (ppr === 0) return 1;
  const t = 1 / tpr;
  const A = 0.06125 * t * Math.exp(-1.2 * (1 - t) * (1 - t));
  const B = 14.76 * t - 9.76 * t * t + 4.58 * t * t * t;
  const C = 90.7 * t - 242.2 * t * t + 42.4 * t * t * t;
  const D = 2.18 + 2.82 * t;
  let y = 0.001;
  for (let i = 0; i < 100; i++) {
    const y2 = y * y;
    const y3 = y2 * y;
    const y4 = y3 * y;
    const om = 1 - y;
    const f = -A * ppr + (y + y2 + y3 - y4) / (om * om * om) - B * y2 + C * Math.pow(y, D);
    const df = (1 + 4 * y + 4 * y2 - 4 * y3 + y4) / (om * om * om * om) - 2 * B * y + C * D * Math.pow(y, D - 1);
    let next = y - f / df;
    if (!(next > 0) || next >= 1) next = (y + (next >= 1 ? 1 : 0)) / 2;
    if (Math.abs(next - y) < 1e-12) {
      y = next;
      break;
    }
    y = next;
  }
  const z = (A * ppr) / y;
  need(Number.isFinite(z) && z > 0.2 && z < 3, "Z-factor did not converge – check the inputs.");
  return z;
}

/** Gas compressibility factor at p (psia) and T (°F). */
export function zFactor(p: number, tF: number, gasSg: number, co2 = 0, h2s = 0): number {
  need(finite(p, tF) && p >= 0, "Need p ≥ 0 (psia).");
  const { tpc, ppc } = pseudoCritical(gasSg, co2, h2s);
  return hallYarborough(p / ppc, (tF + 459.67) / tpc);
}

export type BgUnit = "ft3/scf" | "rb/scf" | "rb/Mscf";
export const BG_UNITS: BgUnit[] = ["ft3/scf", "rb/scf", "rb/Mscf"];

/** Gas formation volume factor. Default ft³/scf (= 0.02827 Z T / p). */
export function gasFvf(p: number, tF: number, gasSg: number, unit: string = "ft3/scf", co2 = 0, h2s = 0): number {
  need(finite(p) && p > 0, "Need p > 0 (psia).");
  const z = zFactor(p, tF, gasSg, co2, h2s);
  const ft3 = (0.02827 * z * (tF + 459.67)) / p;
  const u = String(unit || "ft3/scf").trim().toLowerCase().replace("³", "3");
  if (u === "ft3/scf" || u === "cf/scf" || u === "rcf/scf") return ft3;
  if (u === "rb/scf" || u === "bbl/scf") return ft3 / 5.614583;
  if (u === "rb/mscf" || u === "bbl/mscf") return (ft3 / 5.614583) * 1000;
  throw new FormulaInputError(`unit must be one of ${BG_UNITS.join(", ")}.`);
}

/** Gas density, lbm/ft³. */
export function gasDensity(p: number, tF: number, gasSg: number, z?: number): number {
  const zz = z ?? zFactor(p, tF, gasSg);
  return (p * 28.97 * gasSg) / (zz * 10.7316 * (tF + 459.67));
}

/** Lee–Gonzalez–Eakin gas viscosity, cp. */
export function gasViscosity(p: number, tF: number, gasSg: number, co2 = 0, h2s = 0): number {
  need(finite(p, tF) && p > 0, "Need p > 0 (psia).");
  const z = zFactor(p, tF, gasSg, co2, h2s);
  const T = tF + 459.67;
  const ma = 28.97 * gasSg;
  const rho = gasDensity(p, tF, gasSg, z) / 62.428; // g/cm³
  const K = ((9.4 + 0.02 * ma) * Math.pow(T, 1.5)) / (209 + 19 * ma + T);
  const X = 3.5 + 986 / T + 0.01 * ma;
  const Y = 2.4 - 0.2 * X;
  return 1e-4 * K * Math.exp(X * Math.pow(rho, Y));
}

/* ---------------------------------------------------------------- back-pressure deliverability */

export interface Deliverability {
  C: number;
  n: number;
  /** Absolute open flow = C · Pr^(2n), same rate unit as the test data. */
  aof: number;
  /** R² of the log-log fit. */
  r2: number;
  points: number;
}

/** Fit q = C (Pr² − Pwf²)^n to multi-rate test points (least squares in log-log space). */
export function fitDeliverability(rates: number[], pwfs: number[], pr: number): Deliverability {
  need(finite(pr) && pr > 0, "Need Pr > 0 (psia).");
  const pts: Array<[number, number]> = [];
  for (let i = 0; i < Math.min(rates.length, pwfs.length); i++) {
    const q = rates[i];
    const pwf = pwfs[i];
    if (!finite(q, pwf) || q <= 0 || pwf < 0 || pwf >= pr) continue;
    pts.push([Math.log10(pr * pr - pwf * pwf), Math.log10(q)]);
  }
  need(pts.length >= 2, "Need at least 2 test points with q > 0 and Pwf < Pr.");
  const k = pts.length;
  const mx = pts.reduce((s, p) => s + p[0], 0) / k;
  const my = pts.reduce((s, p) => s + p[1], 0) / k;
  const sxx = pts.reduce((s, p) => s + (p[0] - mx) ** 2, 0);
  const sxy = pts.reduce((s, p) => s + (p[0] - mx) * (p[1] - my), 0);
  const syy = pts.reduce((s, p) => s + (p[1] - my) ** 2, 0);
  need(sxx > 0, "Test points need different Pwf values.");
  const n = sxy / sxx;
  const logC = my - n * mx;
  const C = Math.pow(10, logC);
  const r2 = syy > 0 ? (sxy * sxy) / (sxx * syy) : 1;
  return { C, n, aof: C * Math.pow(pr * pr, n), r2, points: k };
}

/** Rate from the back-pressure equation. */
export function deliverabilityRate(C: number, n: number, pr: number, pwf: number): number {
  need(finite(C, n, pr, pwf) && C > 0 && n > 0 && pr > 0 && pwf >= 0 && pwf <= pr, "Need C > 0, n > 0 and 0 ≤ Pwf ≤ Pr.");
  return C * Math.pow(pr * pr - pwf * pwf, n);
}

/** Flatten an Excel range argument to numbers (non-numbers dropped, positions kept as NaN). */
export function flattenNumbers(v: unknown): number[] {
  const rows = Array.isArray(v) ? v : [[v]];
  const out: number[] = [];
  for (const r of rows as unknown[]) {
    for (const c of Array.isArray(r) ? r : [r]) {
      const n = typeof c === "number" ? c : typeof c === "string" && c.trim() !== "" ? Number(c.replace(/,/g, "")) : NaN;
      out.push(Number.isFinite(n) ? n : NaN);
    }
  }
  return out;
}
