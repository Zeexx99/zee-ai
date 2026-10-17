/**
 * Nodal analysis for oil wells (node at the bottom-hole): inflow (IPR) vs outflow (VLP / tubing curve).
 *
 *  IPR: straight-line PI, Vogel from one test, or composite (PI above the bubble point, Vogel below).
 *  VLP: Beggs & Brill (1973) with the Payne holdup correction for uphill flow, marched from the wellhead
 *       down to the mid-perforation depth in 50 ft steps. PVT: Standing Rs/Bo, Beggs–Robinson oil viscosity,
 *       Hall–Yarborough Z, Lee–Gonzalez–Eakin gas viscosity. Linear temperature profile.
 *
 * Oilfield units: psia, °F, ft, inches, STB/d (liquid = oil + water), scf/STB. Pure – unit tested.
 */
import { FormulaInputError, apiToSg, standingBo, standingRs, standingPb } from "./petroFormulas";
import { gasFvf, gasDensity, gasViscosity, zFactor } from "./gasFormulas";

function need(ok: boolean, msg: string) {
  if (!ok) throw new FormulaInputError(msg);
}
const finite = (...v: number[]) => v.every((x) => typeof x === "number" && Number.isFinite(x));

/* ================================================================ inflow */

export type IprModel = "pi" | "vogel" | "composite";

export interface IprInput {
  model: IprModel;
  /** Average reservoir pressure, psia. */
  pr: number;
  /** Productivity index, STB/d/psi (model "pi", or "composite" when no test is given). */
  pi?: number;
  /** One well test: liquid rate (STB/d) at a flowing pressure (psia). Used by "vogel" and "composite". */
  qTest?: number;
  pwfTest?: number;
  /** Bubble-point pressure for "composite". */
  pb?: number;
}

export interface Ipr {
  /** Liquid rate at Pwf. */
  rate: (pwf: number) => number;
  /** Absolute open flow (rate at Pwf = 0). */
  aof: number;
  /** PI used above the bubble point (straight line part), if any. */
  j?: number;
  label: string;
}

export function buildIpr(i: IprInput): Ipr {
  need(finite(i.pr) && i.pr > 0, "Reservoir pressure must be > 0.");
  const hasTest = finite(i.qTest ?? NaN, i.pwfTest ?? NaN);
  if (hasTest) need((i.qTest as number) > 0 && (i.pwfTest as number) >= 0 && (i.pwfTest as number) < i.pr, "The test needs q > 0 and Pwf below reservoir pressure.");

  if (i.model === "pi") {
    let j = i.pi;
    if (!(j && j > 0) && hasTest) j = (i.qTest as number) / (i.pr - (i.pwfTest as number));
    need(!!j && j > 0, "Give a PI (STB/d/psi) or one well test.");
    const jj = j as number;
    return { rate: (pwf) => Math.max(0, jj * (i.pr - pwf)), aof: jj * i.pr, j: jj, label: `Straight-line PI = ${jj.toPrecision(3)} STB/d/psi` };
  }

  if (i.model === "vogel") {
    need(hasTest, "Vogel needs one well test (rate and flowing pressure).");
    const x = (i.pwfTest as number) / i.pr;
    const qmax = (i.qTest as number) / (1 - 0.2 * x - 0.8 * x * x);
    return {
      rate: (pwf) => {
        const r = Math.min(Math.max(pwf, 0), i.pr) / i.pr;
        return qmax * (1 - 0.2 * r - 0.8 * r * r);
      },
      aof: qmax,
      label: `Vogel, AOF = ${Math.round(qmax)} STB/d`,
    };
  }

  // Composite: straight line down to Pb, Vogel below.
  const pb = i.pb as number;
  need(finite(pb) && pb > 0 && pb <= i.pr, "Composite IPR needs a bubble point between 0 and reservoir pressure.");
  let j = i.pi;
  if (hasTest) {
    const qt = i.qTest as number;
    const pt = i.pwfTest as number;
    if (pt >= pb) j = qt / (i.pr - pt);
    else {
      const r = pt / pb;
      j = qt / (i.pr - pb + (pb / 1.8) * (1 - 0.2 * r - 0.8 * r * r));
    }
  }
  need(!!j && j > 0, "Give a PI or one well test.");
  const jj = j as number;
  const qb = jj * (i.pr - pb);
  const qv = (jj * pb) / 1.8;
  const rate = (pwf: number) => {
    const p = Math.min(Math.max(pwf, 0), i.pr);
    if (p >= pb) return jj * (i.pr - p);
    const r = p / pb;
    return qb + qv * (1 - 0.2 * r - 0.8 * r * r);
  };
  return { rate, aof: qb + qv, j: jj, label: `Composite (PI = ${jj.toPrecision(3)} above Pb ${Math.round(pb)} psia, Vogel below), AOF = ${Math.round(qb + qv)} STB/d` };
}

/** Pwf on the IPR for a liquid rate (inverse by bisection). null above AOF. */
export function iprPwf(ipr: Ipr, pr: number, q: number): number | null {
  if (q < 0 || q > ipr.aof) return null;
  let lo = 0;
  let hi = pr;
  for (let k = 0; k < 60; k++) {
    const mid = (lo + hi) / 2;
    if (ipr.rate(mid) > q) lo = mid;
    else hi = mid;
  }
  return (lo + hi) / 2;
}

/* ================================================================ outflow */

export interface WellInput {
  /** Mid-perforation depth (TVD, vertical well assumed), ft. */
  depth: number;
  /** Tubing inner diameter, inches. */
  tubingId: number;
  /** Wellhead (tubing head) flowing pressure, psia. */
  whp: number;
  /** Producing GOR, scf/STB. */
  gor: number;
  /** Water cut, fraction 0–1. */
  waterCut: number;
  api: number;
  gasSg: number;
  waterSg: number;
  /** Wellhead and bottom-hole temperatures, °F. */
  tWh: number;
  tBh: number;
  /** Absolute roughness, inches (default 0.0006). */
  roughness?: number;
}

const G = 32.174;
const STEP_FT = 50;

/** Beggs–Robinson live-oil viscosity, cp. */
export function oilViscosity(api: number, tF: number, rs: number): number {
  const x = Math.pow(tF, -1.163) * Math.exp(6.9824 - 0.04658 * api);
  const muDead = Math.pow(10, x) - 1;
  const a = 10.715 * Math.pow(rs + 100, -0.515);
  const b = 5.44 * Math.pow(rs + 150, -0.338);
  return a * Math.pow(Math.max(muDead, 0.1), b);
}

/** Water viscosity, cp (simple temperature fit, McCain-type, fresh water). */
export function waterViscosity(tF: number): number {
  return Math.max(0.2, 109.574 * Math.pow(tF, -1.12166));
}

export interface GradientPoint {
  /** psi/ft */
  gradient: number;
  holdup: number;
  regime: "segregated" | "transition" | "intermittent" | "distributed" | "single-phase";
}

/** Beggs & Brill pressure gradient for vertical upward flow at one depth. */
export function beggsBrillGradient(qLiquid: number, p: number, tF: number, w: WellInput): GradientPoint {
  const d = w.tubingId / 12; // ft
  const area = (Math.PI * d * d) / 4;
  const wc = w.waterCut;
  const qo = qLiquid * (1 - wc);
  const qw = qLiquid * wc;
  const pbPsia = standingPb(w.gor, tF, w.api, w.gasSg);
  const rs = p >= pbPsia ? w.gor : Math.min(w.gor, standingRs(Math.max(p, 14.7), tF, w.api, w.gasSg));
  const bo = standingBo(rs, tF, w.api, w.gasSg);
  const go = apiToSg(w.api);
  const freeGas = Math.max(0, qo * (w.gor - rs)); // scf/d
  const bg = freeGas > 0 ? gasFvf(Math.max(p, 14.7), tF, w.gasSg) : 0; // ft3/scf
  const qgIs = (freeGas * bg) / 86400; // ft3/s
  const qoIs = (qo * bo * 5.614583) / 86400;
  const qwIs = (qw * 1.0 * 5.614583) / 86400;
  const qlIs = qoIs + qwIs;
  const vsl = qlIs / area;
  const vsg = qgIs / area;
  const vm = vsl + vsg;

  const rhoO = (62.428 * go + 0.0136 * rs * w.gasSg) / bo;
  const rhoW = 62.428 * w.waterSg;
  const fo = qlIs > 0 ? qoIs / qlIs : 0;
  const rhoL = fo * rhoO + (1 - fo) * rhoW;
  const muL = fo * oilViscosity(w.api, tF, rs) + (1 - fo) * waterViscosity(tF);
  const eps = (w.roughness ?? 0.0006) / w.tubingId;

  const friction = (rho: number, mu: number, v: number) => {
    if (v <= 0) return 0;
    const re = (1488 * rho * v * d) / Math.max(mu, 1e-4);
    if (re < 2300) return 64 / Math.max(re, 1);
    // Chen (1979) explicit Colebrook.
    const a = Math.log10(Math.pow(eps, 1.1098) / 2.8257 + Math.pow(7.149 / re, 0.8981));
    const inv = -2 * Math.log10(eps / 3.7065 - (5.0452 / re) * a);
    return 1 / (inv * inv);
  };

  if (vsg <= 0 || vm <= 0) {
    // Single-phase liquid.
    const f = friction(rhoL, muL, vsl);
    return { gradient: (rhoL + (f * rhoL * vsl * vsl) / (2 * G * d)) / 144, holdup: 1, regime: "single-phase" };
  }

  const z = zFactor(Math.max(p, 14.7), tF, w.gasSg);
  const rhoG = gasDensity(Math.max(p, 14.7), tF, w.gasSg, z);
  const muG = gasViscosity(Math.max(p, 14.7), tF, w.gasSg);
  const lam = Math.min(Math.max(vsl / vm, 1e-6), 1);
  const nfr = (vm * vm) / (G * d);
  const L1 = 316 * Math.pow(lam, 0.302);
  const L2 = 0.0009252 * Math.pow(lam, -2.4684);
  const L3 = 0.1 * Math.pow(lam, -1.4516);
  const L4 = 0.5 * Math.pow(lam, -6.738);

  let regime: GradientPoint["regime"];
  if ((lam < 0.01 && nfr < L1) || (lam >= 0.01 && nfr < L2)) regime = "segregated";
  else if (lam >= 0.01 && nfr >= L2 && nfr <= L3) regime = "transition";
  else if ((lam >= 0.01 && lam < 0.4 && nfr > L3 && nfr <= L1) || (lam >= 0.4 && nfr > L3 && nfr <= L4)) regime = "intermittent";
  else regime = "distributed";

  const sigma = 30; // dyn/cm, liquid surface tension (only enters the inclination correction)
  const nlv = 1.938 * vsl * Math.pow(rhoL / sigma, 0.25);
  const sinTerm = Math.sin((1.8 * Math.PI) / 2) - (1 / 3) * Math.pow(Math.sin((1.8 * Math.PI) / 2), 3); // θ = 90°

  const holdupFor = (r: "segregated" | "intermittent" | "distributed") => {
    const [a, b, c] = r === "segregated" ? [0.98, 0.4846, 0.0868] : r === "intermittent" ? [0.845, 0.5351, 0.0173] : [1.065, 0.5824, 0.0609];
    const h0 = Math.max(lam, (a * Math.pow(lam, b)) / Math.pow(nfr, c));
    let C = 0;
    if (r === "segregated") C = (1 - lam) * Math.log(0.011 * Math.pow(lam, -3.768) * Math.pow(nlv, 3.539) * Math.pow(nfr, -1.614));
    else if (r === "intermittent") C = (1 - lam) * Math.log(2.96 * Math.pow(lam, 0.305) * Math.pow(nlv, -0.4473) * Math.pow(nfr, 0.0978));
    C = Math.max(0, Number.isFinite(C) ? C : 0);
    return h0 * (1 + C * sinTerm);
  };
  let hl: number;
  if (regime === "transition") {
    const A = (L3 - nfr) / (L3 - L2);
    hl = A * holdupFor("segregated") + (1 - A) * holdupFor("intermittent");
  } else hl = holdupFor(regime);
  hl = Math.min(1, Math.max(lam, hl * 0.924)); // Payne et al. correction (uphill)

  const rhoS = rhoL * hl + rhoG * (1 - hl);
  const rhoN = rhoL * lam + rhoG * (1 - lam);
  const muN = muL * lam + muG * (1 - lam);
  const fn = friction(rhoN, muN, vm);
  const y = lam / (hl * hl);
  let S = 0;
  if (y > 1 && y < 1.2) S = Math.log(2.2 * y - 1.2);
  else if (y > 0) {
    const ly = Math.log(y);
    S = ly / (-0.0523 + 3.182 * ly - 0.8725 * ly * ly + 0.01853 * Math.pow(ly, 4));
  }
  const ftp = fn * Math.exp(S);
  const gradient = (rhoS + (ftp * rhoN * vm * vm) / (2 * G * d)) / 144;
  return { gradient, holdup: hl, regime };
}

/** Flowing bottom-hole pressure for a liquid rate: march down from the wellhead. */
export function vlpPwf(qLiquid: number, w: WellInput): number {
  need(finite(qLiquid) && qLiquid > 0, "Rate must be > 0.");
  let p = w.whp;
  let z = 0;
  while (z < w.depth) {
    const dz = Math.min(STEP_FT, w.depth - z);
    const tMid = w.tWh + ((w.tBh - w.tWh) * (z + dz / 2)) / w.depth;
    // Two-pass (predictor–corrector) on the average segment pressure.
    const g1 = beggsBrillGradient(qLiquid, p, tMid, w).gradient;
    const g2 = beggsBrillGradient(qLiquid, p + (g1 * dz) / 2, tMid, w).gradient;
    p += g2 * dz;
    z += dz;
    if (!Number.isFinite(p) || p > 30000) throw new FormulaInputError("The tubing curve did not converge – check tubing size, depth and rates.");
  }
  return p;
}

export function validateWell(w: WellInput) {
  need(finite(w.depth) && w.depth > 100 && w.depth < 40000, "Depth must be 100–40,000 ft.");
  need(finite(w.tubingId) && w.tubingId >= 0.5 && w.tubingId <= 12, "Tubing ID must be 0.5–12 inches.");
  need(finite(w.whp) && w.whp >= 14.7, "Wellhead pressure must be ≥ 14.7 psia.");
  need(finite(w.gor) && w.gor >= 0 && w.gor < 50000, "GOR must be 0–50,000 scf/STB.");
  need(finite(w.waterCut) && w.waterCut >= 0 && w.waterCut <= 1, "Water cut must be 0–1 (or 0–100 %).");
  need(finite(w.api) && w.api > 5 && w.api < 70, "API must be 5–70.");
  need(finite(w.gasSg) && w.gasSg >= 0.55 && w.gasSg <= 1.8, "Gas SG must be 0.55–1.8.");
  need(finite(w.tWh, w.tBh) && w.tWh > 32 && w.tBh > 32 && w.tBh < 450, "Temperatures must be in °F (above 32).");
}

/* ================================================================ nodal */

export interface CurvePoint {
  q: number;
  pwfIpr: number | null;
  pwfVlp: number | null;
}

export interface OperatingPoint {
  q: number;
  pwf: number;
  /** Drawdown, psi */
  drawdown: number;
  /** Oil rate, STB/d */
  oil: number;
  /** False when the crossing is left of the tubing curve's minimum (VLP still falling) – flow may head / load up. */
  stable: boolean;
}

export interface NodalCase {
  label: string;
  well: WellInput;
  curve: CurvePoint[];
  op: OperatingPoint | null;
  /** Why there's no operating point (well won't flow naturally, etc.) */
  note?: string;
}

export interface NodalResult {
  ipr: Ipr;
  rates: number[];
  cases: NodalCase[];
}

export type SensitivityKind = "none" | "whp" | "tubingId" | "waterCut" | "gor";

export const SENSITIVITY_LABELS: Record<SensitivityKind, string> = {
  none: "None",
  whp: "Wellhead pressure (psia)",
  tubingId: "Tubing ID (in)",
  waterCut: "Water cut (%)",
  gor: "GOR (scf/STB)",
};

export function rateGrid(aof: number, points = 30): number[] {
  const top = aof * 0.995;
  return Array.from({ length: points }, (_, k) => +((top * (k + 1)) / points).toPrecision(5));
}

function intersect(curve: CurvePoint[], ipr: Ipr, pr: number, well: WellInput): OperatingPoint | null {
  const diff = curve.map((c) => (c.pwfIpr !== null && c.pwfVlp !== null ? c.pwfIpr - c.pwfVlp : null));
  // Highest-rate crossing where IPR goes from above VLP to below (the stable solution).
  for (let k = diff.length - 1; k > 0; k--) {
    const a = diff[k - 1];
    const b = diff[k];
    if (a === null || b === null) continue;
    if (a >= 0 && b < 0) {
      // Refine by bisection on the rate.
      let lo = curve[k - 1].q;
      let hi = curve[k].q;
      for (let it = 0; it < 30; it++) {
        const mid = (lo + hi) / 2;
        const d = (iprPwf(ipr, pr, mid) ?? 0) - vlpPwf(mid, well);
        if (d >= 0) lo = mid;
        else hi = mid;
      }
      const q = (lo + hi) / 2;
      const pwf = iprPwf(ipr, pr, q) ?? 0;
      const slopeV = (curve[k].pwfVlp as number) - (curve[k - 1].pwfVlp as number);
      return { q, pwf, drawdown: pr - pwf, oil: q * (1 - well.waterCut), stable: slopeV >= 0 };
    }
  }
  return null;
}

export interface NodalInput {
  ipr: IprInput;
  well: WellInput;
  sensitivity?: { kind: SensitivityKind; values: number[] };
  points?: number;
}

export function runNodal(input: NodalInput): NodalResult {
  validateWell(input.well);
  const ipr = buildIpr(input.ipr);
  const pr = input.ipr.pr;
  need(input.well.whp < pr, "Wellhead pressure must be below reservoir pressure.");
  const rates = rateGrid(ipr.aof, input.points ?? 30);
  const variants: Array<{ label: string; well: WellInput }> = [{ label: "Base case", well: input.well }];
  const sens = input.sensitivity;
  if (sens && sens.kind !== "none") {
    for (const v of sens.values.slice(0, 6)) {
      if (!finite(v)) continue;
      const w = { ...input.well };
      if (sens.kind === "whp") w.whp = v;
      else if (sens.kind === "tubingId") w.tubingId = v;
      else if (sens.kind === "waterCut") w.waterCut = v > 1 ? v / 100 : v;
      else if (sens.kind === "gor") w.gor = v;
      try {
        validateWell(w);
      } catch (e) {
        throw new FormulaInputError(`Sensitivity value ${v}: ${(e as Error).message}`);
      }
      const unit = sens.kind === "whp" ? " psia" : sens.kind === "tubingId" ? " in" : sens.kind === "waterCut" ? " %" : " scf/STB";
      const shown = sens.kind === "waterCut" ? +(w.waterCut * 100).toFixed(1) : v;
      variants.push({ label: `${SENSITIVITY_LABELS[sens.kind].replace(/ \(.*\)$/, "")} ${shown}${unit}`, well: w });
    }
  }
  const cases: NodalCase[] = variants.map(({ label, well }) => {
    const curve: CurvePoint[] = rates.map((q) => {
      let pwfVlp: number | null = null;
      try {
        pwfVlp = vlpPwf(q, well);
      } catch {
        pwfVlp = null;
      }
      return { q, pwfIpr: iprPwf(ipr, pr, q), pwfVlp };
    });
    const op = intersect(curve, ipr, pr, well);
    let note: string | undefined;
    if (!op) {
      const minVlp = Math.min(...curve.map((c) => c.pwfVlp ?? Infinity));
      note = minVlp >= pr ? "No natural flow: the tubing needs more pressure than the reservoir has at every rate – the well needs artificial lift." : "No crossing found between the curves in the rate range.";
    }
    return { label, well, curve, op, note };
  });
  return { ipr, rates, cases };
}
