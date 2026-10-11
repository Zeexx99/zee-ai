import { describe, expect, it } from "vitest";
import { analyzeEsp, analyzeEspRow, FLAG, parseCurve, type EspOptions } from "../utils/espAnalysis";

// Per-stage catalog curve at 60 Hz.
const curve = [
  { q: 500, h: 40 },
  { q: 1000, h: 36 },
  { q: 1500, h: 30 },
  { q: 2000, h: 20 },
];
const base: EspOptions = { baseHz: 60, sg: 1, stages: 100, curve, degradationPct: 15 };

describe("analyzeEspRow", () => {
  it("a healthy pump at 60 Hz sits on its curve", () => {
    // 36 ft/stage × 100 stages = 3600 ft → ΔP = 3600 × 0.433 = 1558.8 psi
    const r = analyzeEspRow({ rate: 1000, pip: 500, pd: 2058.8, hz: 60 }, base);
    expect(r.tdh).toBeCloseTo(3600, 0);
    expect(r.headBase).toBeCloseTo(36, 1);
    expect(r.headCurve).toBeCloseTo(36, 6);
    expect(Math.abs(r.headDeviationPct ?? 99)).toBeLessThan(0.5);
    expect(r.range).toBe("In range");
    expect(r.flags).toEqual([]);
  });

  it("affinity correction at 50 Hz", () => {
    // At 50 Hz a point on the curve: Q = 1000 × 50/60, H = 36 × (50/60)² per stage.
    const q = 1000 * (50 / 60);
    const tdh = 36 * (50 / 60) ** 2 * 100;
    const r = analyzeEspRow({ rate: q, pip: 500, pd: 500 + tdh * 0.433, hz: 50 }, base);
    expect(r.rateBase).toBeCloseTo(1000, 0);
    expect(r.headBase).toBeCloseTo(36, 1);
  });

  it("flags head loss, downthrust, overload and low PIP", () => {
    const r = analyzeEspRow({ rate: 300, pip: 200, pd: 1200, hz: 60, amps: 110 }, { ...base, motorAmps: 100, minPip: 300, minRate: 600, maxRate: 1800 });
    expect(r.range).toBe("Downthrust");
    expect(r.flags).toContain(FLAG.down);
    expect(r.flags).toContain(FLAG.overload);
    expect(r.flags).toContain(FLAG.lowPip);
    expect(r.flags).toContain(FLAG.offCurve); // 300 bbl/d is left of the catalog curve
  });

  it("detects degraded head", () => {
    const r = analyzeEspRow({ rate: 1000, pip: 500, pd: 500 + 2600 * 0.433, hz: 60 }, base);
    expect(r.headDeviationPct).toBeLessThan(-15);
    expect(r.flags).toContain(FLAG.headLoss);
  });

  it("water cut in percent with API gives a mixture SG", () => {
    const r = analyzeEspRow({ rate: 1000, pip: 500, pd: 1500, hz: 60, waterCut: "50%" }, { ...base, api: 30 });
    expect(r.sg).toBeCloseTo(0.5 + 0.5 * (141.5 / 161.5), 3);
  });

  it("missing data never throws", () => {
    const r = analyzeEspRow({ rate: "", pip: null, pd: "abc", hz: 0 }, base);
    expect(r.tdh).toBeNull();
    expect(r.rateBase).toBeNull();
  });
});

describe("analyzeEsp summary", () => {
  it("summarises status", () => {
    const rows = [
      { rate: 1000, pip: 500, pd: 2058.8, hz: 60 },
      { rate: 1000, pip: 500, pd: 500 + 2600 * 0.433, hz: 60 },
    ];
    const { summary } = analyzeEsp(rows, base);
    expect(summary.validRows).toBe(2);
    expect(summary.status).toMatch(/below curve/);
  });
});

describe("parseCurve", () => {
  it("skips the header and blanks, sorts by rate", () => {
    expect(parseCurve([["Q", "H"], [1000, 36], ["", ""], [500, 40]])).toEqual([
      { q: 500, h: 40 },
      { q: 1000, h: 36 },
    ]);
  });
});
