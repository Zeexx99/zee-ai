import { describe, expect, it } from "vitest";
import {
  affinity,
  apiToSg,
  convertUnits,
  headOnCurve,
  hydraulicHp,
  liquidSg,
  productivityIndex,
  pumpTdh,
  standingBo,
  standingPb,
  standingRs,
  vogelQmax,
  vogelRate,
} from "../utils/petroFormulas";

describe("Vogel IPR", () => {
  it("rate is q_max at Pwf = 0 and zero at Pwf = Pr", () => {
    const qmax = vogelQmax(3000, 800, 2000);
    expect(vogelRate(3000, 0, 800, 2000)).toBeCloseTo(qmax, 6);
    expect(vogelRate(3000, 3000, 800, 2000)).toBeCloseTo(0, 6);
  });
  it("reproduces the test point", () => {
    expect(vogelRate(3000, 2000, 800, 2000)).toBeCloseTo(800, 6);
  });
  it("rejects Pwf_test ≥ Pr", () => {
    expect(() => vogelQmax(2000, 500, 2500)).toThrow();
  });
  it("PI", () => expect(productivityIndex(1000, 3000, 2000)).toBe(1));
});

describe("Standing PVT", () => {
  it("Pb and Rs are inverse functions", () => {
    const rs = standingRs(2000, 180, 35, 0.75);
    expect(standingPb(rs, 180, 35, 0.75)).toBeCloseTo(2000, 0);
  });
  it("Rs increases with pressure", () => {
    expect(standingRs(3000, 180, 35, 0.75)).toBeGreaterThan(standingRs(1000, 180, 35, 0.75));
  });
  it("Bo is physically reasonable", () => {
    const bo = standingBo(standingRs(2500, 200, 32, 0.8), 200, 32, 0.8);
    expect(bo).toBeGreaterThan(1.1);
    expect(bo).toBeLessThan(1.6);
  });
  it("API → SG", () => expect(apiToSg(10)).toBeCloseTo(1, 6));
});

describe("ESP formulas", () => {
  it("TDH", () => expect(pumpTdh(1500, 500, 1)).toBeCloseTo(2309.47, 1));
  it("hydraulic HP", () => expect(hydraulicHp(1000, 1000)).toBeCloseTo(17.015, 2));
  it("affinity laws", () => {
    expect(affinity(1000, 60, 50, "Q")).toBeCloseTo(833.333, 2);
    expect(affinity(36, 60, 50, "H")).toBeCloseTo(25, 6);
    expect(affinity(216, 60, 50, "P")).toBeCloseTo(125, 6);
    expect(() => affinity(1, 60, 50, "X" as never)).toThrow();
  });
  it("mixture SG", () => {
    expect(liquidSg(1, 30)).toBe(1);
    expect(liquidSg(0, 10)).toBeCloseTo(1, 6);
  });
  it("head interpolation on a curve", () => {
    const curve = [
      { q: 0, h: 50 },
      { q: 1000, h: 40 },
      { q: 2000, h: 20 },
    ];
    expect(headOnCurve(curve, 500)).toBeCloseTo(45, 6);
    expect(headOnCurve(curve, 1500)).toBeCloseTo(30, 6);
    expect(headOnCurve(curve, 2500)).toBeNull();
  });
});

describe("convertUnits", () => {
  it("known id, any case", () => expect(convertUnits(100, "BAR_PSI")).toBeCloseTo(1450.377, 2));
  it("unknown id lists valid ids", () => expect(() => convertUnits(1, "nope")).toThrow(/psi_bar/));
});
