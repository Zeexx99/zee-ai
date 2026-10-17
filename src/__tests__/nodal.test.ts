import { describe, expect, it } from "vitest";
import { buildIpr, iprPwf, rateGrid, runNodal, vlpPwf, type WellInput } from "../utils/nodal";

const well: WellInput = { depth: 8000, tubingId: 2.441, whp: 150, gor: 500, waterCut: 0.5, api: 35, gasSg: 0.7, waterSg: 1.05, tWh: 100, tBh: 200 };

describe("IPR", () => {
  it("straight-line PI from a test", () => {
    const ipr = buildIpr({ model: "pi", pr: 3000, qTest: 1000, pwfTest: 2000 });
    expect(ipr.j).toBeCloseTo(1, 6);
    expect(ipr.aof).toBeCloseTo(3000, 6);
    expect(iprPwf(ipr, 3000, 1500)).toBeCloseTo(1500, 3);
  });
  it("Vogel passes through the test point", () => {
    const ipr = buildIpr({ model: "vogel", pr: 2500, qTest: 800, pwfTest: 1500 });
    expect(ipr.rate(1500)).toBeCloseTo(800, 6);
    expect(ipr.rate(0)).toBeCloseTo(ipr.aof, 6);
  });
  it("composite is linear above Pb and continuous at Pb", () => {
    const ipr = buildIpr({ model: "composite", pr: 3500, pi: 1, pb: 2000 });
    expect(ipr.rate(3000)).toBeCloseTo(500, 6);
    expect(ipr.rate(2000)).toBeCloseTo(1500, 6);
    expect(ipr.rate(1999.9)).toBeCloseTo(1500, 0);
    expect(ipr.aof).toBeCloseTo(1500 + 2000 / 1.8, 6);
  });
  it("needs its inputs", () => {
    expect(() => buildIpr({ model: "vogel", pr: 3000 })).toThrow(/test/);
    expect(() => buildIpr({ model: "composite", pr: 3000, pi: 1 })).toThrow(/bubble point/);
  });
});

describe("VLP and operating point", () => {
  it("tubing curve has the J shape: friction dominates at high rates", () => {
    expect(vlpPwf(4000, well)).toBeGreaterThan(vlpPwf(1000, well));
    expect(vlpPwf(1000, well)).toBeGreaterThan(well.whp);
  });
  it("operating point sits on both curves", () => {
    const r = runNodal({ ipr: { model: "composite", pr: 3500, qTest: 1500, pwfTest: 2000, pb: 2200 }, well });
    const op = r.cases[0].op!;
    expect(op).toBeTruthy();
    expect(Math.abs(vlpPwf(op.q, well) - op.pwf)).toBeLessThan(15);
    expect(op.oil).toBeCloseTo(op.q * 0.5, 6);
  });
  it("sensitivity adds one case per value", () => {
    const r = runNodal({ ipr: { model: "pi", pr: 3500, pi: 1.5 }, well, sensitivity: { kind: "whp", values: [100, 300] }, points: 12 });
    expect(r.cases.map((c) => c.label)).toEqual(["Base case", "Wellhead pressure 100 psia", "Wellhead pressure 300 psia"]);
    const q = r.cases.map((c) => c.op?.q ?? 0);
    expect(q[1]).toBeGreaterThan(q[2]); // lower WHP → more rate
  });
  it("reports a well that can't flow", () => {
    const r = runNodal({ ipr: { model: "pi", pr: 1500, pi: 1 }, well: { ...well, waterCut: 0.95, gor: 50 }, points: 10 });
    expect(r.cases[0].op).toBeNull();
    expect(r.cases[0].note).toMatch(/artificial lift/);
  });
  it("rate grid stays below AOF", () => {
    const g = rateGrid(1000, 10);
    expect(g).toHaveLength(10);
    expect(Math.max(...g)).toBeLessThan(1000);
  });
});
