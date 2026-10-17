import { describe, expect, it } from "vitest";
import { deliverabilityRate, fitDeliverability, flattenNumbers, gasFvf, gasViscosity, hallYarborough, pseudoCritical, zFactor } from "../utils/gasFormulas";

describe("gas toolkit", () => {
  it("Hall–Yarborough matches Standing–Katz / DAK within 1 %", () => {
    // Reference values from the Dranchuk–Abou-Kassem fit of the Standing–Katz chart.
    expect(hallYarborough(2, 1.5)).toBeCloseTo(0.821, 2);
    expect(hallYarborough(3, 2)).toBeCloseTo(0.938, 2);
    expect(hallYarborough(10, 1.3)).toBeCloseTo(1.146, 1);
    expect(hallYarborough(0, 1.5)).toBe(1);
  });

  it("rejects inputs outside the correlation's range", () => {
    expect(() => hallYarborough(2, 0.9)).toThrow(/Tpr/);
    expect(() => zFactor(2000, 200, 0.3)).toThrow(/Gas SG/);
  });

  it("Sutton pseudo-criticals and Wichert–Aziz correction", () => {
    const pc = pseudoCritical(0.7);
    expect(pc.tpc).toBeCloseTo(377.6, 1);
    expect(pc.ppc).toBeCloseTo(663.3, 1);
    const sour = pseudoCritical(0.7, 0.05, 0.1);
    expect(sour.tpc).toBeLessThan(pc.tpc);
  });

  it("Bg in three units and viscosity in a sensible range", () => {
    const ft3 = gasFvf(2000, 200, 0.7);
    expect(ft3).toBeCloseTo(0.0082, 4);
    expect(gasFvf(2000, 200, 0.7, "rb/scf")).toBeCloseTo(ft3 / 5.614583, 8);
    expect(gasFvf(2000, 200, 0.7, "rb/Mscf")).toBeCloseTo((ft3 / 5.614583) * 1000, 6);
    expect(() => gasFvf(2000, 200, 0.7, "m3")).toThrow();
    const mu = gasViscosity(2000, 200, 0.7);
    expect(mu).toBeGreaterThan(0.012);
    expect(mu).toBeLessThan(0.025);
  });

  it("back-pressure fit recovers C and n", () => {
    const pr = 2500;
    const pw = [2200, 1900, 1600, 1300];
    const q = pw.map((p) => deliverabilityRate(0.0015, 0.85, pr, p));
    const d = fitDeliverability(q, pw, pr);
    expect(d.C).toBeCloseTo(0.0015, 6);
    expect(d.n).toBeCloseTo(0.85, 6);
    expect(d.aof).toBeCloseTo(deliverabilityRate(0.0015, 0.85, pr, 0), 3);
    expect(d.points).toBe(4);
    expect(() => fitDeliverability([100], [2000], pr)).toThrow(/2 test points/);
  });

  it("flattens Excel ranges", () => {
    expect(flattenNumbers([[1], ["2"], [""], ["x"]]).map((v) => (Number.isNaN(v) ? "nan" : v))).toEqual([1, 2, "nan", "nan"]);
    expect(flattenNumbers(5)).toEqual([5]);
  });
});
