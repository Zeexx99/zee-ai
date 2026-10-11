import { describe, expect, it } from "vitest";
import { arpsCum, arpsRate, arpsTimeToRate, convertValue, fitArps, fitArpsByGroup, prepareDca, UNIT_CONVERSIONS } from "../utils/petro";

/** Trapezoid integral of the Arps rate – reference for the closed-form cumulative. */
function integrate(f: { qi: number; di: number; b: number }, t: number, steps = 20000) {
  const h = t / steps;
  let s = 0;
  for (let i = 0; i < steps; i++) s += ((arpsRate(f, i * h) + arpsRate(f, (i + 1) * h)) / 2) * h;
  return s;
}

describe("Arps rate and cumulative", () => {
  it("exponential rate", () => {
    expect(arpsRate({ qi: 1000, di: 0.01, b: 0 }, 100)).toBeCloseTo(1000 * Math.exp(-1), 6);
  });
  it("harmonic rate", () => {
    expect(arpsRate({ qi: 1000, di: 0.01, b: 1 }, 100)).toBeCloseTo(500, 6);
  });
  it.each([0, 0.5, 1])("cumulative matches numerical integral (b = %s)", (b) => {
    const f = { qi: 800, di: 0.004, b };
    expect(arpsCum(f, 900) / integrate(f, 900)).toBeCloseTo(1, 4);
  });
  it("time to economic limit is consistent with the rate", () => {
    const f = { qi: 500, di: 0.002, b: 0.6 };
    const t = arpsTimeToRate(f, 50);
    expect(arpsRate(f, t)).toBeCloseTo(50, 6);
    expect(arpsTimeToRate(f, 600)).toBe(0);
  });
});

describe("fitArps", () => {
  it("recovers a known hyperbolic decline", () => {
    const truth = { qi: 1200, di: 0.003, b: 0.7 };
    const t = Array.from({ length: 60 }, (_, i) => i * 30);
    const q = t.map((x) => arpsRate(truth, x));
    const best = fitArps({ t, q })[0];
    expect(best.r2).toBeGreaterThan(0.999);
    expect(best.qi).toBeCloseTo(1200, -1);
    expect(best.b).toBeGreaterThan(0.5);
    expect(best.b).toBeLessThan(0.9);
  });
  it("needs at least 3 points", () => {
    expect(() => fitArps({ t: [0, 1], q: [10, 9] })).toThrow();
  });
});

describe("prepareDca", () => {
  it("drops blanks, text and non-positive rates and shifts time to 0", () => {
    const p = prepareDca([45000, 45030, "", 45060, 45090], [100, 0, 90, "x", 80]);
    expect(p.t).toEqual([0, 90]);
    expect(p.q).toEqual([100, 80]);
    expect(p.t0).toBe(45000);
    expect(p.dropped).toBe(3);
  });
});

describe("fitArpsByGroup", () => {
  it("fits each well separately and reports wells with too few points", () => {
    const f = { qi: 300, di: 0.01, b: 0 };
    const keys: string[] = [];
    const t: number[] = [];
    const q: number[] = [];
    for (let i = 0; i < 10; i++) {
      keys.push("W-1");
      t.push(i * 10);
      q.push(arpsRate(f, i * 10));
    }
    keys.push("W-2", "W-2");
    t.push(0, 10);
    q.push(50, 40);
    const res = fitArpsByGroup(keys, t, q);
    expect(res.map((r) => r.group)).toEqual(["W-1", "W-2"]);
    expect(res[0].best?.r2).toBeGreaterThan(0.999);
    expect(res[1].best).toBeNull();
    expect(res[1].error).toMatch(/3 valid points/);
  });
});

describe("unit conversions", () => {
  const conv = (id: string) => UNIT_CONVERSIONS.find((u) => u.id === id)!;
  it("psi → bar", () => expect(convertValue(1000, conv("psi_bar"))).toBeCloseTo(68.9476, 3));
  it("bbl → m³", () => expect(convertValue("1,000", conv("bbl_m3"))).toBeCloseTo(158.987, 2));
  it("°F → °C", () => expect(convertValue(212, conv("f_c"))).toBe(100));
  it("API ↔ SG round trip", () => expect(convertValue(convertValue(35, conv("api_sg")) as number, conv("sg_api"))).toBeCloseTo(35, 3));
  it("blank and text give empty", () => {
    expect(convertValue("", conv("psi_bar"))).toBe("");
    expect(convertValue("abc", conv("psi_bar"))).toBe("");
  });
});
