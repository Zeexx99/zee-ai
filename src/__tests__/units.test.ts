import { describe, expect, it } from "vitest";
import { convertUnit } from "../utils/units";

describe("ZEE.UNITS", () => {
  it("converts oilfield units", () => {
    expect(convertUnit(100, "psi", "bar")).toBeCloseTo(6.894757, 5);
    expect(convertUnit(1000, "bbl/d", "m3/d")).toBeCloseTo(158.987295, 5);
    expect(convertUnit(1, "MMscf/d", "m3/d")).toBeCloseTo(28316.846592, 4);
    expect(convertUnit(1, "hp", "kW")).toBeCloseTo(0.745699872, 8);
    expect(convertUnit(3281, "ft", "m")).toBeCloseTo(1000.0488, 3);
  });
  it("handles temperature, API and gradients", () => {
    expect(convertUnit(212, "°F", "C")).toBeCloseTo(100, 8);
    expect(convertUnit(10, "API", "SG")).toBeCloseTo(1, 8);
    expect(convertUnit(1, "SG", "API")).toBeCloseTo(10, 8);
    expect(convertUnit(0.433, "psi/ft", "sg")).toBeCloseTo(1, 8);
  });
  it("accepts aliases", () => {
    expect(convertUnit(1, "BPD", "bbl/d")).toBe(1);
    expect(convertUnit(1, "psia", "psig")).toBe(1);
  });
  it("rejects bad input", () => {
    expect(() => convertUnit(1, "psi", "ft")).toThrow(/Can't convert/);
    expect(() => convertUnit(1, "psix", "bar")).toThrow(/Unknown unit/);
  });
});
