import { describe, expect, it } from "vitest";
import { buildFilterCriteria, cfFormula, columnIndexFromLetter, matchHeader, normalizeAggregation, normalizeOperator, resolveColumn, validName } from "../utils/agentArgs";

describe("columns", () => {
  it("letters", () => {
    expect(columnIndexFromLetter("A")).toBe(0);
    expect(columnIndexFromLetter("z")).toBe(25);
    expect(columnIndexFromLetter("AB")).toBe(27);
    expect(columnIndexFromLetter("Oil1")).toBeNull();
  });
  it("resolves by header, letter, position or partial header", () => {
    const header = ["Date", "Well", "Oil rate (bbl/d)", "Water cut"];
    // Range starts in column B (index 1).
    expect(resolveColumn("Well", header, 1)).toBe(1);
    expect(resolveColumn("D", header, 1)).toBe(2);
    expect(resolveColumn("4", header, 1)).toBe(3);
    expect(resolveColumn(1, header, 1)).toBe(0);
    expect(resolveColumn("oil", header, 1)).toBe(2);
    expect(resolveColumn("A", header, 1)).toBeNull(); // left of the range
    expect(resolveColumn("Gas", header, 1)).toBeNull();
  });
  it("prefers an exact header over a letter", () => {
    expect(resolveColumn("B", ["A", "B"], 5)).toBe(1);
  });
});

describe("filters", () => {
  it("values, conditions and top N", () => {
    expect(buildFilterCriteria({ values: ["W-1", 2] })).toEqual({ filterOn: "Values", values: ["W-1", "2"] });
    expect(buildFilterCriteria({ condition: ">100" })).toEqual({ filterOn: "Custom", criterion1: ">100" });
    expect(buildFilterCriteria({ condition: "Shut-in" })).toEqual({ filterOn: "Custom", criterion1: "=Shut-in" });
    expect(buildFilterCriteria({ condition: ">=10", condition2: "<20", operator: "and" })).toEqual({ filterOn: "Custom", criterion1: ">=10", criterion2: "<20", operator: "And" });
    expect(buildFilterCriteria({ top: 5 })).toEqual({ filterOn: "TopItems", criterion1: "5" });
    expect(buildFilterCriteria({ bottom_percent: "10" })).toEqual({ filterOn: "BottomPercent", criterion1: "10" });
    expect(() => buildFilterCriteria({})).toThrow();
  });
});

describe("operators and formulas", () => {
  it("normalizes operator spellings", () => {
    expect(normalizeOperator("greater_than")).toBe("GreaterThan");
    expect(normalizeOperator(">=")).toBe("GreaterThanOrEqual");
    expect(normalizeOperator("Not Between")).toBe("NotBetween");
    expect(normalizeOperator("less_than_or_equal_to")).toBe("LessThanOrEqual");
    expect(normalizeOperator("bigger")).toBeNull();
  });
  it("conditional-format formulas", () => {
    expect(cfFormula(100)).toBe("100");
    expect(cfFormula("50.5")).toBe("50.5");
    expect(cfFormula("=$B$1")).toBe("=$B$1");
    expect(cfFormula('Shut "in"')).toBe('="Shut ""in"""');
    expect(cfFormula("")).toBe("");
  });
  it("pivot aggregations", () => {
    expect(normalizeAggregation("avg")).toBe("Average");
    expect(normalizeAggregation(undefined)).toBe("Sum");
    expect(normalizeAggregation("stdev")).toBe("StandardDeviation");
    expect(normalizeAggregation("COUNT")).toBe("Count");
  });
  it("matches headers", () => {
    expect(matchHeader("oil", ["Well", "Oil (bbl/d)"])).toBe("Oil (bbl/d)");
    expect(matchHeader("WELL", ["Well", "Oil"])).toBe("Well");
    expect(matchHeader("gas", ["Well"])).toBeNull();
  });
  it("defined names", () => {
    expect(validName("OilRate")).toBe(true);
    expect(validName("_tmp.1")).toBe(true);
    expect(validName("A1")).toBe(false);
    expect(validName("XFD100")).toBe(false);
    expect(validName("has space")).toBe(false);
    expect(validName("1abc")).toBe(false);
    expect(validName("r")).toBe(false);
  });
});
