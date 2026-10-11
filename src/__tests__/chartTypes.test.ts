import { describe, expect, it } from "vitest";
import { CHART_TYPES, chartTypeInfo, checkYCount, supportsSecondaryAxis } from "../utils/chartTypes";

describe("chart catalog", () => {
  it("ids are unique", () => {
    expect(new Set(CHART_TYPES.map((t) => t.id)).size).toBe(CHART_TYPES.length);
  });
  it("Y-count rules", () => {
    expect(checkYCount(chartTypeInfo("Pie"), 2)).toMatch(/exactly 1 Y column/);
    expect(checkYCount(chartTypeInfo("StockOHLC"), 3)).toMatch(/exactly 4/);
    expect(checkYCount(chartTypeInfo("Line"), 5)).toBeNull();
    expect(checkYCount(chartTypeInfo("Line"), 0)).toMatch(/at least one/);
  });
  it("secondary axis only for 2-D series charts", () => {
    expect(supportsSecondaryAxis(chartTypeInfo("LineMarkers"))).toBe(true);
    expect(supportsSecondaryAxis(chartTypeInfo("3DColumn"))).toBe(false);
    expect(supportsSecondaryAxis(chartTypeInfo("Pie"))).toBe(false);
  });
});
