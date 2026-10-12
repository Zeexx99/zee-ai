import { describe, expect, it } from "vitest";
import { findEmptyColumns, findEmptyRows, isEmptyLike, parseDateText, parseNumberText, runsDescending, toExcelSerial, toProperCase } from "../utils/clean";

describe("dates", () => {
  it("Excel serial", () => {
    expect(toExcelSerial(2024, 1, 1)).toBe(45292);
    expect(toExcelSerial(2024, 2, 30)).toBeNull();
  });
  it("parses common text dates", () => {
    expect(parseDateText("2024-01-01", "MDY")).toBe(45292);
    expect(parseDateText("01/02/2024", "DMY")).toBe(45323);
    expect(parseDateText("1 Jan 2024", "DMY")).toBe(45292);
    expect(parseDateText("Jan 1, 2024", "MDY")).toBe(45292);
    expect(parseDateText("01/02/2024", "off")).toBeNull();
  });
});

describe("numbers", () => {
  it("parses thousands separators and percents", () => {
    expect(parseNumberText("1,234.50")?.value).toBe(1234.5);
    expect(parseNumberText("45%")?.value).toBeCloseTo(0.45, 9);
  });
  it("leaves identifiers alone", () => {
    expect(parseNumberText("00123")).toBeNull();
    expect(parseNumberText("1234567890123456")).toBeNull();
  });
});

describe("text", () => {
  it("proper case", () => expect(toProperCase("WELL a-12 north")).toBe("Well A-12 North"));
});

describe("empty rows / columns", () => {
  it("empty-like cells", () => {
    for (const v of ["", "  ", 0, "0", "0.00", "NULL", "null", "N/A", "NaN", null]) expect(isEmptyLike(v)).toBe(true);
    for (const v of [1, "x", "0.5", "10", false, "-"]) expect(isEmptyLike(v)).toBe(false);
  });
  it("finds rows that are fully blank / zero / null", () => {
    const v = [
      ["Well", "Oil", "Gas"],
      ["", 0, "NULL"],
      ["A-1", 120, 0],
      [0, 0, 0],
    ];
    expect(findEmptyRows(v)).toEqual([1, 3]);
  });
  it("finds columns that are fully blank / zero / null, ignoring a header row", () => {
    const v = [
      ["Well", "Water", "Gas", ""],
      ["A-1", 0, 5, ""],
      ["A-2", "null", 0, 0],
    ];
    expect(findEmptyColumns(v)).toEqual([1, 3]);
  });
  it("groups runs last-first", () => {
    expect(runsDescending([1, 2, 5, 7, 8, 9])).toEqual([
      { start: 7, count: 3 },
      { start: 5, count: 1 },
      { start: 1, count: 2 },
    ]);
  });
});
