import { describe, expect, it } from "vitest";
import { parseDateText, parseNumberText, toExcelSerial, toProperCase } from "../utils/clean";

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
