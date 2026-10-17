import { describe, expect, it } from "vitest";
import { detectDelimiter, parseDateTime, parseFile, pivotLong, splitLine } from "../utils/fileImport";
import { findHeaderRow, guessKeys, mergeTables, wellKey, type MergeTable } from "../utils/merge";
import { guessAgg, nextPeriod, periodStart, resample } from "../utils/resample";

const serial = (y: number, m: number, d: number) => (Date.UTC(y, m - 1, d) - Date.UTC(1899, 11, 30)) / 86_400_000;

describe("file import", () => {
  it("splits quoted fields and detects the separator", () => {
    expect(splitLine('a,"b, c","say ""hi"""', ",")).toEqual(["a", "b, c", 'say "hi"']);
    expect(detectDelimiter(["Date;Well;Oil", "1/1/2026;W1;5", "2/1/2026;W1;6"])).toBe(";");
    expect(detectDelimiter(["Date\tWell", "x\ty"])).toBe("\t");
  });
  it("parses dates with time, day-first by default", () => {
    expect(parseDateTime("13/10/2026 08:00")).toBeCloseTo(serial(2026, 10, 13) + 8 / 24, 6);
    expect(parseDateTime("03/04/2026")).toBe(serial(2026, 4, 3));
    expect(parseDateTime("03/04/2026", "MDY")).toBe(serial(2026, 3, 4));
    expect(parseDateTime("2026-10-13T20:30:00Z")).toBeCloseTo(serial(2026, 10, 13) + 20.5 / 24, 6);
    expect(parseDateTime("10/13/2026 8:30 PM", "MDY")).toBeCloseTo(serial(2026, 10, 13) + 20.5 / 24, 6);
    expect(parseDateTime("1.5")).toBeNull();
  });
  it("skips metadata lines and converts numbers", () => {
    const p = parseFile("prod.csv", 'Report: Field X\nExported 2026\n\nDate,Well,Oil\n2026-10-01,W-1,100\n2026-10-02,W-1,"1,200"\n2026-10-03,W-1,-999.25\n');
    expect(p.headers).toEqual(["Date", "Well", "Oil"]);
    expect(p.rows.map((r) => r[2])).toEqual([100, 1200, ""]);
    expect(p.dateCols).toEqual([0]);
    expect(p.notes[0]).toMatch(/skipped/);
  });
  it("pivots a SCADA tag list to columns", () => {
    const p = parseFile("scada.csv", "Timestamp;Tag;Value\n13/10/2026 08:00;W1.PIP;812,5\n13/10/2026 08:00;W1.AMPS;45\n13/10/2026 09:00;W1.PIP;800\n");
    expect(p.kind).toBe("scada");
    const w = pivotLong(p);
    expect(w.headers).toEqual(["Timestamp", "W1.PIP", "W1.AMPS"]);
    expect(w.rows).toHaveLength(2);
    expect(w.rows[0].slice(1)).toEqual([812.5, 45]);
    expect(w.rows[1][2]).toBe("");
  });
  it("reads LAS 2.0 with null values", () => {
    const las = "~Version\nVERS. 2.0 : CWLS\nWRAP. NO : one line\n~Well\nNULL. -999.25 : NULL VALUE\nWELL. W-12 : WELL\n~Curve\nDEPT.FT : depth\nGR.API : gamma\n~A\n1000 45.2\n1000.5 -999.25\n";
    const p = parseFile("w12.las", las);
    expect(p.kind).toBe("las");
    expect(p.headers).toEqual(["DEPT (FT)", "GR (API)"]);
    expect(p.rows).toEqual([[1000, 45.2], [1000.5, ""]]);
    expect(p.notes.join(" ")).toMatch(/W-12/);
  });
});

describe("merge well data", () => {
  const prod: MergeTable = { name: "Prod", headers: ["Date", "Well", "Oil"], rows: [[46296, "W-1", 100], [46297, "W-1", 110], [46297, "w 2", 50]], wellCol: 1, dateCol: 0 };
  const tests: MergeTable = { name: "Tests", headers: ["Well", "Test date", "Oil", "WC"], rows: [["W1", 46290, 90, 0.3], ["W2", "2/10/2026", 45, 0.5]], wellCol: 0, dateCol: 1 };
  const master: MergeTable = { name: "Master", headers: ["Well", "Field"], rows: [["W_1", "North"], ["W2", "South"]], wellCol: 0, dateCol: -1 };

  it("matches well names loosely", () => {
    expect(wellKey(" w-01 ")).toBe(wellKey("W 01"));
    expect(wellKey("W_01")).toBe("W01");
  });
  it("as-of join takes the latest test on or before each day, prefixes clashing names", () => {
    const r = mergeTables([prod, tests, master], { how: "left", dateMatch: "asof" });
    expect(r.headers).toEqual(["Well", "Date", "Prod: Oil", "Tests: Oil", "WC", "Field"]);
    expect(r.rows).toEqual([
      ["W-1", 46296, 100, 90, 0.3, "North"],
      ["W-1", 46297, 110, 90, 0.3, "North"],
      ["w 2", 46297, 50, 45, 0.5, "South"],
    ]);
  });
  it("max age drops stale matches", () => {
    const r = mergeTables([prod, tests], { how: "left", dateMatch: "asof", maxAgeDays: 3 });
    expect(r.rows[0][3]).toBe("");
    expect(r.rows[2][3]).toBe(45);
  });
  it("outer exact join keeps every well/date", () => {
    const r = mergeTables([prod, tests], { how: "outer", dateMatch: "exact" });
    expect(r.rows).toHaveLength(4);
    expect(r.stats.wells).toBe(2);
  });
  it("needs a dated first table for as-of", () => {
    expect(() => mergeTables([master, prod], { how: "left", dateMatch: "asof" })).toThrow(/date column/);
  });
  it("guesses keys and header rows", () => {
    expect(guessKeys(["Test date", "Well name", "Oil"])).toEqual({ wellCol: 1, dateCol: 0 });
    expect(findHeaderRow([["Daily production Oct 2026", "", ""], ["Date", "Well", "Oil"], [46296, "W1", 5]])).toBe(1);
  });
});

describe("resample", () => {
  it("period starts: Monday weeks, months, quarters", () => {
    const wed = serial(2026, 10, 14);
    expect(periodStart(wed, "week")).toBe(serial(2026, 10, 12));
    expect(periodStart(wed, "month")).toBe(serial(2026, 10, 1));
    expect(periodStart(wed, "quarter")).toBe(serial(2026, 10, 1));
    expect(periodStart(serial(2026, 8, 20), "quarter")).toBe(serial(2026, 7, 1));
    expect(nextPeriod(serial(2026, 12, 1), "month")).toBe(serial(2027, 1, 1));
  });
  it("guesses rules from headers", () => {
    expect(guessAgg("Oil rate (bbl/d)", [1, 2])).toBe("avg");
    expect(guessAgg("Hours on", [24, 20])).toBe("sum");
    expect(guessAgg("Cum oil", [1, 2])).toBe("last");
    expect(guessAgg("Status", ["on", "off"])).toBe("last");
  });
  it("rolls days up to months per well, with interpolated missing months", () => {
    const rows = [
      [serial(2026, 1, 1), "W1", 100, 24],
      [serial(2026, 1, 2), "W1", 120, 20],
      [serial(2026, 3, 5), "W1", 90, 24],
      [serial(2026, 1, 1), "W2", 10, 24],
    ];
    const r = resample(["Date", "Well", "Oil rate (bbl/d)", "Hours"], rows, { dateCol: 0, groupCol: 1, period: "month", aggs: {}, fill: "linear", countDays: true });
    expect(r.headers).toEqual(["Well", "Monthly (start)", "Oil rate (bbl/d) (average)", "Hours (sum)", "Days with data"]);
    expect(r.rows[0]).toEqual(["W1", serial(2026, 1, 1), 110, 44, 2]);
    expect(r.rows[1][0]).toBe("W1");
    expect(r.added[1]).toBe(true);
    expect(r.rows[1][2]).toBeCloseTo(110 + ((90 - 110) * 31) / 59, 4);
    expect(r.rows[2]).toEqual(["W1", serial(2026, 3, 1), 90, 24, 1]);
    expect(r.rows[3]).toEqual(["W2", serial(2026, 1, 1), 10, 24, 1]);
    expect(r.stats.addedPeriods).toBe(1);
  });
  it("skip rule leaves a column out", () => {
    const r = resample(["Date", "A", "B"], [[serial(2026, 1, 1), 1, 2]], { dateCol: 0, period: "year", aggs: { 2: "skip" }, fill: "none", countDays: false });
    expect(r.headers).toEqual(["Yearly (start)", "A (average)"]);
  });
});
