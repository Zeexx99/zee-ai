import { describe, expect, it } from "vitest";
import { findDuplicateRows, findOutliers, findSpellingVariants, variantKey } from "../utils/dupes";
import { addMonths, detectStep, fillGaps, missingBetween } from "../utils/gaps";
import { detectDateOrder, fixCell, latinDigits, parseLooseDate, parseLooseNumber, planColumn } from "../utils/importFix";
import { buildSheetMap, formulaRefs, sheetMapDigest } from "../utils/sheetMap";
import { collectConstants, isValidName, knownConstant, replaceLiteral, suggestName } from "../utils/constants";
import { toSerial, serialToIso } from "../utils/dailyReport";

const D = (iso: string) => toSerial(iso) as number;

describe("duplicates, variants, outliers", () => {
  const rows = [
    ["W-12", 100],
    ["W12", 100],
    ["w-12", 100],
    ["Well 12", 90],
    ["W-012", 95],
    ["W-13", 200],
    ["", ""],
    ["", ""],
  ];
  it("finds exact duplicate rows (case/space-insensitive), ignoring blank rows", () => {
    expect(findDuplicateRows(rows)).toEqual([{ keep: 0, dupes: [2] }]);
    expect(findDuplicateRows(rows, [1])).toEqual([{ keep: 0, dupes: [1, 2] }]);
  });
  it("groups spelling variants of the same name", () => {
    expect(variantKey("Well 012")).toBe("w12");
    const g = findSpellingVariants(rows, 0);
    expect(g).toHaveLength(1);
    expect(g[0].canonical).toBe("W-12");
    expect(g[0].variants.map((v) => v.value).sort()).toEqual(["W-012", "W-12", "W12", "Well 12", "w-12"].sort());
  });
  it("flags robust outliers", () => {
    const col = [[100], [102], [98], [101], [99], [103], [97], [100], [1000], [-500]];
    const o = findOutliers(col, 0);
    expect(o.map((x) => [x.row, x.direction])).toEqual([
      [8, "high"],
      [9, "low"],
    ]);
    expect(findOutliers([[1], [2], [3]], 0)).toEqual([]);
  });
});

describe("gap filler", () => {
  it("detects daily and monthly steps", () => {
    expect(detectStep([D("2026-01-01"), D("2026-01-02"), D("2026-01-04")])).toEqual({ kind: "day", days: 1 });
    expect(detectStep([D("2026-01-01"), D("2026-02-01"), D("2026-03-01"), D("2026-05-01")])).toEqual({ kind: "month" });
    expect(serialToIso(addMonths(D("2026-01-31"), 1))).toBe("2026-02-28");
    expect(missingBetween(D("2026-01-01"), D("2026-04-01"), { kind: "month" }).map(serialToIso)).toEqual(["2026-02-01", "2026-03-01"]);
  });
  const rows = [
    [D("2026-01-01"), "W-1", 100, "ok"],
    [D("2026-01-02"), "W-1", "", "ok"],
    [D("2026-01-04"), "W-1", 130, "ok"],
    [D("2026-01-01"), "W-2", 50, "ok"],
    [D("2026-01-03"), "W-2", 70, "ok"],
  ];
  it("adds missing dates per well and interpolates", () => {
    const r = fillGaps(rows, { dateCol: 0, groupCol: 1, method: "linear", maxGapSteps: 0 });
    expect(r.stats.addedRows).toBe(2);
    expect(r.rows.map((x) => [serialToIso(x[0] as number), x[1], x[2]])).toEqual([
      ["2026-01-01", "W-1", 100],
      ["2026-01-02", "W-1", 110],
      ["2026-01-03", "W-1", 120],
      ["2026-01-04", "W-1", 130],
      ["2026-01-01", "W-2", 50],
      ["2026-01-02", "W-2", 60],
      ["2026-01-03", "W-2", 70],
    ]);
    expect(r.added).toEqual([false, false, true, false, false, true, false]);
    expect(r.filledCols[1]).toEqual([2]);
    expect(r.rows[2][3]).toBe(""); // text columns stay empty with "linear"
  });
  it("repeats the previous value, or only adds rows with blank", () => {
    const p = fillGaps(rows, { dateCol: 0, groupCol: 1, method: "previous", maxGapSteps: 0 });
    expect(p.rows[2]).toEqual([D("2026-01-03"), "W-1", 100, "ok"]);
    const b = fillGaps(rows, { dateCol: 0, groupCol: 1, method: "blank", maxGapSteps: 0 });
    expect(b.stats.filledCells).toBe(0);
    expect(b.rows[1][2]).toBe("");
  });
  it("skips gaps longer than the limit", () => {
    const r = fillGaps([[D("2026-01-01"), 1], [D("2026-01-02"), 2], [D("2026-01-03"), 3], [D("2026-01-10"), 10]], { dateCol: 0, method: "linear", maxGapSteps: 3 });
    expect(r.stats.addedRows).toBe(0);
    expect(r.stats.skippedGaps).toBe(1);
  });
});

describe("import fixer", () => {
  it("parses loose numbers with units, Arabic digits and negatives", () => {
    expect(parseLooseNumber("1,250.5 bbl/d")).toEqual({ value: 1250.5, unit: "bbl/d" });
    expect(parseLooseNumber("١٢٥٠ psi")).toEqual({ value: 1250, unit: "psi" });
    expect(parseLooseNumber("(12)")).toEqual({ value: -12, unit: "" });
    expect(parseLooseNumber("35 %")).toEqual({ value: 35, unit: "%" });
    expect(parseLooseNumber("1,25")).toEqual({ value: 1.25, unit: "" });
    expect(parseLooseNumber("60Hz")).toEqual({ value: 60, unit: "Hz" });
    expect(parseLooseNumber("W-12")).toBeNull();
    expect(latinDigits("٣٫٥")).toBe("3.5");
  });
  it("parses mixed date formats", () => {
    expect(serialToIso(parseLooseDate("13/10/2026", "DMY") as number)).toBe("2026-10-13");
    expect(serialToIso(parseLooseDate("10/13/2026", "DMY") as number)).toBe("2026-10-13"); // only valid as M/D
    expect(serialToIso(parseLooseDate("2026-10-13 08:00", "DMY") as number)).toBe("2026-10-13");
    expect(serialToIso(parseLooseDate("13-Oct-2026", "DMY") as number)).toBe("2026-10-13");
    expect(serialToIso(parseLooseDate("October 13, 2026", "DMY") as number)).toBe("2026-10-13");
    expect(parseLooseDate("31/02/2026", "DMY")).toBeNull();
    expect(detectDateOrder(["01/13/2026", "02/14/2026"])).toBe("MDY");
  });
  it("plans a column and fixes cells", () => {
    const p = planColumn(0, "Oil", ["1,200 bbl", "1,300 bbl", 1400, "", "900 m3"], []);
    expect(p).toMatchObject({ kind: "number", unit: "bbl", newHeader: "Oil (bbl)" });
    expect(fixCell("1,200 bbl", p)).toEqual({ value: 1200, changed: true });
    expect(fixCell("900 m3", p)).toMatchObject({ changed: false, otherUnit: "m3" });
    const d = planColumn(1, "Date", ["13/10/2026", "2026-10-14", "15.10.2026"], []);
    expect(d.kind).toBe("date");
    expect(fixCell("13/10/2026", d)).toMatchObject({ changed: true, numberFormat: "dd-mmm-yyyy" });
    const t = planColumn(2, "Well", ["W-1 ", "W-2"], []);
    expect(t.kind).toBe("text");
    expect(fixCell("W-1 ", t)).toEqual({ value: "W-1", changed: true });
  });
});

describe("sheet map", () => {
  it("extracts references", () => {
    expect(formulaRefs("=B2*Inputs!$C$3+'Ref Data'!A1+FACTOR*SUM(D:D)")).toEqual({ columns: ["B", "D"], sheets: ["Inputs", "Ref Data"], names: ["FACTOR"] });
  });
  it("groups formulas per column by shape", () => {
    const m = buildSheetMap({
      rowIndex: 0,
      columnIndex: 0,
      formulas: [["Oil", "Water", "WC %"], [100, 20, "=B2/(A2+B2)*100"], [110, 30, "=B3/(A3+B3)*100"], [120, 25, "=B4/(A4+B4)*100"]],
      formulasR1C1: [["Oil", "Water", "WC %"], [100, 20, "=RC[-1]/(RC[-2]+RC[-1])*100"], [110, 30, "=RC[-1]/(RC[-2]+RC[-1])*100"], [120, 25, "=RC[-1]/(RC[-2]+RC[-1])*100"]],
    });
    expect(m.columns.map((c) => c.role)).toEqual(["input", "input", "calculated"]);
    expect(m.groups[0]).toMatchObject({ column: "C", header: "WC %", count: 3, firstRow: 2, lastRow: 4, dependsOn: ["column B", "column A"] });
    expect(sheetMapDigest("Data", "A1:C4", m)).toContain("C \"WC %\" rows 2–4 (3×)");
  });
});

describe("constants → named inputs", () => {
  it("collects literals with counts and headers", () => {
    const c = collectConstants({ rowIndex: 0, columnIndex: 0, headers: ["Rate", "Volume"], formulas: [["Rate", "Volume"], [1, "=A2*5.615"], [2, "=A3*5.615+ROUND(A3,2)"]] });
    expect(c).toEqual([{ value: "5.615", count: 2, cells: ["B2", "B3"], example: "=A2*5.615", header: "Volume" }]);
  });
  it("suggests and validates names", () => {
    expect(knownConstant("5.6150")?.name).toBe("FT3_PER_BBL");
    expect(suggestName("5.615", "Volume", new Set())).toBe("FT3_PER_BBL");
    expect(suggestName("7.25", "Pump eff (%)", new Set(["k_pump_eff"]))).toBe("K_PUMP_EFF_2");
    expect(isValidName("BBL1")).toBe(false);
    expect(isValidName("R1C1")).toBe(false);
    expect(isValidName("FT3_PER_BBL")).toBe(true);
  });
  it("replaces only standalone literals", () => {
    expect(replaceLiteral("=A2*5.615+B5615*5.615", "5.615", "FT3_PER_BBL")).toEqual({ formula: "=A2*FT3_PER_BBL+B5615*FT3_PER_BBL", n: 2 });
    expect(replaceLiteral('=IF(A1>5.615,"5.615",SUM(A1:A5615))', "5.615", "K")).toEqual({ formula: '=IF(A1>K,"5.615",SUM(A1:A5615))', n: 1 });
    expect(replaceLiteral("='Sheet 5.615'!A1*5.615", "5.615", "K").formula).toBe("='Sheet 5.615'!A1*K");
  });
});
