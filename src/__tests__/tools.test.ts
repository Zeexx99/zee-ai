import { describe, expect, it } from "vitest";
import { buildDailySummary, guessDailyMapping, serialLabel, serialToIso, summaryTable, toMarkdownTable, toSerial } from "../utils/dailyReport";
import { analyzeEspLog, guessEspLogMapping } from "../utils/espFailures";
import { auditFormulas, numericLiterals } from "../utils/formulaAudit";
import { findViolations, profileColumns, sanitizeRules } from "../utils/columnRules";

const D = (iso: string) => toSerial(iso) as number;

describe("daily production report", () => {
  const rows = [
    [D("2026-10-11"), "W-1", 1000, 200, 500],
    [D("2026-10-12"), "W-1", 1000, 250, 500],
    [D("2026-10-13"), "W-1", 700, 300, 400],
    [D("2026-10-12"), "W-2", 500, 0, 100],
    [D("2026-10-13"), "W-2", 520, 0, 110],
    [D("2026-10-12"), "W-3", 300, 100, 50],
  ];
  const m = { date: 0, well: 1, oil: 2, water: 3, gas: 4 };
  it("uses the latest date and each well's previous reading", () => {
    const s = buildDailySummary(rows, m);
    expect(serialToIso(s.reportDate)).toBe("2026-10-13");
    expect(s.wells.map((w) => w.well)).toEqual(["W-1", "W-2"]);
    const w1 = s.wells[0];
    expect(w1.oilPrev).toBe(1000);
    expect(w1.oilChangePct).toBeCloseTo(-30, 5);
    expect(w1.waterCut).toBeCloseTo(30, 5);
    expect(s.missingWells).toEqual(["W-3"]);
    expect(s.totals.oil).toBe(1220);
    expect(s.highlights.join(" ")).toContain("W-1: oil down 30%");
    expect(s.highlights.join(" ")).toContain("W-3");
  });
  it("builds table + markdown", () => {
    const s = buildDailySummary(rows, m, D("2026-10-12"));
    const t = summaryTable(s, { oil: "Oil (bbl/d)", water: "Water (bbl/d)", gas: "Gas (Mscf/d)" });
    expect(t[0]).toEqual(["Well", "Oil (bbl/d)", "Water (bbl/d)", "Water cut %", "Gas (Mscf/d)", "GOR", "Prev. oil", "Oil change %"]);
    expect(t[t.length - 1][0]).toBe("Total");
    expect(toMarkdownTable(t).split("\n")[1]).toContain("---");
  });
  it("guesses columns and parses dates", () => {
    expect(guessDailyMapping(["Date", "Well Name", "Oil (BOPD)", "Water cut %", "Water (BWPD)", "Gas (Mscf/d)"])).toEqual({ date: 0, well: 1, oil: 2, water: 4, gas: 5 });
    expect(toSerial(46308)).toBe(46308);
    expect(serialLabel(D("2026-10-13"))).toBe("13 Oct 2026");
    expect(toSerial("not a date")).toBeNull();
  });
});

describe("ESP failure log", () => {
  const asOf = D("2026-10-13");
  const rows = [
    ["W-1", "TD1750", D("2025-01-01"), D("2025-07-01"), "Failed", "", "Motor", "Electrical"],
    ["W-1", "TD1750", D("2025-07-10"), D("2026-01-10"), "Failed", "", "Pump", "Scale"],
    ["W-1", "TD1750", D("2026-01-20"), "", "Running", "", "", ""],
    ["W-2", "DN3000", D("2025-03-01"), D("2026-03-01"), "Pulled – no failure", "", "", ""],
    ["W-3", "DN3000", D("2025-06-01"), D("2026-02-01"), "", "", "Cable / MLE", "Electrical"],
    ["", "", "", "", "", "", "", ""],
  ];
  it("computes run life, MTBF and groups", () => {
    const m = guessEspLogMapping(["Well", "ESP Model", "Install Date", "Pull / Fail Date", "Status", "Run Life (days)", "Failed Component", "Failure Cause"]);
    expect(m).toMatchObject({ well: 0, model: 1, install: 2, end: 3, status: 4, component: 6, cause: 7 });
    const r = analyzeEspLog(rows, m as never, asOf);
    expect(r.installs).toBe(5);
    expect(r.failures).toBe(3); // the pulled one isn't a failure; W-3 counts (date + cause)
    expect(r.running).toBe(1);
    expect(r.byCause[0]).toMatchObject({ name: "Electrical", failures: 2 });
    expect(r.repeatWells).toEqual(["W-1"]);
    expect(r.mtbf).toBe(Math.round(r.totalRunDays / 3));
    expect(r.monthly.find((x) => x.month === "2025-07")).toMatchObject({ failures: 1, installs: 1 });
    expect(r.skipped).toBe(0);
  });
});

describe("formula audit", () => {
  it("finds hard-coded numbers but not references, strings or counts", () => {
    expect(numericLiterals("=A2*5.615+B$3/0.433")).toEqual(["5.615", "0.433"]);
    expect(numericLiterals('=IF(A1>1,"x 45",0)')).toEqual([]);
    expect(numericLiterals("=SUM(A1:A20)*100")).toEqual([]);
    expect(numericLiterals("='Data 2'!C5*7")).toEqual(["7"]);
  });
  it("flags errors, broken refs, inconsistent formulas and typed values", () => {
    const formulas = [["Rate"], ["=B2*2"], ["=B3*2"], ["=B4*2"], ["=B5*3+1"], [55], ["=#REF!*2"]];
    const r1c1 = [["Rate"], ["=RC[1]*2"], ["=RC[1]*2"], ["=RC[1]*2"], ["=RC[1]*3+1"], [55], ["=#REF!*2"]];
    const values = [["Rate"], [2], [4], ["#DIV/0!"], [9], [55], ["#REF!"]];
    const valueTypes = [["String"], ["Double"], ["Double"], ["Error"], ["Double"], ["Double"], ["Error"]];
    const issues = auditFormulas({ rowIndex: 0, columnIndex: 0, formulas, formulasR1C1: r1c1, values, valueTypes });
    const at = (addr: string) => issues.filter((i) => i.address === addr).map((i) => i.kind).sort();
    expect(at("A4")).toEqual(["error"]);
    expect(at("A7")).toContain("brokenRef");
    expect(at("A5")).toContain("inconsistent");
    expect(at("A6")).toEqual(["valueInFormulaColumn"]);
    expect(issues[0].severity).toBe("high");
  });
});

describe("column rules", () => {
  const headers = ["Well", "Water cut %", "Status"];
  const rows = [
    ["W-1", 20, "Producing"],
    ["W-2", 140, "Producing"],
    ["W-3", -5, "Shut-in"],
    ["W-4", "", "Prodcing"],
  ];
  it("profiles columns", () => {
    const p = profileColumns(headers, rows);
    expect(p[1]).toMatchObject({ kind: "number", min: -5, max: 140, blanks: 1 });
    expect(p[2].kind).toBe("text");
    expect(p[2].top[0]).toEqual({ value: "Producing", n: 2 });
  });
  it("sanitizes rules and finds violations", () => {
    const p = profileColumns(headers, rows);
    const rules = sanitizeRules(
      [
        { column: "water cut %", type: "number", min: 0, max: 100, reason: "%" },
        { column: "Status", type: "list", allowed: ["Producing", "Shut-in", "Producing"], reason: "" },
        { column: "Nope", type: "number", min: 0, reason: "" },
        { column: "Well", type: "none", reason: "id" },
      ],
      p
    );
    expect(rules.map((r) => r.column)).toEqual(["Water cut %", "Status"]);
    expect(rules[1].allowed).toEqual(["Producing", "Shut-in"]);
    expect(findViolations(rows, 1, rules[0])).toEqual([1, 2]);
    expect(findViolations(rows, 2, rules[1])).toEqual([3]);
  });
});
