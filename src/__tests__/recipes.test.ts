import { describe, expect, it } from "vitest";
import { aggsByIndex, describeStep, findColumn, newRecording, recipeFromText, recipeToText, recordMerge, recordStep, requireColumn, usesDataSheet, type Recipe } from "../services/recipes";

const sel = (sheet: string, rowIndex = 0, columnIndex = 0, rowCount = 100, columnCount = 5) => ({ sheet, rowIndex, columnIndex, rowCount, columnCount });

describe("recording", () => {
  it("first selection sheet becomes the data sheet; created sheets are referenced by step", () => {
    let r = newRecording();
    r = recordStep(r, { kind: "clean", options: { trimWhitespace: true, removeNonPrintable: true, caseMode: "none", convertNumbers: true, dateOrder: "off" } }, sel("Oct", 2, 0, 500, 6));
    r = recordStep(r, { kind: "resample", dateHeader: "Date", period: "month", aggs: {}, fill: "none", countDays: true }, sel("Oct", 2, 0, 500, 6), "Monthly Oct");
    r = recordStep(r, { kind: "chart", chartKind: "line", title: "" }, sel("Monthly Oct", 0, 1, 13, 3));
    r = recordStep(r, { kind: "aggregate", fn: "SUM" }, sel("Lookup", 0, 0, 1, 1));
    expect(r.dataSheet).toBe("Oct");
    expect(r.steps[0].target).toEqual({ sheet: { kind: "data" }, startRow: 2, startCol: 0, endCol: 5 });
    expect(r.outputs).toEqual({ "Monthly Oct": 1 });
    expect(r.steps[2].target?.sheet).toEqual({ kind: "step", index: 1 });
    expect(r.steps[2].target?.startCol).toBe(1);
    expect(r.steps[3].target).toMatchObject({ sheet: { kind: "fixed", name: "Lookup" }, single: true });
  });
  it("non-selection steps get no target", () => {
    const r = recordStep(newRecording(), { kind: "nodal", input: { ipr: { model: "pi", pr: 3000, pi: 1 }, well: { depth: 8000, tubingId: 2.441, whp: 150, gor: 500, waterCut: 0.5, api: 35, gasSg: 0.7, waterSg: 1.05, tWh: 100, tBh: 200 } } }, sel("X"), "Nodal");
    expect(r.steps[0].target).toBeUndefined();
    expect(r.dataSheet).toBeNull();
  });
  it("merge: first sheet is data, others fixed", () => {
    const r = recordMerge(newRecording(), [{ sheet: "Prod Oct", wellHeader: "Well", dateHeader: "Date" }, { sheet: "Tests", wellHeader: "Well" }], { how: "left", dateMatch: "asof", maxAgeDays: 90 }, "Merged");
    const m = r.steps[0];
    expect(m.kind).toBe("merge");
    if (m.kind === "merge") expect(m.sheets.map((s) => s.sheet)).toEqual([{ kind: "data" }, { kind: "fixed", name: "Tests" }]);
    expect(r.outputs).toEqual({ Merged: 0 });
  });
});

describe("replay helpers", () => {
  it("finds columns by header, loosely", () => {
    expect(findColumn(["Date", "Well Name", "Oil (bbl/d)"], "well name")).toBe(1);
    expect(findColumn(["Date", "WellName"], "Well name")).toBe(1);
    expect(() => requireColumn(["A"], "Date", "date")).toThrow(/Date/);
    expect(aggsByIndex(["Date", "Oil", "Hours"], { Hours: "sum", Missing: "max" })).toEqual({ 2: "sum" });
  });
  it("describes every step kind", () => {
    expect(describeStep({ kind: "removeEmpty", lines: "rows" })).toBe("Remove empty rows");
    expect(describeStep({ kind: "prompt", text: "Summarize", agent: true })).toBe("Agent: “Summarize”");
  });
});

describe("sharing", () => {
  const recipe: Recipe = { id: "x", name: "Monthly ESP", created: 1, steps: [{ kind: "fixImport", target: { sheet: { kind: "data" }, startRow: 0, startCol: 0, endCol: 4 } }, { kind: "productionQc" }] };
  it("round-trips through text", () => {
    const back = recipeFromText(recipeToText(recipe));
    expect(back.name).toBe("Monthly ESP");
    expect(back.steps).toEqual(recipe.steps);
    expect(back.id).not.toBe("x");
    expect(usesDataSheet(back)).toBe(true);
  });
  it("rejects junk and unknown steps", () => {
    expect(() => recipeFromText("hello")).toThrow(/isn't a Zee AI recipe/);
    expect(() => recipeFromText('ZEE-RECIPE v1\n{"name":"x","steps":[{"kind":"teleport"}]}')).toThrow(/Unknown step/);
  });
});
