import { describe, expect, it } from "vitest";
import { FUNCTION_CATALOG, buildFormula, canApplyDirectly, findFunction, guessHeader, isCellRef, missingArgs, toArg } from "../functions/catalog";
import functionsJson from "../../public/functions.json";

const f = (id: string) => findFunction(id)!;

describe("function catalog", () => {
  it("every catalog entry is registered in functions.json with the same parameter count", () => {
    const meta = new Map((functionsJson as { functions: { id: string; parameters: unknown[] }[] }).functions.map((m) => [m.id, m.parameters.length]));
    for (const spec of FUNCTION_CATALOG) expect(meta.get(spec.id), spec.id).toBe(spec.args.length);
  });

  it("builds per-row formulas from selection cells", () => {
    expect(buildFormula(f("CLASSIFY"), { categories: "A, B" }, ["B2"])).toBe('=ZEE.CLASSIFY(B2, "A, B")');
    expect(buildFormula(f("TRANSLATE"), {}, ["B2"])).toBe('=ZEE.TRANSLATE(B2, "English")');
    expect(buildFormula(f("CLEAN"), {}, ["B2"])).toBe("=ZEE.CLEAN(B2)");
    expect(buildFormula(f("EXPLAIN"), {}, ["C3"])).toBe("=ZEE.EXPLAIN(FORMULATEXT(C3))");
  });

  it("typed inputs override and don't use up selection columns", () => {
    expect(buildFormula(f("VOGEL"), { Pr: "$B$1" }, ["A2", "B2", "C2"])).toBe("=ZEE.VOGEL($B$1, A2, B2, C2)");
    expect(missingArgs(f("VOGEL"), { Pr: "B1" }, true, 3)).toHaveLength(0);
    expect(missingArgs(f("VOGEL"), {}, true, 3).map((a) => a.name)).toEqual(["Pwf_test"]);
  });

  it("range functions and quoting", () => {
    expect(buildFormula(f("FLAG"), { rule: 'say "hi"' }, [], "$A$2:$D$9")).toBe('=ZEE.FLAG($A$2:$D$9, "say ""hi""")');
    expect(buildFormula(f("UNITS"), { value: "100", from: "psi", to: "bar" })).toBe('=ZEE.UNITS(100, "psi", "bar")');
    expect(toArg("=B2*2", "input")).toBe("B2*2");
    expect(toArg("pump tripped", "input")).toBe('"pump tripped"');
  });

  it("detects references", () => {
    expect(isCellRef("'My Sheet'!A1:B2")).toBe(true);
    expect(isCellRef("A:A")).toBe(true);
    expect(isCellRef("pump")).toBe(false);
  });

  it("decides when a click can apply directly", () => {
    expect(canApplyDirectly(f("CLEAN"), 1)).toBe(true);
    expect(canApplyDirectly(f("CLASSIFY"), 1)).toBe(false);
    expect(canApplyDirectly(f("VOGEL"), 4)).toBe(true);
    expect(canApplyDirectly(f("VOGEL"), 2)).toBe(false);
    expect(canApplyDirectly(f("FORMULA"), 3)).toBe(false);
  });

  it("guesses header rows", () => {
    expect(guessHeader([["Well", "Rate"], ["A", 5]])).toBe(true);
    expect(guessHeader([[1, 2], [3, 4]])).toBe(false);
  });
});
