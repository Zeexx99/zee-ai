import { describe, expect, it } from "vitest";
import { evaluateFeatures, highestExcelApi, TOOL_FEATURE, FEATURES } from "../qol/requirements";
import { agentTools } from "../services/agentService";
import { chunkRows } from "../services/excelService";

const upTo = (max: number) => (set: string, v: string) => (set === "ExcelApi" ? Number(v.split(".")[1]) <= max : false);

describe("Excel version gating", () => {
  it("evaluates features against an API level", () => {
    const f = Object.fromEntries(evaluateFeatures(upTo(8)).map((x) => [x.id, x.ok]));
    expect(f.core).toBe(true);
    expect(f.validation).toBe(true);
    expect(f.formatUndo).toBe(false);
    expect(f.filterFind).toBe(false);
    expect(f.functions).toBe(false);
  });
  it("finds the highest ExcelApi version", () => {
    expect(highestExcelApi(upTo(12))).toBe("1.12");
    expect(highestExcelApi(() => false)).toBeNull();
    expect(highestExcelApi(() => {
      throw new Error("no Office");
    })).toBeNull();
  });
  it("every gated tool exists and maps to a known feature", () => {
    const names = new Set(agentTools({ webSearch: true }).map((t) => t.name));
    const ids = new Set(FEATURES.map((f) => f.id));
    for (const [tool, feat] of Object.entries(TOOL_FEATURE)) {
      expect(names.has(tool), tool).toBe(true);
      expect(ids.has(feat)).toBe(true);
    }
  });
  it("agent tools can be filtered by support", () => {
    const all = agentTools({ webSearch: false });
    const some = agentTools({ webSearch: false }, (t) => t !== "create_pivot_table");
    expect(some.length).toBe(all.length - 1);
  });
});

describe("chunked reads and writes", () => {
  it("splits big blocks into ~20k-cell row chunks", () => {
    expect(chunkRows(10)).toBe(2000);
    expect(chunkRows(30_000)).toBe(1);
    expect(chunkRows(0)).toBe(20_000);
  });
});
