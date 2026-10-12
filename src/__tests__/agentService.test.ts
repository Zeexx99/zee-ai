import { describe, expect, it } from "vitest";
import { WRITE_TOOLS, agentTools, convertColumn, describeCall } from "../services/agentService";

const READ_TOOLS = new Set(["get_workbook_overview", "get_selection", "read_range", "find_cells", "fit_decline_curve", "esp_analysis", "web_search"]);

describe("agent tools", () => {
  it("every tool is either a read tool or needs approval", () => {
    for (const t of agentTools({ webSearch: true })) {
      expect(READ_TOOLS.has(t.name) || WRITE_TOOLS.has(t.name)).toBe(true);
    }
    for (const w of WRITE_TOOLS) expect(agentTools({ webSearch: false }).some((t) => t.name === w)).toBe(true);
  });

  it("web_search only when enabled", () => {
    expect(agentTools({ webSearch: false }).some((t) => t.name === "web_search")).toBe(false);
    expect(agentTools({ webSearch: true }).some((t) => t.name === "web_search")).toBe(true);
  });

  it("tool names are unique and every tool has a description", () => {
    const names = agentTools({ webSearch: true }).map((t) => t.name);
    expect(new Set(names).size).toBe(names.length);
    for (const t of agentTools({ webSearch: true })) expect(describeCall({ name: t.name, args: {} }).length).toBeGreaterThan(3);
  });

  it("describes new tools for the approval card", () => {
    expect(describeCall({ name: "sort_range", args: { address: "A1:D20", by: [{ column: "Oil", ascending: false }] } })).toBe("Sort A1:D20 by Oil ↓");
    expect(describeCall({ name: "filter_range", args: { address: "A1:D20", column: "Well", values: ["W-1", "W-2"] } })).toBe("Filter A1:D20 on Well: W-1, W-2");
    expect(describeCall({ name: "freeze_panes", args: { rows: 1, columns: 0 } })).toBe("Freeze 1 row(s), 0 column(s)");
    expect(describeCall({ name: "freeze_panes", args: { rows: 0, columns: 0 } })).toBe("Unfreeze panes");
    expect(describeCall({ name: "delete_rows_columns", args: { address: "7:9", kind: "rows", sheet: "Data" } })).toBe("Delete rows Data!7:9");
  });
});

describe("convert_units column", () => {
  it("keeps the header, converts numbers, skips text", () => {
    const r = convertColumn(["Pressure", 100, "200", "", "n/a"], "psi", "bar");
    expect(r.header).toBe("Pressure");
    expect(r.out[0]).toBe("Pressure (bar)");
    expect(r.out[1] as number).toBeCloseTo(6.894757, 5);
    expect(r.out[2] as number).toBeCloseTo(13.789515, 5);
    expect(r.out[3]).toBe("");
    expect(r.out[4]).toBe("");
    expect(r.converted).toBe(2);
    expect(r.skipped).toBe(1);
  });
  it("works without a header and rejects unknown units", () => {
    expect(convertColumn([32, 212], "F", "C").out).toEqual([0, 100]);
    expect(() => convertColumn([1], "furlong", "m")).toThrow();
  });
});
