import { describe, expect, it } from "vitest";
import { canRevertAlone, stepSheets, timeAgo } from "../utils/undoHistory";
import { describeLines } from "../utils/clean";

describe("change history", () => {
  const steps = [
    { kind: "cells", sheetName: "Data" },
    { kind: "addedChart", sheetName: "Charts" },
    { kind: "addedName" },
    { kind: "cells", sheetName: "Data" },
    { kind: "multi", steps: [{ kind: "format", sheetName: "Summary" }, { kind: "addedTable" }] },
  ];
  it("knows which sheets a step touches", () => {
    expect(stepSheets(steps[0])).toEqual(["Data"]);
    expect(stepSheets(steps[2])).toEqual([]);
    expect(stepSheets(steps[4])).toEqual(["Summary"]);
    expect(stepSheets({ kind: "weird" })).toBeNull();
  });
  it("allows reverting a step alone only when no newer step touches its sheet", () => {
    expect(canRevertAlone(steps, 0)).toBe(false); // step 3 changed "Data" later
    expect(canRevertAlone(steps, 1)).toBe(true);
    expect(canRevertAlone(steps, 2)).toBe(true);
    expect(canRevertAlone(steps, 3)).toBe(true);
    expect(canRevertAlone(steps, 4)).toBe(true); // newest
    expect(canRevertAlone([...steps, { kind: "unknown" }], 1)).toBe(false);
    expect(canRevertAlone(steps, 9)).toBe(false);
  });
  it("formats relative times", () => {
    const now = Date.UTC(2026, 9, 13, 12, 0, 0);
    expect(timeAgo(now - 10_000, now)).toBe("just now");
    expect(timeAgo(now - 5 * 60_000, now)).toBe("5 min ago");
    expect(timeAgo(now - 3 * 3600_000, now)).toBe("3 h ago");
  });
});

describe("describeLines", () => {
  it("groups consecutive lines and shortens long lists", () => {
    expect(describeLines([4, 8, 11, 12, 13, 19], (i) => String(i + 1))).toBe("5, 9, 12–14, 20");
    expect(describeLines([0, 2, 4, 6, 8, 10, 12, 14], (i) => String(i + 1), 3)).toBe("1, 3, 5 and 5 more");
    expect(describeLines([1, 2], (i) => "ABC"[i])).toBe("B–C");
  });
});
