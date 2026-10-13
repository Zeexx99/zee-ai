import { describe, expect, it } from "vitest";
import { compareVersions, isNewerBuild, notesToShow } from "../qol/version";
import { addToMap, dayKey, estimateCost, formatCost, formatTokens, usageForMonth } from "../qol/usageLog";
import { addToHistory, recallStep } from "../qol/promptHistory";
import { buildSettingsExport, parseSettingsExport } from "../qol/settingsTransfer";
import { sortByUse } from "../qol/recent";
import { buildDiagnostics, errorHelp } from "../qol/diagnostics";

describe("version", () => {
  it("compares dotted versions numerically", () => {
    expect(compareVersions("1.10.0", "1.9.2")).toBe(1);
    expect(compareVersions("1.2", "1.2.0")).toBe(0);
    expect(compareVersions("1.2.0", "1.3.0")).toBe(-1);
  });
  it("shows release notes newer than the last seen version only", () => {
    const notes = [
      { version: "1.4.0", items: ["d"] },
      { version: "1.3.0", items: ["c"] },
      { version: "1.2.0", items: ["b"] },
    ];
    expect(notesToShow("1.2.0", "1.4.0", notes).map((n) => n.version)).toEqual(["1.4.0", "1.3.0"]);
    expect(notesToShow("1.4.0", "1.4.0", notes)).toEqual([]);
    expect(notesToShow("", "1.4.0", notes)).toEqual([]); // first install
    expect(notesToShow("0", "1.3.0", notes).map((n) => n.version)).toEqual(["1.3.0", "1.2.0"]);
  });
  it("detects a newer online build (ignores clock noise under a minute)", () => {
    expect(isNewerBuild({ version: "1", build: "2026-10-13T10:05:00Z" }, "2026-10-13T10:00:00Z")).toBe(true);
    expect(isNewerBuild({ version: "1", build: "2026-10-13T10:00:30Z" }, "2026-10-13T10:00:00Z")).toBe(false);
    expect(isNewerBuild({ version: "1", build: "2026-10-12T10:00:00Z" }, "2026-10-13T10:00:00Z")).toBe(false);
    expect(isNewerBuild(null, "2026-10-13T10:00:00Z")).toBe(false);
    expect(isNewerBuild({ version: "1", build: "x" }, "")).toBe(false);
  });
});

describe("usage log", () => {
  const d1 = new Date(2026, 9, 13, 9, 0);
  const d2 = new Date(2026, 9, 14, 9, 0);
  it("adds per day and totals a month", () => {
    let m = addToMap({}, { promptTokens: 100, outputTokens: 20 }, d1);
    m = addToMap(m, { promptTokens: 50, outputTokens: 5 }, d1);
    m = addToMap(m, { promptTokens: 10, outputTokens: 1 }, d2);
    expect(m[dayKey(d1)]).toEqual({ date: "2026-10-13", requests: 2, promptTokens: 150, outputTokens: 25 });
    const month = usageForMonth(d2, m);
    expect(month.requests).toBe(3);
    expect(month.promptTokens).toBe(160);
  });
  it("drops days older than ~2 months", () => {
    let m = {};
    for (let i = 0; i < 70; i++) m = addToMap(m, { promptTokens: 1, outputTokens: 1 }, new Date(2026, 0, 1 + i));
    expect(Object.keys(m).length).toBe(62);
    expect(Object.keys(m).sort()[0]).toBe("2026-01-09");
  });
  it("estimates cost only when prices are set", () => {
    expect(estimateCost({ promptTokens: 1_000_000, outputTokens: 0 }, { inPerM: 0, outPerM: 0 })).toBeNull();
    expect(estimateCost({ promptTokens: 2_000_000, outputTokens: 1_000_000 }, { inPerM: 0.3, outPerM: 2.5 })).toBeCloseTo(3.1, 5);
    expect(formatCost(0.004)).toBe("<$0.01");
    expect(formatCost(3.1)).toBe("$3.10");
    expect(formatTokens(45_200)).toBe("45.2k");
    expect(formatTokens(1_500_000)).toBe("1.5M");
  });
});

describe("prompt history", () => {
  it("keeps newest first without duplicates", () => {
    expect(addToHistory(["b", "a"], "a")).toEqual(["a", "b"]);
    expect(addToHistory(["a"], "  ")).toEqual(["a"]);
  });
  it("recalls older and newer prompts and returns to the draft", () => {
    const h = ["third", "second", "first"];
    let s = recallStep(h, -1, "older", "draft");
    expect(s).toEqual({ index: 0, text: "third" });
    s = recallStep(h, s.index, "older", "draft");
    expect(s.text).toBe("second");
    s = recallStep(h, 2, "older", "draft");
    expect(s.index).toBe(2); // stays on the oldest
    s = recallStep(h, 1, "newer", "draft");
    expect(s.text).toBe("third");
    s = recallStep(h, 0, "newer", "draft");
    expect(s).toEqual({ index: -1, text: "draft" });
  });
});

describe("settings export / import", () => {
  const defaults = { apiKey: "", model: "m", costGuardTokens: 20000, temperature: null as number | null, team: null as object | null, customActions: [] as unknown[], trackChanges: false };
  it("leaves secrets out unless asked", () => {
    const json = buildSettingsExport({ ...defaults, apiKey: "SECRET", model: "x" }, { includeSecrets: false, prompts: ["p1"] });
    expect(json).not.toContain("SECRET");
    const withKey = buildSettingsExport({ ...defaults, apiKey: "SECRET" }, { includeSecrets: true });
    expect(withKey).toContain("SECRET");
    const r = parseSettingsExport(json, defaults);
    expect(r.patch).toMatchObject({ model: "x", costGuardTokens: 20000 });
    expect("apiKey" in r.patch).toBe(false);
    expect(r.prompts).toEqual(["p1"]);
  });
  it("applies only known keys with matching types (cross-add-in import)", () => {
    const file = JSON.stringify({ zeeSettingsExport: 1, app: "excel", version: "1", exportedAt: "", settings: { model: "g", agentMode: true, costGuardTokens: "lots", temperature: 0.4, team: { a: 1 }, customActions: [1] } });
    const r = parseSettingsExport(file, defaults);
    expect(r.patch).toEqual({ model: "g", temperature: 0.4, team: { a: 1 }, customActions: [1] });
    expect(r.skipped.sort()).toEqual(["agentMode", "costGuardTokens"]);
    expect(r.from).toBe("excel");
  });
  it("rejects files that aren't exports", () => {
    expect(() => parseSettingsExport("{nope", defaults)).toThrow("not valid JSON");
    expect(() => parseSettingsExport('{"a":1}', defaults)).toThrow("isn't a Zee AI settings file");
  });
});

describe("most used first", () => {
  it("sorts by count, then recency, keeping unused items in order", () => {
    const items = ["a", "b", "c", "d"];
    const counts = { c: { n: 3, last: 1 }, b: { n: 3, last: 5 }, d: { n: 1, last: 9 } };
    expect(sortByUse(items, (x) => x, counts)).toEqual(["b", "c", "d", "a"]);
    expect(sortByUse(items, (x) => x, {})).toEqual(items);
  });
});

describe("diagnostics", () => {
  it("gives a hint per error kind", () => {
    expect(errorHelp("Network error – check your internet connection", "network").retryable).toBe(true);
    expect(errorHelp("Gemini error 429: Resource has been exhausted").hint).toContain("rate limit");
    expect(errorHelp("Your Gemini API key is invalid. Check it in Settings.", "auth").hint).toContain("Test connection");
    expect(errorHelp('Model "x" isn\'t available for this account.', "notFound").hint).toContain("another model");
    expect(errorHelp("Something odd").hint).toBeNull();
  });
  it("never includes secrets", () => {
    const r = buildDiagnostics({ error: "boom", kind: "server", model: "gemini-3.8-flash", authMode: "apiKey", extra: { Mode: "chat", Empty: "" } });
    expect(r).toContain("Error: boom");
    expect(r).toContain("Model: gemini-3.8-flash");
    expect(r).not.toContain("Empty:");
  });
});
