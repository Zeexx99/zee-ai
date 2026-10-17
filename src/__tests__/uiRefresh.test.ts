import { describe, expect, it } from "vitest";
import { DEFAULT_SETTINGS, normalizeSettings, MAX_PINNED_TOOLS } from "../services/settingsService";
import { searchCommands, togglePin, toolCommands } from "../components/commands";
import type { QuickActionHandlers } from "../components/QuickActions";
import { rampFrom } from "../theme";
import { orderStatus } from "../components/StatusStrip";
import { agentProgressLabel, formatMessageTime } from "../components/MessageBubble";

describe("dark theme by default", () => {
  it("new installs are dark", () => {
    expect(DEFAULT_SETTINGS.theme).toBe("dark");
    expect(normalizeSettings({}).theme).toBe("dark");
  });
  it("moves the old 'Match Office' default to dark once", () => {
    const s = normalizeSettings({ theme: "auto" });
    expect(s.theme).toBe("dark");
    expect(s.themeMigrated).toBe(true);
  });
  it("keeps a light choice made before the switch", () => {
    expect(normalizeSettings({ theme: "light" }).theme).toBe("light");
  });
  it("respects Match Office chosen after the switch", () => {
    expect(normalizeSettings({ theme: "auto", themeMigrated: true }).theme).toBe("auto");
  });
  it("repairs bad values", () => {
    const s = normalizeSettings({ theme: "neon" as never, accent: "pink" as never, density: "tiny" as never, pinnedTools: "x" as never, themeMigrated: true });
    expect(s.theme).toBe("dark");
    expect(s.accent).toBe("violet");
    expect(s.density).toBe("comfortable");
    expect(s.pinnedTools).toEqual([]);
  });
});

const noop = () => undefined;
const handlers = new Proxy({}, { get: () => noop }) as unknown as QuickActionHandlers;

describe("command search", () => {
  const cmds = toolCommands(handlers);
  it("has unique ids", () => {
    expect(new Set(cmds.map((c) => c.id)).size).toBe(cmds.length);
  });
  it("finds tools by name, keyword and abbreviation", () => {
    expect(searchCommands(cmds, "nodal")[0].id).toBe("petro.nodal");
    expect(searchCommands(cmds, "ipr").map((c) => c.id)).toContain("petro.nodal");
    expect(searchCommands(cmds, "resample")[0].id).toBe("data.resample");
    expect(searchCommands(cmds, "esp").length).toBeGreaterThanOrEqual(3);
    expect(searchCommands(cmds, "mtbf")[0].id).toBe("petro.espLogAnalyze");
  });
  it("needs every word to match", () => {
    expect(searchCommands(cmds, "esp nodal")).toEqual([]);
  });
  it("returns everything for an empty query", () => {
    expect(searchCommands(cmds, "  ").length).toBe(cmds.length);
  });
});

describe("pins", () => {
  it("adds, removes and caps", () => {
    expect(togglePin([], "a", 3)).toEqual(["a"]);
    expect(togglePin(["a", "b"], "a", 3)).toEqual(["b"]);
    expect(togglePin(["a", "b", "c"], "d", 3)).toEqual(["b", "c", "d"]);
    expect(MAX_PINNED_TOOLS).toBe(6);
  });
});

describe("accent ramps", () => {
  it("builds a 16-step ramp centred on the colour", () => {
    const r = rampFrom("#00897B");
    expect(r[80]).toBe("#00897B");
    expect(Object.keys(r)).toHaveLength(16);
    for (const v of Object.values(r)) expect(v).toMatch(/^#[0-9A-F]{6}$/);
    // darker below, lighter above
    expect(parseInt(r[10].slice(1, 3), 16) + parseInt(r[10].slice(3, 5), 16)).toBeLessThan(parseInt(r[160].slice(1, 3), 16) + parseInt(r[160].slice(3, 5), 16));
  });
});

describe("status strip", () => {
  it("puts warnings before info and keeps order otherwise", () => {
    const out = orderStatus([
      { intent: "info" as const, k: 1 },
      { intent: "warning" as const, k: 2 },
      { intent: "warning" as const, k: 3 },
    ]);
    expect(out.map((o) => o.k)).toEqual([2, 3, 1]);
  });
});

describe("chat timestamps & agent progress", () => {
  it("shows the time today and the date on other days", () => {
    const now = new Date(2026, 9, 17, 18, 30).getTime();
    expect(formatMessageTime(new Date(2026, 9, 17, 9, 5).getTime(), now)).toBe("09:05");
    expect(formatMessageTime(new Date(2026, 9, 12, 14, 0).getTime(), now)).toBe("12 Oct 14:00");
    expect(formatMessageTime(new Date(2025, 0, 3, 8, 0).getTime(), now)).toBe("3 Jan 2025 08:00");
  });
  it("labels the running step", () => {
    expect(agentProgressLabel([{ status: "done" }, { status: "running" }], true)).toBe("Working · step 2");
    expect(agentProgressLabel([{ status: "done" }, { status: "done" }], true)).toBe("Working · step 2 done, thinking…");
    expect(agentProgressLabel([{ status: "done" }], false)).toBeNull();
    expect(agentProgressLabel([], true)).toBeNull();
  });
});
