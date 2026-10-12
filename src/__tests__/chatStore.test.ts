import { describe, expect, it } from "vitest";
import { historyCandidates } from "../services/chatStore";

const msgs = (n: number, promptLen = 10) =>
  Array.from({ length: n }, (_, i) => ({ id: i, role: i % 2 ? "assistant" : "user", status: "done", promptText: i % 2 ? undefined : "x".repeat(promptLen) }));

describe("chat history trimming", () => {
  it("first tries the full history (max 60, no streaming messages)", () => {
    const list = [...msgs(70), { id: 99, role: "assistant", status: "streaming" }];
    const first = [...historyCandidates(list)][0];
    expect(first).toHaveLength(60);
    expect(first.some((m) => m.status === "streaming")).toBe(false);
  });

  it("then shortens old prompts, then drops the oldest messages in pairs", () => {
    const all = [...historyCandidates(msgs(20, 5000))];
    expect(all[0][0].promptText!.length).toBe(5000);
    expect(all[1][0].promptText!.length).toBe(2001); // 2000 + "…"
    expect(all[1][all[1].length - 2].promptText!.length).toBe(5000); // the latest turns keep their context
    const lengths = all.slice(2).map((l) => l.length);
    expect(lengths[0]).toBe(10);
    expect(lengths.every((n) => n % 2 === 0)).toBe(true);
    expect(lengths[lengths.length - 1]).toBe(2);
    // The newest messages are always kept.
    expect(all[all.length - 1].map((m) => m.id)).toEqual([18, 19]);
  });
});
