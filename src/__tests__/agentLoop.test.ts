import { describe, expect, it } from "vitest";
import { runAgentLoop, type ActionLog, type AgentLoopDeps } from "../services/agentLoop";
import { GeminiError, type AgentStepResult, type GeminiContent } from "../services/geminiService";
import type { ToolCall } from "../services/agentService";

const textStep = (text: string): AgentStepResult => ({ content: { role: "model", parts: [{ text }] }, text, calls: [], usage: { promptTokens: 10, outputTokens: 5, totalTokens: 15 } });
const callStep = (calls: ToolCall[], text = ""): AgentStepResult => ({
  content: { role: "model", parts: calls.map((c) => ({ functionCall: { name: c.name, args: c.args } })) },
  text,
  calls,
  usage: { promptTokens: 10, outputTokens: 5, totalTokens: 15 },
});

function harness(steps: Array<AgentStepResult | Error>, opts: Partial<AgentLoopDeps> = {}) {
  const actions: ActionLog[] = [];
  const executed: string[] = [];
  const approvals: string[][] = [];
  const texts: string[] = [];
  let i = 0;
  const deps: AgentLoopDeps = {
    step: async (_c, onText) => {
      const s = steps[Math.min(i++, steps.length - 1)];
      if (s instanceof Error) throw s;
      onText(s.text);
      return s;
    },
    execute: async (c) => {
      executed.push(c.name);
      return c.name === "bad" ? { error: "boom" } : { ok: true };
    },
    isWrite: (n) => n.startsWith("set") || n === "bad",
    describe: (c) => c.name,
    approve: async (calls) => {
      approvals.push(calls.map((c) => c.name));
      return true;
    },
    autoApprove: () => false,
    onText: (t) => texts.push(t),
    startAction: (e) => actions.push(e) - 1,
    finishAction: (idx, e) => {
      actions[idx] = e;
    },
    onUsage: () => undefined,
    signal: new AbortController().signal,
    maxSteps: 5,
    sleep: async () => undefined,
    ...opts,
  };
  return { deps, actions, executed, approvals, texts, calls: () => i };
}

describe("agent loop", () => {
  it("runs read tools without approval and finishes with the model's answer", async () => {
    const h = harness([callStep([{ name: "read_range", args: {} }]), textStep("Total is 42.")]);
    const contents: GeminiContent[] = [{ role: "user", parts: [{ text: "sum" }] }];
    const r = await runAgentLoop(contents, h.deps);
    expect(r.text).toBe("Total is 42.");
    expect(r.hitStepLimit).toBe(false);
    expect(r.usage.totalTokens).toBe(30);
    expect(h.approvals).toEqual([]);
    expect(h.actions).toEqual([{ description: "read_range", status: "read", detail: undefined }]);
    // user, model(call), user(function response), model(answer)
    expect(contents.map((c) => c.role)).toEqual(["user", "model", "user", "model"]);
    expect(contents[2].parts[0].functionResponse?.name).toBe("read_range");
  });

  it("asks once per batch of writes; rejected writes are not executed", async () => {
    const h = harness([callStep([{ name: "read_range", args: {} }, { name: "set_cells", args: {} }]), textStep("OK")], { approve: async () => false });
    await runAgentLoop([{ role: "user", parts: [{ text: "x" }] }], h.deps);
    expect(h.executed).toEqual(["read_range"]);
    expect(h.actions.map((a) => a.status)).toEqual(["read", "rejected"]);
  });

  it("skips approval with auto-approve and reports tool errors", async () => {
    const h = harness([callStep([{ name: "set_cells", args: {} }, { name: "bad", args: {} }]), textStep("done")], { autoApprove: () => true });
    await runAgentLoop([{ role: "user", parts: [{ text: "x" }] }], h.deps);
    expect(h.approvals).toEqual([]);
    expect(h.actions).toEqual([
      { description: "set_cells", status: "done", detail: undefined },
      { description: "bad", status: "error", detail: "boom" },
    ]);
  });

  it("retries a step once after a dropped connection", async () => {
    const h = harness([new GeminiError("dropped", "network"), textStep("recovered")]);
    const r = await runAgentLoop([{ role: "user", parts: [{ text: "x" }] }], h.deps);
    expect(r.text).toBe("recovered");
    expect(h.calls()).toBe(2);
  });

  it("does not retry auth or quota errors", async () => {
    const h = harness([new GeminiError("bad key", "auth"), textStep("never")]);
    await expect(runAgentLoop([{ role: "user", parts: [{ text: "x" }] }], h.deps)).rejects.toThrow("bad key");
    expect(h.calls()).toBe(1);
  });

  it("stops at the step limit and offers Continue", async () => {
    const h = harness([callStep([{ name: "read_range", args: {} }], "Working…")], { maxSteps: 3 });
    const r = await runAgentLoop([{ role: "user", parts: [{ text: "x" }] }], h.deps);
    expect(r.hitStepLimit).toBe(true);
    expect(r.text).toBe("Working…");
    expect(h.calls()).toBe(3);
  });

  it("keeps the last explanation when a later step has no text", async () => {
    const h = harness([callStep([{ name: "read_range", args: {} }], "Reading the data."), textStep("")]);
    const r = await runAgentLoop([{ role: "user", parts: [{ text: "x" }] }], h.deps);
    expect(r.text).toBe("Reading the data.");
  });
});
