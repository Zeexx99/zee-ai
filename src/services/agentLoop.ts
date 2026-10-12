/**
 * The agent loop, separated from React so it can be unit-tested:
 * ask Gemini → run read tools → ask the user to approve write tools → send results back → repeat.
 */
import { GeminiError, type AgentStepResult, type GeminiContent, type Usage } from "./geminiService";
import type { ToolCall } from "./agentService";

export interface ActionLog {
  description: string;
  status: "running" | "done" | "rejected" | "error" | "read";
  detail?: string;
}

export interface AgentLoopDeps {
  /** One model step (streams text through onText). */
  step: (contents: GeminiContent[], onText: (text: string) => void) => Promise<AgentStepResult>;
  /** Run a tool; must not throw (errors come back as { error }). */
  execute: (call: ToolCall) => Promise<Record<string, unknown>>;
  isWrite: (name: string) => boolean;
  describe: (call: ToolCall) => string;
  /** Ask the user about a batch of write calls. Resolve true to apply. Not called when auto-approve is on. */
  approve: (calls: ToolCall[]) => Promise<boolean>;
  autoApprove: () => boolean;
  /** Text of the current step, while it streams. */
  onText: (text: string) => void;
  /** Add an action row; returns its index so it can be updated when the tool finishes. */
  startAction: (entry: ActionLog) => number;
  finishAction: (index: number, entry: ActionLog) => void;
  onUsage: (u: Usage) => void;
  afterWrites?: () => void;
  signal: AbortSignal;
  maxSteps: number;
  /** Retries of a failed model step on network/server errors (default 1). */
  stepRetries?: number;
  sleep?: (ms: number) => Promise<void>;
}

export interface AgentLoopResult {
  text: string;
  usage: Usage;
  /** True when the step limit was reached with work possibly left (offer "Continue"). */
  hitStepLimit: boolean;
}

const RETRYABLE: GeminiError["kind"][] = ["network", "server"];
const defaultSleep = (ms: number) => new Promise<void>((r) => setTimeout(r, ms));

/** Run one model step, retrying once on a dropped connection or a server hiccup. */
async function stepWithRetry(deps: AgentLoopDeps, contents: GeminiContent[], onText: (t: string) => void): Promise<AgentStepResult> {
  const retries = deps.stepRetries ?? 1;
  for (let attempt = 0; ; attempt++) {
    try {
      return await deps.step(contents, onText);
    } catch (e) {
      const retryable = e instanceof GeminiError && RETRYABLE.includes(e.kind);
      if (!retryable || attempt >= retries || deps.signal.aborted) throw e;
      await (deps.sleep ?? defaultSleep)(1500 * (attempt + 1));
      if (deps.signal.aborted) throw new GeminiError("Stopped.", "aborted");
    }
  }
}

export async function runAgentLoop(contents: GeminiContent[], deps: AgentLoopDeps): Promise<AgentLoopResult> {
  const total: Usage = { promptTokens: 0, outputTokens: 0, totalTokens: 0 };
  let lastText = "";
  for (let step = 0; step < deps.maxSteps; step++) {
    const res = await stepWithRetry(deps, contents, (t) => deps.onText(t));
    if (res.usage) {
      total.promptTokens += res.usage.promptTokens;
      total.outputTokens += res.usage.outputTokens;
      total.totalTokens += res.usage.totalTokens;
      deps.onUsage(res.usage);
    }
    contents.push(res.content);
    if (res.text.trim()) lastText = res.text;
    deps.onText(lastText);
    if (!res.calls.length) return { text: lastText, usage: total, hitStepLimit: false };

    const calls: ToolCall[] = res.calls;
    const writes = calls.filter((c) => deps.isWrite(c.name));
    let approved = true;
    if (writes.length && !deps.autoApprove()) {
      approved = await deps.approve(writes);
      if (deps.signal.aborted) throw new GeminiError("Stopped.", "aborted");
    }

    const responses: GeminiContent = { role: "user", parts: [] };
    for (const c of calls) {
      if (deps.signal.aborted) throw new GeminiError("Stopped.", "aborted");
      const isWrite = deps.isWrite(c.name);
      const description = deps.describe(c);
      let result: Record<string, unknown>;
      if (isWrite && !approved) {
        result = { error: "The user rejected this change. Do not retry it; ask what they want instead." };
        const i = deps.startAction({ description, status: "rejected" });
        deps.finishAction(i, { description, status: "rejected" });
      } else {
        const i = deps.startAction({ description, status: "running" });
        result = await deps.execute(c);
        deps.finishAction(i, { description, status: result.error ? "error" : isWrite ? "done" : "read", detail: result.error ? String(result.error) : undefined });
      }
      responses.parts.push({ functionResponse: { name: c.name, response: result, ...(c.id ? { id: c.id } : {}) } });
    }
    if (writes.length && approved) deps.afterWrites?.();
    contents.push(responses);
  }
  return { text: lastText, usage: total, hitStepLimit: true };
}
