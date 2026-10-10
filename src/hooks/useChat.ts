import { useCallback, useEffect, useRef, useState } from "react";
import { agentStep, estimateTokens, GeminiError, streamChat, type ChatTurn, type GeminiContent, type Usage } from "../services/geminiService";
import type { ZeeSettings } from "../services/settingsService";
import { SYSTEM_PROMPT } from "../services/prompts";
import { AGENT_SYSTEM_PROMPT, AGENT_TOOLS, WRITE_TOOLS, describeCall, executeTool, previewCall, type ToolCall } from "../services/agentService";
import { clearChat, getDocId, loadChat, saveChat } from "../services/chatStore";
import { invalidateOverview } from "../services/excelService";

export interface ActionLog {
  description: string;
  status: "done" | "rejected" | "error" | "read";
  detail?: string;
}

export interface PendingApproval {
  calls: Array<{ description: string; preview: string[][] | null }>;
}

export interface ChatMessage {
  id: string;
  role: "user" | "assistant";
  /** What the user sees. */
  text: string;
  /** What was actually sent (display text + sheet context). */
  promptText?: string;
  /** e.g. "Sheet1!A1:D20" – shown as a chip on user messages. */
  contextLabel?: string;
  status: "streaming" | "done" | "error";
  mode?: "chat" | "agent";
  model?: string;
  usage?: Usage;
  actions?: ActionLog[];
  pending?: PendingApproval;
}

export interface SendInput {
  display: string;
  prompt: string;
  contextLabel?: string;
  /** Force plain chat even when agent mode is on (quick actions that only need an answer). */
  forceChat?: boolean;
}

export interface ChatHookOptions {
  getSettings: () => ZeeSettings;
  onUsage: (u: Usage) => void;
  /** Ask the user before sending a large prompt. Resolve true to continue. */
  confirmLarge: (estimatedTokens: number) => Promise<boolean>;
}

const uid = () => Math.random().toString(36).slice(2, 10) + Date.now().toString(36);
const MAX_AGENT_STEPS = 12;

/** Build alternating user/model history; older turns drop their bulky sheet context. */
function buildHistory(messages: ChatMessage[], next: string): ChatTurn[] {
  const turns: ChatTurn[] = [];
  for (let i = 0; i < messages.length; i++) {
    const m = messages[i];
    const prev = messages[i - 1];
    if (m.role === "assistant" && m.status === "done" && prev?.role === "user") {
      const actions = m.actions?.filter((a) => a.status === "done").map((a) => a.description);
      const modelText = actions?.length ? `${m.text}\n\n(Actions performed: ${actions.join("; ")})` : m.text;
      turns.push({ role: "user", text: prev.promptText ?? prev.text }, { role: "model", text: modelText || "(no text)" });
    }
  }
  const recent = turns.slice(-16);
  const compacted = recent.map((t, i) => {
    if (t.role !== "user" || i >= recent.length - 4) return t;
    const idx = t.text.search(/\n\n\[(Excel selection|Workbook overview)\]/);
    return idx > 0 ? { ...t, text: `${t.text.slice(0, idx)}\n\n(sheet context omitted)` } : t;
  });
  return [...compacted, { role: "user", text: next }];
}

export function useChat({ getSettings, onUsage, confirmLarge }: ChatHookOptions) {
  const docIdRef = useRef<string>("default");
  const [messages, setMessages] = useState<ChatMessage[]>([]);
  const [busy, setBusy] = useState(false);
  const messagesRef = useRef<ChatMessage[]>([]);
  const abortRef = useRef<AbortController | null>(null);
  const approvalRef = useRef<((approved: boolean) => void) | null>(null);

  // Load this workbook's history once.
  useEffect(() => {
    docIdRef.current = getDocId();
    const saved = loadChat<ChatMessage>(docIdRef.current).map((m) => (m.pending ? { ...m, pending: undefined } : m));
    messagesRef.current = saved;
    setMessages(saved);
  }, []);

  const update = useCallback((fn: (m: ChatMessage[]) => ChatMessage[]) => {
    setMessages((prev) => {
      const next = fn(prev);
      messagesRef.current = next;
      saveChat(docIdRef.current, next);
      return next;
    });
  }, []);

  const patch = useCallback(
    (id: string, p: Partial<ChatMessage> | ((m: ChatMessage) => Partial<ChatMessage>)) => {
      update((list) => list.map((x) => (x.id === id ? { ...x, ...(typeof p === "function" ? p(x) : p) } : x)));
    },
    [update]
  );

  /** Shared start: cost guard, add messages, controller. Returns null if cancelled. */
  const begin = useCallback(
    async (input: SendInput, mode: "chat" | "agent") => {
      if (abortRef.current) return null;
      const settings = getSettings();
      const history = buildHistory(messagesRef.current, input.prompt);
      const tokens = estimateTokens(history.map((h) => h.text).join("\n") + (mode === "agent" ? AGENT_SYSTEM_PROMPT : SYSTEM_PROMPT));
      if (settings.costGuardTokens > 0 && tokens > settings.costGuardTokens) {
        if (!(await confirmLarge(tokens))) return null;
      }
      const userMsg: ChatMessage = { id: uid(), role: "user", text: input.display, promptText: input.prompt, contextLabel: input.contextLabel, status: "done", mode };
      const asstId = uid();
      update((m) => [...m, userMsg, { id: asstId, role: "assistant", text: "", status: "streaming", model: settings.model, mode, actions: mode === "agent" ? [] : undefined }]);
      const controller = new AbortController();
      abortRef.current = controller;
      setBusy(true);
      return { settings, history, asstId, controller };
    },
    [confirmLarge, getSettings, update]
  );

  const finish = useCallback(() => {
    abortRef.current = null;
    approvalRef.current = null;
    setBusy(false);
  }, []);

  const fail = useCallback(
    (asstId: string, e: unknown) => {
      const err = e instanceof GeminiError ? e : new GeminiError((e as Error)?.message ?? String(e));
      patch(asstId, (x) => {
        if (err.kind === "aborted") return { text: x.text ? `${x.text}\n\n_…stopped._` : "_Stopped._", status: "done", pending: undefined };
        return { text: err.message, status: "error", pending: undefined };
      });
    },
    [patch]
  );

  /* ---------------------------------------------------------------- plain chat (streaming) */
  const sendChat = useCallback(
    async (input: SendInput) => {
      const ctx = await begin(input, "chat");
      if (!ctx) return;
      const { settings, history, asstId, controller } = ctx;
      let pending: string | null = null;
      let frame = 0;
      const flush = () => {
        frame = 0;
        if (pending === null) return;
        const text = pending;
        pending = null;
        patch(asstId, { text });
      };
      try {
        const result = await streamChat({
          settings,
          history,
          systemInstruction: SYSTEM_PROMPT,
          signal: controller.signal,
          onText: (full) => {
            pending = full;
            if (!frame) frame = requestAnimationFrame(flush);
          },
        });
        if (frame) cancelAnimationFrame(frame);
        patch(asstId, { text: result.text || "_(empty response)_", status: "done", usage: result.usage });
        if (result.usage) onUsage(result.usage);
      } catch (e) {
        if (frame) cancelAnimationFrame(frame);
        fail(asstId, e);
      } finally {
        finish();
      }
    },
    [begin, fail, finish, onUsage, patch]
  );

  /* ---------------------------------------------------------------- agent mode (tools) */
  const sendAgent = useCallback(
    async (input: SendInput) => {
      const ctx = await begin(input, "agent");
      if (!ctx) return;
      const { settings, history, asstId, controller } = ctx;
      const contents: GeminiContent[] = history.map((t) => ({ role: t.role, parts: [{ text: t.text }] }));
      const total: Usage = { promptTokens: 0, outputTokens: 0, totalTokens: 0 };
      const log = (entry: ActionLog) => patch(asstId, (x) => ({ actions: [...(x.actions ?? []), entry] }));
      try {
        for (let step = 0; step < MAX_AGENT_STEPS; step++) {
          const res = await agentStep({ settings, contents, systemInstruction: AGENT_SYSTEM_PROMPT, tools: AGENT_TOOLS, signal: controller.signal });
          if (res.usage) {
            total.promptTokens += res.usage.promptTokens;
            total.outputTokens += res.usage.outputTokens;
            total.totalTokens += res.usage.totalTokens;
            onUsage(res.usage);
          }
          contents.push(res.content);
          if (res.text) patch(asstId, { text: res.text });
          if (!res.calls.length) {
            patch(asstId, { text: res.text || "_Done._", status: "done", usage: total });
            return;
          }

          const calls: ToolCall[] = res.calls;
          const writes = calls.filter((c) => WRITE_TOOLS.has(c.name));
          let approved = true;
          if (writes.length && !getSettings().autoApprove) {
            patch(asstId, { pending: { calls: writes.map((c) => ({ description: describeCall(c), preview: previewCall(c) })) } });
            approved = await new Promise<boolean>((resolve) => {
              approvalRef.current = resolve;
              controller.signal.addEventListener("abort", () => resolve(false), { once: true });
            });
            approvalRef.current = null;
            patch(asstId, { pending: undefined });
            if (controller.signal.aborted) throw new GeminiError("Stopped.", "aborted");
          }

          const responses: GeminiContent = { role: "user", parts: [] };
          for (const c of calls) {
            const isWrite = WRITE_TOOLS.has(c.name);
            let result: Record<string, unknown>;
            if (isWrite && !approved) {
              result = { error: "The user rejected this change. Do not retry it; ask what they want instead." };
              log({ description: describeCall(c), status: "rejected" });
            } else {
              result = await executeTool(c);
              log({ description: describeCall(c), status: result.error ? "error" : isWrite ? "done" : "read", detail: result.error ? String(result.error) : undefined });
            }
            responses.parts.push({ functionResponse: { name: c.name, response: result, ...(c.id ? { id: c.id } : {}) } });
          }
          if (writes.length) invalidateOverview();
          contents.push(responses);
        }
        patch(asstId, (x) => ({ text: `${x.text}\n\n_Stopped after ${MAX_AGENT_STEPS} steps. Ask me to continue if needed._`, status: "done", usage: total }));
      } catch (e) {
        fail(asstId, e);
      } finally {
        finish();
      }
    },
    [begin, fail, finish, getSettings, onUsage, patch]
  );

  const send = useCallback(
    (input: SendInput) => (getSettings().agentMode && !input.forceChat ? sendAgent(input) : sendChat(input)),
    [getSettings, sendAgent, sendChat]
  );

  const respondToApproval = useCallback((approved: boolean) => approvalRef.current?.(approved), []);
  const stop = useCallback(() => abortRef.current?.abort(), []);
  const clear = useCallback(() => {
    abortRef.current?.abort();
    clearChat(docIdRef.current);
    update(() => []);
  }, [update]);

  return { messages, busy, send, stop, clear, respondToApproval };
}
