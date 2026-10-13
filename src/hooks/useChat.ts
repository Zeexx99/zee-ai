import { useCallback, useEffect, useRef, useState } from "react";
import { agentStep, estimateTokens, GeminiError, streamChat, turnToContent, type Attachment, type ChatTurn, type GeminiContent, type Source, type Usage } from "../services/geminiService";
import type { ZeeSettings } from "../services/settingsService";
import { SYSTEM_PROMPT } from "../services/prompts";
import { AGENT_SYSTEM_PROMPT, WRITE_TOOLS, agentTools, describeCall, executeTool, previewCall } from "../services/agentService";
import { runAgentLoop, type ActionLog } from "../services/agentLoop";
import { clearChat, getDocId, loadChat, saveChat } from "../services/chatStore";
import { invalidateOverview, writeTargetHasData } from "../services/excelService";
import { dryRunPreview, type DryRunPreview } from "../services/dataTools";

export type { ActionLog };

/** What chat history keeps about an attachment (the file itself is never stored). */
export interface AttachmentInfo {
  name: string;
  mimeType: string;
  size: number;
}

export interface PendingApproval {
  calls: Array<{
    description: string;
    preview: string[][] | null;
    /** Address whose existing data would be replaced. */
    overwrites?: string | null;
    /** Dry run: the affected cells before → after (Settings → Agent mode). */
    dryRun?: DryRunPreview | null;
  }>;
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
  /** Files sent with a user message (names only). */
  attachments?: AttachmentInfo[];
  /** Web pages a grounded answer used. */
  sources?: Source[];
  /** The agent stopped at its step limit – show a Continue button. */
  canContinue?: boolean;
  /** GeminiError kind of a failed answer (drives the hint under the error). */
  errorKind?: string;
}

export interface SendInput {
  display: string;
  prompt: string;
  contextLabel?: string;
  /** Force plain chat even when agent mode is on (quick actions that only need an answer). */
  forceChat?: boolean;
  /** Force agent mode (e.g. "Continue" after the step limit). */
  forceAgent?: boolean;
  /** Images / PDFs to send with this message. */
  attachments?: Attachment[];
}

export interface ChatHookOptions {
  getSettings: () => ZeeSettings;
  onUsage: (u: Usage) => void;
  /** Ask the user before sending a large prompt. Resolve true to continue. */
  confirmLarge: (estimatedTokens: number) => Promise<boolean>;
}

const uid = () => Math.random().toString(36).slice(2, 10) + Date.now().toString(36);
export const MAX_AGENT_STEPS = 15;

/**
 * Build alternating user/model history; older turns drop their bulky sheet context.
 * Only the new message carries attachments – earlier ones are mentioned by name.
 */
export function buildHistory(messages: ChatMessage[], next: string, attachments?: Attachment[]): ChatTurn[] {
  const turns: ChatTurn[] = [];
  for (let i = 0; i < messages.length; i++) {
    const m = messages[i];
    const prev = messages[i - 1];
    if (m.role === "assistant" && m.status === "done" && prev?.role === "user") {
      const actions = m.actions?.filter((a) => a.status === "done").map((a) => a.description);
      const modelText = actions?.length ? `${m.text}\n\n(Actions performed: ${actions.join("; ")})` : m.text;
      const files = prev.attachments?.length ? `\n\n(Attached earlier: ${prev.attachments.map((a) => a.name).join(", ")} – no longer available)` : "";
      turns.push({ role: "user", text: (prev.promptText ?? prev.text) + files }, { role: "model", text: modelText || "(no text)" });
    }
  }
  const recent = turns.slice(-16);
  const compacted = recent.map((t, i) => {
    if (t.role !== "user" || i >= recent.length - 4) return t;
    const idx = t.text.search(/\n\n\[(Excel selection|Workbook overview)\]/);
    return idx > 0 ? { ...t, text: `${t.text.slice(0, idx)}\n\n(sheet context omitted)` } : t;
  });
  return [...compacted, { role: "user", text: next, ...(attachments?.length ? { attachments } : {}) }];
}

/** Rough token cost of attachments for the cost guard (images ≈ 1.1k tokens, PDFs ≈ 258 per page, ~1 page per 60 KB). */
export function attachmentTokens(files: Attachment[] | undefined): number {
  return (files ?? []).reduce((n, f) => n + (f.mimeType === "application/pdf" ? Math.max(258, Math.round(f.size / 60_000) * 258) : 1100), 0);
}

export function useChat({ getSettings, onUsage, confirmLarge }: ChatHookOptions) {
  const docIdRef = useRef<string>("default");
  const [messages, setMessages] = useState<ChatMessage[]>([]);
  const [busy, setBusy] = useState(false);
  const messagesRef = useRef<ChatMessage[]>([]);
  const abortRef = useRef<AbortController | null>(null);
  const approvalRef = useRef<((approved: boolean) => void) | null>(null);
  const saveTimer = useRef(0);

  // Persist history, debounced: streaming patches the message ~60×/s and serialising the whole
  // chat into localStorage on every frame made long conversations sluggish.
  const scheduleSave = useCallback((next: ChatMessage[], immediate = false) => {
    window.clearTimeout(saveTimer.current);
    if (immediate) saveChat(docIdRef.current, next);
    else saveTimer.current = window.setTimeout(() => saveChat(docIdRef.current, messagesRef.current), 400);
  }, []);
  useEffect(() => () => window.clearTimeout(saveTimer.current), []);

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
      scheduleSave(next, !next.some((m) => m.status === "streaming"));
      return next;
    });
  }, [scheduleSave]);

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
      const history = buildHistory(messagesRef.current, input.prompt, input.attachments);
      const tokens = estimateTokens(history.map((h) => h.text).join("\n") + (mode === "agent" ? AGENT_SYSTEM_PROMPT : SYSTEM_PROMPT)) + attachmentTokens(input.attachments);
      if (settings.costGuardTokens > 0 && tokens > settings.costGuardTokens) {
        if (!(await confirmLarge(tokens))) return null;
      }
      const attachments = input.attachments?.map(({ name, mimeType, size }) => ({ name, mimeType, size }));
      const userMsg: ChatMessage = { id: uid(), role: "user", text: input.display, promptText: input.prompt, contextLabel: input.contextLabel, status: "done", mode, ...(attachments?.length ? { attachments } : {}) };
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
      const err =
        e instanceof GeminiError
          ? e
          : (e as Error)?.name === "AbortError"
            ? new GeminiError("Stopped.", "aborted")
            : new GeminiError((e as Error)?.message ?? String(e));
      patch(asstId, (x) => {
        if (err.kind === "aborted") return { text: x.text ? `${x.text}\n\n_…stopped._` : "_Stopped._", status: "done", pending: undefined };
        return { text: err.message, status: "error", pending: undefined, errorKind: err.kind };
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
          useBuiltinTools: true,
          onText: (full) => {
            pending = full;
            if (!frame) frame = requestAnimationFrame(flush);
          },
        });
        if (frame) cancelAnimationFrame(frame);
        patch(asstId, { text: result.text || "_(empty response)_", status: "done", usage: result.usage, ...(result.sources ? { sources: result.sources } : {}) });
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
      const contents: GeminiContent[] = history.map(turnToContent);
      const tools = agentTools(settings);
      let actionCount = 0;

      // Stream the model's text, at most once per animation frame.
      let pendingText: string | null = null;
      let frame = 0;
      const flush = () => {
        frame = 0;
        if (pendingText === null) return;
        const text = pendingText;
        pendingText = null;
        patch(asstId, { text });
      };
      const onText = (t: string) => {
        pendingText = t;
        if (!frame) frame = requestAnimationFrame(flush);
      };

      try {
        const result = await runAgentLoop(contents, {
          step: (c, cb) => agentStep({ settings, contents: c, systemInstruction: AGENT_SYSTEM_PROMPT, tools, signal: controller.signal, onText: cb }),
          execute: (call) => executeTool(call, { settings, signal: controller.signal }),
          isWrite: (name) => WRITE_TOOLS.has(name),
          describe: describeCall,
          autoApprove: () => getSettings().autoApprove,
          approve: async (writes) => {
            const overwrites = await Promise.all(writes.map((c) => (c.name === "set_cells" ? writeTargetHasData(c.args).catch(() => null) : Promise.resolve(null))));
            const dry = getSettings().agentDryRun ? await Promise.all(writes.map((c) => dryRunPreview(c.name, c.args ?? {}))) : writes.map(() => null);
            patch(asstId, { pending: { calls: writes.map((c, i) => ({ description: describeCall(c), preview: previewCall(c), overwrites: overwrites[i], dryRun: dry[i] })) } });
            const ok = await new Promise<boolean>((resolve) => {
              approvalRef.current = resolve;
              controller.signal.addEventListener("abort", () => resolve(false), { once: true });
            });
            approvalRef.current = null;
            patch(asstId, { pending: undefined });
            return ok;
          },
          onText,
          // Track indexes locally: React may run state updaters later, so we can't read them back.
          startAction: (entry) => {
            const index = actionCount++;
            patch(asstId, (x) => ({ actions: [...(x.actions ?? []), entry] }));
            return index;
          },
          finishAction: (index, entry) => patch(asstId, (x) => ({ actions: (x.actions ?? []).map((a, i) => (i === index ? entry : a)) })),
          onUsage,
          afterWrites: invalidateOverview,
          signal: controller.signal,
          maxSteps: MAX_AGENT_STEPS,
        });
        if (frame) cancelAnimationFrame(frame);
        patch(asstId, {
          text: result.hitStepLimit
            ? `${result.text}\n\n_Paused after ${MAX_AGENT_STEPS} steps – click **Continue** to let me keep going._`
            : result.text || "_Done._",
          status: "done",
          usage: result.usage,
          canContinue: result.hitStepLimit,
        });
      } catch (e) {
        if (frame) cancelAnimationFrame(frame);
        fail(asstId, e);
      } finally {
        finish();
      }
    },
    [begin, fail, finish, getSettings, onUsage, patch]
  );

  const send = useCallback(
    (input: SendInput) => ((getSettings().agentMode || input.forceAgent) && !input.forceChat ? sendAgent(input) : sendChat(input)),
    [getSettings, sendAgent, sendChat]
  );

  /** "Continue" after the agent paused at its step limit. */
  const continueAgent = useCallback(() => {
    update((list) => list.map((m) => (m.canContinue ? { ...m, canContinue: false } : m)));
    return send({ display: "Continue", prompt: "Continue the task where you left off. Check the workbook state first if needed.", forceAgent: true });
  }, [send, update]);

  /**
   * Regenerate / Retry: removes the answer and its prompt and sends the same prompt again (same
   * sheet context as the first time, same mode). Files attached to it aren't kept, so those can't be resent.
   */
  const retry = useCallback(
    async (asstId: string) => {
      if (abortRef.current) return;
      const list = messagesRef.current;
      const i = list.findIndex((m) => m.id === asstId);
      if (i < 1 || list[i].role !== "assistant" || list[i - 1].role !== "user") return;
      const user = list[i - 1];
      const next = list.filter((_, k) => k !== i && k !== i - 1);
      messagesRef.current = next; // begin() builds the history from the ref right away
      update(() => next);
      const mode = list[i].mode ?? user.mode;
      await send({
        display: user.text,
        prompt: user.promptText ?? user.text,
        contextLabel: user.contextLabel,
        forceAgent: mode === "agent",
        forceChat: mode === "chat",
      });
    },
    [send, update]
  );

  const respondToApproval = useCallback((approved: boolean) => approvalRef.current?.(approved), []);
  const stop = useCallback(() => abortRef.current?.abort(), []);
  const clear = useCallback(() => {
    abortRef.current?.abort();
    clearChat(docIdRef.current);
    update(() => []);
  }, [update]);

  return { messages, busy, send, retry, stop, clear, respondToApproval, continueAgent };
}
