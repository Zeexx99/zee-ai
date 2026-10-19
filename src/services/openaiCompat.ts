/**
 * SHARED MODULE – keep host-agnostic and import-free (identical copy in both Zee AI add-ins:
 * Word src/shared/openaiCompat.ts, Excel src/services/openaiCompat.ts).
 *
 * "Other API key" support: talks to any OpenAI-compatible Chat Completions API (Groq, OpenRouter,
 * OpenAI, Mistral, DeepSeek, Together, a local LM Studio / Ollama server…).
 *
 * The Gemini client keeps building Gemini-style request bodies. This module translates them to
 * `POST {baseUrl}/chat/completions` and translates the answers (plain or streamed) back into
 * Gemini-style responses, so chat, JSON output and agent tool calls work unchanged.
 * Gemini-only features (Google Search grounding, code execution, PDF input) are dropped.
 */

export const DEFAULT_OTHER_BASE_URL = "https://api.groq.com/openai/v1";
export const DEFAULT_OTHER_MODEL = "openai/gpt-oss-120b";

export interface OtherProvider {
  id: string;
  label: string;
  baseUrl: string;
  /** Where to create a key. */
  keyUrl: string;
  /** Model filled in when the provider is picked. */
  model: string;
}

/** Presets offered in Settings. "custom" keeps whatever URL the person typed. */
export const OTHER_PROVIDERS: OtherProvider[] = [
  { id: "groq", label: "Groq", baseUrl: "https://api.groq.com/openai/v1", keyUrl: "https://console.groq.com/keys", model: "openai/gpt-oss-120b" },
  { id: "openrouter", label: "OpenRouter", baseUrl: "https://openrouter.ai/api/v1", keyUrl: "https://openrouter.ai/settings/keys", model: "" },
  { id: "openai", label: "OpenAI", baseUrl: "https://api.openai.com/v1", keyUrl: "https://platform.openai.com/api-keys", model: "" },
  { id: "mistral", label: "Mistral", baseUrl: "https://api.mistral.ai/v1", keyUrl: "https://console.mistral.ai/api-keys", model: "" },
  { id: "deepseek", label: "DeepSeek", baseUrl: "https://api.deepseek.com/v1", keyUrl: "https://platform.deepseek.com/api_keys", model: "" },
];

/** Trims, removes trailing slashes and a pasted "/chat/completions" suffix. */
export function normalizeBaseUrl(url: string): string {
  return url
    .trim()
    .replace(/\/+$/, "")
    .replace(/\/chat\/completions$/i, "")
    .replace(/\/+$/, "");
}

/** Preset matching a base URL ("custom" when none does). */
export function providerOf(baseUrl: string): OtherProvider | undefined {
  const b = normalizeBaseUrl(baseUrl).toLowerCase();
  return OTHER_PROVIDERS.find((p) => p.baseUrl.toLowerCase() === b);
}

/** Friendly provider name for messages ("Groq", or the host name for a custom URL). */
export function providerName(baseUrl: string): string {
  const p = providerOf(baseUrl);
  if (p) return p.label;
  try {
    return new URL(normalizeBaseUrl(baseUrl)).hostname || "the provider";
  } catch {
    return "the provider";
  }
}

/* ------------------------------------------------------------------ Gemini-style shapes (structural) */

interface GPart {
  text?: string;
  thought?: boolean;
  inlineData?: { mimeType: string; data: string };
  functionCall?: { name: string; args?: Record<string, unknown>; id?: string };
  functionResponse?: { name: string; response: Record<string, unknown>; id?: string };
}
interface GContent {
  role: "user" | "model";
  parts: GPart[];
}
export interface GeminiStyleBody {
  systemInstruction?: { parts?: Array<{ text?: string }> };
  contents: GContent[];
  generationConfig?: Record<string, unknown>;
  tools?: Array<Record<string, unknown>>;
  toolConfig?: unknown;
}
export interface CompatResponse {
  candidates?: Array<{ content?: { role?: string; parts?: GPart[] }; finishReason?: string }>;
  usageMetadata?: { promptTokenCount?: number; candidatesTokenCount?: number; totalTokenCount?: number };
  error?: { message?: string };
}

/* ------------------------------------------------------------------ request */

/** Gemini schema (type: "OBJECT" …) -> JSON Schema (type: "object" …). */
export function toJsonSchema(schema: unknown): unknown {
  if (Array.isArray(schema)) return schema.map(toJsonSchema);
  if (!schema || typeof schema !== "object") return schema;
  const out: Record<string, unknown> = {};
  for (const [k, v] of Object.entries(schema as Record<string, unknown>)) {
    if (k === "nullable" || k === "propertyOrdering") continue;
    if (k === "type" && typeof v === "string") out.type = v.toLowerCase();
    else if (k === "properties" && v && typeof v === "object") {
      out.properties = Object.fromEntries(Object.entries(v as Record<string, unknown>).map(([pk, pv]) => [pk, toJsonSchema(pv)]));
    } else out[k] = toJsonSchema(v);
  }
  if (out.type === "object" && !out.properties) out.properties = {};
  return out;
}

type OAMessage = Record<string, unknown>;

/** Builds the Chat Completions request body from a Gemini-style body. */
export function toOpenAIRequest(body: GeminiStyleBody, model: string, stream: boolean): Record<string, unknown> {
  const messages: OAMessage[] = [];
  const cfg = body.generationConfig ?? {};
  let system = (body.systemInstruction?.parts ?? []).map((p) => p.text ?? "").join("\n").trim();
  const wantsJson = cfg.responseMimeType === "application/json";
  if (wantsJson) {
    const schema = cfg.responseSchema ? `\nIt must match this JSON schema:\n${JSON.stringify(toJsonSchema(cfg.responseSchema))}` : "";
    system += `${system ? "\n\n" : ""}Reply with one valid JSON value only – no Markdown fences, no explanation.${schema}`;
  }
  if (system) messages.push({ role: "system", content: system });

  let callSeq = 0;
  let pending: Array<{ id: string; name: string }> = [];
  for (const c of body.contents ?? []) {
    const parts = c.parts ?? [];
    if (c.role === "model") {
      const text = parts.filter((p) => !p.thought && typeof p.text === "string").map((p) => p.text).join("");
      const calls = parts
        .filter((p) => p.functionCall)
        .map((p) => ({ id: p.functionCall!.id || `call_${++callSeq}`, name: p.functionCall!.name, args: p.functionCall!.args ?? {} }));
      pending = calls.map(({ id, name }) => ({ id, name }));
      const msg: OAMessage = { role: "assistant", content: text || (calls.length ? null : "") };
      if (calls.length) msg.tool_calls = calls.map((k) => ({ id: k.id, type: "function", function: { name: k.name, arguments: JSON.stringify(k.args) } }));
      messages.push(msg);
      continue;
    }
    // User turn: tool results first (they must follow the assistant's tool_calls), then text / images.
    for (const p of parts) {
      if (!p.functionResponse) continue;
      const fr = p.functionResponse;
      let id = fr.id;
      const i = pending.findIndex((x) => (id ? x.id === id : x.name === fr.name));
      if (i >= 0) {
        id = pending[i].id;
        pending.splice(i, 1);
      }
      messages.push({ role: "tool", tool_call_id: id || `call_${++callSeq}`, content: JSON.stringify(fr.response ?? {}) });
    }
    const pieces: Array<Record<string, unknown>> = [];
    let hasImage = false;
    for (const p of parts) {
      if (p.functionResponse || p.thought) continue;
      if (p.inlineData) {
        if (p.inlineData.mimeType.startsWith("image/")) {
          hasImage = true;
          pieces.push({ type: "image_url", image_url: { url: `data:${p.inlineData.mimeType};base64,${p.inlineData.data}` } });
        } else {
          pieces.push({ type: "text", text: `[An attached ${p.inlineData.mimeType} file was left out: this provider can't read it. Switch to Gemini for PDFs.]` });
        }
      } else if (typeof p.text === "string" && p.text) pieces.push({ type: "text", text: p.text });
    }
    if (!pieces.length) continue;
    messages.push({ role: "user", content: hasImage ? pieces : pieces.map((x) => x.text as string).join("\n\n") });
  }

  const req: Record<string, unknown> = { model: model.trim(), messages, stream };
  if (stream) req.stream_options = { include_usage: true };
  if (typeof cfg.temperature === "number") req.temperature = cfg.temperature;
  if (typeof cfg.maxOutputTokens === "number") req.max_tokens = cfg.maxOutputTokens;
  if (wantsJson) req.response_format = { type: "json_object" };

  const fns = (body.tools ?? []).flatMap((t) => (Array.isArray(t.functionDeclarations) ? (t.functionDeclarations as Array<Record<string, unknown>>) : []));
  if (fns.length) {
    req.tools = fns.map((f) => ({
      type: "function",
      function: { name: f.name, description: f.description ?? "", parameters: f.parameters ? toJsonSchema(f.parameters) : { type: "object", properties: {} } },
    }));
    req.tool_choice = "auto";
  }
  return req;
}

/* ------------------------------------------------------------------ response */

interface OAUsage {
  prompt_tokens?: number;
  completion_tokens?: number;
  total_tokens?: number;
}
interface OAToolCall {
  index?: number;
  id?: string;
  function?: { name?: string; arguments?: string };
}

function usageOf(u: OAUsage | undefined): CompatResponse["usageMetadata"] | undefined {
  if (!u) return undefined;
  return { promptTokenCount: u.prompt_tokens ?? 0, candidatesTokenCount: u.completion_tokens ?? 0, totalTokenCount: u.total_tokens ?? (u.prompt_tokens ?? 0) + (u.completion_tokens ?? 0) };
}

function finishOf(reason: string | null | undefined): string | undefined {
  if (!reason) return undefined;
  if (reason === "length") return "MAX_TOKENS";
  if (reason === "content_filter") return "SAFETY";
  return "STOP";
}

function parseArgs(raw: string | undefined): Record<string, unknown> | null {
  if (!raw || !raw.trim()) return {};
  try {
    const v = JSON.parse(raw) as unknown;
    return v && typeof v === "object" && !Array.isArray(v) ? (v as Record<string, unknown>) : {};
  } catch {
    return null;
  }
}

/** Removes <think>…</think> blocks some reasoning models put in the answer text. */
export function stripThink(text: string): string {
  return text.replace(/<think>[\s\S]*?(<\/think>|$)/g, "").replace(/^\s+/, "");
}

function callParts(calls: OAToolCall[]): { parts: GPart[]; malformed: boolean } {
  const parts: GPart[] = [];
  let bad = 0;
  for (const c of calls) {
    const name = c.function?.name;
    if (!name) continue;
    const args = parseArgs(c.function?.arguments);
    if (args === null) {
      bad++;
      continue;
    }
    parts.push({ functionCall: { name, args, ...(c.id ? { id: c.id } : {}) } });
  }
  return { parts, malformed: bad > 0 && parts.length === 0 };
}

/** Non-streamed Chat Completions JSON -> Gemini-style response. */
export function fromOpenAIResponse(json: unknown): CompatResponse {
  const j = json as { choices?: Array<{ message?: { content?: string | null; tool_calls?: OAToolCall[] }; finish_reason?: string }>; usage?: OAUsage; error?: { message?: string } };
  if (j?.error) return { error: { message: j.error.message ?? "Provider error" } };
  const choice = j?.choices?.[0];
  const parts: GPart[] = [];
  const text = stripThink(choice?.message?.content ?? "");
  if (text) parts.push({ text });
  const calls = callParts(choice?.message?.tool_calls ?? []);
  parts.push(...calls.parts);
  return {
    candidates: [{ content: { role: "model", parts }, finishReason: calls.malformed ? "MALFORMED_FUNCTION_CALL" : finishOf(choice?.finish_reason) }],
    usageMetadata: usageOf(j?.usage),
  };
}

/**
 * Decodes a streamed Chat Completions response (Server-Sent Events) into Gemini-style chunks.
 * Text arrives chunk by chunk; tool calls are assembled and emitted by finish().
 */
export class OpenAIStreamDecoder {
  private buf = "";
  private text = "";
  private inThink = false;
  private calls = new Map<number, OAToolCall>();
  private finish?: string;
  private usage?: OAUsage;

  push(raw: string): CompatResponse[] {
    this.buf += raw.replace(/\r\n/g, "\n");
    const out: CompatResponse[] = [];
    let nl: number;
    while ((nl = this.buf.indexOf("\n")) >= 0) {
      const line = this.buf.slice(0, nl).trim();
      this.buf = this.buf.slice(nl + 1);
      const c = this.line(line);
      if (c) out.push(c);
    }
    return out;
  }

  /** Call once the stream has ended. */
  end(): CompatResponse[] {
    const out: CompatResponse[] = [];
    const rest = this.line(this.buf.trim());
    this.buf = "";
    if (rest) out.push(rest);
    const tail = this.text && !this.inThink && !"<think>".startsWith(this.text) ? this.text : "";
    this.text = "";
    const calls = callParts([...this.calls.entries()].sort((a, b) => a[0] - b[0]).map(([, c]) => c));
    const parts: GPart[] = [...(tail ? [{ text: tail }] : []), ...calls.parts];
    out.push({
      candidates: [{ content: { role: "model", parts }, finishReason: calls.malformed ? "MALFORMED_FUNCTION_CALL" : (this.finish ?? "STOP") }],
      usageMetadata: usageOf(this.usage),
    });
    return out;
  }

  private line(line: string): CompatResponse | null {
    if (!line.startsWith("data:")) return null;
    const data = line.slice(5).trim();
    if (!data || data === "[DONE]") return null;
    let j: {
      choices?: Array<{ delta?: { content?: string | null; tool_calls?: OAToolCall[] }; finish_reason?: string | null }>;
      usage?: OAUsage;
      x_groq?: { usage?: OAUsage };
      error?: { message?: string };
    };
    try {
      j = JSON.parse(data);
    } catch {
      return null;
    }
    if (j.error) return { error: { message: j.error.message ?? "Provider stream error" } };
    if (j.usage) this.usage = j.usage;
    if (j.x_groq?.usage) this.usage = j.x_groq.usage;
    const ch = j.choices?.[0];
    if (!ch) return null;
    if (ch.finish_reason) this.finish = finishOf(ch.finish_reason);
    for (const t of ch.delta?.tool_calls ?? []) {
      const i = t.index ?? 0;
      const cur = this.calls.get(i) ?? { id: undefined, function: { name: "", arguments: "" } };
      if (t.id) cur.id = t.id;
      if (t.function?.name) cur.function!.name = (cur.function!.name ?? "") + t.function.name;
      if (t.function?.arguments) cur.function!.arguments = (cur.function!.arguments ?? "") + t.function.arguments;
      this.calls.set(i, cur);
    }
    const visible = this.visibleText(ch.delta?.content ?? "");
    return visible ? { candidates: [{ content: { role: "model", parts: [{ text: visible }] } }] } : null;
  }

  /** Streams text while hiding <think>…</think> blocks that may be split across chunks. */
  private visibleText(delta: string): string {
    this.text += delta;
    let out = "";
    for (;;) {
      if (this.inThink) {
        const close = this.text.indexOf("</think>");
        if (close < 0) {
          this.text = this.text.slice(-8);
          return out;
        }
        this.text = this.text.slice(close + 8).replace(/^\s+/, "");
        this.inThink = false;
        continue;
      }
      const open = this.text.indexOf("<think>");
      if (open >= 0) {
        out += this.text.slice(0, open);
        this.text = this.text.slice(open + 7);
        this.inThink = true;
        continue;
      }
      // Hold back a possible partial "<think" at the end.
      let keep = 0;
      for (let k = Math.min(6, this.text.length); k > 0; k--) {
        if ("<think>".startsWith(this.text.slice(-k))) {
          keep = k;
          break;
        }
      }
      out += this.text.slice(0, this.text.length - keep);
      this.text = this.text.slice(this.text.length - keep);
      return out;
    }
  }
}

/** Model IDs from GET {baseUrl}/models. */
export function modelIdsFromList(json: unknown): string[] {
  const j = json as { data?: Array<{ id?: string; active?: boolean }>; models?: Array<{ id?: string; name?: string }> };
  const ids = (j?.data ?? []).filter((m) => m.active !== false).map((m) => m.id ?? "");
  const alt = (j?.models ?? []).map((m) => m.id ?? m.name ?? "");
  return Array.from(new Set([...ids, ...alt].filter((id) => id && !/(whisper|tts|embed|guard|orpheus|playai|transcri|moderation|dall-e|image)/i.test(id)))).sort();
}
