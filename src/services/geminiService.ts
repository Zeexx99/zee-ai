/**
 * Gemini client (direct REST, no SDK) supporting two auth paths:
 *  - API key  -> Gemini Developer API  (generativelanguage.googleapis.com)
 *  - Google sign-in -> Vertex AI       (aiplatform.googleapis.com, user's GCP project quota)
 * Streaming uses `streamGenerateContent?alt=sse` and is parsed incrementally.
 */
import type { ZeeSettings } from "./settingsService";
import { getValidGoogleToken } from "./authService";
import { SseParser } from "../utils/sse";

export interface ChatTurn {
  role: "user" | "model";
  text: string;
}

export interface Usage {
  promptTokens: number;
  outputTokens: number;
  totalTokens: number;
}

export type GeminiErrorKind = "config" | "auth" | "quota" | "notFound" | "safety" | "network" | "server" | "aborted" | "unknown";

export class GeminiError extends Error {
  readonly kind: GeminiErrorKind;
  readonly status?: number;
  constructor(message: string, kind: GeminiErrorKind = "unknown", status?: number) {
    super(message);
    this.name = "GeminiError";
    this.kind = kind;
    this.status = status;
  }
}

interface Endpoint {
  url: (method: string) => string;
  headers: Record<string, string>;
  label: string;
}

const GEMINI_API = "https://generativelanguage.googleapis.com/v1beta";

function resolveEndpoint(s: ZeeSettings, model: string): Endpoint {
  const m = encodeURIComponent(model.trim());
  if (!m) throw new GeminiError("Choose a model in Settings.", "config");
  if (s.authMode === "proxy") {
    const base = s.proxyUrl.trim().replace(/\/+$/, "");
    if (!base) throw new GeminiError("Enter your team proxy URL in Settings.", "config");
    const headers: Record<string, string> = { "Content-Type": "application/json" };
    if (s.proxyToken.trim()) headers.Authorization = `Bearer ${s.proxyToken.trim()}`;
    return { url: (method) => `${base}/v1beta/models/${m}:${method}`, headers, label: "team proxy" };
  }
  if (s.authMode === "apiKey") {
    const key = s.apiKey.trim();
    if (!key) throw new GeminiError("Add your Gemini API key in Settings (gear icon).", "config");
    return {
      url: (method) => `${GEMINI_API}/models/${m}:${method}`,
      headers: { "Content-Type": "application/json", "x-goog-api-key": key },
      label: "Gemini API",
    };
  }
  const token = getValidGoogleToken();
  if (!token) throw new GeminiError("Your Google sign-in has expired. Open Settings and sign in again.", "auth");
  const project = s.gcpProjectId.trim();
  const location = (s.gcpLocation.trim() || "global").toLowerCase();
  if (!project) throw new GeminiError("Enter your Google Cloud project ID in Settings.", "config");
  const host = location === "global" ? "aiplatform.googleapis.com" : `${location}-aiplatform.googleapis.com`;
  return {
    url: (method) =>
      `https://${host}/v1/projects/${encodeURIComponent(project)}/locations/${location}/publishers/google/models/${m}:${method}`,
    headers: { "Content-Type": "application/json", Authorization: `Bearer ${token}`, "x-goog-user-project": project },
    label: `Vertex AI (${project})`,
  };
}

function buildBody(history: ChatTurn[], systemInstruction: string, s: ZeeSettings, extraConfig: Record<string, unknown> = {}) {
  return buildRawBody(
    history.map((t) => ({ role: t.role, parts: [{ text: t.text }] })),
    systemInstruction,
    s,
    extraConfig
  );
}

function buildRawBody(contents: GeminiContent[], systemInstruction: string, s: ZeeSettings, extraConfig: Record<string, unknown> = {}) {
  const generationConfig: Record<string, unknown> = { ...extraConfig };
  if (s.temperature !== null && s.temperature !== undefined) generationConfig.temperature = s.temperature;
  return {
    systemInstruction: { parts: [{ text: systemInstruction }] },
    contents,
    generationConfig,
  };
}

/** Raw Gemini content (used by agent mode to round-trip function calls + thought signatures). */
export interface GeminiPart {
  text?: string;
  thought?: boolean;
  thoughtSignature?: string;
  functionCall?: { name: string; args?: Record<string, unknown>; id?: string };
  functionResponse?: { name: string; response: Record<string, unknown>; id?: string };
}
export interface GeminiContent {
  role: "user" | "model";
  parts: GeminiPart[];
}

/** Rough token estimate (~4 characters per token) used by the cost guard. */
export function estimateTokens(text: string): number {
  return Math.ceil(text.length / 4);
}

interface ApiErrorBody {
  error?: { code?: number; message?: string; status?: string; details?: Array<Record<string, unknown>> };
}

function retryDelayMs(body: ApiErrorBody | null): number | null {
  const info = body?.error?.details?.find((d) => String(d["@type"] ?? "").includes("RetryInfo"));
  const delay = info?.retryDelay;
  if (typeof delay === "string") {
    const sec = parseFloat(delay);
    if (Number.isFinite(sec)) return Math.min(sec * 1000, 30_000);
  }
  return null;
}

/** Map an HTTP error to a user-friendly GeminiError. */
export function toGeminiError(status: number, bodyText: string, model: string): GeminiError {
  let body: ApiErrorBody | null = null;
  try {
    body = JSON.parse(bodyText) as ApiErrorBody;
  } catch {
    /* not JSON */
  }
  const apiMsg = body?.error?.message ?? bodyText.slice(0, 300);
  if (status === 400 && /api key/i.test(apiMsg)) return new GeminiError("Your Gemini API key is invalid. Check it in Settings.", "auth", status);
  if (status === 401) return new GeminiError("Authentication failed – sign in again or re-check your key.", "auth", status);
  if (status === 403)
    return new GeminiError(
      `Permission denied: ${apiMsg} (Check that the API is enabled for your project and your account has the "Vertex AI User" role.)`,
      "auth",
      status
    );
  if (status === 404) return new GeminiError(`Model "${model}" isn't available for this account. Pick another model in Settings (or click Refresh).`, "notFound", status);
  if (status === 429)
    return new GeminiError(
      "Rate limit or quota exceeded. Wait a minute, switch to a Flash-Lite model, or check billing/quota in Google AI Studio / Cloud console.",
      "quota",
      status
    );
  if (status >= 500) return new GeminiError(`Gemini is temporarily unavailable (${status}). Please try again.`, "server", status);
  return new GeminiError(`Gemini error ${status}: ${apiMsg}`, "unknown", status);
}

function sleep(ms: number, signal?: AbortSignal): Promise<void> {
  return new Promise((resolve, reject) => {
    if (signal?.aborted) {
      reject(new GeminiError("Stopped.", "aborted"));
      return;
    }
    const onAbort = () => {
      clearTimeout(t);
      reject(new GeminiError("Stopped.", "aborted"));
    };
    const t = setTimeout(() => {
      signal?.removeEventListener("abort", onAbort);
      resolve();
    }, ms);
    signal?.addEventListener("abort", onAbort, { once: true });
  });
}

/** fetch with automatic retry on 429 / 5xx (before any streamed bytes are consumed). */
async function fetchWithRetry(url: string, init: RequestInit, model: string, retries = 2): Promise<Response> {
  for (let attempt = 0; ; attempt++) {
    let res: Response;
    try {
      res = await fetch(url, init);
    } catch (e) {
      if ((e as Error).name === "AbortError") throw new GeminiError("Stopped.", "aborted");
      throw new GeminiError("Network error – check your internet connection or firewall.", "network");
    }
    if (res.ok) return res;
    const text = await res.text();
    const retryable = res.status === 429 || res.status === 500 || res.status === 503;
    if (retryable && attempt < retries) {
      let body: ApiErrorBody | null = null;
      try {
        body = JSON.parse(text) as ApiErrorBody;
      } catch {
        /* ignore */
      }
      const wait = retryDelayMs(body) ?? 1000 * 2 ** attempt + Math.random() * 400;
      await sleep(wait, init.signal ?? undefined);
      continue;
    }
    throw toGeminiError(res.status, text, model);
  }
}

interface Candidate {
  content?: { role?: string; parts?: GeminiPart[] };
  finishReason?: string;
}
interface GenerateResponse {
  candidates?: Candidate[];
  promptFeedback?: { blockReason?: string };
  usageMetadata?: { promptTokenCount?: number; candidatesTokenCount?: number; totalTokenCount?: number; thoughtsTokenCount?: number };
  error?: { message?: string };
}

function textOf(r: GenerateResponse): string {
  return (r.candidates?.[0]?.content?.parts ?? [])
    .filter((p) => !p.thought && typeof p.text === "string")
    .map((p) => p.text)
    .join("");
}

function usageOf(r: GenerateResponse): Usage | undefined {
  const u = r.usageMetadata;
  if (!u) return undefined;
  const out = (u.candidatesTokenCount ?? 0) + (u.thoughtsTokenCount ?? 0);
  return { promptTokens: u.promptTokenCount ?? 0, outputTokens: out, totalTokens: u.totalTokenCount ?? (u.promptTokenCount ?? 0) + out };
}

export interface StreamResult {
  text: string;
  usage?: Usage;
  finishReason?: string;
}

export interface StreamOptions {
  settings: ZeeSettings;
  history: ChatTurn[];
  systemInstruction: string;
  signal?: AbortSignal;
  onText: (fullText: string) => void;
}

/** Stream a chat completion. Resolves with the full text when the stream ends. */
export async function streamChat(opts: StreamOptions): Promise<StreamResult> {
  const { settings, signal } = opts;
  const model = settings.model;
  const ep = resolveEndpoint(settings, model);
  const res = await fetchWithRetry(
    `${ep.url("streamGenerateContent")}?alt=sse`,
    { method: "POST", headers: ep.headers, body: JSON.stringify(buildBody(opts.history, opts.systemInstruction, settings)), signal },
    model
  );
  if (!res.body) throw new GeminiError("Streaming is not supported in this environment.", "unknown");

  const reader = res.body.getReader();
  const decoder = new TextDecoder();
  const parser = new SseParser();
  let full = "";
  let usage: Usage | undefined;
  let finishReason: string | undefined;

  const handle = (data: string) => {
    let chunk: GenerateResponse;
    try {
      chunk = JSON.parse(data) as GenerateResponse;
    } catch {
      return;
    }
    if (chunk.error) throw new GeminiError(chunk.error.message ?? "Gemini stream error", "server");
    if (chunk.promptFeedback?.blockReason)
      throw new GeminiError(`The request was blocked by Gemini safety filters (${chunk.promptFeedback.blockReason}).`, "safety");
    const delta = textOf(chunk);
    if (delta) {
      full += delta;
      opts.onText(full);
    }
    usage = usageOf(chunk) ?? usage;
    finishReason = chunk.candidates?.[0]?.finishReason ?? finishReason;
  };

  try {
    for (;;) {
      const { done, value } = await reader.read();
      if (done) break;
      parser.push(decoder.decode(value, { stream: true })).forEach(handle);
    }
    parser.push(decoder.decode()).forEach(handle);
    parser.flush().forEach(handle);
  } catch (e) {
    if ((e as Error).name === "AbortError") throw new GeminiError("Stopped.", "aborted");
    throw e;
  }

  if (!full && finishReason === "SAFETY") throw new GeminiError("The response was blocked by Gemini safety filters.", "safety");
  if (finishReason === "MAX_TOKENS") {
    full += "\n\n_…response truncated (max output tokens reached)._";
    opts.onText(full);
  }
  return { text: full, usage, finishReason };
}

/** One-shot JSON generation (used for bulk row transforms). */
export async function generateJson<T>(opts: {
  settings: ZeeSettings;
  prompt: string;
  systemInstruction: string;
  schema: Record<string, unknown>;
  signal?: AbortSignal;
}): Promise<{ data: T; usage?: Usage }> {
  const { settings } = opts;
  const ep = resolveEndpoint(settings, settings.model);
  const body = buildBody([{ role: "user", text: opts.prompt }], opts.systemInstruction, settings, {
    responseMimeType: "application/json",
    responseSchema: opts.schema,
  });
  const res = await fetchWithRetry(
    ep.url("generateContent"),
    { method: "POST", headers: ep.headers, body: JSON.stringify(body), signal: opts.signal },
    settings.model,
    3
  );
  const json = (await res.json()) as GenerateResponse;
  const text = textOf(json);
  try {
    return { data: JSON.parse(text) as T, usage: usageOf(json) };
  } catch {
    throw new GeminiError("Gemini returned malformed JSON. Try again or use a Pro model.", "unknown");
  }
}

/** List generateContent-capable Gemini models for an API key. */
export async function listModels(apiKey: string): Promise<string[]> {
  if (!apiKey.trim()) throw new GeminiError("Enter an API key first.", "config");
  let res: Response;
  try {
    res = await fetch(`${GEMINI_API}/models?pageSize=1000`, { headers: { "x-goog-api-key": apiKey.trim() } });
  } catch {
    throw new GeminiError("Network error – check your internet connection.", "network");
  }
  if (!res.ok) throw toGeminiError(res.status, await res.text(), "");
  const json = (await res.json()) as { models?: Array<{ name: string; supportedGenerationMethods?: string[] }> };
  return (json.models ?? [])
    .filter((m) => m.name.includes("gemini") && (m.supportedGenerationMethods ?? []).includes("generateContent"))
    .map((m) => m.name.replace(/^models\//, ""))
    .filter((id) => !/(tts|embedding|image|live|audio|transcribe)/i.test(id))
    .sort();
}

/** Validates credentials + model access cheaply. Returns a success message or throws GeminiError. */
export async function testConnection(s: ZeeSettings): Promise<string> {
  if (s.authMode === "apiKey") {
    const models = await listModels(s.apiKey);
    const hasModel = models.includes(s.model);
    return `Connected to Gemini API – ${models.length} models available.${hasModel ? "" : ` Note: "${s.model}" is not in your list.`}`;
  }
  const ep = resolveEndpoint(s, s.model);
  const res = await fetchWithRetry(
    ep.url("countTokens"),
    { method: "POST", headers: ep.headers, body: JSON.stringify({ contents: [{ role: "user", parts: [{ text: "ping" }] }] }) },
    s.model,
    0
  );
  await res.json();
  return `Connected to ${ep.label} – model "${s.model}" is reachable.`;
}

/** One-shot plain text generation (used by the =ZEE.ASK() worksheet function). */
export async function generateText(opts: { settings: ZeeSettings; prompt: string; systemInstruction: string; signal?: AbortSignal }): Promise<{ text: string; usage?: Usage }> {
  const { settings } = opts;
  const ep = resolveEndpoint(settings, settings.model);
  const res = await fetchWithRetry(
    ep.url("generateContent"),
    { method: "POST", headers: ep.headers, body: JSON.stringify(buildBody([{ role: "user", text: opts.prompt }], opts.systemInstruction, settings)), signal: opts.signal },
    settings.model,
    3
  );
  const json = (await res.json()) as GenerateResponse;
  if (json.promptFeedback?.blockReason) throw new GeminiError(`Blocked by safety filters (${json.promptFeedback.blockReason}).`, "safety");
  return { text: textOf(json).trim(), usage: usageOf(json) };
}

export interface FunctionDeclaration {
  name: string;
  description: string;
  parameters?: Record<string, unknown>;
}

export interface AgentStepResult {
  content: GeminiContent; // the model turn, verbatim (keeps thought signatures)
  text: string;
  calls: Array<{ name: string; args: Record<string, unknown>; id?: string }>;
  usage?: Usage;
  finishReason?: string;
}

/** One agent step: send the conversation + tool declarations, get text and/or function calls back. */
export async function agentStep(opts: {
  settings: ZeeSettings;
  contents: GeminiContent[];
  systemInstruction: string;
  tools: FunctionDeclaration[];
  signal?: AbortSignal;
}): Promise<AgentStepResult> {
  const { settings } = opts;
  const ep = resolveEndpoint(settings, settings.model);
  const body = {
    ...buildRawBody(opts.contents, opts.systemInstruction, settings),
    tools: [{ functionDeclarations: opts.tools }],
    toolConfig: { functionCallingConfig: { mode: "AUTO" } },
  };
  const res = await fetchWithRetry(
    ep.url("generateContent"),
    { method: "POST", headers: ep.headers, body: JSON.stringify(body), signal: opts.signal },
    settings.model,
    2
  );
  const json = (await res.json()) as GenerateResponse;
  if (json.promptFeedback?.blockReason) throw new GeminiError(`Blocked by safety filters (${json.promptFeedback.blockReason}).`, "safety");
  const cand = json.candidates?.[0];
  const parts = cand?.content?.parts ?? [];
  const calls = parts
    .filter((p) => p.functionCall)
    .map((p) => ({ name: p.functionCall!.name, args: p.functionCall!.args ?? {}, id: p.functionCall!.id }));
  if (!parts.length && cand?.finishReason === "SAFETY") throw new GeminiError("The response was blocked by safety filters.", "safety");
  return {
    content: { role: "model", parts: parts.length ? parts : [{ text: "" }] },
    text: textOf(json),
    calls,
    usage: usageOf(json),
    finishReason: cand?.finishReason,
  };
}
