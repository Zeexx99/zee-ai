/**
 * Excel worksheet functions (shared runtime):
 *   =ZEE.ASK(prompt, [context1], [context2], ...)
 * e.g. =ZEE.ASK("Classify this ESP failure reason as electrical, mechanical or reservoir", B2)
 *
 * Results are cached (same model + prompt + context → no new API call), and calls are
 * throttled so filling down 1,000 rows doesn't trip Gemini rate limits.
 */
import { loadSettings, isConfigured } from "../services/settingsService";
import { generateText, GeminiError } from "../services/geminiService";

declare const CustomFunctions:
  | {
      associate: (id: string, fn: (...args: any[]) => unknown) => void;
      Error: new (code: string, message?: string) => Error;
      ErrorCode: Record<string, string>;
    }
  | undefined;

const FN_SYSTEM_PROMPT = `You are an Excel worksheet function. Answer with ONLY the value that should appear in the cell:
no explanations, no markdown, no quotes, no labels. Keep it short (ideally a word, number or one sentence).
If a number is asked for, return just the number. If the answer cannot be determined, return "N/A".`;

const CACHE_KEY = "zeeai.fncache.v1";
const CACHE_MAX = 2000;
const MAX_CONCURRENT = 4;

let cache: Record<string, string> | null = null;
function getCache(): Record<string, string> {
  if (cache) return cache;
  try {
    cache = JSON.parse(localStorage.getItem(CACHE_KEY) ?? "{}") as Record<string, string>;
  } catch {
    cache = {};
  }
  return cache;
}
let saveTimer = 0;
function putCache(key: string, value: string) {
  const c = getCache();
  c[key] = value;
  const keys = Object.keys(c);
  if (keys.length > CACHE_MAX) keys.slice(0, keys.length - CACHE_MAX).forEach((k) => delete c[k]);
  window.clearTimeout(saveTimer);
  saveTimer = window.setTimeout(() => {
    try {
      localStorage.setItem(CACHE_KEY, JSON.stringify(c));
    } catch {
      /* quota – keep in memory only */
    }
  }, 1000);
}

export function clearFunctionCache() {
  cache = {};
  try {
    localStorage.removeItem(CACHE_KEY);
  } catch {
    /* ignore */
  }
}

function hash(s: string): string {
  let h1 = 0xdeadbeef ^ s.length;
  let h2 = 0x41c6ce57 ^ s.length;
  for (let i = 0; i < s.length; i++) {
    const ch = s.charCodeAt(i);
    h1 = Math.imul(h1 ^ ch, 2654435761);
    h2 = Math.imul(h2 ^ ch, 1597334677);
  }
  h1 = Math.imul(h1 ^ (h1 >>> 16), 2246822507) ^ Math.imul(h2 ^ (h2 >>> 13), 3266489909);
  h2 = Math.imul(h2 ^ (h2 >>> 16), 2246822507) ^ Math.imul(h1 ^ (h1 >>> 13), 3266489909);
  return (4294967296 * (2097151 & h2) + (h1 >>> 0)).toString(36);
}

/* Simple semaphore. */
let active = 0;
const waiting: Array<() => void> = [];
async function withSlot<T>(fn: () => Promise<T>): Promise<T> {
  if (active >= MAX_CONCURRENT) await new Promise<void>((r) => waiting.push(r));
  active++;
  try {
    return await fn();
  } finally {
    active--;
    waiting.shift()?.();
  }
}

function contextToText(ctx: unknown[]): string {
  return ctx
    .map((arg) => {
      if (!Array.isArray(arg)) return String(arg ?? "");
      return (arg as unknown[][]).map((row) => (Array.isArray(row) ? row.map((v) => String(v ?? "")).join("\t") : String(row ?? ""))).join("\n");
    })
    .filter((t) => t.trim() !== "")
    .join("\n---\n");
}

function fnError(message: string): Error {
  if (typeof CustomFunctions !== "undefined") return new CustomFunctions.Error(CustomFunctions.ErrorCode.notAvailable ?? "#N/A", message);
  return new Error(message);
}

async function ask(prompt: string, ...rest: unknown[]): Promise<string> {
  // Excel passes the repeating "context" parameter as ONE array of matrices, then the invocation object.
  const repeated = rest.find((r) => Array.isArray(r)) as unknown[][][] | undefined;
  const maybeInvocation = rest.find((r) => r && typeof r === "object" && !Array.isArray(r)) as { onCanceled?: () => void } | undefined;
  const flat: unknown[] = repeated ?? [];

  const settings = loadSettings();
  if (!isConfigured(settings)) throw fnError("Open Zee AI → Settings and connect Gemini first.");
  const contextText = contextToText(flat);
  const fullPrompt = contextText ? `${prompt}\n\nInput:\n${contextText}` : String(prompt ?? "");
  if (!fullPrompt.trim()) return "";
  const key = hash(`${settings.model}|${fullPrompt}`);
  const hit = getCache()[key];
  if (hit !== undefined) return hit;

  const controller = new AbortController();
  if (maybeInvocation && typeof maybeInvocation === "object") maybeInvocation.onCanceled = () => controller.abort();
  try {
    const { text } = await withSlot(() => generateText({ settings, prompt: fullPrompt, systemInstruction: FN_SYSTEM_PROMPT, signal: controller.signal }));
    const value = text.replace(/^["'`]+|["'`]+$/g, "").slice(0, 32000);
    putCache(key, value);
    return value;
  } catch (e) {
    const msg = e instanceof GeminiError ? e.message : (e as Error).message;
    throw fnError(msg.slice(0, 250));
  }
}

let registered = false;
export function registerCustomFunctions() {
  if (registered || typeof CustomFunctions === "undefined") return;
  try {
    CustomFunctions.associate("ASK", ask);
    registered = true;
  } catch {
    /* Custom functions not supported on this host/version. */
  }
}
