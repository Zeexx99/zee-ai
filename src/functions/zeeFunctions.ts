/**
 * Excel worksheet functions (shared runtime).
 *
 * AI:     ZEE.ASK, EXTRACT, CLASSIFY, CLEAN, TRANSLATE, SUMMARIZE, FLAG, FORMULA, EXPLAIN
 * Local:  ZEE.UNITS + the petroleum functions (VOGEL, PI, ARPS, TDH, …) – no API call, free.
 *
 * Anti-spam rules for the AI functions:
 *  - Cache: same model + task + input → answer comes from the cache, no new call. Recalc is free.
 *  - Batch: per-cell calls made within ~100 ms with the same task (e.g. a filled-down column)
 *    are merged into ONE Gemini request of up to 25 rows.
 *  - Throttle: at most 4 requests in flight; Gemini 429s are retried with backoff by geminiService.
 *  - Empty input → empty result, no call.
 */
import { loadSettings, isConfigured } from "../services/settingsService";
import { generateJson, generateText, GeminiError } from "../services/geminiService";
import { convertUnit } from "../utils/units";
import {
  FormulaInputError,
  affinity,
  arpsCumAt,
  arpsRateAt,
  convertUnits,
  hydraulicHp,
  liquidSg,
  productivityIndex,
  pumpTdh,
  standingBo,
  standingPb,
  standingRs,
  vogelQmax,
  vogelRate,
  type AffinityKind,
} from "../utils/petroFormulas";

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
const CACHE_MAX = 5000;
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
  const p = String(prompt ?? "").trim();
  if (!p && !contextText.trim()) return "";
  // Filled-down =ZEE.ASK("same prompt", B2) cells are merged into one request per ~25 rows.
  if (p && contextText && !contextText.includes("\n")) return batched(p, contextText);

  const fullPrompt = contextText ? `${p}\n\nInput:\n${contextText}` : p;
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
    throw aiError(e);
  }
}

/* ---------------------------------------------------------------- batching (one request for many cells) */

const BATCH_SYSTEM_PROMPT = `You are an Excel worksheet function applied to many cells at once.
Apply the task to EACH input independently. Return ONLY a JSON array of strings with exactly one output per input, in the same order.
Each output is just the cell value: no explanations, no markdown, no quotes, no labels. Use "" when the input is empty or the answer cannot be determined.`;

const BATCH_MAX = 25;
const BATCH_WAIT_MS = 100;

interface Pending {
  input: string;
  key: string;
  resolve: (v: string) => void;
  reject: (e: Error) => void;
}
const queues = new Map<string, { items: Pending[]; timer: number }>();

function aiError(e: unknown): Error {
  const msg = e instanceof GeminiError ? e.message : (e as Error)?.message ?? String(e);
  return fnError(msg.slice(0, 250));
}

async function flush(task: string) {
  const q = queues.get(task);
  if (!q) return;
  queues.delete(task);
  window.clearTimeout(q.timer);
  const settings = loadSettings();
  // Same input twice in one batch (e.g. repeated values down a column) → ask once.
  const unique = [...new Map(q.items.map((p) => [p.key, p])).values()];
  const chunks: Pending[][] = [];
  for (let i = 0; i < unique.length; i += BATCH_MAX) chunks.push(unique.slice(i, i + BATCH_MAX));
  const settle = (key: string, fn: (x: Pending) => void) => q.items.filter((x) => x.key === key).forEach(fn);
  // Chunks run in parallel, but withSlot() keeps at most MAX_CONCURRENT requests in flight.
  await Promise.all(
    chunks.map(async (chunk) => {
      try {
        const { data } = await withSlot(() =>
          generateJson<unknown[]>({
            settings,
            systemInstruction: BATCH_SYSTEM_PROMPT,
            prompt: `Task: ${task}\n\nInputs (JSON array of ${chunk.length}):\n${JSON.stringify(chunk.map((p) => p.input))}`,
            schema: { type: "ARRAY", items: { type: "STRING" } },
          })
        );
        const arr = Array.isArray(data) ? data : [];
        chunk.forEach((p, j) => {
          const v = cleanAnswer(arr[j] === null || arr[j] === undefined ? "" : String(arr[j]));
          // Don't cache when Gemini returned the wrong number of items – the next recalc retries.
          if (arr.length === chunk.length) putCache(p.key, v);
          settle(p.key, (x) => x.resolve(v));
        });
      } catch (e) {
        const err = aiError(e);
        chunk.forEach((p) => settle(p.key, (x) => x.reject(err)));
      }
    })
  );
}

function cleanAnswer(text: string): string {
  return text
    .replace(/^```[a-z]*\s*|\s*```$/gi, "")
    .replace(/^["'`]+|["'`]+$/g, "")
    .trim()
    .slice(0, 32000);
}

/** Run `task` on one cell's input, batched with other cells that use the same task. */
function batched(task: string, input: string): Promise<string> {
  const settings = loadSettings();
  if (!isConfigured(settings)) return Promise.reject(fnError("Open Zee AI → Settings and connect Gemini first."));
  if (!input.trim()) return Promise.resolve("");
  const key = hash(`${settings.model}|b|${task}|${input}`);
  const hit = getCache()[key];
  if (hit !== undefined) return Promise.resolve(hit);
  return new Promise<string>((resolve, reject) => {
    let q = queues.get(task);
    if (!q) {
      q = { items: [], timer: 0 };
      queues.set(task, q);
    }
    q.items.push({ input, key, resolve, reject });
    window.clearTimeout(q.timer);
    if (q.items.length >= BATCH_MAX * 4) void flush(task);
    else q.timer = window.setTimeout(() => void flush(task), BATCH_WAIT_MS);
  });
}

/** One cached, throttled call for functions that answer once for a whole range. */
async function single(prompt: string, system: string): Promise<string> {
  const settings = loadSettings();
  if (!isConfigured(settings)) throw fnError("Open Zee AI → Settings and connect Gemini first.");
  const key = hash(`${settings.model}|s|${system.length}|${prompt}`);
  const hit = getCache()[key];
  if (hit !== undefined) return hit;
  try {
    const { text } = await withSlot(() => generateText({ settings, prompt, systemInstruction: system }));
    const v = cleanAnswer(text);
    putCache(key, v);
    return v;
  } catch (e) {
    throw aiError(e);
  }
}

const str = (v: unknown) => (v === null || v === undefined ? "" : String(v)).trim();
/** Excel passes a scalar for a cell; be defensive if it hands us a 1×1 matrix. */
const scalar = (v: unknown): unknown => (Array.isArray(v) ? (Array.isArray(v[0]) ? v[0][0] : v[0]) : v);

function tableText(range: unknown, maxRows = 500): { text: string; rows: number } {
  const m = Array.isArray(range) ? (range as unknown[][]) : [[range]];
  const rows = m.slice(0, maxRows).map((r) => (Array.isArray(r) ? r : [r]).map((v) => str(v).replace(/\t|\r?\n/g, " ")).join("\t"));
  return { text: rows.join("\n"), rows: m.length };
}

function extract(text: unknown, what: unknown): Promise<string> {
  const w = str(what);
  if (!w) return Promise.reject(fnError('Say what to extract, e.g. =ZEE.EXTRACT(B2, "pump model").'));
  return batched(`Extract the ${w} from the input text. Return only the extracted value exactly as it appears (or normalized if obviously needed), or "" if it is not present.`, str(scalar(text)));
}

async function classify(text: unknown, categories: unknown): Promise<string> {
  const cats = str(categories)
    .split(/[,;\n|]/)
    .map((c) => c.trim())
    .filter(Boolean);
  if (cats.length < 2) throw fnError('Give at least two categories, e.g. "Electrical, Mechanical, Reservoir".');
  const raw = await batched(`Classify the input into exactly ONE of these categories: ${cats.join(" | ")}. Return the category name exactly as written. If none fits, return the closest one.`, str(scalar(text)));
  if (!raw) return "";
  const lower = raw.toLowerCase();
  return cats.find((c) => c.toLowerCase() === lower) ?? cats.find((c) => lower.includes(c.toLowerCase())) ?? cats.find((c) => c.toLowerCase().includes(lower)) ?? "N/A";
}

function clean(text: unknown, style: unknown): Promise<string> {
  const s = str(style);
  return batched(
    `Clean and standardize the value: fix spacing, obvious typos and casing, and make the formatting consistent${s ? `. Style: ${s}` : ""}. Keep the meaning; do not add information.`,
    str(scalar(text))
  );
}

function translate(text: unknown, language: unknown): Promise<string> {
  const lang = str(language) || "English";
  return batched(`Translate the input into ${lang}. Keep numbers, units, well names and equipment codes unchanged. Return only the translation.`, str(scalar(text)));
}

function explain(formula: unknown): Promise<string> {
  const f = str(scalar(formula));
  if (!f) return Promise.resolve("");
  return batched("Explain what this Excel formula does in one or two plain sentences for a non-expert. If the input is not a formula, describe the value.", f);
}

async function summarize(range: unknown, maxWords: unknown): Promise<string> {
  const { text, rows } = tableText(range);
  if (!text.trim()) return "";
  const n = Math.max(5, Math.min(200, Number(maxWords) || 30));
  return single(
    `Summarize this spreadsheet data (${rows} rows, tab-separated${rows > 500 ? ", first 500 shown" : ""}) in at most ${n} words. Mention the key numbers or patterns.\n\n${text}`,
    "You are an Excel worksheet function. Answer with ONLY the summary text: one short paragraph, no markdown, no preamble."
  );
}

async function formula(description: unknown, sample: unknown): Promise<string> {
  const d = str(description);
  if (!d) throw fnError("Describe the formula you want.");
  const s = sample === null || sample === undefined ? "" : tableText(sample, 20).text;
  const out = await single(
    `Write ONE Excel formula that does: ${d}${s ? `\n\nSample of the data it works on (tab-separated):\n${s}` : ""}`,
    "Return ONLY a single Excel formula starting with =, using English function names and commas. No explanation, no code fences."
  );
  const line = out.split(/\r?\n/).find((l) => l.trim().startsWith("=")) ?? out;
  // Returned as text (a custom function result is never evaluated), so the user can copy it.
  return line.trim();
}

async function flag(range: unknown, rule: unknown): Promise<string[][]> {
  const r = str(rule);
  if (!r) throw fnError('Describe the rule, e.g. "current jumps more than 20%".');
  const m = Array.isArray(range) ? (range as unknown[][]) : [[range]];
  if (m.length > 1000) throw fnError("Select at most 1,000 rows for ZEE.FLAG.");
  const { text } = tableText(m, 1000);
  if (!text.trim()) return m.map(() => [""]);
  const settings = loadSettings();
  if (!isConfigured(settings)) throw fnError("Open Zee AI → Settings and connect Gemini first.");
  const key = hash(`${settings.model}|flag|${r}|${text}`);
  let joined = getCache()[key];
  if (joined === undefined) {
    try {
      const { data } = await withSlot(() =>
        generateJson<unknown[]>({
          settings,
          systemInstruction: `You check spreadsheet rows against a rule. Rows may need comparing with neighbouring rows (trends, jumps, drops).
Return ONLY a JSON array of strings, exactly one per row in order: "" if the row is fine, otherwise a very short reason (max 8 words).`,
          prompt: `Rule: ${r}\n\nRows (${m.length}, tab-separated, in order):\n${text}`,
          schema: { type: "ARRAY", items: { type: "STRING" } },
        })
      );
      const arr = Array.isArray(data) ? data.map((v) => cleanAnswer(str(v))) : [];
      joined = JSON.stringify(arr);
      putCache(key, joined);
    } catch (e) {
      throw aiError(e);
    }
  }
  const arr = JSON.parse(joined) as string[];
  return m.map((_, i) => [arr[i] ? `⚠ ${arr[i]}` : ""]);
}

function units(value: unknown, from: unknown, to: unknown): number {
  const v = scalar(value);
  const n = typeof v === "number" ? v : Number(String(v ?? "").replace(/,/g, ""));
  return convertUnit(n, str(from), str(to));
}

/** Wrap a pure formula: bad input becomes #NUM! with the reason as the cell's error tooltip. */
function numeric(fn: (...args: any[]) => number) {
  return (...args: unknown[]) => {
    try {
      const r = fn(...args);
      if (!Number.isFinite(r)) throw new FormulaInputError("Result is not a finite number – check the inputs.");
      return r;
    } catch (e) {
      const msg = (e as Error).message ?? String(e);
      if (typeof CustomFunctions !== "undefined") throw new CustomFunctions.Error(CustomFunctions.ErrorCode.invalidNumber ?? "#NUM!", msg.slice(0, 250));
      throw e;
    }
  };
}

/** Petroleum worksheet functions (no AI, instant, free). Metadata lives in public/functions.json. */
const PETRO_FUNCTIONS: Record<string, (...args: any[]) => unknown> = {
  VOGEL: numeric((pr: number, pwf: number, qTest: number, pwfTest: number) => vogelRate(pr, pwf, qTest, pwfTest)),
  VOGEL_QMAX: numeric((pr: number, qTest: number, pwfTest: number) => vogelQmax(pr, qTest, pwfTest)),
  PI: numeric((q: number, pr: number, pwf: number) => productivityIndex(q, pr, pwf)),
  RS: numeric((p: number, t: number, api: number, gasSg: number) => standingRs(p, t, api, gasSg)),
  PB: numeric((rs: number, t: number, api: number, gasSg: number) => standingPb(rs, t, api, gasSg)),
  BO: numeric((p: number, t: number, api: number, gasSg: number) => standingBo(standingRs(p, t, api, gasSg), t, api, gasSg)),
  ARPS: numeric((qi: number, di: number, b: number, t: number) => arpsRateAt(qi, di, b, t)),
  ARPS_CUM: numeric((qi: number, di: number, b: number, t: number) => arpsCumAt(qi, di, b, t)),
  TDH: numeric((pd: number, pi: number, sg: number) => pumpTdh(pd, pi, sg)),
  HHP: numeric((q: number, dp: number) => hydraulicHp(q, dp)),
  AFFINITY: numeric((v: number, f1: number, f2: number, kind: string) => affinity(v, f1, f2, (kind || "Q") as AffinityKind)),
  SG_MIX: numeric((wc: number, api: number, waterSg?: number | null) => liquidSg(wc, api, waterSg ?? 1)),
  CONVERT: numeric((v: number, id: string) => convertUnits(v, id)),
  UNITS: numeric((v: unknown, from: string, to: string) => units(v, from, to)),
};

/** AI worksheet functions. Metadata lives in public/functions.json. */
const AI_FUNCTIONS: Record<string, (...args: any[]) => unknown> = {
  ASK: ask,
  EXTRACT: (text: unknown, what: unknown) => extract(text, what),
  CLASSIFY: (text: unknown, cats: unknown) => classify(text, cats),
  CLEAN: (text: unknown, style: unknown) => clean(text, style),
  TRANSLATE: (text: unknown, lang: unknown) => translate(text, lang),
  SUMMARIZE: (range: unknown, maxWords: unknown) => summarize(range, maxWords),
  FLAG: (range: unknown, rule: unknown) => flag(range, rule),
  FORMULA: (desc: unknown, sample: unknown) => formula(desc, sample),
  EXPLAIN: (f: unknown) => explain(f),
};

let registered = false;
export function registerCustomFunctions() {
  if (registered || typeof CustomFunctions === "undefined") return;
  try {
    for (const [id, fn] of Object.entries({ ...PETRO_FUNCTIONS, ...AI_FUNCTIONS })) {
      try {
        CustomFunctions.associate(id, fn);
      } catch {
        /* metadata not loaded yet for this id (old cached functions.json) – skip */
      }
    }
    registered = true;
  } catch {
    /* Custom functions not supported on this host/version. */
  }
}
