/** Token usage per day (kept for 62 days on this PC) and an optional cost estimate. */
import { APP } from "./app";
import { readJson, writeJson } from "./store";

export interface DayUsage {
  /** YYYY-MM-DD (local time). */
  date: string;
  requests: number;
  promptTokens: number;
  outputTokens: number;
}

export interface Prices {
  /** USD per 1 million input (prompt) tokens. 0 = unknown. */
  inPerM: number;
  /** USD per 1 million output tokens. 0 = unknown. */
  outPerM: number;
}

const KEY = `${APP.storagePrefix}.usageByDay`;
const KEEP_DAYS = 62;

const pad = (n: number) => String(n).padStart(2, "0");
export const dayKey = (d: Date = new Date()) => `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;

type UsageMap = Record<string, DayUsage>;

const empty = (date: string): DayUsage => ({ date, requests: 0, promptTokens: 0, outputTokens: 0 });

/** Adds one request to `map` (pure – returns a new map, oldest days dropped). */
export function addToMap(map: UsageMap, u: { promptTokens: number; outputTokens: number }, now: Date = new Date()): UsageMap {
  const date = dayKey(now);
  const cur = map[date] ?? empty(date);
  const next: UsageMap = {
    ...map,
    [date]: {
      date,
      requests: cur.requests + 1,
      promptTokens: cur.promptTokens + (Number.isFinite(u.promptTokens) ? u.promptTokens : 0),
      outputTokens: cur.outputTokens + (Number.isFinite(u.outputTokens) ? u.outputTokens : 0),
    },
  };
  const days = Object.keys(next).sort();
  for (const d of days.slice(0, Math.max(0, days.length - KEEP_DAYS))) delete next[d];
  return next;
}

export function recordUsage(u: { promptTokens: number; outputTokens: number }, now: Date = new Date()): DayUsage {
  const map = addToMap(readJson<UsageMap>(KEY, {}), u, now);
  writeJson(KEY, map);
  return map[dayKey(now)];
}

export function usageForDay(now: Date = new Date()): DayUsage {
  const d = dayKey(now);
  return readJson<UsageMap>(KEY, {})[d] ?? empty(d);
}

/** Totals for the calendar month of `now`. */
export function usageForMonth(now: Date = new Date(), map: UsageMap = readJson<UsageMap>(KEY, {})): DayUsage {
  const prefix = dayKey(now).slice(0, 7);
  return Object.values(map)
    .filter((u) => u && typeof u.date === "string" && u.date.startsWith(prefix))
    .reduce((t, u) => ({ ...t, requests: t.requests + u.requests, promptTokens: t.promptTokens + u.promptTokens, outputTokens: t.outputTokens + u.outputTokens }), empty(prefix));
}

/** Estimated cost in USD, or null when no prices are set. */
export function estimateCost(u: { promptTokens: number; outputTokens: number }, p: Prices): number | null {
  if (!(p.inPerM > 0) && !(p.outPerM > 0)) return null;
  return (u.promptTokens * Math.max(0, p.inPerM || 0) + u.outputTokens * Math.max(0, p.outPerM || 0)) / 1_000_000;
}

export function formatTokens(n: number): string {
  if (n >= 1_000_000) return `${(n / 1_000_000).toFixed(n >= 10_000_000 ? 0 : 1)}M`;
  if (n >= 1000) return `${(n / 1000).toFixed(n >= 100_000 ? 0 : 1)}k`;
  return String(Math.round(n));
}

export function formatCost(usd: number): string {
  if (usd < 0.01) return usd > 0 ? "<$0.01" : "$0";
  return `$${usd < 10 ? usd.toFixed(2) : usd.toFixed(1)}`;
}
