/** "Most used first": counts how often each quick action / template is used (on this PC). */
import { APP } from "./app";
import { readJson, writeJson } from "./store";

const KEY = `${APP.storagePrefix}.useCounts`;

export type UseCounts = Record<string, { n: number; last: number }>;
type Counts = UseCounts;

export function loadUseCounts(): Counts {
  const c = readJson<Counts>(KEY, {});
  return c && typeof c === "object" ? c : {};
}

/** Records one use of `key` and returns the updated counts. */
export function bumpUse(key: string, now = Date.now()): Counts {
  const c = loadUseCounts();
  const cur = c[key];
  const next: Counts = { ...c, [key]: { n: (cur?.n ?? 0) + 1, last: now } };
  // Keep the store small.
  const keys = Object.keys(next);
  if (keys.length > 200) keys.sort((a, b) => next[a].last - next[b].last).slice(0, keys.length - 200).forEach((k) => delete next[k]);
  writeJson(KEY, next);
  return next;
}

/**
 * Most-used first (ties: most recent first); items never used keep their original order after them.
 * Stable and pure.
 */
export function sortByUse<T>(items: T[], keyOf: (t: T) => string, counts: Counts): T[] {
  return items
    .map((item, i) => ({ item, i, c: counts[keyOf(item)] }))
    .sort((a, b) => {
      const an = a.c?.n ?? 0;
      const bn = b.c?.n ?? 0;
      if (an !== bn) return bn - an;
      const al = a.c?.last ?? 0;
      const bl = b.c?.last ?? 0;
      if (al !== bl) return bl - al;
      return a.i - b.i;
    })
    .map((x) => x.item);
}
