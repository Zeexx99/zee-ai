/** Recently sent prompts (↑ / ↓ in the chat box) and the unsent draft. */
import { APP } from "./app";
import { readJson, readText, writeJson, writeText } from "./store";

const HISTORY_KEY = `${APP.storagePrefix}.promptHistory`;
const DRAFT_KEY = `${APP.storagePrefix}.draft`;
const MAX = 50;

export function loadHistory(): string[] {
  const list = readJson<unknown>(HISTORY_KEY, []);
  return Array.isArray(list) ? list.filter((p): p is string => typeof p === "string" && p.trim() !== "") : [];
}

/** Newest first, no duplicates (pure). */
export function addToHistory(list: string[], text: string): string[] {
  const t = text.trim();
  if (!t) return list;
  return [t, ...list.filter((p) => p !== t)].slice(0, MAX);
}

export function pushHistory(text: string): string[] {
  const next = addToHistory(loadHistory(), text);
  writeJson(HISTORY_KEY, next);
  return next;
}

/**
 * One step of ↑ / ↓ recall. `index` -1 = the text being typed (not from history).
 * ↑ goes to older prompts, ↓ back to newer ones and finally to the draft.
 */
export function recallStep(history: string[], index: number, dir: "older" | "newer", draft: string): { index: number; text: string } {
  if (dir === "older") {
    const next = Math.min(index + 1, history.length - 1);
    return next < 0 ? { index: -1, text: draft } : { index: next, text: history[next] };
  }
  const next = index - 1;
  return next < 0 ? { index: -1, text: draft } : { index: next, text: history[next] };
}

export function loadDraft(): string {
  return readText(DRAFT_KEY);
}

export function saveDraft(text: string): void {
  writeText(DRAFT_KEY, text.slice(0, 20_000));
}
