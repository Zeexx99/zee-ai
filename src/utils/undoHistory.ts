/**
 * Change history helpers (pure – unit tested): which sheets an undo step touches, and whether one
 * step can be reverted on its own without disturbing newer steps.
 */

export interface StepLike {
  kind: string;
  sheetName?: string;
  steps?: StepLike[];
}

/**
 * Sheets a step affects. [] = workbook-level things that don't move cells (names, table objects);
 * null = unknown (treated as "touches everything").
 */
export function stepSheets(s: StepLike): string[] | null {
  if (s.kind === "multi") {
    const all: string[] = [];
    for (const st of s.steps ?? []) {
      const sh = stepSheets(st);
      if (sh === null) return null;
      all.push(...sh);
    }
    return Array.from(new Set(all));
  }
  if (s.kind === "addedName" || s.kind === "addedTable") return [];
  return typeof s.sheetName === "string" ? [s.sheetName] : null;
}

/**
 * True when step `i` (0 = oldest) can be reverted alone: no newer step touches the same sheet(s).
 * Reverting an older step on a sheet that changed later could put old values over newer edits or
 * shift rows / columns that newer steps refer to, so those must be undone in order instead.
 */
export function canRevertAlone(steps: StepLike[], i: number): boolean {
  if (i < 0 || i >= steps.length) return false;
  if (i === steps.length - 1) return true;
  const mine = stepSheets(steps[i]);
  if (mine === null) return false;
  if (!mine.length) return true;
  for (let j = i + 1; j < steps.length; j++) {
    const later = stepSheets(steps[j]);
    if (later === null) return false;
    if (later.some((s) => mine.includes(s))) return false;
  }
  return true;
}

/** "just now", "5 min ago", "2 h ago", or a time. */
export function timeAgo(at: number, now: number = Date.now()): string {
  const s = Math.max(0, Math.round((now - at) / 1000));
  if (s < 45) return "just now";
  const m = Math.round(s / 60);
  if (m < 60) return `${m} min ago`;
  const h = Math.round(m / 60);
  if (h < 12) return `${h} h ago`;
  return new Date(at).toLocaleString();
}
