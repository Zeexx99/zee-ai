/** Saved prompts (per user, in this add-in's local storage). */
const KEY = "zeeai.prompts.v1";

export const DEFAULT_PROMPTS: string[] = [
  "Summarize production for each well in the selection: latest oil/water/gas rate, change vs the previous month and the water-cut trend.",
  "Find wells whose oil rate dropped more than 20% month over month and list them with the drop.",
  "Build a monthly summary table (oil, water, gas, water cut, GOR) from the selection.",
  "Check this data for errors: gaps in dates, duplicates, negative or stuck values.",
];

export function loadPrompts(): string[] {
  try {
    const raw = localStorage.getItem(KEY);
    if (!raw) return [...DEFAULT_PROMPTS];
    const list = JSON.parse(raw) as unknown;
    return Array.isArray(list) ? list.filter((p): p is string => typeof p === "string" && p.trim() !== "") : [...DEFAULT_PROMPTS];
  } catch {
    return [...DEFAULT_PROMPTS];
  }
}

export function savePrompts(list: string[]): void {
  try {
    localStorage.setItem(KEY, JSON.stringify(list.slice(0, 50)));
  } catch {
    /* storage unavailable – keep for this session only */
  }
}
