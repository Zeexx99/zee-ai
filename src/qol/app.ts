/**
 * Per-add-in details for the shared quality-of-life modules in this folder.
 * Everything else in src/qol is identical in the Excel and Word add-ins – only this file differs.
 */

export const APP = {
  id: "excel" as const,
  name: "Zee AI",
  host: "Excel",
  /** Where the add-in keeps its settings (used to tell a first install from an update). */
  settingsKey: "zeeai.settings.v1",
  /** Prefix for every localStorage key the qol modules use. */
  storagePrefix: "zeeai.excel",
};

export interface ReleaseNote {
  version: string;
  items: string[];
}

/** Shown once after an update ("What's new"). Newest first. */
export const RELEASE_NOTES: ReleaseNote[] = [
  {
    version: "1.6.0",
    items: [
      "Petroleum → Daily production report: per-well rates, change vs the previous reading, water cut, GOR, totals, highlights and a chart – then hand it to Word.",
      "Petroleum → ESP failure log: create a ready-made log sheet, or analyze yours (run life, MTBF, causes, components, models, repeat wells, monthly trend, charts) with an AI summary.",
      "Analyze → Formula audit: error values, #REF!, inconsistent formulas, typed values in formula columns and hard-coded numbers, with clickable links.",
      "Analyze → Suggest data validation (AI): Zee proposes limits per column (e.g. water cut 0–100) – review them, highlight the values that break them and optionally add Excel validation.",
      "The agent can now run the ESP failure analysis and the formula audit too.",
      "Offline queue: requests that fail because the connection dropped are sent again automatically when it's back.",
    ],
  },
  {
    version: "1.5.0",
    items: [
      "Remove empty rows / columns and Clean now show a preview (“will remove 412 rows”) before anything changes.",
      "Big scans show a progress bar with a Cancel button.",
      "Tables in chat answers: insert at the selected cell, or as a formatted Excel table with filters.",
      "Change history (clock icon next to ↶): see Zee AI’s recent changes and revert one, or go back to any point.",
      "Live selection chip (e.g. Sheet1!A1:F200) above the chat box – click × to leave it out of the next message.",
      "Press ↑ / ↓ in the chat box to recall earlier prompts; favourite prompts appear as one-click chips.",
      "Every message: copy, edit & resend, regenerate. Esc stops a running answer.",
      "Errors now explain what to do, with Retry and “Copy diagnostics”.",
      "Usage meter: today’s tokens (and estimated cost if you enter your prices in Settings).",
      "Settings → Backup & updates: export / import settings to another PC, and “Reload latest version”.",
      "A banner tells you when a newer version is online, and this “What’s new” note appears after each update.",
    ],
  },
];
