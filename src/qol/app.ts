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
    version: "1.9.0",
    items: [
      "Dark theme is now the default (light and Match Office are still in Settings → Appearance, or ⋯ → Light mode). The pane no longer flashes white when it opens.",
      "Search all tools with Ctrl+K (or the 🔍 at the start of the toolbar). Click ☆ next to a tool to pin it to the toolbar as a one-click button (up to 6).",
      "Tidier header: Undo, New chat and Settings stay visible; history, export and the theme switch moved into ⋯.",
      "Toolbar shows ‹ › arrows when it doesn't fit; notices fold into one line with \"+N more\".",
      "New welcome screen with petroleum starters (daily production, ESP readings, well test QC) and your recent prompts.",
      "A \"Latest\" button brings you back down when you scroll up during an answer; times show on hover; agent runs show which step they're on.",
      "Agent and Selection are now toggle chips; today's usage moved into a small icon.",
      "Settings → Appearance: accent colour (violet, teal, blue, orange) and compact density. Windows high-contrast mode is followed automatically.",
    ],
  },
  {
    version: "1.8.0",
    items: [
      "Petroleum → Nodal analysis: IPR (PI, Vogel or composite) against the tubing curve (Beggs & Brill) – operating point, curves and chart on a new sheet, plus what-if on wellhead pressure, tubing size, water cut or GOR.",
      "Gas-well functions: =ZEE.Z (Z-factor), =ZEE.BG, =ZEE.GAS_VISC, =ZEE.DELIV (fits C and n from a multi-rate test) and =ZEE.DELIV_Q.",
      "New Data menu → Import file: CSV / TXT, LAS well logs and SCADA exports (tag lists can become one column per tag), with a preview first.",
      "Data → Merge well data: joins production, well tests, ESP readings… by well and date – even “latest test on or before each day”.",
      "Data → Resample: daily → weekly / monthly / quarterly / yearly per well, with a rule per column (average, sum, last…).",
      "Recipes: record what you do once (clean, fill gaps, resample, QC, charts, prompts…) and replay it on next month's data with one click. Share recipes with colleagues as text.",
      "Big sheets are now read and written in chunks, so tables of 100k+ rows work instead of hitting Excel's size limit.",
      "Older Excel versions: features they can't run are greyed out with a note instead of failing. Settings → This Excel lists what's available.",
      "The agent can run nodal analysis, merge sheets and resample time series too.",
    ],
  },
  {
    version: "1.7.0",
    items: [
      "Clean → Find duplicates & outliers: duplicate rows, spellings that differ (\"Well-12\" vs \"W12\") and unusual numbers – highlight them, unify spellings or delete duplicates. Undoable.",
      "Clean → Fill gaps in time series: adds the missing dates per well and fills values (straight line, previous value, or dates only) on a new sheet, with filled cells marked.",
      "Clean → Fix imported data: numbers stored as text, Arabic digits, units stuck to numbers (\"1,250 bbl\") and mixed date formats – with a preview first. Undoable.",
      "Analyze → Explain this sheet (AI) and Add a sheet guide tab: inputs, calculated columns, formula patterns and what feeds what.",
      "Analyze → Move hard-coded numbers to inputs: puts numbers like 5.615 on an Inputs sheet as named cells and rewrites the formulas. Also offered after a formula audit.",
      "Agent dry run: the approval card now shows the real cells before → after (Settings → Agent mode, on by default).",
      "Big-data mode: selections from 100,000 rows are sent as statistics over every row plus a sample instead of being cut off (Settings → Sheet context). The agent gets a profile_range tool for the same.",
    ],
  },
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
