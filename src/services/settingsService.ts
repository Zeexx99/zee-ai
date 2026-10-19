/**
 * Settings persistence.
 *
 * Note: Office.context.roamingSettings exists only in Outlook add-ins. In Excel the options are
 * Office.context.document.settings (saved INSIDE the workbook – wrong place for an API key) or
 * browser storage. Zee AI therefore keeps settings in localStorage of the add-in's own origin,
 * which is per user, per machine and never written into the workbook.
 */

import { DEFAULT_OTHER_BASE_URL, DEFAULT_OTHER_MODEL } from "./openaiCompat";

/** "openai" = Other API key: any OpenAI-compatible provider (Groq, OpenRouter, OpenAI, Mistral…). */
export type AuthMode = "apiKey" | "openai" | "google" | "proxy";

/** "auto" follows the Office theme (or the OS colour scheme outside Office). Default: dark. */
export type ThemeMode = "auto" | "light" | "dark";

/** Brand colour used for buttons, links, focus rings and your chat bubbles. */
export type AccentColor = "violet" | "teal" | "blue" | "orange";
export const ACCENTS: Array<{ id: AccentColor; label: string; hex: string }> = [
  { id: "violet", label: "Violet", hex: "#5B42E0" },
  { id: "teal", label: "Teal", hex: "#00897B" },
  { id: "blue", label: "Blue", hex: "#1F6FD1" },
  { id: "orange", label: "Orange", hex: "#D9640F" },
];

/** Spacing and text size of the task pane. */
export type Density = "comfortable" | "compact";

export interface ModelInfo {
  id: string;
  label: string;
  hint: string;
}

/**
 * Default models (checked against Google's model list in Oct 2026).
 * Gemini 1.5 models are shut down; Gemini 2.5 remains for projects that already had access.
 * Use "Refresh" in Settings to pull the live list for your key.
 */
export const DEFAULT_MODELS: ModelInfo[] = [
  { id: "gemini-3.8-flash", label: "Gemini 3.8 Flash", hint: "Recommended default – fast and smart for daily sheet work" },
  { id: "gemini-3.1-pro-preview", label: "Gemini 3.1 Pro (preview)", hint: "Complex reasoning and advanced data modeling" },
  { id: "gemini-3.5-flash-lite", label: "Gemini 3.5 Flash-Lite", hint: "Cheapest and fastest – bulk row-by-row tasks" },
  { id: "gemini-2.5-pro", label: "Gemini 2.5 Pro (legacy)", hint: "Only for projects that already had 2.5 access" },
  { id: "gemini-2.5-flash", label: "Gemini 2.5 Flash (legacy)", hint: "Only for projects that already had 2.5 access" },
];

export interface ZeeSettings {
  authMode: AuthMode;
  apiKey: string;
  model: string;
  /** Extra model IDs fetched from the API with "Refresh". */
  fetchedModels: string[];
  /** Other API key (OpenAI-compatible provider). */
  otherBaseUrl: string;
  otherApiKey: string;
  /** Model remembered for the other provider while Gemini is selected. */
  otherModel: string;
  /** Model IDs fetched from the other provider with "Refresh". */
  otherFetchedModels: string[];
  /** Gemini model remembered while the other provider is selected. */
  geminiModel: string;
  googleClientId: string;
  gcpProjectId: string;
  gcpLocation: string;
  includeSelection: boolean;
  maxContextCells: number;
  /** null = let the model use its default (recommended for Gemini 3). */
  temperature: number | null;
  /** Team proxy (keeps the Gemini key on a server). */
  proxyUrl: string;
  proxyToken: string;
  /** Send a compact map of all sheets with each chat message. */
  includeWorkbookOverview: boolean;
  /** Ask before sending prompts larger than this many (estimated) tokens. */
  costGuardTokens: number;
  /** Agent mode: let Gemini act on the workbook through tools. */
  agentMode: boolean;
  /** Agent mode: apply changes without asking (not recommended). */
  autoApprove: boolean;
  /** Appearance of the task pane. */
  theme: ThemeMode;
  /** Let Gemini search the web (Google Search grounding) – chat answers cite sources; the agent gets a web_search tool. */
  webSearch: boolean;
  /** Let Gemini write and run Python for calculations in chat (code execution). */
  codeExecution: boolean;
  /** Optional prices (USD per 1M tokens) for the usage meter's cost estimate. 0 = unknown. */
  priceInPerM: number;
  priceOutPerM: number;
  /** Agent approval card shows the cells before → after for each write (dry run). */
  agentDryRun: boolean;
  /** Very large selections are sent as column statistics + a sample instead of truncated cells. */
  bigDataMode: boolean;
  /** Rows from which big-data mode kicks in. */
  bigDataRows: number;
  /** Brand colour (Settings → Appearance). */
  accent: AccentColor;
  /** Comfortable (default) or compact spacing. */
  density: Density;
  /** Tool ids pinned as one-click buttons on the toolbar (star them in Ctrl+K search). */
  pinnedTools: string[];
  /** Set once the one-time "dark by default" switch (v1.9) has been applied on this PC. */
  themeMigrated: boolean;
}

export const DEFAULT_SETTINGS: ZeeSettings = {
  authMode: "apiKey",
  apiKey: "",
  model: DEFAULT_MODELS[0].id,
  fetchedModels: [],
  otherBaseUrl: DEFAULT_OTHER_BASE_URL,
  otherApiKey: "",
  otherModel: DEFAULT_OTHER_MODEL,
  otherFetchedModels: [],
  geminiModel: DEFAULT_MODELS[0].id,
  googleClientId: "",
  gcpProjectId: "",
  gcpLocation: "global",
  includeSelection: true,
  maxContextCells: 2000,
  temperature: null,
  proxyUrl: "",
  proxyToken: "",
  includeWorkbookOverview: true,
  costGuardTokens: 20000,
  agentMode: true,
  autoApprove: false,
  theme: "dark",
  webSearch: false,
  codeExecution: false,
  priceInPerM: 0,
  priceOutPerM: 0,
  agentDryRun: true,
  bigDataMode: true,
  bigDataRows: 100_000,
  accent: "violet",
  density: "comfortable",
  pinnedTools: [],
  themeMigrated: true,
};

export const MAX_PINNED_TOOLS = 6;

/**
 * Normalises stored settings (pure – unit-tested).
 * One-time switch to dark: PCs that had the old default ("Match Office") move to dark once.
 * After that, whatever the person picks (light, dark or Match Office) is kept.
 */
export function normalizeSettings(parsed: Partial<ZeeSettings>): ZeeSettings {
  const merged: ZeeSettings = { ...DEFAULT_SETTINGS, ...parsed };
  if (!parsed.themeMigrated) {
    if (parsed.theme === undefined || parsed.theme === "auto") merged.theme = "dark";
    merged.themeMigrated = true;
  }
  if (!["auto", "light", "dark"].includes(merged.theme)) merged.theme = "dark";
  if (!["apiKey", "openai", "google", "proxy"].includes(merged.authMode)) merged.authMode = "apiKey";
  if (typeof merged.otherBaseUrl !== "string" || !merged.otherBaseUrl.trim()) merged.otherBaseUrl = DEFAULT_OTHER_BASE_URL;
  if (typeof merged.otherApiKey !== "string") merged.otherApiKey = "";
  if (!Array.isArray(merged.otherFetchedModels)) merged.otherFetchedModels = [];
  if (typeof merged.model !== "string" || !merged.model.trim()) merged.model = merged.authMode === "openai" ? merged.otherModel || DEFAULT_OTHER_MODEL : DEFAULT_MODELS[0].id;
  if (!ACCENTS.some((a) => a.id === merged.accent)) merged.accent = "violet";
  if (merged.density !== "compact") merged.density = "comfortable";
  merged.pinnedTools = Array.isArray(merged.pinnedTools) ? merged.pinnedTools.filter((t) => typeof t === "string").slice(0, MAX_PINNED_TOOLS) : [];
  if (!Number.isFinite(merged.priceInPerM) || merged.priceInPerM < 0) merged.priceInPerM = 0;
  if (!Number.isFinite(merged.priceOutPerM) || merged.priceOutPerM < 0) merged.priceOutPerM = 0;
  if (!Number.isFinite(merged.bigDataRows) || merged.bigDataRows < 1000) merged.bigDataRows = DEFAULT_SETTINGS.bigDataRows;
  return merged;
}

export const STORAGE_KEY = "zeeai.settings.v1";

export function loadSettings(): ZeeSettings {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return { ...DEFAULT_SETTINGS };
    const parsed = JSON.parse(raw) as Partial<ZeeSettings>;
    const merged = normalizeSettings(parsed);
    // Persist the one-time dark switch right away so it never runs twice.
    if (!parsed.themeMigrated) saveSettings(merged);
    return merged;
  } catch {
    return { ...DEFAULT_SETTINGS };
  }
}

export function saveSettings(s: ZeeSettings): void {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(s));
  } catch {
    // Storage can be unavailable (private mode / policy). Settings then live for this session only.
  }
}

export function modelLabel(id: string): string {
  return DEFAULT_MODELS.find((m) => m.id === id)?.label ?? id;
}

export function modelHint(id: string): string | undefined {
  return DEFAULT_MODELS.find((m) => m.id === id)?.hint;
}

/**
 * Settings patch for switching connection. Gemini and the other provider each keep their own model,
 * so switching back and forth restores the model that was picked for that side.
 */
export function switchAuthMode(s: ZeeSettings, mode: AuthMode): Partial<ZeeSettings> {
  if (mode === s.authMode) return {};
  if (mode === "openai") return { authMode: mode, geminiModel: s.model, model: s.otherModel.trim() || DEFAULT_OTHER_MODEL };
  if (s.authMode === "openai") return { authMode: mode, otherModel: s.model, model: s.geminiModel.trim() || DEFAULT_MODELS[0].id };
  return { authMode: mode };
}

export function isConfigured(s: ZeeSettings): boolean {
  if (s.authMode === "apiKey") return s.apiKey.trim().length > 0;
  if (s.authMode === "openai") return s.otherApiKey.trim().length > 0 && s.otherBaseUrl.trim().length > 0;
  if (s.authMode === "proxy") return s.proxyUrl.trim().length > 0;
  return s.googleClientId.trim().length > 0 && s.gcpProjectId.trim().length > 0;
}
