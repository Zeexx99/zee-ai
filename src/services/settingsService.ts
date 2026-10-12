/**
 * Settings persistence.
 *
 * Note: Office.context.roamingSettings exists only in Outlook add-ins. In Excel the options are
 * Office.context.document.settings (saved INSIDE the workbook – wrong place for an API key) or
 * browser storage. Zee AI therefore keeps settings in localStorage of the add-in's own origin,
 * which is per user, per machine and never written into the workbook.
 */

export type AuthMode = "apiKey" | "google" | "proxy";

/** "auto" follows the Office theme (or the OS colour scheme outside Office). */
export type ThemeMode = "auto" | "light" | "dark";

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
}

export const DEFAULT_SETTINGS: ZeeSettings = {
  authMode: "apiKey",
  apiKey: "",
  model: DEFAULT_MODELS[0].id,
  fetchedModels: [],
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
  theme: "auto",
  webSearch: false,
  codeExecution: false,
};

const STORAGE_KEY = "zeeai.settings.v1";

export function loadSettings(): ZeeSettings {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return { ...DEFAULT_SETTINGS };
    const parsed = JSON.parse(raw) as Partial<ZeeSettings>;
    const merged = { ...DEFAULT_SETTINGS, ...parsed };
    if (!["auto", "light", "dark"].includes(merged.theme)) merged.theme = "auto";
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

export function isConfigured(s: ZeeSettings): boolean {
  if (s.authMode === "apiKey") return s.apiKey.trim().length > 0;
  if (s.authMode === "proxy") return s.proxyUrl.trim().length > 0;
  return s.googleClientId.trim().length > 0 && s.gcpProjectId.trim().length > 0;
}
