/** Export / import settings to move them to another PC (file or clipboard). */
import { APP } from "./app";
import { VERSION } from "./version";

/** Keys never exported unless the user ticks "include API key / proxy token". */
export const SECRET_KEYS = ["apiKey", "proxyToken"];

export interface SettingsExport {
  zeeSettingsExport: 1;
  app: string;
  version: string;
  exportedAt: string;
  settings: Record<string, unknown>;
  /** Saved (favourite) prompts. */
  prompts?: string[];
}

export function buildSettingsExport(settings: object, opts: { includeSecrets: boolean; prompts?: string[] }): string {
  const out: Record<string, unknown> = {};
  for (const [k, v] of Object.entries(settings)) {
    if (!opts.includeSecrets && SECRET_KEYS.includes(k)) continue;
    out[k] = v;
  }
  const data: SettingsExport = {
    zeeSettingsExport: 1,
    app: APP.id,
    version: VERSION,
    exportedAt: new Date().toISOString(),
    settings: out,
    ...(opts.prompts?.length ? { prompts: opts.prompts } : {}),
  };
  return JSON.stringify(data, null, 2);
}

export interface ImportResult<T> {
  patch: Partial<T>;
  prompts?: string[];
  /** Which add-in made the file ("word" / "excel"). */
  from: string;
  /** Keys in the file that this add-in doesn't use or that had the wrong type. */
  skipped: string[];
}

/**
 * Parses an export file. Only keys this add-in knows (present in `defaults`) with a matching type are
 * applied, so a Word export can be imported into Excel (shared keys such as the API key and model move
 * across, Word-only options are skipped).
 */
export function parseSettingsExport<T extends object>(text: string, defaults: T): ImportResult<T> {
  let data: Partial<SettingsExport>;
  try {
    data = JSON.parse(text) as Partial<SettingsExport>;
  } catch {
    throw new Error("That isn't a Zee AI settings file (not valid JSON).");
  }
  if (!data || data.zeeSettingsExport !== 1 || typeof data.settings !== "object" || !data.settings) {
    throw new Error("That isn't a Zee AI settings file. Use “Export settings” to create one.");
  }
  const patch: Record<string, unknown> = {};
  const skipped: string[] = [];
  const known = defaults as Record<string, unknown>;
  for (const [k, v] of Object.entries(data.settings)) {
    if (!(k in known)) {
      skipped.push(k);
      continue;
    }
    const d = known[k];
    const ok =
      d === null
        ? v === null || typeof v === "number" || typeof v === "object"
        : Array.isArray(d)
          ? Array.isArray(v)
          : typeof v === typeof d || (typeof d === "object" && (v === null || typeof v === "object"));
    if (ok) patch[k] = v;
    else skipped.push(k);
  }
  const prompts = Array.isArray(data.prompts) ? data.prompts.filter((p): p is string => typeof p === "string" && p.trim() !== "") : undefined;
  return { patch: patch as Partial<T>, prompts, from: typeof data.app === "string" ? data.app : "unknown", skipped };
}

/** Offers text as a file download. Returns false where downloads are blocked. */
export function downloadFile(filename: string, text: string, type = "application/json"): boolean {
  try {
    const url = URL.createObjectURL(new Blob([text], { type }));
    const a = document.createElement("a");
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 5000);
    return true;
  } catch {
    return false;
  }
}
