/** Friendly error hints and a "Copy diagnostics" report (never includes keys or document text). */
import { APP } from "./app";
import { BUILD, VERSION } from "./version";

export interface ErrorHelp {
  /** One-line advice shown under the error. */
  hint: string | null;
  /** Retrying the same request may well work. */
  retryable: boolean;
}

export function errorHelp(message: string, kind?: string): ErrorHelp {
  const m = message.toLowerCase();
  if (kind === "aborted" || /^stopped\.?$/.test(m)) return { hint: null, retryable: true };
  if (kind === "network" || /network|failed to fetch|internet|firewall|timed? ?out/.test(m))
    return {
      hint: "Check your internet connection or VPN. If the connection dropped, this is sent again automatically when it's back – or press Retry. A company firewall may block generativelanguage.googleapis.com – the team proxy avoids that.",
      retryable: true,
    };
  if (kind === "quota" || /\b429\b|quota|rate limit|resource.?exhausted|too many requests/.test(m))
    return { hint: "You've hit Gemini's rate limit or quota. Wait a minute and Retry, or pick a lighter model (Flash-Lite) in Settings.", retryable: true };
  if (kind === "auth" || /api key|authenticat|sign in|unauthori[sz]ed|\b401\b|\b403\b|permission/.test(m))
    return { hint: "Open Settings → Connection and click “Test connection”. Re-paste the API key or sign in again.", retryable: false };
  if (kind === "notFound" || /\b404\b|isn't available|not found|not supported for/.test(m))
    return { hint: "Choose another model in Settings (click ⟳ next to the model list to load the ones your key can use).", retryable: false };
  if (kind === "safety" || /safety|blocked/.test(m)) return { hint: "Gemini's safety filter blocked this. Rephrase the request or send less of the document.", retryable: false };
  if (kind === "config" || /settings/.test(m)) return { hint: "Finish the setup in Settings, then Retry.", retryable: false };
  if (kind === "server" || /\b5\d\d\b|unavailable|overloaded|internal error/.test(m)) return { hint: "Gemini is busy right now. Retry in a few seconds.", retryable: true };
  if (/malformed json/.test(m)) return { hint: "The model returned an incomplete answer. Retry, or use a Pro model for long documents.", retryable: true };
  if (/document changed|no longer exists|deleted or moved/.test(m)) return { hint: "The document changed since this answer. Select the text again and Retry.", retryable: true };
  return { hint: null, retryable: true };
}

/** Host details from Office.js (empty outside Office). */
function officeInfo(): string {
  try {
    const d = Office?.context?.diagnostics;
    if (!d) return "not in Office";
    return `${String(d.host)} ${d.version} on ${String(d.platform)}`;
  } catch {
    return "not in Office";
  }
}

/** Plain-text report to paste to Claude / IT. Contains no API keys, tokens or document text. */
export function buildDiagnostics(info: { error?: string; kind?: string; model?: string; authMode?: string; extra?: Record<string, string | number | undefined> }): string {
  const lines = [
    `${APP.name} diagnostics`,
    `Time: ${new Date().toISOString()}`,
    `Version: ${VERSION}${BUILD ? ` (built ${BUILD})` : ""}`,
    `Office: ${officeInfo()}`,
    `Page: ${typeof location !== "undefined" ? location.origin + location.pathname : "?"}`,
    `Browser: ${typeof navigator !== "undefined" ? navigator.userAgent : "?"}`,
    `Online: ${typeof navigator !== "undefined" ? String(navigator.onLine) : "?"}`,
  ];
  if (info.model) lines.push(`Model: ${info.model}`);
  if (info.authMode) lines.push(`Connection: ${info.authMode}`);
  if (info.kind) lines.push(`Error kind: ${info.kind}`);
  if (info.error) lines.push(`Error: ${info.error}`);
  for (const [k, v] of Object.entries(info.extra ?? {})) if (v !== undefined && v !== "") lines.push(`${k}: ${v}`);
  return lines.join("\n");
}
