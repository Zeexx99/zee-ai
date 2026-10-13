import { createContext, useContext } from "react";
import type { Cell } from "../utils/markdown";
import type { ChatMessage } from "../hooks/useChat";

/** Actions that chat output can trigger (implemented in App, which owns confirm + toasts). */
export interface InsertActions {
  insertFormulas: (formulas: string[]) => Promise<void>;
  /** asTable = formatted Excel table with filters (otherwise plain cells with a bold header). */
  insertTable: (matrix: Cell[][], asTable?: boolean) => Promise<void>;
  /** Agent mode: approve or reject the pending batch of workbook changes. */
  respondToApproval: (approved: boolean) => void;
  /** Agent mode: keep going after the step limit. */
  continueAgent: () => void;
  /** Run an answer again (Regenerate / Retry). */
  retry: (assistantId: string) => void;
  /** Put a prompt back into the chat box to edit and resend it. */
  editPrompt: (text: string) => void;
  /** Copy a diagnostics report for a failed answer. */
  copyDiagnostics: (message: ChatMessage) => void;
  /** A request is running (Regenerate is disabled meanwhile). */
  busy: boolean;
}

export const InsertActionsContext = createContext<InsertActions | null>(null);

export function useInsertActions(): InsertActions {
  const ctx = useContext(InsertActionsContext);
  if (!ctx) throw new Error("InsertActionsContext missing");
  return ctx;
}

/** Open a URL in the user's browser (Office dialog API inside Excel, a new tab elsewhere). */
export function openExternal(url: string) {
  try {
    Office.context.ui.openBrowserWindow(url);
  } catch {
    window.open(url, "_blank", "noopener");
  }
}

export async function copyText(text: string): Promise<boolean> {
  try {
    await navigator.clipboard.writeText(text);
    return true;
  } catch {
    // Clipboard API is often blocked in Office webviews – fall back to execCommand.
    const ta = document.createElement("textarea");
    ta.value = text;
    ta.style.position = "fixed";
    ta.style.opacity = "0";
    document.body.appendChild(ta);
    ta.select();
    const ok = document.execCommand("copy");
    ta.remove();
    return ok;
  }
}
