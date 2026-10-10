/**
 * Google sign-in (OAuth 2.0) through the Office Dialog API.
 *
 * Flow: task pane -> displayDialogAsync(/auth/start.html) -> Google consent -> /auth/callback.html
 * -> Office.context.ui.messageParent(token) -> task pane stores the short-lived access token and
 * uses it as a Bearer token for Vertex AI (billing/quota go to the user's Google Cloud project).
 */

export interface GoogleSession {
  accessToken: string;
  expiresAt: number; // epoch ms
  email?: string;
}

const TOKEN_KEY = "zeeai.google.session.v1";

export function getGoogleSession(): GoogleSession | null {
  try {
    const raw = localStorage.getItem(TOKEN_KEY);
    if (!raw) return null;
    const s = JSON.parse(raw) as GoogleSession;
    return s.accessToken ? s : null;
  } catch {
    return null;
  }
}

/** Returns a token valid for at least another minute, or null. */
export function getValidGoogleToken(): string | null {
  const s = getGoogleSession();
  if (!s || s.expiresAt - Date.now() < 60_000) return null;
  return s.accessToken;
}

export function signOutGoogle(): void {
  const s = getGoogleSession();
  try {
    localStorage.removeItem(TOKEN_KEY);
  } catch {
    /* ignore */
  }
  if (s?.accessToken) {
    // Best effort token revocation.
    fetch(`https://oauth2.googleapis.com/revoke?token=${encodeURIComponent(s.accessToken)}`, {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
    }).catch(() => undefined);
  }
}

interface DialogPayload {
  ok: boolean;
  accessToken?: string;
  expiresIn?: number;
  state?: string;
  error?: string;
}

function randomState(): string {
  const bytes = new Uint8Array(16);
  crypto.getRandomValues(bytes);
  return Array.from(bytes, (b) => b.toString(16).padStart(2, "0")).join("");
}

async function fetchEmail(token: string): Promise<string | undefined> {
  try {
    const r = await fetch("https://openidconnect.googleapis.com/v1/userinfo", {
      headers: { Authorization: `Bearer ${token}` },
    });
    if (!r.ok) return undefined;
    const j = (await r.json()) as { email?: string };
    return j.email;
  } catch {
    return undefined;
  }
}

const DIALOG_ERRORS: Record<number, string> = {
  12004: "The sign-in page domain is not trusted. Check AppDomains in manifest.xml.",
  12005: "The sign-in page must be served over HTTPS.",
  12007: "A sign-in dialog is already open.",
  12009: "The sign-in popup was blocked. Allow pop-ups for Office and try again.",
};

export function signInWithGoogle(clientId: string): Promise<GoogleSession> {
  return new Promise((resolve, reject) => {
    if (typeof Office === "undefined" || !Office.context?.ui) {
      reject(new Error("Google sign-in only works when Zee AI runs inside Excel."));
      return;
    }
    if (!clientId.trim()) {
      reject(new Error("Enter your Google OAuth Client ID first."));
      return;
    }
    const state = randomState();
    const url = new URL(`auth/start.html?client_id=${encodeURIComponent(clientId.trim())}&state=${state}`, window.location.href).toString();

    Office.context.ui.displayDialogAsync(url, { height: 70, width: 35, promptBeforeOpen: false }, (result) => {
      if (result.status === Office.AsyncResultStatus.Failed) {
        const code = result.error.code;
        reject(new Error(DIALOG_ERRORS[code] ?? `Could not open sign-in dialog (${code}): ${result.error.message}`));
        return;
      }
      const dialog = result.value;
      let settled = false;

      dialog.addEventHandler(Office.EventType.DialogMessageReceived, async (arg) => {
        if (!("message" in arg)) return;
        settled = true;
        dialog.close();
        let payload: DialogPayload;
        try {
          payload = JSON.parse(String(arg.message)) as DialogPayload;
        } catch {
          reject(new Error("Unexpected response from sign-in dialog."));
          return;
        }
        if (!payload.ok || !payload.accessToken) {
          reject(new Error(payload.error === "access_denied" ? "Sign-in was cancelled." : `Google sign-in failed: ${payload.error ?? "unknown error"}`));
          return;
        }
        if (payload.state !== state) {
          reject(new Error("Sign-in response failed the security check (state mismatch). Please try again."));
          return;
        }
        const session: GoogleSession = {
          accessToken: payload.accessToken,
          expiresAt: Date.now() + (payload.expiresIn ?? 3600) * 1000,
        };
        session.email = await fetchEmail(session.accessToken);
        try {
          localStorage.setItem(TOKEN_KEY, JSON.stringify(session));
        } catch {
          /* ignore */
        }
        resolve(session);
      });

      dialog.addEventHandler(Office.EventType.DialogEventReceived, (arg) => {
        if (settled) return;
        settled = true;
        const code = "error" in arg ? arg.error : 0;
        reject(new Error(code === 12006 ? "Sign-in window was closed." : `Sign-in dialog error (${code}).`));
      });
    });
  });
}
