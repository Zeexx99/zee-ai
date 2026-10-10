/**
 * Per-workbook chat history. Each workbook gets a random id stored in its document settings
 * (a tiny, harmless value saved with the file); messages live in localStorage under that id.
 */
const DOC_ID_KEY = "zeeai.docId";
const MAX_MESSAGES = 60;
const MAX_PROMPT_CHARS = 12_000;

let cachedDocId: string | null = null;

export function getDocId(): string {
  if (cachedDocId) return cachedDocId;
  try {
    const settings = Office.context?.document?.settings;
    if (settings) {
      let id = settings.get(DOC_ID_KEY) as string | null;
      if (!id) {
        id = `wb-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`;
        settings.set(DOC_ID_KEY, id);
        settings.saveAsync();
      }
      cachedDocId = id;
      return id;
    }
  } catch {
    /* not in Office */
  }
  cachedDocId = "default";
  return cachedDocId;
}

const keyFor = (docId: string) => `zeeai.chat.${docId}`;

export function loadChat<T>(docId: string): T[] {
  try {
    const raw = localStorage.getItem(keyFor(docId));
    return raw ? (JSON.parse(raw) as T[]) : [];
  } catch {
    return [];
  }
}

export function saveChat<T extends { promptText?: string; status?: string }>(docId: string, messages: T[]): void {
  try {
    const trimmed = messages
      .filter((m) => m.status !== "streaming")
      .slice(-MAX_MESSAGES)
      .map((m) => (m.promptText && m.promptText.length > MAX_PROMPT_CHARS ? { ...m, promptText: m.promptText.slice(0, MAX_PROMPT_CHARS) } : m));
    localStorage.setItem(keyFor(docId), JSON.stringify(trimmed));
  } catch {
    // Quota exceeded: drop history for this workbook rather than failing.
    try {
      localStorage.removeItem(keyFor(docId));
    } catch {
      /* ignore */
    }
  }
}

export function clearChat(docId: string): void {
  try {
    localStorage.removeItem(keyFor(docId));
  } catch {
    /* ignore */
  }
}
