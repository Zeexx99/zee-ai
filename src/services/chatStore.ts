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

type StoredMessage = { promptText?: string; status?: string };

/**
 * Smaller and smaller versions of the history to try when storage is full:
 * 1) as is, 2) older prompts trimmed to 2,000 chars, 3+) dropping the oldest half each time.
 * Pure – unit-tested.
 */
export function* historyCandidates<T extends StoredMessage>(messages: T[]): Generator<T[], void, unknown> {
  const base = messages
    .filter((m) => m.status !== "streaming")
    .slice(-MAX_MESSAGES)
    .map((m) => (m.promptText && m.promptText.length > MAX_PROMPT_CHARS ? { ...m, promptText: m.promptText.slice(0, MAX_PROMPT_CHARS) } : m));
  yield base;
  let list = base.map((m, i) => (i < base.length - 4 && m.promptText && m.promptText.length > 2000 ? { ...m, promptText: `${m.promptText.slice(0, 2000)}…` } : m));
  yield list;
  while (list.length > 2) {
    let cut = Math.floor(list.length / 2);
    // Keep user/assistant pairs together.
    if (cut % 2 === 1) cut++;
    list = list.slice(Math.min(cut, list.length - 2));
    yield list;
  }
}

export function saveChat<T extends StoredMessage>(docId: string, messages: T[]): void {
  for (const candidate of historyCandidates(messages)) {
    try {
      localStorage.setItem(keyFor(docId), JSON.stringify(candidate));
      return;
    } catch {
      // Quota exceeded – try a smaller history (oldest messages go first).
    }
  }
  try {
    localStorage.removeItem(keyFor(docId));
  } catch {
    /* ignore */
  }
}

export function clearChat(docId: string): void {
  try {
    localStorage.removeItem(keyFor(docId));
  } catch {
    /* ignore */
  }
}
