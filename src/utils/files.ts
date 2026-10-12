/** Attachment rules for chat (images + PDFs sent inline to Gemini). */

export const ACCEPTED_TYPES = ["image/png", "image/jpeg", "image/webp", "image/heic", "image/heif", "application/pdf"];
export const ACCEPT_ATTR = ACCEPTED_TYPES.join(",");
export const MAX_FILES = 5;
/** Gemini's inline request limit is 20 MB; base64 adds a third, so keep raw files under ~14 MB in total. */
export const MAX_TOTAL_BYTES = 14_000_000;

export interface FileLike {
  name: string;
  type: string;
  size: number;
}

/** MIME type for a file, falling back to the extension (Office webviews sometimes report ""). */
export function mimeOf(f: FileLike): string {
  if (f.type) return f.type === "image/jpg" ? "image/jpeg" : f.type;
  const ext = f.name.toLowerCase().split(".").pop() ?? "";
  return { png: "image/png", jpg: "image/jpeg", jpeg: "image/jpeg", webp: "image/webp", heic: "image/heic", heif: "image/heif", pdf: "application/pdf" }[ext] ?? "";
}

/**
 * Decide which new files can be added next to the ones already attached.
 * Returns the accepted files and a message for the rejected ones (or null).
 */
export function pickFiles<T extends FileLike>(incoming: T[], current: Array<{ size: number }>): { accepted: T[]; error: string | null } {
  const accepted: T[] = [];
  const problems: string[] = [];
  let total = current.reduce((n, f) => n + f.size, 0);
  let count = current.length;
  for (const f of incoming) {
    const mime = mimeOf(f);
    if (!ACCEPTED_TYPES.includes(mime)) {
      problems.push(`${f.name}: only images (PNG, JPG, WebP, HEIC) and PDFs`);
      continue;
    }
    if (count >= MAX_FILES) {
      problems.push(`${f.name}: at most ${MAX_FILES} files per message`);
      continue;
    }
    if (total + f.size > MAX_TOTAL_BYTES) {
      problems.push(`${f.name}: files must add up to less than ${Math.round(MAX_TOTAL_BYTES / 1_000_000)} MB`);
      continue;
    }
    accepted.push(f);
    total += f.size;
    count++;
  }
  return { accepted, error: problems.length ? problems.join("; ") : null };
}

/** "data:image/png;base64,AAAA" -> "AAAA" */
export function stripDataUrl(dataUrl: string): string {
  const i = dataUrl.indexOf(",");
  return i >= 0 ? dataUrl.slice(i + 1) : dataUrl;
}

/** Read a browser File as base64 (no data: prefix). */
export function readAsBase64(file: Blob): Promise<string> {
  return new Promise((resolve, reject) => {
    const r = new FileReader();
    r.onload = () => resolve(stripDataUrl(String(r.result)));
    r.onerror = () => reject(r.error ?? new Error("Could not read the file."));
    r.readAsDataURL(file);
  });
}
