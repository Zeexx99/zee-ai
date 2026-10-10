/** Helpers to turn assistant output (markdown tables / code blocks / CSV) into Excel-ready matrices. */

export type Cell = string | number | boolean;

/** Coerce a text cell into the most useful Excel value. Formulas stay as text starting with "=". */
export function coerceCell(raw: string): Cell {
  const s = raw.trim();
  if (s === "") return "";
  if (s.startsWith("=")) return s;
  if (/^(true|false)$/i.test(s)) return s.toUpperCase() === "TRUE";
  if (/^[-+]?0\d/.test(s)) return s; // keep leading-zero identifiers as text
  if (/^[-+]?(\d{1,3}(,\d{3})+|\d+)(\.\d+)?$/.test(s) && s.replace(/\D/g, "").length <= 15) {
    return Number(s.replace(/,/g, ""));
  }
  return s;
}

/** Make every row the same width. */
export function normalizeMatrix(rows: Cell[][]): Cell[][] {
  const width = rows.reduce((w, r) => Math.max(w, r.length), 0);
  return rows.map((r) => {
    const out = r.slice();
    while (out.length < width) out.push("");
    return out;
  });
}

/** Parse CSV/TSV text (handles quoted fields). */
export function parseDelimited(text: string): Cell[][] {
  const lines = text.replace(/\r\n/g, "\n").trim();
  if (!lines) return [];
  const firstLine = lines.split("\n")[0];
  const delim = firstLine.includes("\t") ? "\t" : firstLine.includes(";") && !firstLine.includes(",") ? ";" : ",";
  const rows: string[][] = [];
  let row: string[] = [];
  let field = "";
  let inQuotes = false;
  for (let i = 0; i < lines.length; i++) {
    const ch = lines[i];
    if (inQuotes) {
      if (ch === '"' && lines[i + 1] === '"') {
        field += '"';
        i++;
      } else if (ch === '"') inQuotes = false;
      else field += ch;
    } else if (ch === '"') inQuotes = true;
    else if (ch === delim) {
      row.push(field);
      field = "";
    } else if (ch === "\n") {
      row.push(field);
      rows.push(row);
      row = [];
      field = "";
    } else field += ch;
  }
  row.push(field);
  rows.push(row);
  return normalizeMatrix(rows.map((r) => r.map(coerceCell)));
}

/** Minimal hast-like node shape (avoids a hard dependency on @types/hast). */
export interface HastLike {
  type: string;
  tagName?: string;
  value?: string;
  properties?: Record<string, unknown>;
  children?: HastLike[];
}

export function hastText(node: HastLike | undefined): string {
  if (!node) return "";
  if (node.type === "text") return node.value ?? "";
  if (node.tagName === "br") return "\n";
  return (node.children ?? []).map(hastText).join("");
}

/** Convert a rendered markdown <table> node to a matrix (header row first). */
export function hastTableToMatrix(table: HastLike): Cell[][] {
  const rows: Cell[][] = [];
  const walk = (n: HastLike) => {
    if (n.tagName === "tr") {
      rows.push(
        (n.children ?? [])
          .filter((c) => c.tagName === "td" || c.tagName === "th")
          .map((c) => coerceCell(hastText(c)))
      );
      return;
    }
    (n.children ?? []).forEach(walk);
  };
  walk(table);
  return normalizeMatrix(rows);
}

/** Split a formula code block into individual formulas (one per non-empty line starting with "="). */
export function formulasFromCode(code: string): string[] {
  const lines = code
    .split(/\r?\n/)
    .map((l) => l.trim())
    .filter(Boolean);
  if (lines.length === 0) return [];
  if (lines.every((l) => l.startsWith("="))) return lines;
  // A single multi-line formula (e.g. a formatted LET): join it.
  if (lines[0].startsWith("=")) return [lines.join(" ")];
  return [];
}
