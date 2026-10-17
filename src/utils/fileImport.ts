/**
 * File import: CSV / TSV / TXT (delimiter auto-detected), LAS 2.0 well logs and SCADA / historian exports
 * (metadata lines on top, or "long" Tag–Timestamp–Value files that can be pivoted to one column per tag).
 * Pure (works on the file's text) – unit tested.
 */
import { latinDigits, parseLooseNumber } from "./importFix";

export type FileKind = "csv" | "las" | "scada";

export interface ParsedFile {
  kind: FileKind;
  headers: string[];
  rows: unknown[][];
  /** Column indexes holding Excel date/time serials (formatted on write). */
  dateCols: number[];
  /** Lines/info worth showing in the preview (LAS well info, skipped metadata). */
  notes: string[];
  /** Long-format SCADA file that can be pivoted: tag, time and value columns. */
  long?: { tagCol: number; timeCol: number; valueCol: number };
}

export interface ParseOptions {
  delimiter?: string;
  /** Rows to skip before the header (auto when undefined). */
  skipRows?: number;
  /** Day-first (DMY, default) or month-first dates in text. */
  dateOrder?: "DMY" | "MDY";
}

const EPOCH = Date.UTC(1899, 11, 30);

export function detectKind(fileName: string, text: string): FileKind {
  const ext = fileName.toLowerCase().split(".").pop() ?? "";
  if (ext === "las" || /^\s*~V/im.test(text.slice(0, 2000))) return "las";
  return "csv";
}

/* ---------------------------------------------------------------- delimited text */

export function detectDelimiter(lines: string[]): string {
  const cands = [",", ";", "\t", "|"];
  let best = ",";
  let bestScore = -1;
  for (const d of cands) {
    const counts = lines.slice(0, 30).filter((l) => l.trim()).map((l) => splitLine(l, d).length);
    if (!counts.length) continue;
    const mode = counts.sort((a, b) => counts.filter((x) => x === b).length - counts.filter((x) => x === a).length)[0];
    const agree = counts.filter((c) => c === mode).length;
    const score = mode > 1 ? agree * mode : 0;
    if (score > bestScore) {
      bestScore = score;
      best = d;
    }
  }
  return best;
}

/** Split one line honouring double quotes ("a, b" and "" escapes). */
export function splitLine(line: string, delim: string): string[] {
  const out: string[] = [];
  let cur = "";
  let q = false;
  for (let i = 0; i < line.length; i++) {
    const ch = line[i];
    if (q) {
      if (ch === '"') {
        if (line[i + 1] === '"') {
          cur += '"';
          i++;
        } else q = false;
      } else cur += ch;
    } else if (ch === '"' && cur.trim() === "") {
      q = true;
      cur = "";
    } else if (ch === delim) {
      out.push(cur);
      cur = "";
    } else cur += ch;
  }
  out.push(cur);
  return out.map((s) => s.trim());
}

/** Split text into records, keeping newlines inside quoted fields. */
export function splitRecords(text: string): string[] {
  const out: string[] = [];
  let cur = "";
  let q = false;
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (ch === '"') q = !q;
    if (!q && (ch === "\n" || ch === "\r")) {
      if (ch === "\r" && text[i + 1] === "\n") i++;
      out.push(cur);
      cur = "";
    } else cur += ch;
  }
  if (cur) out.push(cur);
  return out;
}

/** Date or date-time text → Excel serial (with time fraction). null when not a date. */
export function parseDateTime(raw: string, order: "DMY" | "MDY" = "DMY"): number | null {
  const s = latinDigits(raw).trim().replace(/\s+/g, " ");
  let y: number, mo: number, d: number;
  let rest = "";
  let m = s.match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})(?:[T ](.*))?$/);
  if (m) {
    [y, mo, d] = [+m[1], +m[2], +m[3]];
    rest = m[4] ?? "";
  } else {
    m = s.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{2}|\d{4})(?: (.*))?$/);
    if (!m) return null;
    const a = +m[1];
    const b = +m[2];
    const dmy = order === "DMY" ? !(a <= 12 && b > 12) : a > 12;
    [d, mo] = dmy ? [a, b] : [b, a];
    y = +m[3] < 100 ? 2000 + +m[3] : +m[3];
    rest = m[4] ?? "";
  }
  if (mo < 1 || mo > 12 || d < 1 || d > 31) return null;
  let frac = 0;
  if (rest) {
    const t = rest.replace(/(Z|[+-]\d{2}:?\d{2})$/, "").match(/^(\d{1,2}):(\d{2})(?::(\d{2})(?:\.\d+)?)?\s*([AaPp][Mm])?$/);
    if (!t) return null;
    let h = +t[1];
    if (t[4]) h = (h % 12) + (/p/i.test(t[4]) ? 12 : 0);
    frac = (h * 3600 + +t[2] * 60 + (t[3] ? +t[3] : 0)) / 86400;
  }
  const serial = (Date.UTC(y, mo - 1, d) - EPOCH) / 86_400_000;
  return Math.round((serial + frac) * 1e8) / 1e8;
}

function convertCell(s: string, order: "DMY" | "MDY"): { v: unknown; date: boolean } {
  const t = s.trim();
  if (t === "") return { v: "", date: false };
  if (/^(null|nan|n\/a|#n\/a|-999(\.0+)?|-999\.25)$/i.test(t)) return { v: "", date: false };
  const dt = /\d[-/.]\d/.test(t) ? parseDateTime(t, order) : null;
  if (dt !== null) return { v: dt, date: true };
  if (/^[+-]?[\d٠-٩۰-۹.,]+(e[+-]?\d+)?$/i.test(t)) {
    const n = parseLooseNumber(t);
    if (n && !n.unit) return { v: n.value, date: false };
    const plain = Number(t);
    if (Number.isFinite(plain)) return { v: plain, date: false };
  }
  return { v: t, date: false };
}

/** Header row = first row (after metadata) whose cells are mostly text and the next row is wider/numeric. */
export function guessHeaderRow(cells: string[][]): number {
  const width = (r: string[]) => r.filter((c) => c !== "").length;
  const target = Math.max(...cells.slice(0, 60).map(width), 0);
  for (let i = 0; i < Math.min(cells.length, 60); i++) {
    const r = cells[i];
    if (width(r) < Math.max(2, target * 0.6)) continue;
    const texty = r.filter((c) => c !== "" && !/^[+-]?\d+([.,]\d+)?$/.test(c)).length;
    if (texty >= width(r) * 0.6) return i;
    return -1; // first wide row is data → no header
  }
  return -1;
}

export function parseDelimited(text: string, o: ParseOptions = {}): ParsedFile {
  const records = splitRecords(text.replace(/^﻿/, ""));
  const delim = o.delimiter ?? detectDelimiter(records);
  const cells = records.map((l) => splitLine(l, delim));
  const notes: string[] = [];
  let h = o.skipRows ?? guessHeaderRow(cells);
  let headers: string[];
  let body: string[][];
  if (h < 0) {
    const w = Math.max(...cells.map((r) => r.length));
    headers = Array.from({ length: w }, (_, j) => `Column ${j + 1}`);
    body = cells;
  } else {
    if (h > 0) notes.push(`${h} line${h === 1 ? "" : "s"} above the header skipped${o.skipRows === undefined ? " (looked like file info)" : ""}: ${records.slice(0, Math.min(h, 3)).join(" | ").slice(0, 160)}`);
    headers = cells[h].map((c, j) => c || `Column ${j + 1}`);
    body = cells.slice(h + 1);
  }
  h = Math.max(h, 0);
  body = body.filter((r) => r.some((c) => c !== ""));
  const order = o.dateOrder ?? guessOrder(body.flat());
  const width = Math.max(headers.length, ...body.map((r) => r.length));
  while (headers.length < width) headers.push(`Column ${headers.length + 1}`);
  const dateHits = new Array(width).fill(0);
  const rows = body.map((r) =>
    Array.from({ length: width }, (_, j) => {
      const c = convertCell(r[j] ?? "", order);
      if (c.date) dateHits[j]++;
      return c.v;
    })
  );
  const dateCols = dateHits.map((n, j) => ({ n, j })).filter((x) => x.n > 0 && x.n >= rows.filter((r) => r[x.j] !== "").length * 0.6).map((x) => x.j);
  const long = detectLong(headers, rows, dateCols);
  return { kind: long ? "scada" : "csv", headers, rows, dateCols, notes, ...(long ? { long } : {}) };
}

function guessOrder(texts: string[]): "DMY" | "MDY" {
  let dmy = 0;
  let mdy = 0;
  for (const t of texts.slice(0, 5000)) {
    const m = t.match(/^(\d{1,2})[-/.](\d{1,2})[-/.]\d{2,4}/);
    if (!m) continue;
    if (+m[1] > 12) dmy++;
    else if (+m[2] > 12) mdy++;
  }
  return mdy > dmy ? "MDY" : "DMY";
}

/* ---------------------------------------------------------------- SCADA long format */

function detectLong(headers: string[], rows: unknown[][], dateCols: number[]): ParsedFile["long"] | undefined {
  const find = (re: RegExp) => headers.findIndex((h) => re.test(h));
  const tagCol = find(/^(tag|tag ?name|point|point ?name|signal|parameter|variable|name|description)$/i);
  const valueCol = find(/^(value|val|reading|pv|avg|average|data)$/i);
  let timeCol = find(/^(time|timestamp|date ?time|date|time ?stamp|ts)$/i);
  if (timeCol < 0 && dateCols.length) timeCol = dateCols[0];
  if (tagCol < 0 || valueCol < 0 || timeCol < 0) return undefined;
  const tags = new Set(rows.slice(0, 5000).map((r) => String(r[tagCol] ?? "")));
  if (tags.size < 2 || tags.size > 500) return undefined;
  return { tagCol, timeCol, valueCol };
}

/** Long (time, tag, value) → wide (time + one column per tag). Same timestamps merge into one row. */
export function pivotLong(p: ParsedFile, roundMinutes = 0): ParsedFile {
  if (!p.long) return p;
  const { tagCol, timeCol, valueCol } = p.long;
  const tags: string[] = [];
  const tagIdx = new Map<string, number>();
  const byTime = new Map<number | string, unknown[]>();
  const step = roundMinutes > 0 ? roundMinutes / 1440 : 0;
  for (const r of p.rows) {
    const tag = String(r[tagCol] ?? "").trim();
    if (!tag) continue;
    let t = r[timeCol] as number | string;
    if (typeof t === "number" && step) t = Math.round(t / step) * step;
    if (!tagIdx.has(tag)) {
      tagIdx.set(tag, tags.length);
      tags.push(tag);
    }
    const row = byTime.get(t) ?? [];
    row[tagIdx.get(tag)!] = r[valueCol];
    byTime.set(t, row);
  }
  const times = Array.from(byTime.keys()).sort((a, b) => (typeof a === "number" && typeof b === "number" ? a - b : String(a).localeCompare(String(b))));
  const rows = times.map((t) => {
    const vals = byTime.get(t)!;
    return [typeof t === "number" ? Math.round(t * 1e8) / 1e8 : t, ...tags.map((_, k) => vals[k] ?? "")];
  });
  return {
    kind: "scada",
    headers: [p.headers[timeCol], ...tags],
    rows,
    dateCols: p.dateCols.includes(timeCol) ? [0] : [],
    notes: [...p.notes, `Pivoted ${p.rows.length.toLocaleString()} tag readings into ${rows.length.toLocaleString()} rows × ${tags.length} tags.`],
  };
}

/* ---------------------------------------------------------------- LAS 2.0 */

export function parseLas(text: string): ParsedFile {
  const lines = text.replace(/^﻿/, "").split(/\r?\n/);
  let section = "";
  let nullValue: number | null = -999.25;
  let wrap = false;
  const curves: string[] = [];
  const info: string[] = [];
  const data: number[] = [];
  for (const raw of lines) {
    const line = raw.trimEnd();
    if (!line.trim() || line.trim().startsWith("#")) continue;
    if (line.trim().startsWith("~")) {
      section = line.trim()[1]?.toUpperCase() ?? "";
      continue;
    }
    if (section === "A") {
      for (const tok of line.trim().split(/[\s,]+/)) data.push(Number(tok));
      continue;
    }
    // MNEM.UNIT   DATA : DESCRIPTION
    const m = line.match(/^\s*([^.\s][^.]*?)\s*\.(\S*)\s*(.*?)\s*:\s*(.*)$/);
    if (!m) continue;
    const [, mnem, unit, value, desc] = m;
    if (section === "V" && /^WRAP$/i.test(mnem)) wrap = /^YES$/i.test(value);
    if (section === "W") {
      if (/^NULL$/i.test(mnem) && Number.isFinite(Number(value))) nullValue = Number(value);
      if (value && /^(WELL|FLD|COMP|UWI|API|STRT|STOP|STEP|LOC|SRVC|DATE)$/i.test(mnem)) info.push(`${desc || mnem}: ${value}${unit && !/^(WELL|FLD|COMP|UWI|API|LOC|SRVC|DATE)$/i.test(mnem) ? ` ${unit}` : ""}`);
    }
    if (section === "C") curves.push(unit ? `${mnem} (${unit})` : mnem);
  }
  if (!curves.length) throw new Error("No ~Curve section found – is this a LAS 2.0 file?");
  const width = curves.length;
  const rows: unknown[][] = [];
  for (let i = 0; i + width <= data.length; i += width) {
    rows.push(data.slice(i, i + width).map((v) => (!Number.isFinite(v) || (nullValue !== null && Math.abs(v - nullValue) < 1e-9) ? "" : v)));
  }
  if (!rows.length) throw new Error("The ~ASCII data section is empty.");
  const notes = [...info.slice(0, 8)];
  if (wrap) notes.push("Wrapped LAS (WRAP = YES) – values were re-joined.");
  if (data.length % width) notes.push(`${data.length % width} trailing values didn't fill a full row and were dropped.`);
  return { kind: "las", headers: curves, rows, dateCols: [], notes };
}

export function parseFile(fileName: string, text: string, o: ParseOptions = {}): ParsedFile {
  return detectKind(fileName, text) === "las" ? parseLas(text) : parseDelimited(text, o);
}

export const IMPORT_ACCEPT = ".csv,.txt,.tsv,.dat,.las,.prn,.log";
/** Bigger files would freeze the pane while parsing / writing. */
export const IMPORT_MAX_BYTES = 60_000_000;
