/**
 * Data menu: import a file (CSV / LAS / SCADA), merge well data from several sheets, resample time series.
 * Results always go to a NEW sheet (undo = remove that sheet), the source data is never changed.
 */
import { ExcelUserError, addSheet, invalidateOverview, loadChunked, run, setSnapshot, writeChunked } from "./excelService";
import { findHeaderRow, mergeTables, type MergeOptions, type MergeResult, type MergeTable } from "../utils/merge";
import { resample, type ResampleOptions, type ResampleResult } from "../utils/resample";
import type { ParsedFile } from "../utils/fileImport";
import type { HeaderedSelection } from "./toolActions";
import { BIG_TABLE_CELLS } from "./toolActions";
import { a1, columnLetter } from "../utils/a1";

export type Progress = (fraction: number) => void;

/* ================================================================ shared: write a table to a new sheet */

export interface TableWrite {
  sheetBase: string;
  headers: string[];
  rows: unknown[][];
  /** Columns holding Excel date serials. */
  dateCols: number[];
  /** Rows to tint (e.g. inserted periods). */
  marked?: boolean[];
  /** Also turn it into an Excel table (filter buttons). Skipped above 100k rows to stay quick. */
  asTable?: boolean;
}

const MAX_MARKS = 3000;

export async function writeTableToNewSheet(t: TableWrite, onProgress?: Progress): Promise<{ sheet: string; address: string }> {
  if (!t.rows.length) throw new ExcelUserError("Nothing to write – the result has no rows.");
  const width = t.headers.length;
  if (t.rows.length + 1 > 1_048_576) throw new ExcelUserError(`The result has ${t.rows.length.toLocaleString()} rows – more than one Excel sheet can hold.`);
  if (width > 16_384) throw new ExcelUserError("The result has more columns than Excel allows.");
  const sheet = await addSheet(t.sheetBase);
  const data = t.rows.map((r) => Array.from({ length: width }, (_, j) => sanitize(r[j])));
  const address = await run(async (ctx) => {
    const ws = ctx.workbook.worksheets.getItem(sheet);
    const head = ws.getRangeByIndexes(0, 0, 1, width);
    head.values = [t.headers];
    head.format.font.bold = true;
    head.format.fill.color = "#EDEBFB";
    await writeChunked(ctx, ws, 1, 0, data, "values", (done, total) => onProgress?.(done / total));
    for (const j of t.dateCols) {
      const hasTime = t.rows.some((r) => typeof r[j] === "number" && Math.abs((r[j] as number) % 1) > 1e-9);
      const fmt = hasTime ? "dd-mmm-yyyy hh:mm" : "dd-mmm-yyyy";
      await writeChunked(ctx, ws, 1, j, data.map(() => [fmt]), "numberFormat");
    }
    await ctx.sync();
    if (t.marked) {
      let n = 0;
      for (let i = 0; i < t.marked.length && n < MAX_MARKS; i++) {
        if (!t.marked[i]) continue;
        ws.getRangeByIndexes(1 + i, 0, 1, width).format.fill.color = "#DDEBF7";
        if (++n % 500 === 0) await ctx.sync();
      }
    }
    const all = ws.getRangeByIndexes(0, 0, data.length + 1, width);
    if (t.asTable && data.length <= 100_000) {
      try {
        const table = ws.tables.add(all, true);
        table.style = "TableStyleLight9";
        await ctx.sync();
      } catch {
        /* table creation is cosmetic */
      }
    }
    ws.freezePanes.freezeRows(1);
    // Auto-fit on a sample: fitting 100k+ rows is slow and the widths come out the same.
    ws.getRangeByIndexes(0, 0, Math.min(data.length + 1, 2000), width).format.autofitColumns();
    ws.getRange("A1").select();
    all.load("address");
    await ctx.sync();
    return all.address;
  });
  invalidateOverview();
  return { sheet, address };
}

/** Excel rejects undefined / NaN / Infinity and very long strings. */
function sanitize(v: unknown): string | number | boolean {
  if (v === null || v === undefined) return "";
  if (typeof v === "number") return Number.isFinite(v) ? v : "";
  if (typeof v === "boolean") return v;
  const s = String(v);
  // A text that starts with "=" would become a formula when written as a value – keep it as text.
  return (s.length > 32_000 ? s.slice(0, 32_000) : s).replace(/^=/, "'=");
}

/* ================================================================ sheets as tables */

export interface SheetInfo {
  name: string;
  rows: number;
  headers: string[];
  headerRow: number;
}

/** Every non-empty sheet with its header row (first 20 rows are scanned for it). */
export async function listSheetTables(): Promise<SheetInfo[]> {
  return run(async (ctx) => {
    const sheets = ctx.workbook.worksheets;
    sheets.load("items/name,items/visibility");
    await ctx.sync();
    const used = sheets.items.map((ws) => {
      const u = ws.getUsedRangeOrNullObject(true);
      u.load("rowCount,columnCount,rowIndex,columnIndex");
      return { ws, u };
    });
    await ctx.sync();
    const tops = used.map(({ ws, u }) => {
      if (u.isNullObject) return null;
      const top = ws.getRangeByIndexes(u.rowIndex, u.columnIndex, Math.min(20, u.rowCount), Math.min(u.columnCount, 200));
      top.load("values");
      return { ws, u, top };
    });
    await ctx.sync();
    return tops
      .filter((x): x is NonNullable<typeof x> => !!x && x.u.rowCount >= 2)
      .map(({ ws, u, top }) => {
        const h = findHeaderRow(top.values);
        return {
          name: ws.name,
          rows: u.rowCount - h - 1,
          headerRow: h,
          headers: top.values[h].map((v, j) => String(v ?? "").trim() || `Column ${columnLetter(u.columnIndex + j)}`),
        };
      });
  });
}

/** Read a whole sheet (used range) as header + rows, in chunks. */
export async function readSheetTable(name: string, onProgress?: Progress): Promise<{ headers: string[]; rows: unknown[][] }> {
  return run(async (ctx) => {
    const ws = ctx.workbook.worksheets.getItem(name);
    const u = ws.getUsedRangeOrNullObject(true);
    u.load("rowCount,columnCount,rowIndex,columnIndex");
    await ctx.sync();
    if (u.isNullObject) throw new ExcelUserError(`Sheet "${name}" is empty.`);
    if (u.rowCount * u.columnCount > BIG_TABLE_CELLS) throw new ExcelUserError(`Sheet "${name}" has more than ${BIG_TABLE_CELLS.toLocaleString()} cells – too large to merge in the task pane.`);
    const got = await loadChunked(ctx, ws, u.rowIndex, u.columnIndex, u.rowCount, u.columnCount, ["values"], (d, t) => onProgress?.(d / t));
    const h = findHeaderRow(got.values.slice(0, 20));
    return {
      headers: got.values[h].map((v, j) => String(v ?? "").trim() || `Column ${columnLetter(u.columnIndex + j)}`),
      rows: got.values.slice(h + 1),
    };
  });
}

/* ================================================================ merge */

export interface MergePick {
  sheet: string;
  wellCol: number;
  dateCol: number;
}

export interface MergeRun {
  sheet: string;
  result: MergeResult;
}

export async function runMerge(picks: MergePick[], o: MergeOptions, onProgress?: Progress): Promise<MergeRun> {
  if (picks.length < 2) throw new ExcelUserError("Pick at least two sheets to merge.");
  const tables: MergeTable[] = [];
  for (let i = 0; i < picks.length; i++) {
    const p = picks[i];
    const t = await readSheetTable(p.sheet, (f) => onProgress?.((i + f) / (picks.length + 1)));
    if (p.wellCol < 0 || p.wellCol >= t.headers.length) throw new ExcelUserError(`Choose the well column for "${p.sheet}".`);
    tables.push({ name: p.sheet, headers: t.headers, rows: t.rows, wellCol: p.wellCol, dateCol: p.dateCol });
  }
  let result: MergeResult;
  try {
    result = mergeTables(tables, o);
  } catch (e) {
    throw new ExcelUserError((e as Error).message);
  }
  const { sheet } = await writeTableToNewSheet(
    { sheetBase: "Merged", headers: result.headers, rows: result.rows, dateCols: result.dateCol >= 0 ? [result.dateCol] : [], asTable: true },
    (f) => onProgress?.((picks.length + f) / (picks.length + 1))
  );
  return { sheet, result };
}

/* ================================================================ resample */

export async function runResample(sel: HeaderedSelection, o: ResampleOptions, onProgress?: Progress): Promise<{ sheet: string; result: ResampleResult }> {
  const result = resample(sel.headers, sel.rows, o);
  if (!result.rows.length) throw new ExcelUserError("No rows with a valid date were found in the date column.");
  const dateOut = o.groupCol !== undefined ? 1 : 0;
  const { sheet } = await writeTableToNewSheet(
    { sheetBase: `${o.period === "day" ? "Daily" : o.period === "week" ? "Weekly" : o.period === "month" ? "Monthly" : o.period === "quarter" ? "Quarterly" : "Yearly"} ${sel.block.sheetName}`, headers: result.headers, rows: result.rows, dateCols: [dateOut], marked: result.added },
    onProgress
  );
  return { sheet, result };
}

/* ================================================================ import */

export type ImportTarget = "newSheet" | "activeCell";

export async function runImport(fileName: string, p: ParsedFile, target: ImportTarget, onProgress?: Progress): Promise<{ sheet: string; address: string }> {
  const base = fileName.replace(/\.[^.]+$/, "").slice(0, 28) || "Import";
  if (target === "newSheet") return writeTableToNewSheet({ sheetBase: base, headers: p.headers, rows: p.rows, dateCols: p.dateCols }, onProgress);
  // At the active cell: refuse to overwrite anything.
  const width = p.headers.length;
  const data = [p.headers, ...p.rows].map((r) => Array.from({ length: width }, (_, j) => sanitize(r[j])));
  const out = await run(async (ctx) => {
    const cell = ctx.workbook.getActiveCell();
    cell.load("rowIndex,columnIndex");
    const ws = cell.worksheet;
    ws.load("name");
    await ctx.sync();
    if (cell.rowIndex + data.length > 1_048_576) throw new ExcelUserError("Not enough rows below the active cell – import to a new sheet instead.");
    const area = ws.getRangeByIndexes(cell.rowIndex, cell.columnIndex, data.length, width);
    const existing = await loadChunked(ctx, ws, cell.rowIndex, cell.columnIndex, data.length, width, ["formulas"]);
    if (existing.formulas.some((r) => r.some((v) => v !== "" && v !== null)))
      throw new ExcelUserError(`The ${data.length.toLocaleString()} × ${width} cells from ${a1(cell.rowIndex, cell.columnIndex)} aren't empty – pick an empty spot or import to a new sheet.`);
    await writeChunked(ctx, ws, cell.rowIndex, cell.columnIndex, data, "values", (d, t) => onProgress?.(d / t));
    ws.getRangeByIndexes(cell.rowIndex, cell.columnIndex, 1, width).format.font.bold = true;
    for (const j of p.dateCols) await writeChunked(ctx, ws, cell.rowIndex + 1, cell.columnIndex + j, p.rows.map(() => ["dd-mmm-yyyy hh:mm"]), "numberFormat");
    area.load("address");
    await ctx.sync();
    // Undo = put the (empty) cells back. Kept for imports up to 200k cells.
    if (data.length * width <= 200_000) setSnapshot({ kind: "cells", sheetName: ws.name, address: area.address, formulas: existing.formulas, label: `import ${fileName}` });
    return { sheet: ws.name, address: area.address };
  });
  invalidateOverview();
  return out;
}
