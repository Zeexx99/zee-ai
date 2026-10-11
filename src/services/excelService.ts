/**
 * Office.js helpers for Zee AI.
 *
 * Principles:
 *  - Every operation runs inside a single Excel.run batch with as few context.sync() calls as possible.
 *  - Selections are clipped to the sheet's used range, so selecting whole columns (A:A) is safe.
 *  - Reads are capped (maxCells) and large writes are chunked to stay under Office payload limits.
 *  - Before every write we keep a snapshot so the user can "Undo last Zee AI change"
 *    (Office.js changes are not on Excel's own undo stack).
 */
import { a1, a1Range, columnLetter, quoteSheet, stripSheet, EXCEL_MAX_ROWS } from "../utils/a1";
import { cleanCell, type CleanOptions } from "../utils/clean";
import { coerceCell, type Cell } from "../utils/markdown";
import { convertValue, type UnitConversion } from "../utils/petro";

/* ------------------------------------------------------------------ errors */

export class ExcelUserError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "ExcelUserError";
  }
}

/** Thrown when a write would replace existing data; retry with overwrite=true after the user confirms. */
export class OverwriteRequiredError extends Error {
  readonly address: string;
  constructor(address: string) {
    super(`${stripSheet(address)} already contains data.`);
    this.name = "OverwriteRequiredError";
    this.address = address;
  }
}

export function isExcelAvailable(): boolean {
  try {
    return typeof Office !== "undefined" && typeof Excel !== "undefined" && Office.context?.host === Office.HostType.Excel;
  } catch {
    return false;
  }
}

function toUserError(e: unknown): Error {
  if (e instanceof ExcelUserError || e instanceof OverwriteRequiredError) return e;
  const code = (e as { code?: string })?.code;
  const msg = (e as Error)?.message ?? String(e);
  switch (code) {
    case "InvalidSelection":
      return new ExcelUserError("Select a single block of cells (multi-area selections aren't supported).");
    case "AccessDenied":
    case "InvalidOperation":
      return new ExcelUserError(`Excel refused the change – the sheet may be protected or a cell is being edited. (${msg})`);
    case "RequestPayloadSizeLimitExceeded":
    case "ResponsePayloadSizeLimitExceeded":
      return new ExcelUserError("The selection is too large for one operation. Select fewer cells.");
    case "ItemNotFound":
      return new ExcelUserError("The sheet or range no longer exists.");
    case "InsertDeleteConflict":
      return new ExcelUserError("Can't insert a column here because it would split a merged cell, table or array.");
    default:
      return new ExcelUserError(msg);
  }
}

async function run<T>(fn: (ctx: Excel.RequestContext) => Promise<T>): Promise<T> {
  if (!isExcelAvailable()) throw new ExcelUserError("Open Zee AI inside Excel to work with your sheet.");
  try {
    return await Excel.run(fn);
  } catch (e) {
    throw toUserError(e);
  }
}

/* ---------------------------------------------------------------- snapshot / undo (multi-step) */

export type Snapshot =
  | { kind: "cells"; sheetName: string; address: string; formulas: unknown[][]; numberFormat?: unknown[][]; label: string }
  | { kind: "insertedColumn"; sheetName: string; columnAddress: string; label: string }
  | { kind: "addedSheet"; sheetName: string; label: string }
  | { kind: "addedChart"; sheetName: string; chartName: string; label: string }
  | { kind: "addedTable"; tableName: string; label: string };

const UNDO_LIMIT = 25;
const undoStack: Snapshot[] = [];
const snapshotListeners = new Set<(label: string | null, depth: number) => void>();

function notifyUndo() {
  const top = undoStack[undoStack.length - 1];
  snapshotListeners.forEach((l) => l(top?.label ?? null, undoStack.length));
}

function setSnapshot(s: Snapshot) {
  undoStack.push(s);
  if (undoStack.length > UNDO_LIMIT) undoStack.shift();
  notifyUndo();
}

export function onUndoAvailabilityChange(listener: (label: string | null, depth: number) => void): () => void {
  snapshotListeners.add(listener);
  listener(undoStack[undoStack.length - 1]?.label ?? null, undoStack.length);
  return () => {
    snapshotListeners.delete(listener);
  };
}

async function applyUndo(ctx: Excel.RequestContext, snap: Snapshot) {
  switch (snap.kind) {
    case "cells": {
      const r = ctx.workbook.worksheets.getItem(snap.sheetName).getRange(stripSheet(snap.address));
      if (snap.numberFormat) r.numberFormat = snap.numberFormat;
      r.formulas = snap.formulas;
      break;
    }
    case "insertedColumn":
      ctx.workbook.worksheets.getItem(snap.sheetName).getRange(snap.columnAddress).delete(Excel.DeleteShiftDirection.left);
      break;
    case "addedSheet":
      ctx.workbook.worksheets.getItem(snap.sheetName).delete();
      break;
    case "addedChart":
      ctx.workbook.worksheets.getItem(snap.sheetName).charts.getItem(snap.chartName).delete();
      break;
    case "addedTable":
      // Keep the data, remove only the table object.
      ctx.workbook.tables.getItem(snap.tableName).convertToRange();
      break;
  }
  await ctx.sync();
}

/** Undo the most recent Zee AI change (call repeatedly for multiple steps). */
export async function undoLastChange(): Promise<string> {
  const snap = undoStack[undoStack.length - 1];
  if (!snap) throw new ExcelUserError("Nothing to undo.");
  try {
    await run((ctx) => applyUndo(ctx, snap));
  } finally {
    undoStack.pop();
    notifyUndo();
  }
  return snap.label;
}

/* ---------------------------------------------------------------- reading */

interface Block {
  rowIndex: number;
  columnIndex: number;
  rowCount: number;
  columnCount: number;
  address: string;
}

interface ClippedSelection {
  sheet: Excel.Worksheet;
  sheetName: string;
  selectionAddress: string;
  block: Block;
  isEmpty: boolean;
}

function blockOf(r: Excel.Range): Block {
  return { rowIndex: r.rowIndex, columnIndex: r.columnIndex, rowCount: r.rowCount, columnCount: r.columnCount, address: r.address };
}

/** Selected range intersected with the used range (handles A:A or Ctrl+A selections safely). */
async function loadClippedSelection(ctx: Excel.RequestContext): Promise<ClippedSelection> {
  const sel = ctx.workbook.getSelectedRange();
  sel.load("address,rowIndex,columnIndex,rowCount,columnCount");
  const sheet = sel.worksheet;
  sheet.load("name");
  const used = sheet.getUsedRangeOrNullObject(true);
  used.load("address");
  await ctx.sync();

  const base = { sheet, sheetName: sheet.name, selectionAddress: sel.address };
  if (used.isNullObject) return { ...base, block: blockOf(sel), isEmpty: true };

  const inter = sel.getIntersectionOrNullObject(used);
  inter.load("address,rowIndex,columnIndex,rowCount,columnCount");
  await ctx.sync();
  if (inter.isNullObject) return { ...base, block: blockOf(sel), isEmpty: true };
  return { ...base, block: blockOf(inter), isEmpty: false };
}

export interface SelectionContext {
  sheetName: string;
  address: string; // clipped address, e.g. Sheet1!A1:D20
  selectionAddress: string; // what the user actually selected
  rowIndex: number;
  columnIndex: number;
  rowCount: number;
  columnCount: number;
  values: unknown[][];
  text: string[][];
  formulas: unknown[][];
  truncated: boolean;
  isEmpty: boolean;
}

export async function getSelectionContext(maxCells: number): Promise<SelectionContext> {
  const cap = Math.max(50, Math.floor(maxCells));
  return run(async (ctx) => {
    const c = await loadClippedSelection(ctx);
    const b = c.block;
    const result: SelectionContext = {
      sheetName: c.sheetName,
      address: b.address,
      selectionAddress: c.selectionAddress,
      rowIndex: b.rowIndex,
      columnIndex: b.columnIndex,
      rowCount: b.rowCount,
      columnCount: b.columnCount,
      values: [],
      text: [],
      formulas: [],
      truncated: false,
      isEmpty: c.isEmpty,
    };
    if (c.isEmpty) return result;

    let readRows = b.rowCount;
    let readCols = b.columnCount;
    if (readRows * readCols > cap) {
      readCols = Math.min(b.columnCount, 50);
      readRows = Math.max(1, Math.min(b.rowCount, Math.floor(cap / readCols)));
      result.truncated = true;
    }
    const r = c.sheet.getRangeByIndexes(b.rowIndex, b.columnIndex, readRows, readCols);
    r.load("values,text,formulas");
    await ctx.sync();
    result.values = r.values;
    result.text = r.text;
    result.formulas = r.formulas;
    return result;
  });
}

function cellForPrompt(value: unknown, text: string | undefined, formula: unknown): string {
  let shown = text !== undefined && text !== "" && !/^#+$/.test(text) ? text : value === null || value === undefined ? "" : String(value);
  if (typeof formula === "string" && formula.startsWith("=")) shown = `${formula} → ${shown}`;
  shown = shown.replace(/\|/g, "\\|").replace(/\r?\n/g, " ⏎ ");
  return shown.length > 200 ? `${shown.slice(0, 197)}…` : shown;
}

/** Render the selection as a compact grid with real row numbers and column letters. */
export function formatSelectionForPrompt(s: SelectionContext): string {
  if (s.isEmpty) return `[Excel selection]\nSheet "${s.sheetName}", ${stripSheet(s.selectionAddress)} – the selected cells are empty.`;
  const rows = s.values.length;
  const cols = rows ? s.values[0].length : 0;
  const header = ["#", ...Array.from({ length: cols }, (_, j) => columnLetter(s.columnIndex + j))];
  const lines = [`| ${header.join(" | ")} |`, `|${header.map(() => "---").join("|")}|`];
  for (let i = 0; i < rows; i++) {
    const cells = Array.from({ length: cols }, (_, j) => cellForPrompt(s.values[i][j], s.text[i]?.[j], s.formulas[i]?.[j]));
    lines.push(`| ${s.rowIndex + i + 1} | ${cells.join(" | ")} |`);
  }
  const note = s.truncated
    ? `\n(Truncated: showing the first ${rows} rows × ${cols} columns of ${s.rowCount} × ${s.columnCount}. Increase "Max cells sent" in Settings to send more.)`
    : "";
  return `[Excel selection]\nSheet "${s.sheetName}", range ${stripSheet(s.address)} (${s.rowCount} rows × ${s.columnCount} columns). Formulas are shown as =FORMULA → displayed value.\n${lines.join("\n")}${note}`;
}

/** Light-weight address of the current selection (for the context chip). */
export async function getSelectedAddress(): Promise<string> {
  return run(async (ctx) => {
    const r = ctx.workbook.getSelectedRange();
    r.load("address");
    await ctx.sync();
    return r.address;
  });
}

export interface ActiveCellInfo {
  address: string;
  formula: string;
  value: unknown;
  hasFormula: boolean;
}

export async function getActiveCellInfo(): Promise<ActiveCellInfo> {
  return run(async (ctx) => {
    const cell = ctx.workbook.getActiveCell();
    cell.load("address,formulas,values");
    await ctx.sync();
    const f = cell.formulas[0][0];
    const formula = typeof f === "string" ? f : String(f ?? "");
    return { address: cell.address, formula, value: cell.values[0][0], hasFormula: formula.startsWith("=") };
  });
}

/** Header row + a few sample rows of the used range – gives the model column meanings. */
export async function getSheetOverview(sampleRows = 5, maxCols = 40): Promise<string> {
  return run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getActiveWorksheet();
    sheet.load("name");
    const used = sheet.getUsedRangeOrNullObject(true);
    used.load("address,rowIndex,columnIndex,rowCount,columnCount");
    await ctx.sync();
    if (used.isNullObject) return `[Sheet overview]\nSheet "${sheet.name}" is empty.`;
    const rows = Math.min(used.rowCount, sampleRows + 1);
    const cols = Math.min(used.columnCount, maxCols);
    const sample = sheet.getRangeByIndexes(used.rowIndex, used.columnIndex, rows, cols);
    sample.load("text");
    await ctx.sync();
    const letters = Array.from({ length: cols }, (_, j) => columnLetter(used.columnIndex + j));
    const lines = [`| # | ${letters.join(" | ")} |`, `|${["#", ...letters].map(() => "---").join("|")}|`];
    sample.text.forEach((r, i) => lines.push(`| ${used.rowIndex + i + 1} | ${r.map((t) => String(t).replace(/\|/g, "\\|").slice(0, 60)).join(" | ")} |`));
    return `[Sheet overview]\nSheet "${sheet.name}", used range ${stripSheet(used.address)} (${used.rowCount} rows × ${used.columnCount} columns). First rows:\n${lines.join("\n")}`;
  });
}

/* ---------------------------------------------------------------- writing */

function hasContent(formulas: unknown[][]): boolean {
  return formulas.some((r) => r.some((v) => v !== "" && v !== null && v !== undefined));
}

export interface WriteOptions {
  overwrite?: boolean;
  /** Bold the first row (table header). */
  header?: boolean;
  label?: string;
}

/** Write a 2D block (values and/or formulas) with its top-left corner at the active cell. */
export async function writeMatrixAtActiveCell(matrix: Cell[][], opts: WriteOptions = {}): Promise<string> {
  if (!matrix.length || !matrix[0].length) throw new ExcelUserError("Nothing to insert.");
  const width = Math.max(...matrix.map((r) => r.length));
  const data = matrix.map((r) => [...r, ...Array<Cell>(width - r.length).fill("")]);
  return run(async (ctx) => {
    const cell = ctx.workbook.getActiveCell();
    cell.load("rowIndex,columnIndex");
    const sheet = cell.worksheet;
    sheet.load("name");
    await ctx.sync();
    if (cell.rowIndex + data.length > EXCEL_MAX_ROWS) throw new ExcelUserError("Not enough rows below the active cell.");

    const target = sheet.getRangeByIndexes(cell.rowIndex, cell.columnIndex, data.length, width);
    target.load("address,formulas,numberFormat");
    await ctx.sync();
    if (!opts.overwrite && hasContent(target.formulas)) throw new OverwriteRequiredError(target.address);

    setSnapshot({
      kind: "cells",
      sheetName: sheet.name,
      address: target.address,
      formulas: target.formulas,
      numberFormat: target.numberFormat,
      label: opts.label ?? `insert at ${stripSheet(target.address)}`,
    });
    target.formulas = data;
    if (opts.header && data.length > 1) target.getRow(0).format.font.bold = true;
    if (width > 1 || data.length > 1) target.format.autofitColumns();
    target.select();
    await ctx.sync();
    return target.address;
  });
}

/** Insert one formula (or a vertical list of formulas) at the active cell. */
export function insertFormulas(formulas: string[], opts: WriteOptions = {}): Promise<string> {
  return writeMatrixAtActiveCell(
    formulas.map((f) => [f]),
    { ...opts, label: opts.label ?? "insert formula" }
  );
}

/* ---------------------------------------------------------------- quick math */

export type AggregateFn = "SUM" | "AVERAGE" | "COUNT" | "COUNTA" | "UNIQUE" | "MIN" | "MAX";

export const AGGREGATE_LABELS: Record<AggregateFn, string> = {
  SUM: "Sum",
  AVERAGE: "Average",
  COUNT: "Count numbers",
  COUNTA: "Count non-empty",
  UNIQUE: "Count unique",
  MIN: "Min",
  MAX: "Max",
};

function aggregateFormula(fn: AggregateFn, ref: string, horizontal: boolean): string {
  if (fn === "UNIQUE") return `=COUNTA(UNIQUE(FILTER(${ref},${ref}<>"")${horizontal ? ",TRUE" : ""}))`;
  return `=${fn}(${ref})`;
}

export interface AggregateResult {
  address: string;
  formula: string;
  value: unknown;
}

/**
 * AutoSum-style insert: a formula below each selected column (or to the right of a single row).
 */
export async function insertAggregate(fn: AggregateFn, overwrite = false): Promise<AggregateResult[]> {
  return run(async (ctx) => {
    const c = await loadClippedSelection(ctx);
    if (c.isEmpty) throw new ExcelUserError("Select the cells you want to calculate first.");
    const b = c.block;
    if (b.rowCount * b.columnCount < 2) throw new ExcelUserError("Select at least two cells (a column, row or block).");

    const horizontal = b.rowCount === 1;
    let target: Excel.Range;
    let formulas: string[][];
    if (horizontal) {
      target = c.sheet.getRangeByIndexes(b.rowIndex, b.columnIndex + b.columnCount, 1, 1);
      formulas = [[aggregateFormula(fn, a1Range(b.rowIndex, b.columnIndex, 1, b.columnCount), true)]];
    } else {
      if (b.rowIndex + b.rowCount >= EXCEL_MAX_ROWS) throw new ExcelUserError("There's no empty row below the selection.");
      const cols = Math.min(b.columnCount, 200);
      target = c.sheet.getRangeByIndexes(b.rowIndex + b.rowCount, b.columnIndex, 1, cols);
      formulas = [Array.from({ length: cols }, (_, j) => aggregateFormula(fn, a1Range(b.rowIndex, b.columnIndex + j, b.rowCount, 1), false))];
    }
    target.load("address,formulas,numberFormat");
    await ctx.sync();
    if (!overwrite && hasContent(target.formulas)) throw new OverwriteRequiredError(target.address);

    setSnapshot({
      kind: "cells",
      sheetName: c.sheetName,
      address: target.address,
      formulas: target.formulas,
      numberFormat: target.numberFormat,
      label: `${AGGREGATE_LABELS[fn]} in ${stripSheet(target.address)}`,
    });
    target.formulas = formulas;
    target.format.font.bold = true;
    target.load("values");
    await ctx.sync();

    const firstRow = target.values[0];
    const anchorRow = horizontal ? b.rowIndex : b.rowIndex + b.rowCount;
    const anchorCol = horizontal ? b.columnIndex + b.columnCount : b.columnIndex;
    return firstRow.map((v, j) => ({ address: a1(anchorRow, anchorCol + j), formula: formulas[0][j], value: v }));
  });
}

/* ---------------------------------------------------------------- clean */

export interface CleanSummary {
  address: string;
  cellsChanged: number;
  cellsScanned: number;
  undoAvailable: boolean;
}

const CLEAN_MAX_CELLS = 200_000;
const UNDO_MAX_CELLS = 50_000;

export async function cleanSelection(options: CleanOptions): Promise<CleanSummary> {
  return run(async (ctx) => {
    const c = await loadClippedSelection(ctx);
    if (c.isEmpty) throw new ExcelUserError("The selected cells are empty.");
    const b = c.block;
    const total = b.rowCount * b.columnCount;
    if (total > CLEAN_MAX_CELLS) throw new ExcelUserError(`Select fewer than ${CLEAN_MAX_CELLS.toLocaleString()} cells to clean at once.`);

    const keepUndo = total <= UNDO_MAX_CELLS;
    const allFormulas: unknown[][] = [];
    const allFormats: unknown[][] = [];
    const chunkRows = Math.max(1, Math.floor(20_000 / b.columnCount));
    let changed = 0;

    for (let start = 0; start < b.rowCount; start += chunkRows) {
      const rows = Math.min(chunkRows, b.rowCount - start);
      const r = c.sheet.getRangeByIndexes(b.rowIndex + start, b.columnIndex, rows, b.columnCount);
      r.load("formulas,numberFormat");
      await ctx.sync();
      if (keepUndo) {
        allFormulas.push(...r.formulas);
        allFormats.push(...r.numberFormat);
      }
      let chunkChanged = 0;
      const newFormats = r.numberFormat.map((row) => row.slice());
      const newFormulas = r.formulas.map((row, i) =>
        row.map((v, j) => {
          const res = cleanCell(v, options);
          if (res.changed) {
            chunkChanged++;
            if (res.numberFormat) newFormats[i][j] = res.numberFormat;
          }
          return res.value;
        })
      );
      if (chunkChanged > 0) {
        r.numberFormat = newFormats;
        r.formulas = newFormulas;
        changed += chunkChanged;
      }
    }
    if (changed > 0) {
      await ctx.sync();
      if (keepUndo)
        setSnapshot({
          kind: "cells",
          sheetName: c.sheetName,
          address: b.address,
          formulas: allFormulas,
          numberFormat: allFormats,
          label: `clean ${stripSheet(b.address)}`,
        });
    }
    return { address: b.address, cellsChanged: changed, cellsScanned: total, undoAvailable: keepUndo && changed > 0 };
  });
}

/* ---------------------------------------------------------------- split / extract */

export interface ColumnSource {
  sheetName: string;
  rowIndex: number;
  columnIndex: number;
  values: string[];
  /** Text of the first row when treated as a header. */
  headerText?: string;
}

/** Read a single selected column (clipped to used range) for extraction. */
export async function readColumnSelection(firstRowIsHeader: boolean, maxRows = 5000): Promise<ColumnSource> {
  return run(async (ctx) => {
    const c = await loadClippedSelection(ctx);
    if (c.isEmpty) throw new ExcelUserError("The selected cells are empty.");
    const b = c.block;
    if (b.columnCount !== 1) throw new ExcelUserError("Select a single column of values to split or extract from.");
    if (b.rowCount > maxRows) throw new ExcelUserError(`Select at most ${maxRows.toLocaleString()} rows at a time.`);
    const r = c.sheet.getRangeByIndexes(b.rowIndex, b.columnIndex, b.rowCount, 1);
    r.load("text");
    await ctx.sync();
    const all = r.text.map((row) => String(row[0] ?? ""));
    if (firstRowIsHeader && all.length > 1) {
      return { sheetName: c.sheetName, rowIndex: b.rowIndex + 1, columnIndex: b.columnIndex, values: all.slice(1), headerText: all[0] };
    }
    return { sheetName: c.sheetName, rowIndex: b.rowIndex, columnIndex: b.columnIndex, values: all };
  });
}

/** Insert a new column right of the source column and write results there (as text). */
export async function insertResultColumn(src: ColumnSource, results: string[], label: string): Promise<string> {
  return run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getItem(src.sheetName);
    const newCol = src.columnIndex + 1;
    sheet.getRangeByIndexes(0, newCol, 1, 1).getEntireColumn().insert(Excel.InsertShiftDirection.right);

    if (src.headerText !== undefined && src.rowIndex > 0) {
      const h = sheet.getRangeByIndexes(src.rowIndex - 1, newCol, 1, 1);
      h.values = [[`${src.headerText || "Value"} – ${label}`]];
      h.format.font.bold = true;
    }
    const out = sheet.getRangeByIndexes(src.rowIndex, newCol, results.length, 1);
    out.numberFormat = results.map(() => ["@"]); // keep as text: "+1 555…" must not become a formula
    out.values = results.map((v) => [v]);
    sheet.getRangeByIndexes(0, newCol, 1, 1).getEntireColumn().format.autofitColumns();
    await ctx.sync();

    const letter = columnLetter(newCol);
    setSnapshot({ kind: "insertedColumn", sheetName: src.sheetName, columnAddress: `${letter}:${letter}`, label: `${label} column ${letter}` });
    return `${src.sheetName}!${letter}${src.rowIndex + 1}`;
  });
}

/* ---------------------------------------------------------------- formatting */

export interface FormatOptions {
  bold?: boolean;
  fillColor?: string;
  fontColor?: string;
  numberFormat?: string;
  autofit?: boolean;
  borders?: boolean;
}

/** Apply simple formatting to the current selection. */
export async function formatSelection(o: FormatOptions): Promise<string> {
  return run(async (ctx) => {
    const r = ctx.workbook.getSelectedRange();
    r.load("address,rowCount,columnCount");
    await ctx.sync();
    if (o.bold !== undefined) r.format.font.bold = o.bold;
    if (o.fillColor) r.format.fill.color = o.fillColor;
    if (o.fontColor) r.format.font.color = o.fontColor;
    if (o.numberFormat) r.numberFormat = Array.from({ length: r.rowCount }, () => Array<string>(r.columnCount).fill(o.numberFormat as string));
    if (o.borders) {
      ([
        Excel.BorderIndex.edgeTop,
        Excel.BorderIndex.edgeBottom,
        Excel.BorderIndex.edgeLeft,
        Excel.BorderIndex.edgeRight,
        Excel.BorderIndex.insideHorizontal,
        Excel.BorderIndex.insideVertical,
      ] as const).forEach((i) => {
        r.format.borders.getItem(i).style = Excel.BorderLineStyle.continuous;
      });
    }
    if (o.autofit) r.format.autofitColumns();
    await ctx.sync();
    return r.address;
  });
}

/** Convert the selection (clipped to data) into an Excel Table with headers. */
export async function formatSelectionAsTable(): Promise<string> {
  return run(async (ctx) => {
    const c = await loadClippedSelection(ctx);
    if (c.isEmpty) throw new ExcelUserError("Select the data you want to turn into a table.");
    const table = c.sheet.tables.add(c.block.address, true);
    table.style = "TableStyleMedium2";
    table.getRange().format.autofitColumns();
    table.load("name");
    await ctx.sync();
    setSnapshot({ kind: "addedTable", tableName: table.name, label: `table ${table.name}` });
    return table.name;
  });
}

export async function autofitSelection(): Promise<string> {
  return formatSelection({ autofit: true });
}

/* ================================================================ agent / workbook-level helpers */

/** Resolve "Sheet1!A1:B3", "'My Sheet'!A1" or "A1:B3" (+ optional sheet name) to a Range. */
function resolveRange(ctx: Excel.RequestContext, address: string, sheet?: string): Excel.Range {
  const addr = address.trim();
  const bang = addr.lastIndexOf("!");
  if (bang > 0) {
    const sheetName = addr.slice(0, bang).replace(/^'(.*)'$/, "$1").replace(/''/g, "'");
    return ctx.workbook.worksheets.getItem(sheetName).getRange(addr.slice(bang + 1));
  }
  const ws = sheet ? ctx.workbook.worksheets.getItem(sheet) : ctx.workbook.worksheets.getActiveWorksheet();
  return ws.getRange(addr);
}

let overviewCache: { at: number; text: string } | null = null;

/** Compact map of the whole workbook: sheets, used ranges, header rows, tables, named ranges. */
export async function getWorkbookOverview(maxChars = 3500, maxAgeMs = 20_000): Promise<string> {
  if (overviewCache && Date.now() - overviewCache.at < maxAgeMs) return overviewCache.text;
  const text = await run(async (ctx) => {
    const sheets = ctx.workbook.worksheets;
    sheets.load("items/name,items/visibility");
    const active = sheets.getActiveWorksheet();
    active.load("name");
    const tables = ctx.workbook.tables;
    tables.load("items/name");
    const names = ctx.workbook.names;
    names.load("items/name,items/formula");
    await ctx.sync();

    const list = sheets.items.slice(0, 30);
    const used = list.map((ws) => {
      const u = ws.getUsedRangeOrNullObject(true);
      u.load("address,rowCount,columnCount,rowIndex,columnIndex");
      return u;
    });
    const tableRanges = tables.items.slice(0, 30).map((t) => {
      const r = t.getRange();
      r.load("address");
      return r;
    });
    await ctx.sync();

    const headers = used.map((u, i) => {
      if (u.isNullObject) return null;
      const h = list[i].getRangeByIndexes(u.rowIndex, u.columnIndex, 1, Math.min(u.columnCount, 25));
      h.load("text");
      return h;
    });
    await ctx.sync();

    const lines: string[] = [`[Workbook overview] Active sheet: "${active.name}"`];
    list.forEach((ws, i) => {
      const u = used[i];
      const hidden = ws.visibility !== Excel.SheetVisibility.visible ? " (hidden)" : "";
      if (u.isNullObject) {
        lines.push(`- Sheet "${ws.name}"${hidden}: empty`);
        return;
      }
      const hdr = headers[i]?.text?.[0]?.map((t) => String(t).slice(0, 25)).join(" | ") ?? "";
      lines.push(`- Sheet "${ws.name}"${hidden}: ${stripSheet(u.address)} (${u.rowCount}×${u.columnCount}); first row: ${hdr}`);
    });
    if (tables.items.length) lines.push(`Tables: ${tables.items.slice(0, 30).map((t, i) => `${t.name} = ${tableRanges[i]?.address ?? "?"}`).join("; ")}`);
    if (names.items.length) lines.push(`Named ranges: ${names.items.slice(0, 30).map((n) => `${n.name} → ${n.formula}`).join("; ")}`);
    const out = lines.join("\n");
    return out.length > maxChars ? `${out.slice(0, maxChars)}\n…(overview truncated)` : out;
  });
  overviewCache = { at: Date.now(), text };
  return text;
}

export function invalidateOverview() {
  overviewCache = null;
}

/** Read any range by address (for the agent). Capped by maxCells. */
export async function readRangeText(address: string, sheet?: string, maxCells = 2000): Promise<string> {
  return run(async (ctx) => {
    const target = resolveRange(ctx, address, sheet);
    target.load("address,rowIndex,columnIndex,rowCount,columnCount");
    const ws = target.worksheet;
    ws.load("name");
    await ctx.sync();
    let rows = target.rowCount;
    let cols = target.columnCount;
    let truncated = false;
    if (rows * cols > maxCells) {
      cols = Math.min(cols, 50);
      rows = Math.max(1, Math.min(rows, Math.floor(maxCells / cols)));
      truncated = true;
    }
    // Clip whole-column / whole-sheet references to the used range.
    const used = ws.getUsedRangeOrNullObject(true);
    used.load("rowIndex,rowCount");
    await ctx.sync();
    let startRow = target.rowIndex;
    if (!used.isNullObject && target.rowCount > 5000) {
      startRow = Math.max(target.rowIndex, used.rowIndex);
      rows = Math.min(rows, used.rowIndex + used.rowCount - startRow);
    }
    if (rows <= 0) return `${target.address} is empty.`;
    const r = ws.getRangeByIndexes(startRow, target.columnIndex, rows, cols);
    r.load("values,text,formulas,address");
    await ctx.sync();
    const sc: SelectionContext = {
      sheetName: ws.name,
      address: r.address,
      selectionAddress: target.address,
      rowIndex: startRow,
      columnIndex: target.columnIndex,
      rowCount: truncated ? target.rowCount : rows,
      columnCount: truncated ? target.columnCount : cols,
      values: r.values,
      text: r.text,
      formulas: r.formulas,
      truncated,
      isEmpty: false,
    };
    return formatSelectionForPrompt(sc).replace("[Excel selection]", "[Range]");
  });
}

/** Write a block of values/formulas starting at `address` (top-left). Strings starting with "=" are formulas. */
export async function setCells(address: string, rows: unknown[][], sheet?: string, label?: string): Promise<string> {
  if (!rows.length) throw new ExcelUserError("No values to write.");
  const width = Math.max(...rows.map((r) => r.length));
  const data = rows.map((r) => {
    const out = r.map((v) => (typeof v === "string" ? coerceCell(v) : (v as Cell) ?? ""));
    while (out.length < width) out.push("");
    return out;
  });
  return run(async (ctx) => {
    const start = resolveRange(ctx, address, sheet).getCell(0, 0);
    start.load("rowIndex,columnIndex");
    const ws = start.worksheet;
    ws.load("name");
    await ctx.sync();
    const target = ws.getRangeByIndexes(start.rowIndex, start.columnIndex, data.length, width);
    target.load("address,formulas,numberFormat");
    await ctx.sync();
    setSnapshot({ kind: "cells", sheetName: ws.name, address: target.address, formulas: target.formulas, numberFormat: target.numberFormat, label: label ?? `write ${stripSheet(target.address)}` });
    target.formulas = data;
    await ctx.sync();
    invalidateOverview();
    return target.address;
  });
}

/**
 * For the agent approval card: would a set_cells call replace existing data?
 * Returns the address of the target block if it already holds something, else null.
 */
export async function writeTargetHasData(args: Record<string, unknown>): Promise<string | null> {
  const rows = Array.isArray(args.values) ? (args.values as unknown[]).map((r) => (Array.isArray(r) ? r : [r])) : [];
  if (!rows.length || typeof args.address !== "string") return null;
  const width = Math.max(1, ...rows.map((r) => r.length));
  return run(async (ctx) => {
    const start = resolveRange(ctx, args.address as string, typeof args.sheet === "string" ? args.sheet : undefined).getCell(0, 0);
    start.load("rowIndex,columnIndex");
    const ws = start.worksheet;
    await ctx.sync();
    const target = ws.getRangeByIndexes(start.rowIndex, start.columnIndex, rows.length, width);
    target.load("address,formulas");
    await ctx.sync();
    return hasContent(target.formulas) ? target.address : null;
  });
}

export interface RangeFormat extends FormatOptions {
  wrap?: boolean;
}

/** Format any range (for the agent). Formatting is not tracked by undo. */
export async function formatRangeAt(address: string, o: RangeFormat, sheet?: string): Promise<string> {
  return run(async (ctx) => {
    const r = resolveRange(ctx, address, sheet);
    r.load("address,rowCount,columnCount");
    await ctx.sync();
    if (o.bold !== undefined) r.format.font.bold = o.bold;
    if (o.fillColor) r.format.fill.color = o.fillColor;
    if (o.fontColor) r.format.font.color = o.fontColor;
    if (o.wrap !== undefined) r.format.wrapText = o.wrap;
    if (o.numberFormat) r.numberFormat = Array.from({ length: r.rowCount }, () => Array<string>(r.columnCount).fill(o.numberFormat as string));
    if (o.borders) {
      ([
        Excel.BorderIndex.edgeTop,
        Excel.BorderIndex.edgeBottom,
        Excel.BorderIndex.edgeLeft,
        Excel.BorderIndex.edgeRight,
        Excel.BorderIndex.insideHorizontal,
        Excel.BorderIndex.insideVertical,
      ] as const).forEach((i) => {
        r.format.borders.getItem(i).style = Excel.BorderLineStyle.continuous;
      });
    }
    if (o.autofit) r.format.autofitColumns();
    await ctx.sync();
    return r.address;
  });
}

export type ChartKind = "line" | "column" | "bar" | "scatter" | "scatterLines" | "pie" | "area";
const CHART_TYPES: Record<ChartKind, string> = {
  line: "Line",
  column: "ColumnClustered",
  bar: "BarClustered",
  scatter: "XYScatter",
  scatterLines: "XYScatterLines",
  pie: "Pie",
  area: "Area",
};

/** Create a chart from a data range (header row = series names). */
export async function createChart(dataAddress: string, kind: ChartKind, title?: string, sheet?: string): Promise<string> {
  return run(async (ctx) => {
    const data = resolveRange(ctx, dataAddress, sheet);
    data.load("address,rowIndex,columnIndex,columnCount");
    const ws = data.worksheet;
    ws.load("name");
    await ctx.sync();
    const chart = ws.charts.add(CHART_TYPES[kind] as Excel.ChartType, data, Excel.ChartSeriesBy.auto);
    if (title) {
      chart.title.text = title;
      chart.title.visible = true;
    }
    const anchor = ws.getRangeByIndexes(data.rowIndex, data.columnIndex + data.columnCount + 1, 1, 1);
    chart.setPosition(anchor);
    chart.width = 480;
    chart.height = 300;
    chart.load("name");
    await ctx.sync();
    setSnapshot({ kind: "addedChart", sheetName: ws.name, chartName: chart.name, label: `chart "${title ?? chart.name}"` });
    return `${chart.name} on ${ws.name}`;
  });
}

/** Turn a range into an Excel table. */
export async function createTableAt(address: string, hasHeaders: boolean, sheet?: string): Promise<string> {
  return run(async (ctx) => {
    const r = resolveRange(ctx, address, sheet);
    r.load("address");
    await ctx.sync();
    const table = ctx.workbook.tables.add(r.address, hasHeaders);
    table.style = "TableStyleMedium2";
    table.load("name");
    await ctx.sync();
    setSnapshot({ kind: "addedTable", tableName: table.name, label: `table ${table.name}` });
    invalidateOverview();
    return table.name;
  });
}

/** Add a worksheet with a unique name; returns the final name. */
export async function addSheet(baseName: string, activate = true): Promise<string> {
  return run(async (ctx) => {
    const sheets = ctx.workbook.worksheets;
    sheets.load("items/name");
    await ctx.sync();
    const existing = new Set(sheets.items.map((w) => w.name.toLowerCase()));
    const clean = baseName.replace(/[\\/?*[\]:]/g, " ").trim().slice(0, 28) || "Sheet";
    let name = clean;
    for (let i = 2; existing.has(name.toLowerCase()); i++) name = `${clean} ${i}`;
    const ws = sheets.add(name);
    if (activate) ws.activate();
    await ctx.sync();
    setSnapshot({ kind: "addedSheet", sheetName: name, label: `sheet "${name}"` });
    invalidateOverview();
    return name;
  });
}

/** Write a matrix onto a sheet at a cell (no snapshot; used for freshly created report sheets). */
export async function writeBlock(sheetName: string, topLeft: string, rows: Cell[][], opts: { header?: boolean; numberFormat?: string } = {}): Promise<string> {
  if (!rows.length) return "";
  const width = Math.max(...rows.map((r) => r.length));
  const data = rows.map((r) => [...r, ...Array<Cell>(width - r.length).fill("")]);
  return run(async (ctx) => {
    const ws = ctx.workbook.worksheets.getItem(sheetName);
    const start = ws.getRange(topLeft);
    start.load("rowIndex,columnIndex");
    await ctx.sync();
    const target = ws.getRangeByIndexes(start.rowIndex, start.columnIndex, data.length, width);
    target.values = data;
    if (opts.header) {
      const h = target.getRow(0);
      h.format.font.bold = true;
      h.format.fill.color = "#EDEBFB";
    }
    target.format.autofitColumns();
    target.load("address");
    await ctx.sync();
    return target.address;
  });
}

export interface SelectionBlock {
  sheetName: string;
  address: string;
  rowIndex: number;
  columnIndex: number;
  values: unknown[][];
  text: string[][];
}

/** Read the (used-range-clipped) selection as raw values + text. */
export async function readSelectionBlock(maxCells = 200_000): Promise<SelectionBlock> {
  return run(async (ctx) => {
    const c = await loadClippedSelection(ctx);
    if (c.isEmpty) throw new ExcelUserError("The selected cells are empty.");
    const b = c.block;
    if (b.rowCount * b.columnCount > maxCells) throw new ExcelUserError(`Select fewer than ${maxCells.toLocaleString()} cells.`);
    const r = c.sheet.getRangeByIndexes(b.rowIndex, b.columnIndex, b.rowCount, b.columnCount);
    r.load("values,text");
    await ctx.sync();
    return { sheetName: c.sheetName, address: b.address, rowIndex: b.rowIndex, columnIndex: b.columnIndex, values: r.values, text: r.text };
  });
}

/** Convert units of the selected column(s) into new columns inserted to the right. */
export async function insertConvertedColumns(conv: UnitConversion, firstRowIsHeader: boolean): Promise<{ address: string; converted: number }> {
  const blk = await readSelectionBlock(100_000);
  const cols = blk.values[0]?.length ?? 0;
  if (cols > 20) throw new ExcelUserError("Select at most 20 columns to convert.");
  let converted = 0;
  const out: Cell[][] = blk.values.map((row, i) =>
    row.map((v) => {
      if (i === 0 && firstRowIsHeader) return `${String(v ?? "").trim() || "Value"} (${conv.suffix})`;
      const r = convertValue(v, conv);
      if (r !== "") converted++;
      return r;
    })
  );
  const address = await run(async (ctx) => {
    const ws = ctx.workbook.worksheets.getItem(blk.sheetName);
    const firstNew = blk.columnIndex + cols;
    const insertRange = ws.getRangeByIndexes(0, firstNew, 1, cols).getEntireColumn();
    insertRange.insert(Excel.InsertShiftDirection.right);
    const target = ws.getRangeByIndexes(blk.rowIndex, firstNew, out.length, cols);
    target.values = out;
    if (firstRowIsHeader) target.getRow(0).format.font.bold = true;
    target.format.autofitColumns();
    await ctx.sync();
    const colAddr = `${columnLetter(firstNew)}:${columnLetter(firstNew + cols - 1)}`;
    setSnapshot({ kind: "insertedColumn", sheetName: blk.sheetName, columnAddress: colAddr, label: `${conv.label} columns ${colAddr}` });
    return `${quoteSheet(blk.sheetName)}!${colAddr}`;
  });
  invalidateOverview();
  return { address, converted };
}

/** Raw values of a range (first column flattened if asked). Capped by maxCells. */
export async function readRangeValues(address: string, sheet?: string, maxCells = 50_000): Promise<unknown[][]> {
  return run(async (ctx) => {
    const target = resolveRange(ctx, address, sheet);
    const ws = target.worksheet;
    const used = ws.getUsedRangeOrNullObject(true);
    used.load("address");
    target.load("rowCount,columnCount");
    await ctx.sync();
    const r = used.isNullObject ? target : target.getIntersectionOrNullObject(used);
    r.load("rowCount,columnCount");
    await ctx.sync();
    if (r.isNullObject) return [];
    if (r.rowCount * r.columnCount > maxCells) throw new ExcelUserError(`Range is too large (max ${maxCells.toLocaleString()} cells).`);
    r.load("values");
    await ctx.sync();
    return r.values;
  });
}
