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
import { cleanCell, findEmptyRows, looksLikeHeaderRow, markFilledColumns, runsDescending, type CleanOptions } from "../utils/clean";
import { coerceCell, type Cell } from "../utils/markdown";
import { convertValue, type UnitConversion } from "../utils/petro";
import { chartTypeInfo, checkYCount, supportsLogScale, supportsSecondaryAxis } from "../utils/chartTypes";
import { canRevertAlone } from "../utils/undoHistory";

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

/** @internal Shared with workbookOps.ts. */
export async function run<T>(fn: (ctx: Excel.RequestContext) => Promise<T>): Promise<T> {
  if (!isExcelAvailable()) throw new ExcelUserError("Open Zee AI inside Excel to work with your sheet.");
  try {
    return await Excel.run(fn);
  } catch (e) {
    throw toUserError(e);
  }
}

/**
 * Auto-fit EVERY used column of the sheet that holds `r` (not just the cells written).
 * Zee AI calls this after every write so all columns always fit their content. Queued; the caller syncs.
 */
export function autofitAllOf(r: Excel.Range) {
  r.worksheet.getUsedRange(true).getEntireColumn().format.autofitColumns();
}

/* ---------------------------------------------------------------- snapshot / undo (multi-step) */

export type Snapshot =
  | { kind: "cells"; sheetName: string; address: string; formulas: unknown[][]; numberFormat?: unknown[][]; label: string }
  | { kind: "insertedColumn"; sheetName: string; columnAddress: string; label: string }
  | { kind: "addedSheet"; sheetName: string; label: string }
  | { kind: "addedChart"; sheetName: string; chartName: string; label: string }
  | { kind: "addedTable"; tableName: string; label: string }
  | { kind: "insertedRows"; sheetName: string; rowAddress: string; label: string }
  /** Whole rows/columns that were deleted: re-insert them and put their cell contents back. */
  | { kind: "deletedLines"; sheetName: string; lines: "rows" | "columns"; start: number; count: number; contentAddress: string | null; formulas: unknown[][] | null; label: string }
  /** Per-cell formatting (fill, font, borders, alignment, wrap) + number formats. */
  | { kind: "format"; sheetName: string; address: string; props: Excel.SettableCellProperties[][]; numberFormat: unknown[][]; label: string }
  | { kind: "addedConditionalFormat"; sheetName: string; address: string; id: string; label: string }
  | { kind: "dataValidation"; sheetName: string; address: string; rule: Excel.DataValidationRule | null; label: string }
  | { kind: "addedName"; name: string; label: string }
  | { kind: "autoFilter"; sheetName: string; label: string }
  | { kind: "freeze"; sheetName: string; previous: string | null; label: string }
  /** Several snapshots taken by one operation; undone in reverse order. */
  | { kind: "multi"; steps: Snapshot[]; label: string };

const UNDO_LIMIT = 25;
interface UndoEntry {
  snap: Snapshot;
  /** When the change was made (ms). */
  at: number;
}
const undoStack: UndoEntry[] = [];
const snapshotListeners = new Set<(label: string | null, depth: number) => void>();

function notifyUndo() {
  const top = undoStack[undoStack.length - 1];
  snapshotListeners.forEach((l) => l(top?.snap.label ?? null, undoStack.length));
}

/** @internal Shared with workbookOps.ts – push one undo step. */
export function setSnapshot(s: Snapshot) {
  undoStack.push({ snap: s, at: Date.now() });
  if (undoStack.length > UNDO_LIMIT) undoStack.shift();
  notifyUndo();
}

export function onUndoAvailabilityChange(listener: (label: string | null, depth: number) => void): () => void {
  snapshotListeners.add(listener);
  listener(undoStack[undoStack.length - 1]?.snap.label ?? null, undoStack.length);
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
    case "insertedRows":
      ctx.workbook.worksheets.getItem(snap.sheetName).getRange(snap.rowAddress).delete(Excel.DeleteShiftDirection.up);
      break;
    case "deletedLines": {
      const ws = ctx.workbook.worksheets.getItem(snap.sheetName);
      const lines =
        snap.lines === "rows"
          ? ws.getRangeByIndexes(snap.start, 0, snap.count, 1).getEntireRow()
          : ws.getRangeByIndexes(0, snap.start, 1, snap.count).getEntireColumn();
      lines.insert(snap.lines === "rows" ? Excel.InsertShiftDirection.down : Excel.InsertShiftDirection.right);
      if (snap.contentAddress && snap.formulas) ws.getRange(snap.contentAddress).formulas = snap.formulas;
      break;
    }
    case "format": {
      const r = ctx.workbook.worksheets.getItem(snap.sheetName).getRange(stripSheet(snap.address));
      r.setCellProperties(snap.props);
      r.numberFormat = snap.numberFormat;
      break;
    }
    case "addedConditionalFormat":
      ctx.workbook.worksheets.getItem(snap.sheetName).getRange(stripSheet(snap.address)).conditionalFormats.getItem(snap.id).delete();
      break;
    case "dataValidation": {
      const dv = ctx.workbook.worksheets.getItem(snap.sheetName).getRange(stripSheet(snap.address)).dataValidation;
      dv.clear();
      if (snap.rule) dv.rule = snap.rule;
      break;
    }
    case "addedName":
      ctx.workbook.names.getItem(snap.name).delete();
      break;
    case "autoFilter":
      ctx.workbook.worksheets.getItem(snap.sheetName).autoFilter.remove();
      break;
    case "freeze": {
      const fp = ctx.workbook.worksheets.getItem(snap.sheetName).freezePanes;
      fp.unfreeze();
      if (snap.previous) fp.freezeAt(stripSheet(snap.previous));
      break;
    }
    case "multi":
      for (const step of [...snap.steps].reverse()) await applyUndo(ctx, step);
      break;
  }
  await ctx.sync();
}

/* ---------------------------------------------------------------- format snapshots (for undo) */

/** Largest range whose per-cell formatting is saved for undo (bigger ranges are formatted without undo). */
export const FORMAT_UNDO_MAX_CELLS = 20_000;

const FORMAT_LOAD: Excel.CellPropertiesLoadOptions = {
  format: {
    fill: { color: true, pattern: true },
    font: { bold: true, italic: true, color: true, name: true, size: true, underline: true, strikethrough: true },
    borders: { style: true, color: true, weight: true },
    wrapText: true,
    horizontalAlignment: true,
    verticalAlignment: true,
    indentLevel: true,
  },
};

/**
 * Turn loaded cell properties into settable ones. A cell without fill reports a white colour,
 * so it is restored as "no fill" (pattern None) instead of painting it white.
 */
export function toSettableFormat(props: Excel.CellProperties[][]): Excel.SettableCellProperties[][] {
  return props.map((row) =>
    row.map((c) => {
      const f = c.format ?? {};
      const fill = f.fill && String(f.fill.pattern ?? "") === "None" ? { pattern: "None" as Excel.FillPattern } : f.fill;
      return { format: { ...f, fill } };
    })
  );
}

/**
 * @internal Queue a read of a range's formatting and return a function that builds the snapshot after
 * the next sync. Returns null when the range is too large to snapshot. `r` must have rowCount/columnCount loaded.
 */
export function prepareFormatSnapshot(r: Excel.Range, sheetName: string, label: string): (() => Snapshot) | null {
  if (r.rowCount * r.columnCount > FORMAT_UNDO_MAX_CELLS) return null;
  const props = r.getCellProperties(FORMAT_LOAD);
  r.load("address,numberFormat");
  return () => ({ kind: "format", sheetName, address: r.address, props: toSettableFormat(props.value), numberFormat: r.numberFormat, label });
}

/** Undo the most recent Zee AI change (call repeatedly for multiple steps). */
export async function undoLastChange(): Promise<string> {
  const snap = undoStack[undoStack.length - 1]?.snap;
  if (!snap) throw new ExcelUserError("Nothing to undo.");
  try {
    await run((ctx) => applyUndo(ctx, snap));
  } catch (e) {
    // Only forget the step when its target is gone; otherwise (protected sheet, cell in edit mode…)
    // keep it so the user can fix the cause and press undo again.
    if (e instanceof ExcelUserError && /no longer exists/i.test(e.message)) {
      undoStack.pop();
      notifyUndo();
      throw new ExcelUserError(`Couldn't undo "${snap.label}": the sheet, chart or range no longer exists. It was removed from the undo list.`);
    }
    throw new ExcelUserError(`Couldn't undo "${snap.label}" – ${(e as Error).message} Fix that (e.g. leave cell edit mode or unprotect the sheet) and press undo again.`);
  }
  undoStack.pop();
  notifyUndo();
  return snap.label;
}

/* ---------------------------------------------------------------- change history */

export interface HistoryItem {
  /** Position in the stack (0 = oldest). */
  index: number;
  label: string;
  at: number;
  /** Can be reverted without undoing newer changes (no newer change touches the same sheet). */
  alone: boolean;
}

/** Zee AI's recent changes, newest first. */
export function getChangeHistory(): HistoryItem[] {
  const steps = undoStack.map((e) => e.snap);
  return undoStack.map((e, index) => ({ index, label: e.snap.label, at: e.at, alone: canRevertAlone(steps, index) })).reverse();
}

/** Revert one change that no newer change depends on (see HistoryItem.alone). */
export async function revertChange(index: number): Promise<string> {
  const entry = undoStack[index];
  if (!entry) throw new ExcelUserError("That change is no longer in the history.");
  if (!canRevertAlone(undoStack.map((e) => e.snap), index))
    throw new ExcelUserError(`Newer changes touch the same sheet – use “Undo back to here” for "${entry.snap.label}".`);
  try {
    await run((ctx) => applyUndo(ctx, entry.snap));
  } catch (e) {
    if (e instanceof ExcelUserError && /no longer exists/i.test(e.message)) {
      undoStack.splice(undoStack.indexOf(entry), 1);
      notifyUndo();
      throw new ExcelUserError(`Couldn't revert "${entry.snap.label}": the sheet, chart or range no longer exists. It was removed from the history.`);
    }
    throw e;
  }
  undoStack.splice(undoStack.indexOf(entry), 1);
  notifyUndo();
  return entry.snap.label;
}

/** Undo every change from the newest back to (and including) `index`. Returns how many were undone. */
export async function undoBackTo(index: number): Promise<number> {
  let n = 0;
  while (undoStack.length > index) {
    await undoLastChange();
    n++;
  }
  return n;
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

export function hasContent(formulas: unknown[][]): boolean {
  return formulas.some((r) => r.some((v) => v !== "" && v !== null && v !== undefined));
}

export interface WriteOptions {
  overwrite?: boolean;
  /** Bold the first row (table header). */
  header?: boolean;
  label?: string;
  /** Turn the written block into a formatted Excel table (filters, banded rows). Needs a header row. */
  asTable?: boolean;
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
    target.load("address,formulas,numberFormat,rowCount,columnCount");
    await ctx.sync();
    if (!opts.overwrite && hasContent(target.formulas)) throw new OverwriteRequiredError(target.address);

    const label = opts.label ?? `insert at ${stripSheet(target.address)}`;
    const cellsSnap: Snapshot = {
      kind: "cells",
      sheetName: sheet.name,
      address: target.address,
      formulas: target.formulas,
      numberFormat: target.numberFormat,
      label,
    };
    if (opts.asTable && data.length > 1) {
      // Formatted Excel table: undo removes the table, restores the cells and their old formatting.
      const formatSnap = prepareFormatSnapshot(target, sheet.name, label);
      await ctx.sync();
      const before = formatSnap?.();
      target.formulas = data;
      const table = sheet.tables.add(target.address, true);
      table.style = "TableStyleMedium2";
      table.load("name");
      autofitAllOf(target);
      target.select();
      await ctx.sync();
      setSnapshot({ kind: "multi", steps: [...(before ? [before] : []), cellsSnap, { kind: "addedTable", tableName: table.name, label }], label });
      return target.address;
    }
    setSnapshot(cellsSnap);
    target.formulas = data;
    if (opts.header && data.length > 1) target.getRow(0).format.font.bold = true;
    autofitAllOf(target);
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
  /** Dry run: a few example changes (address, before → after). */
  samples?: Array<{ address: string; before: string; after: string }>;
}

/** Progress reporting + cancel for long scans. */
export interface ProgressOptions {
  /** fraction 0..1 and a short status text. */
  onProgress?: (fraction: number, text: string) => void;
  signal?: AbortSignal;
}

function checkCancelled(signal?: AbortSignal) {
  if (signal?.aborted) throw new ExcelUserError("Cancelled – nothing was changed.");
}

const CLEAN_MAX_CELLS = 200_000;
const UNDO_MAX_CELLS = 50_000;
/** Remove empty rows/columns scans in chunks, so it can handle much bigger sheets than Clean. */
const REMOVE_EMPTY_MAX_CELLS = 10_000_000;

const shortCell = (v: unknown) => {
  const t = v === null || v === undefined ? "" : String(v);
  return t.length > 30 ? `${t.slice(0, 28)}…` : t;
};

/**
 * Clean the selection. With `dryRun` nothing is written: the result tells how many cells would change
 * and shows a few examples (used for the preview before cleaning).
 */
export async function cleanSelection(options: CleanOptions, opts: ProgressOptions & { dryRun?: boolean } = {}): Promise<CleanSummary> {
  return run(async (ctx) => {
    const c = await loadClippedSelection(ctx);
    if (c.isEmpty) throw new ExcelUserError("The selected cells are empty.");
    const b = c.block;
    const total = b.rowCount * b.columnCount;
    if (total > CLEAN_MAX_CELLS) throw new ExcelUserError(`Select fewer than ${CLEAN_MAX_CELLS.toLocaleString()} cells to clean at once.`);

    const dry = !!opts.dryRun;
    const keepUndo = !dry && total <= UNDO_MAX_CELLS;
    const allFormulas: unknown[][] = [];
    const allFormats: unknown[][] = [];
    const samples: NonNullable<CleanSummary["samples"]> = [];
    const chunkRows = Math.max(1, Math.floor(20_000 / b.columnCount));
    let changed = 0;

    for (let start = 0; start < b.rowCount; start += chunkRows) {
      checkCancelled(opts.signal);
      const rows = Math.min(chunkRows, b.rowCount - start);
      opts.onProgress?.(start / b.rowCount, `${dry ? "Checking" : "Cleaning"} rows ${(start + 1).toLocaleString()}–${(start + rows).toLocaleString()} of ${b.rowCount.toLocaleString()}…`);
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
            if (dry && samples.length < 6) samples.push({ address: a1(b.rowIndex + start + i, b.columnIndex + j), before: shortCell(v), after: shortCell(res.value) });
          }
          return res.value;
        })
      );
      if (chunkChanged > 0 && !dry) {
        r.numberFormat = newFormats;
        r.formulas = newFormulas;
      }
      changed += chunkChanged;
    }
    opts.onProgress?.(1, dry ? "Preview ready" : "Done");
    if (changed > 0 && !dry) {
      autofitAllOf(c.sheet.getRangeByIndexes(b.rowIndex, b.columnIndex, 1, 1));
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
    return { address: b.address, cellsChanged: changed, cellsScanned: total, undoAvailable: keepUndo && changed > 0, ...(dry ? { samples } : {}) };
  });
}

/* ---------------------------------------------------------------- remove empty rows / columns */

export interface RemoveEmptySummary {
  address: string;
  removed: number;
  scanned: number;
  undoAvailable: boolean;
}

/** Result of the scan: what "Remove empty rows / columns" would delete (shown as a preview first). */
export interface EmptyLinesPlan {
  kind: "rows" | "columns";
  sheetName: string;
  /** The sheet's used range when scanned (the apply step checks it didn't change). */
  address: string;
  block: Block;
  /** 0-based offsets inside the block. */
  hits: number[];
  /** Rows / columns checked. */
  scanned: number;
  undo: { formulas: unknown[][]; numberFormat: unknown[][] } | null;
}

/**
 * Finds every row (or column) on the active sheet whose cells are all blank, 0 or a null word
 * (NULL, N/A, NaN…). Always scans the sheet's whole data area – no selection needed. Nothing changes.
 */
export async function scanEmptyLines(kind: "rows" | "columns", opts: ProgressOptions = {}): Promise<EmptyLinesPlan> {
  return run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getActiveWorksheet();
    sheet.load("name");
    const used = sheet.getUsedRangeOrNullObject(true);
    used.load("address,rowIndex,columnIndex,rowCount,columnCount");
    await ctx.sync();
    if (used.isNullObject) throw new ExcelUserError("This sheet is empty.");
    const b: Block = blockOf(used);
    const total = b.rowCount * b.columnCount;
    if (total > REMOVE_EMPTY_MAX_CELLS)
      throw new ExcelUserError(`This sheet has ${total.toLocaleString()} cells – the limit is ${REMOVE_EMPTY_MAX_CELLS.toLocaleString()}.`);

    // Read the sheet in row chunks (~20,000 cells each) so big sheets stay under Office's payload limit.
    const keepUndo = total <= UNDO_MAX_CELLS;
    const undoFormulas: unknown[][] = [];
    const undoFormats: unknown[][] = [];
    const emptyRows: number[] = [];
    const filled = Array<boolean>(b.columnCount).fill(false);
    const chunkRows = Math.max(1, Math.floor(20_000 / b.columnCount));
    for (let start = 0; start < b.rowCount; start += chunkRows) {
      checkCancelled(opts.signal);
      const rows = Math.min(chunkRows, b.rowCount - start);
      opts.onProgress?.(start / b.rowCount, `Scanning rows ${(start + 1).toLocaleString()}–${(start + rows).toLocaleString()} of ${b.rowCount.toLocaleString()}…`);
      const r = sheet.getRangeByIndexes(b.rowIndex + start, b.columnIndex, rows, b.columnCount);
      r.load(keepUndo ? "values,formulas,numberFormat" : "values");
      await ctx.sync();
      if (keepUndo) {
        undoFormulas.push(...r.formulas);
        undoFormats.push(...r.numberFormat);
      }
      if (kind === "rows") {
        for (const i of findEmptyRows(r.values)) emptyRows.push(start + i);
      } else {
        const skipHeader = start === 0 && b.rowCount > 1 && looksLikeHeaderRow(r.values[0]);
        markFilledColumns(r.values, filled, skipHeader);
      }
    }
    opts.onProgress?.(1, "Scan complete");
    const hits = kind === "rows" ? emptyRows : filled.flatMap((f, j) => (f ? [] : [j]));
    return {
      kind,
      sheetName: sheet.name,
      address: b.address,
      block: b,
      hits,
      scanned: kind === "rows" ? b.rowCount : b.columnCount,
      undo: keepUndo ? { formulas: undoFormulas, numberFormat: undoFormats } : null,
    };
  });
}

/** Deletes what scanEmptyLines found (after the user confirmed the preview). */
export async function applyEmptyLines(plan: EmptyLinesPlan, opts: ProgressOptions = {}): Promise<RemoveEmptySummary> {
  const { kind, block: b, hits } = plan;
  if (!hits.length) return { address: plan.address, removed: 0, scanned: plan.scanned, undoAvailable: false };
  return run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getItemOrNullObject(plan.sheetName);
    const used = sheet.getUsedRangeOrNullObject(true);
    used.load("address");
    await ctx.sync();
    if (sheet.isNullObject || used.isNullObject || used.address !== plan.address)
      throw new ExcelUserError("The sheet changed since the preview – run Remove empty rows / columns again.");

    // Delete bottom-up / right-to-left so earlier indexes stay valid; sync in batches.
    // The used range holds all the sheet's data, so whole rows/columns can go (much faster on big sheets).
    const segs = runsDescending(hits);
    let queued = 0;
    for (const seg of segs) {
      if (kind === "rows") sheet.getRangeByIndexes(b.rowIndex + seg.start, 0, seg.count, 1).getEntireRow().delete(Excel.DeleteShiftDirection.up);
      else sheet.getRangeByIndexes(0, b.columnIndex + seg.start, 1, seg.count).getEntireColumn().delete(Excel.DeleteShiftDirection.left);
      if (++queued % 500 === 0) {
        opts.onProgress?.(queued / segs.length, `Removing… ${queued.toLocaleString()} of ${segs.length.toLocaleString()} blocks`);
        await ctx.sync();
      }
    }
    await ctx.sync();

    const left = sheet.getUsedRangeOrNullObject(true);
    left.load("address");
    await ctx.sync();
    if (!left.isNullObject) {
      left.getEntireColumn().format.autofitColumns();
      await ctx.sync();
    }

    if (plan.undo)
      setSnapshot({
        kind: "cells",
        sheetName: plan.sheetName,
        address: b.address,
        formulas: plan.undo.formulas,
        numberFormat: plan.undo.numberFormat,
        label: `remove empty ${kind} ${stripSheet(b.address)}`,
      });
    return { address: b.address, removed: hits.length, scanned: plan.scanned, undoAvailable: !!plan.undo };
  });
}

/** Scan + delete in one go (no preview). */
export async function removeEmptyLines(kind: "rows" | "columns", opts: ProgressOptions = {}): Promise<RemoveEmptySummary> {
  return applyEmptyLines(await scanEmptyLines(kind, opts), opts);
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
    autofitAllOf(sheet.getRangeByIndexes(0, newCol, 1, 1));
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
    if (o.autofit) autofitAllOf(r);
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
    autofitAllOf(table.getRange());
    table.load("name");
    await ctx.sync();
    setSnapshot({ kind: "addedTable", tableName: table.name, label: `table ${table.name}` });
    return table.name;
  });
}

/** Auto-fit all columns of the active sheet, whatever is selected. */
export async function autofitAllColumns(): Promise<string> {
  return run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getActiveWorksheet();
    sheet.load("name");
    const used = sheet.getUsedRangeOrNullObject(true);
    used.load("address");
    await ctx.sync();
    if (used.isNullObject) throw new ExcelUserError("This sheet is empty – nothing to auto-fit.");
    used.getEntireColumn().format.autofitColumns();
    await ctx.sync();
    return `${sheet.name} – all columns`;
  });
}

/** Kept for older callers: now fits every column on the sheet. */
export const autofitSelection = autofitAllColumns;

/* ================================================================ agent / workbook-level helpers */

/** Resolve "Sheet1!A1:B3", "'My Sheet'!A1" or "A1:B3" (+ optional sheet name) to a Range. */
export function resolveRange(ctx: Excel.RequestContext, address: string, sheet?: string): Excel.Range {
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

/**
 * Format any range. With `trackUndo` (agent mode) the previous per-cell formatting is saved so ↶ can
 * restore it (ranges up to FORMAT_UNDO_MAX_CELLS cells). Report builders on fresh sheets skip it.
 */
export async function formatRangeAt(address: string, o: RangeFormat, sheet?: string, trackUndo = false): Promise<string> {
  return run(async (ctx) => {
    const r = resolveRange(ctx, address, sheet);
    r.load("address,rowCount,columnCount");
    const ws = r.worksheet;
    ws.load("name");
    await ctx.sync();
    if (trackUndo) {
      const build = prepareFormatSnapshot(r, ws.name, `format ${stripSheet(r.address)}`);
      if (build) {
        await ctx.sync();
        setSnapshot(build());
      }
    }
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
    if (o.autofit) autofitAllOf(r);
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
    autofitAllOf(target);
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

/** Like readSelectionBlock, for any address (clipped to the used range; used by agent tools). */
export async function readBlockAt(address: string, sheet?: string, maxCells = 200_000): Promise<SelectionBlock> {
  return run(async (ctx) => {
    const target = resolveRange(ctx, address, sheet);
    const ws = target.worksheet;
    ws.load("name");
    const used = ws.getUsedRangeOrNullObject(true);
    used.load("address");
    await ctx.sync();
    if (used.isNullObject) throw new ExcelUserError(`${address} is empty.`);
    const r = target.getIntersectionOrNullObject(used);
    r.load("address,rowIndex,columnIndex,rowCount,columnCount");
    await ctx.sync();
    if (r.isNullObject) throw new ExcelUserError(`${address} is empty.`);
    if (r.rowCount * r.columnCount > maxCells) throw new ExcelUserError(`Range is too large (max ${maxCells.toLocaleString()} cells).`);
    r.load("values,text");
    await ctx.sync();
    return { sheetName: ws.name, address: r.address, rowIndex: r.rowIndex, columnIndex: r.columnIndex, values: r.values, text: r.text };
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
    autofitAllOf(target);
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

/* ================================================================ Draw: plot chosen columns */

export interface PlotColumn {
  /** Absolute column index on the sheet. */
  index: number;
  letter: string;
  /** Header text, or "Column X" when there is no header row. */
  name: string;
  /** Most sampled data cells are numbers (dates count as numbers). */
  numeric: boolean;
}

export interface PlotSource {
  sheetName: string;
  /** Where the data comes from, e.g. "A1:F250" – the selection if it has 2+ columns, otherwise the used range. */
  address: string;
  columns: PlotColumn[];
  hasHeader: boolean;
  /** Absolute row indexes of the first and last data rows (header excluded). */
  firstDataRow: number;
  lastDataRow: number;
}

const isNum = (v: unknown) => typeof v === "number" || (typeof v === "string" && v.trim() !== "" && Number.isFinite(Number(v.replace(/,/g, ""))));

/** Columns available for plotting: the selection (if it spans 2+ columns) or the active sheet's used range. */
export async function getPlotSource(): Promise<PlotSource> {
  return run(async (ctx) => {
    const c = await loadClippedSelection(ctx);
    let block = c.block;
    if (c.isEmpty || block.columnCount < 2 || block.rowCount < 2) {
      const used = c.sheet.getUsedRangeOrNullObject(true);
      used.load("address,rowIndex,columnIndex,rowCount,columnCount");
      await ctx.sync();
      if (used.isNullObject) throw new ExcelUserError("This sheet is empty. Open the sheet with your data first.");
      block = blockOf(used);
    }
    if (block.columnCount < 2) throw new ExcelUserError("The data needs at least two columns: one for the X axis and one to plot.");
    const cols = Math.min(block.columnCount, 200);
    const sampleRows = Math.min(block.rowCount, 12);
    const sample = c.sheet.getRangeByIndexes(block.rowIndex, block.columnIndex, sampleRows, cols);
    sample.load("values");
    await ctx.sync();

    const first = sample.values[0] ?? [];
    const textCells = first.filter((v) => typeof v === "string" && v.trim() !== "" && !isNum(v)).length;
    const hasHeader = block.rowCount > 1 && textCells >= Math.ceil(first.filter((v) => v !== "" && v !== null).length / 2) && textCells > 0;
    const body = hasHeader ? sample.values.slice(1) : sample.values;

    const columns: PlotColumn[] = Array.from({ length: cols }, (_, j) => {
      const letter = columnLetter(block.columnIndex + j);
      const raw = hasHeader ? String(first[j] ?? "").trim() : "";
      const filled = body.map((r) => r[j]).filter((v) => v !== "" && v !== null && v !== undefined);
      return {
        index: block.columnIndex + j,
        letter,
        name: raw || `Column ${letter}`,
        numeric: filled.length > 0 && filled.filter(isNum).length >= Math.ceil(filled.length * 0.6),
      };
    });
    return {
      sheetName: c.sheetName,
      address: stripSheet(block.address),
      columns,
      hasHeader,
      firstDataRow: block.rowIndex + (hasHeader ? 1 : 0),
      lastDataRow: block.rowIndex + block.rowCount - 1,
    };
  });
}

export interface PlotRequest {
  source: PlotSource;
  xIndex: number;
  yIndexes: number[];
  /** Excel.ChartType value from utils/chartTypes. */
  chartType: string;
  title?: string;
  /** Y columns to draw against a right-hand axis (2-D series charts only). */
  secondaryYIndexes?: number[];
  /** Logarithmic value axes (e.g. production rate on a semi-log decline plot). */
  logLeft?: boolean;
  logRight?: boolean;
}

const HELPER_MAX_ROWS = 50_000;

/** Chart several Y columns against one X column, in any Excel chart type. */
export async function plotColumns(req: PlotRequest): Promise<{ title: string; helperSheet?: string; warnings: string[] }> {
  const { source: src } = req;
  const info = chartTypeInfo(req.chartType);
  const yIdx = [...req.yIndexes].sort((a, b) => a - b); // sheet order (matters for stock charts)
  const yProblem = checkYCount(info, yIdx.length);
  if (yProblem) throw new ExcelUserError(yProblem);
  const rows = src.lastDataRow - src.firstDataRow + 1;
  if (rows < 1) throw new ExcelUserError("There are no data rows under the header.");
  const colByIndex = new Map(src.columns.map((c) => [c.index, c]));
  const nameOf = (i: number) => colByIndex.get(i)?.name ?? columnLetter(i);
  const xName = nameOf(req.xIndex);
  const yNames = yIdx.map(nameOf);
  const autoTitle = info.useX ? `${yNames.join(", ")} vs ${xName}` : yNames.join(", ");
  const title = (req.title?.trim() || autoTitle).slice(0, 120);

  // Range-based types need one contiguous block: [X?, Y1, Y2, ...] in that order.
  const blockCols = info.mode === "range" ? (info.useX ? [req.xIndex, ...yIdx] : yIdx) : [];
  const contiguous = blockCols.every((c, k) => k === 0 || c === blockCols[k - 1] + 1);
  let helperSheet: string | undefined;
  if (info.mode === "range" && !contiguous) {
    if (rows > HELPER_MAX_ROWS) throw new ExcelUserError(`For ${info.label}, put the columns side by side (X first) or select at most ${HELPER_MAX_ROWS.toLocaleString()} rows.`);
    helperSheet = await addSheet("Chart data", false);
  }

  return run(async (ctx) => {
    const ws = ctx.workbook.worksheets.getItem(src.sheetName);
    let chart: Excel.Chart;

    if (info.mode === "series") {
      const xRange = ws.getRangeByIndexes(src.firstDataRow, req.xIndex, rows, 1);
      const yRanges = yIdx.map((i) => ws.getRangeByIndexes(src.firstDataRow, i, rows, 1));
      chart = ws.charts.add(info.id as Excel.ChartType, yRanges[0], Excel.ChartSeriesBy.columns);
      const first = chart.series.getItemAt(0);
      first.name = yNames[0];
      first.setXAxisValues(xRange);
      yIdx.slice(1).forEach((_, k) => {
        const s = chart.series.add(yNames[k + 1]);
        s.setValues(yRanges[k + 1]);
        s.setXAxisValues(xRange);
      });
    } else {
      const withHeader = src.hasHeader || !!helperSheet;
      const top = withHeader ? src.firstDataRow - (src.hasHeader ? 1 : 0) : src.firstDataRow;
      let block: Excel.Range;
      if (!helperSheet) {
        // Include the header row so Excel names the series.
        block = ws.getRangeByIndexes(top, blockCols[0], rows + (src.hasHeader ? 1 : 0), blockCols.length);
      } else {
        // Linked copy (formulas pointing at the original cells) so the chart stays live.
        const hs = ctx.workbook.worksheets.getItem(helperSheet);
        hs.getRangeByIndexes(0, 0, 1, blockCols.length).values = [blockCols.map(nameOf)];
        hs.getRangeByIndexes(0, 0, 1, blockCols.length).format.font.bold = true;
        const CHUNK = 5000;
        for (let start = 0; start < rows; start += CHUNK) {
          const n = Math.min(CHUNK, rows - start);
          const formulas = Array.from({ length: n }, (_, r) =>
            blockCols.map((c) => `=${quoteSheet(src.sheetName)}!${a1(src.firstDataRow + start + r, c)}`)
          );
          hs.getRangeByIndexes(1 + start, 0, n, blockCols.length).formulas = formulas;
          await ctx.sync();
        }
        autofitAllOf(hs.getRangeByIndexes(0, 0, rows + 1, blockCols.length));
        block = hs.getRangeByIndexes(0, 0, rows + 1, blockCols.length);
      }
      chart = ws.charts.add(info.id as Excel.ChartType, block, Excel.ChartSeriesBy.columns);
    }

    chart.title.text = title;
    chart.title.visible = true;
    const lastCol = Math.max(...src.columns.map((c) => c.index));
    chart.setPosition(ws.getRangeByIndexes(Math.max(0, src.firstDataRow - (src.hasHeader ? 1 : 0)), lastCol + 2, 1, 1));
    chart.width = 520;
    chart.height = 320;
    chart.load("name");
    await ctx.sync();
    setSnapshot({ kind: "addedChart", sheetName: src.sheetName, chartName: chart.name, label: `chart "${title.slice(0, 40)}"` });

    // Cosmetics in a separate batch: some chart types (treemap, funnel, map…) have no legend/axes,
    // and a failure here must not lose the chart.
    try {
      chart.legend.visible = yIdx.length > 1 || !info.axes;
      chart.legend.position = Excel.ChartLegendPosition.bottom;
      await ctx.sync();
    } catch {
      /* no legend on this type */
    }
    if (info.axes) {
      try {
        if (info.useX) {
          chart.axes.categoryAxis.title.text = xName;
          chart.axes.categoryAxis.title.visible = true;
        }
        if (yIdx.length === 1) {
          chart.axes.valueAxis.title.text = yNames[0];
          chart.axes.valueAxis.title.visible = true;
        }
        await ctx.sync();
      } catch {
        /* axis titles not supported for this type */
      }
    }
    // Right-hand axis and log scales (newer chart APIs – report instead of failing).
    const warnings: string[] = [];
    const secondary = supportsSecondaryAxis(info) ? yIdx.filter((i) => req.secondaryYIndexes?.includes(i)) : [];
    if (secondary.length && secondary.length < yIdx.length) {
      try {
        yIdx.forEach((i, k) => {
          if (secondary.includes(i)) chart.series.getItemAt(k).axisGroup = Excel.ChartAxisGroup.secondary;
        });
        await ctx.sync();
        const right = chart.axes.getItem(Excel.ChartAxisType.value, Excel.ChartAxisGroup.secondary);
        right.title.text = secondary.map(nameOf).join(", ");
        right.title.visible = true;
        const left = chart.axes.valueAxis;
        left.title.text = yIdx.filter((i) => !secondary.includes(i)).map(nameOf).join(", ");
        left.title.visible = true;
        await ctx.sync();
      } catch {
        warnings.push("right-hand axis needs a newer Excel version");
      }
    }
    if (supportsLogScale(info) && (req.logLeft || (req.logRight && secondary.length))) {
      try {
        if (req.logLeft) chart.axes.valueAxis.scaleType = Excel.ChartAxisScaleType.logarithmic;
        if (req.logRight && secondary.length && secondary.length < yIdx.length)
          chart.axes.getItem(Excel.ChartAxisType.value, Excel.ChartAxisGroup.secondary).scaleType = Excel.ChartAxisScaleType.logarithmic;
        await ctx.sync();
      } catch {
        warnings.push("log scale could not be applied (zero/negative values or older Excel)");
      }
    }
    return { title, helperSheet, warnings };
  });
}

/** Raw values of chosen columns over the data rows of a PlotSource (one array per column). */
export async function readSourceColumns(src: PlotSource, indices: number[]): Promise<unknown[][]> {
  const rows = src.lastDataRow - src.firstDataRow + 1;
  if (rows < 1) return indices.map(() => []);
  if (rows * indices.length > 500_000) throw new ExcelUserError("Too many rows – select at most about 100,000 rows.");
  return run(async (ctx) => {
    const ws = ctx.workbook.worksheets.getItem(src.sheetName);
    const ranges = indices.map((i) => {
      const r = ws.getRangeByIndexes(src.firstDataRow, i, rows, 1);
      r.load("values");
      return r;
    });
    await ctx.sync();
    return ranges.map((r) => r.values.map((row) => row[0]));
  });
}

export interface ScatterSeriesSpec {
  name: string;
  /** Addresses on `sheetName` (no sheet prefix), single column each. */
  x: string;
  y: string;
  style: "points" | "line" | "smooth";
}

/** Scatter chart built from explicit series (e.g. pump curve + operating points). Not undo-tracked (lives on a report sheet). */
export async function createScatterSeriesChart(sheetName: string, title: string, series: ScatterSeriesSpec[], anchor: string, axisTitles?: { x?: string; y?: string }): Promise<void> {
  if (!series.length) return;
  const STYLE: Record<ScatterSeriesSpec["style"], string> = { points: "XYScatter", line: "XYScatterLinesNoMarkers", smooth: "XYScatterSmoothNoMarkers" };
  await run(async (ctx) => {
    const ws = ctx.workbook.worksheets.getItem(sheetName);
    const chart = ws.charts.add(STYLE[series[0].style] as Excel.ChartType, ws.getRange(series[0].y), Excel.ChartSeriesBy.columns);
    const first = chart.series.getItemAt(0);
    first.name = series[0].name;
    first.setXAxisValues(ws.getRange(series[0].x));
    const added = series.slice(1).map((s) => {
      const ser = chart.series.add(s.name);
      ser.setValues(ws.getRange(s.y));
      ser.setXAxisValues(ws.getRange(s.x));
      return ser;
    });
    chart.title.text = title;
    chart.setPosition(anchor);
    chart.width = 520;
    chart.height = 320;
    await ctx.sync();
    try {
      added.forEach((ser, k) => (ser.chartType = STYLE[series[k + 1].style] as Excel.ChartType));
      chart.legend.position = Excel.ChartLegendPosition.bottom;
      if (axisTitles?.x) {
        chart.axes.categoryAxis.title.text = axisTitles.x;
        chart.axes.categoryAxis.title.visible = true;
      }
      if (axisTitles?.y) {
        chart.axes.valueAxis.title.text = axisTitles.y;
        chart.axes.valueAxis.title.visible = true;
      }
      await ctx.sync();
    } catch {
      /* cosmetic – older Excel */
    }
  });
}

/* ---------------------------------------------------------------- =ZEE.* function picker */

export interface FunctionTarget {
  sheetName: string;
  /** Clipped selection block (0-based). */
  rowIndex: number;
  columnIndex: number;
  rowCount: number;
  columnCount: number;
  /** True when the selected cells hold no data → insert into the active cell instead. */
  isEmpty: boolean;
  /** First few rows of values (for header detection and previews). */
  sample: unknown[][];
  activeAddress: string;
}

/** Where a function picked in the pane would apply. */
export async function getFunctionTarget(): Promise<FunctionTarget> {
  return run(async (ctx) => {
    const c = await loadClippedSelection(ctx);
    const active = ctx.workbook.getActiveCell();
    active.load("address");
    const b = c.block;
    let sample: unknown[][] = [];
    if (!c.isEmpty) {
      const r = c.sheet.getRangeByIndexes(b.rowIndex, b.columnIndex, Math.min(b.rowCount, 3), Math.min(b.columnCount, 50));
      r.load("values");
      await ctx.sync();
      sample = r.values;
      if (sample.every((row) => row.every((v) => v === "" || v === null))) sample = [];
    } else {
      await ctx.sync();
    }
    return {
      sheetName: c.sheetName,
      rowIndex: b.rowIndex,
      columnIndex: b.columnIndex,
      rowCount: b.rowCount,
      columnCount: b.columnCount,
      isEmpty: c.isEmpty || sample.length === 0,
      sample,
      activeAddress: active.address,
    };
  });
}

/**
 * Write a column of formulas (and optional header text) starting at (rowIndex, columnIndex).
 * `reserveRows` lets a spilling formula (e.g. ZEE.FLAG) check that the cells it will spill into are empty.
 */
export async function writeFormulaColumn(
  sheetName: string,
  rowIndex: number,
  columnIndex: number,
  cells: string[],
  opts: WriteOptions & { reserveRows?: number } = {}
): Promise<string> {
  if (!cells.length) throw new ExcelUserError("Nothing to insert.");
  return run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getItem(sheetName);
    const checkRows = Math.max(cells.length, opts.reserveRows ?? 0);
    if (rowIndex + checkRows > EXCEL_MAX_ROWS) throw new ExcelUserError("Not enough rows below the selection.");
    const check = sheet.getRangeByIndexes(rowIndex, columnIndex, checkRows, 1);
    check.load("address,formulas,numberFormat");
    await ctx.sync();
    if (!opts.overwrite && hasContent(check.formulas)) throw new OverwriteRequiredError(check.address);

    setSnapshot({
      kind: "cells",
      sheetName,
      address: check.address,
      formulas: check.formulas,
      numberFormat: check.numberFormat,
      label: opts.label ?? `insert formulas at ${stripSheet(check.address)}`,
    });
    if (checkRows > cells.length) sheet.getRangeByIndexes(rowIndex + cells.length, columnIndex, checkRows - cells.length, 1).clear(Excel.ClearApplyTo.contents);
    const target = sheet.getRangeByIndexes(rowIndex, columnIndex, cells.length, 1);
    // Formulas must not inherit a Text ("@") format, or Excel shows them as text.
    target.numberFormat = cells.map(() => ["General"]);
    target.formulas = cells.map((f) => [f]);
    if (opts.header) target.getCell(0, 0).format.font.bold = true;
    autofitAllOf(target);
    target.getCell(opts.header && cells.length > 1 ? 1 : 0, 0).select();
    await ctx.sync();
    return target.address;
  });
}

/** Replace formulas in the selection with their current values (stops further AI calls). */
export async function freezeSelection(): Promise<{ address: string; frozen: number }> {
  return run(async (ctx) => {
    const c = await loadClippedSelection(ctx);
    if (c.isEmpty) throw new ExcelUserError("Select the cells with =ZEE formulas you want to freeze.");
    const b = c.block;
    if (b.rowCount * b.columnCount > 200_000) throw new ExcelUserError("Select fewer cells (max 200,000) to freeze.");
    const r = c.sheet.getRangeByIndexes(b.rowIndex, b.columnIndex, b.rowCount, b.columnCount);
    r.load("address,formulas,values,numberFormat");
    await ctx.sync();
    let frozen = 0;
    const out = r.formulas.map((row, i) =>
      row.map((f, j) => {
        if (typeof f === "string" && f.startsWith("=")) {
          frozen++;
          return r.values[i][j];
        }
        // Spilled cells show a value but no formula of their own – keep their value too.
        return f === "" && r.values[i][j] !== "" ? r.values[i][j] : f;
      })
    );
    if (!frozen) throw new ExcelUserError("No formulas in the selection – nothing to freeze.");
    setSnapshot({ kind: "cells", sheetName: c.sheetName, address: r.address, formulas: r.formulas, numberFormat: r.numberFormat, label: `freeze ${stripSheet(r.address)}` });
    r.values = out as unknown[][] as (string | number | boolean)[][];
    await ctx.sync();
    return { address: r.address, frozen };
  });
}
