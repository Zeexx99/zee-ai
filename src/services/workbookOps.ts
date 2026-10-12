/**
 * Extra workbook operations used by agent mode: insert/delete rows and columns, clear, copy, sort,
 * filter, find/replace, freeze panes, conditional formatting, data validation, named ranges and
 * PivotTables. Every change records an undo snapshot (see excelService.ts).
 */
import { columnLetter, stripSheet } from "../utils/a1";
import { buildFilterCriteria, cfFormula, matchHeader, normalizeAggregation, normalizeOperator, resolveColumn, validName, type FilterArgs } from "../utils/agentArgs";
import {
  ExcelUserError,
  addSheet,
  autofitAllOf,
  hasContent,
  invalidateOverview,
  prepareFormatSnapshot,
  resolveRange,
  run,
  setSnapshot,
  type Snapshot,
} from "./excelService";

/** Cap on cells saved for undo by destructive operations (delete, clear, replace). */
const UNDO_CELLS_MAX = 200_000;

async function sheetNameOf(ctx: Excel.RequestContext, r: Excel.Range): Promise<string> {
  const ws = r.worksheet;
  ws.load("name");
  await ctx.sync();
  return ws.name;
}

/* ---------------------------------------------------------------- rows & columns */

/** Insert whole rows above / columns left of `address`. `count` defaults to the size of the address. */
export async function insertLines(address: string, lines: "rows" | "columns", count?: number, sheet?: string): Promise<string> {
  return run(async (ctx) => {
    const at = resolveRange(ctx, address, sheet);
    at.load("rowIndex,columnIndex,rowCount,columnCount");
    const sheetName = await sheetNameOf(ctx, at);
    const n = Math.max(1, Math.min(10_000, Math.round(count ?? (lines === "rows" ? at.rowCount : at.columnCount))));
    const ws = ctx.workbook.worksheets.getItem(sheetName);
    if (lines === "rows") {
      ws.getRangeByIndexes(at.rowIndex, 0, n, 1).getEntireRow().insert(Excel.InsertShiftDirection.down);
      await ctx.sync();
      const rowAddress = `${at.rowIndex + 1}:${at.rowIndex + n}`;
      setSnapshot({ kind: "insertedRows", sheetName, rowAddress, label: `insert rows ${rowAddress}` });
      invalidateOverview();
      return `${sheetName}!${rowAddress}`;
    }
    ws.getRangeByIndexes(0, at.columnIndex, 1, n).getEntireColumn().insert(Excel.InsertShiftDirection.right);
    await ctx.sync();
    const columnAddress = `${columnLetter(at.columnIndex)}:${columnLetter(at.columnIndex + n - 1)}`;
    setSnapshot({ kind: "insertedColumn", sheetName, columnAddress, label: `insert columns ${columnAddress}` });
    invalidateOverview();
    return `${sheetName}!${columnAddress}`;
  });
}

/** Delete the whole rows / columns that `address` touches. Cell contents are kept for undo. */
export async function deleteLines(address: string, lines: "rows" | "columns", sheet?: string): Promise<string> {
  return run(async (ctx) => {
    const at = resolveRange(ctx, address, sheet);
    at.load("rowIndex,columnIndex,rowCount,columnCount");
    const sheetName = await sheetNameOf(ctx, at);
    const ws = ctx.workbook.worksheets.getItem(sheetName);
    const used = ws.getUsedRangeOrNullObject(true);
    used.load("rowIndex,columnIndex,rowCount,columnCount");
    await ctx.sync();

    const start = lines === "rows" ? at.rowIndex : at.columnIndex;
    const count = lines === "rows" ? at.rowCount : at.columnCount;
    if (count > 100_000) throw new ExcelUserError("Too many rows/columns to delete at once.");
    let contentAddress: string | null = null;
    let formulas: unknown[][] | null = null;
    if (!used.isNullObject) {
      const block =
        lines === "rows"
          ? ws.getRangeByIndexes(start, used.columnIndex, count, used.columnCount)
          : ws.getRangeByIndexes(used.rowIndex, start, used.rowCount, count);
      const content = block.getIntersectionOrNullObject(used);
      content.load("address,rowCount,columnCount");
      await ctx.sync();
      if (!content.isNullObject) {
        if (content.rowCount * content.columnCount > UNDO_CELLS_MAX) throw new ExcelUserError("These rows/columns hold too much data to delete safely (no undo possible). Delete them in Excel instead.");
        content.load("formulas");
        await ctx.sync();
        if (hasContent(content.formulas)) {
          contentAddress = stripSheet(content.address);
          formulas = content.formulas;
        }
      }
    }
    const target = lines === "rows" ? ws.getRangeByIndexes(start, 0, count, 1).getEntireRow() : ws.getRangeByIndexes(0, start, 1, count).getEntireColumn();
    target.delete(lines === "rows" ? Excel.DeleteShiftDirection.up : Excel.DeleteShiftDirection.left);
    await ctx.sync();
    const label = lines === "rows" ? `delete rows ${start + 1}:${start + count}` : `delete columns ${columnLetter(start)}:${columnLetter(start + count - 1)}`;
    setSnapshot({ kind: "deletedLines", sheetName, lines, start, count, contentAddress, formulas, label });
    invalidateOverview();
    return `${label} on ${sheetName}`;
  });
}

/* ---------------------------------------------------------------- clear / copy */

export async function clearRange(address: string, what: "contents" | "formats" | "all", sheet?: string): Promise<string> {
  return run(async (ctx) => {
    const r = resolveRange(ctx, address, sheet);
    r.load("address,rowCount,columnCount");
    const sheetName = await sheetNameOf(ctx, r);
    if (r.rowCount * r.columnCount > UNDO_CELLS_MAX) throw new ExcelUserError("Range is too large to clear with undo. Use a smaller range.");
    const steps: Array<() => Snapshot> = [];
    if (what !== "formats") {
      r.load("formulas,numberFormat");
      steps.push(() => ({ kind: "cells", sheetName, address: r.address, formulas: r.formulas, numberFormat: r.numberFormat, label: "" }));
    }
    if (what !== "contents") {
      const fmt = prepareFormatSnapshot(r, sheetName, "");
      if (fmt) steps.push(fmt);
    }
    await ctx.sync();
    const label = `clear ${what} ${stripSheet(r.address)}`;
    setSnapshot({ kind: "multi", steps: steps.map((s) => s()), label });
    r.clear(what === "contents" ? Excel.ClearApplyTo.contents : what === "formats" ? Excel.ClearApplyTo.formats : Excel.ClearApplyTo.all);
    await ctx.sync();
    invalidateOverview();
    return r.address;
  });
}

export async function copyRange(source: string, destination: string, mode: "all" | "values" | "formulas" | "formats", sheet?: string): Promise<string> {
  return run(async (ctx) => {
    const src = resolveRange(ctx, source, sheet);
    src.load("address,rowCount,columnCount");
    const destStart = resolveRange(ctx, destination, sheet).getCell(0, 0);
    destStart.load("rowIndex,columnIndex");
    const ws = destStart.worksheet;
    ws.load("name");
    await ctx.sync();
    if (src.rowCount * src.columnCount > UNDO_CELLS_MAX) throw new ExcelUserError("Source range is too large to copy with undo.");
    const dest = ws.getRangeByIndexes(destStart.rowIndex, destStart.columnIndex, src.rowCount, src.columnCount);
    dest.load("address,rowCount,columnCount,formulas,numberFormat");
    await ctx.sync();
    const fmt = mode === "all" || mode === "formats" ? prepareFormatSnapshot(dest, ws.name, "") : null;
    await ctx.sync();
    const steps: Snapshot[] = [{ kind: "cells", sheetName: ws.name, address: dest.address, formulas: dest.formulas, numberFormat: dest.numberFormat, label: "" }];
    if (fmt) steps.push(fmt());
    setSnapshot({ kind: "multi", steps, label: `copy to ${stripSheet(dest.address)}` });
    const type = { all: Excel.RangeCopyType.all, values: Excel.RangeCopyType.values, formulas: Excel.RangeCopyType.formulas, formats: Excel.RangeCopyType.formats }[mode];
    dest.copyFrom(src, type);
    autofitAllOf(dest);
    await ctx.sync();
    invalidateOverview();
    return dest.address;
  });
}

/* ---------------------------------------------------------------- sort / filter */

export async function sortRange(address: string, by: Array<{ column: unknown; ascending?: boolean }>, hasHeaders: boolean, sheet?: string): Promise<string> {
  if (!by.length) throw new ExcelUserError("Say which column to sort by.");
  return run(async (ctx) => {
    const r = resolveRange(ctx, address, sheet);
    r.load("address,rowCount,columnCount,columnIndex,formulas,numberFormat");
    const sheetName = await sheetNameOf(ctx, r);
    if (r.rowCount * r.columnCount > UNDO_CELLS_MAX) throw new ExcelUserError("Range is too large to sort with undo.");
    const header = r.formulas[0] ?? [];
    const fields = by.map((b) => {
      const key = resolveColumn(b.column, header, r.columnIndex);
      if (key === null) throw new ExcelUserError(`Column "${String(b.column)}" is not in ${stripSheet(r.address)}.`);
      return { key, ascending: b.ascending !== false };
    });
    setSnapshot({ kind: "cells", sheetName, address: r.address, formulas: r.formulas, numberFormat: r.numberFormat, label: `sort ${stripSheet(r.address)}` });
    r.sort.apply(fields, false, hasHeaders, Excel.SortOrientation.rows);
    await ctx.sync();
    return r.address;
  });
}

/** Apply (or clear) an AutoFilter on a range. Only one AutoFilter can exist per sheet. */
export async function filterRange(address: string, column: unknown, criteria: FilterArgs & { clear?: boolean }, sheet?: string): Promise<string> {
  return run(async (ctx) => {
    const r = resolveRange(ctx, address, sheet);
    r.load("address,columnIndex,values");
    const sheetName = await sheetNameOf(ctx, r);
    const af = ctx.workbook.worksheets.getItem(sheetName).autoFilter;
    af.load("enabled");
    await ctx.sync();
    if (criteria.clear) {
      if (af.enabled) af.remove();
      await ctx.sync();
      return "filter removed (not undoable)";
    }
    const idx = resolveColumn(column, r.values[0] ?? [], r.columnIndex);
    if (idx === null) throw new ExcelUserError(`Column "${String(column)}" is not in ${stripSheet(r.address)}.`);
    let crit;
    try {
      crit = buildFilterCriteria(criteria);
    } catch (e) {
      throw new ExcelUserError((e as Error).message);
    }
    const wasEnabled = af.enabled;
    af.apply(r, idx, crit as Excel.FilterCriteria);
    await ctx.sync();
    if (!wasEnabled) setSnapshot({ kind: "autoFilter", sheetName, label: `filter ${stripSheet(r.address)}` });
    return `${stripSheet(r.address)} filtered on column ${idx + 1}${wasEnabled ? " (added to the existing filter; undo won't remove it)" : ""}`;
  });
}

/* ---------------------------------------------------------------- find / replace */

export interface FindHit {
  sheet: string;
  addresses: string[];
  total: number;
}

export async function findCells(text: string, opts: { sheet?: string; matchCase?: boolean; wholeCell?: boolean; max?: number } = {}): Promise<FindHit[]> {
  if (!text) throw new ExcelUserError("Nothing to search for.");
  const max = opts.max ?? 50;
  return run(async (ctx) => {
    const sheets = ctx.workbook.worksheets;
    sheets.load("items/name");
    await ctx.sync();
    const list = opts.sheet ? sheets.items.filter((w) => w.name.toLowerCase() === opts.sheet!.toLowerCase()) : sheets.items;
    if (!list.length) throw new ExcelUserError(`Sheet "${opts.sheet}" not found.`);
    const found = list.map((ws) => {
      const areas = ws.findAllOrNullObject(text, { completeMatch: !!opts.wholeCell, matchCase: !!opts.matchCase });
      areas.load("address,cellCount");
      return { ws, areas };
    });
    await ctx.sync();
    const out: FindHit[] = [];
    let left = max;
    for (const { ws, areas } of found) {
      if (areas.isNullObject) continue;
      const addrs = areas.address.split(",").map(stripSheet);
      out.push({ sheet: ws.name, addresses: addrs.slice(0, Math.max(0, left)), total: areas.cellCount });
      left -= addrs.length;
    }
    return out;
  });
}

/** Replace text in cell values/formulas on one sheet (or all sheets). The used range is saved for undo. */
export async function replaceText(find: string, replacement: string, opts: { sheet?: string; matchCase?: boolean; wholeCell?: boolean } = {}): Promise<number> {
  if (!find) throw new ExcelUserError("Nothing to replace.");
  return run(async (ctx) => {
    const sheets = ctx.workbook.worksheets;
    sheets.load("items/name");
    await ctx.sync();
    const list = opts.sheet ? sheets.items.filter((w) => w.name.toLowerCase() === opts.sheet!.toLowerCase()) : sheets.items;
    if (!list.length) throw new ExcelUserError(`Sheet "${opts.sheet}" not found.`);
    const used = list.map((ws) => {
      const u = ws.getUsedRangeOrNullObject(true);
      u.load("address,rowCount,columnCount");
      return { ws, u };
    });
    await ctx.sync();
    const withData = used.filter(({ u }) => !u.isNullObject);
    const cells = withData.reduce((n, { u }) => n + u.rowCount * u.columnCount, 0);
    if (cells > UNDO_CELLS_MAX) throw new ExcelUserError("Too much data to replace with undo. Pick one sheet, or use Excel's Find & Replace (Ctrl+H).");
    withData.forEach(({ u }) => u.load("formulas,numberFormat"));
    await ctx.sync();
    const counts = withData.map(({ ws }) => ws.replaceAll(find, replacement, { completeMatch: !!opts.wholeCell, matchCase: !!opts.matchCase }));
    await ctx.sync();
    const total = counts.reduce((n, c) => n + c.value, 0);
    if (total > 0) {
      setSnapshot({
        kind: "multi",
        steps: withData.map(({ ws, u }) => ({ kind: "cells" as const, sheetName: ws.name, address: u.address, formulas: u.formulas, numberFormat: u.numberFormat, label: "" })),
        label: `replace "${find}"`,
      });
      invalidateOverview();
    }
    return total;
  });
}

/* ---------------------------------------------------------------- freeze panes */

export async function freezePanes(rows: number, columns: number, sheet?: string): Promise<string> {
  return run(async (ctx) => {
    const ws = sheet ? ctx.workbook.worksheets.getItem(sheet) : ctx.workbook.worksheets.getActiveWorksheet();
    ws.load("name");
    const loc = ws.freezePanes.getLocationOrNullObject();
    loc.load("address");
    await ctx.sync();
    const previous = loc.isNullObject ? null : loc.address;
    const r = Math.max(0, Math.round(rows || 0));
    const c = Math.max(0, Math.round(columns || 0));
    ws.freezePanes.unfreeze();
    if (r && c) ws.freezePanes.freezeAt(ws.getRangeByIndexes(0, 0, r, c));
    else if (r) ws.freezePanes.freezeRows(r);
    else if (c) ws.freezePanes.freezeColumns(c);
    await ctx.sync();
    setSnapshot({ kind: "freeze", sheetName: ws.name, previous, label: r || c ? `freeze panes on ${ws.name}` : `unfreeze ${ws.name}` });
    return r || c ? `Froze ${r} row(s) and ${c} column(s) on ${ws.name}` : `Unfroze panes on ${ws.name}`;
  });
}

/* ---------------------------------------------------------------- conditional formatting */

export interface CfArgs {
  rule: string;
  operator?: unknown;
  value1?: unknown;
  value2?: unknown;
  text?: unknown;
  formula?: unknown;
  rank?: unknown;
  percent?: unknown;
  bottom?: unknown;
  fill_color?: unknown;
  font_color?: unknown;
  bold?: unknown;
  min_color?: unknown;
  mid_color?: unknown;
  max_color?: unknown;
  bar_color?: unknown;
}

function applyCfFormat(fmt: Excel.ConditionalRangeFormat, a: CfArgs) {
  const fill = typeof a.fill_color === "string" && a.fill_color ? a.fill_color : !a.font_color && !a.bold ? "#FFC7CE" : "";
  if (fill) fmt.fill.color = fill;
  if (typeof a.font_color === "string" && a.font_color) fmt.font.color = a.font_color;
  else if (!a.fill_color && !a.bold) fmt.font.color = "#9C0006";
  if (a.bold === true) fmt.font.bold = true;
}

export async function addConditionalFormat(address: string, a: CfArgs, sheet?: string): Promise<string> {
  return run(async (ctx) => {
    const r = resolveRange(ctx, address, sheet);
    r.load("address");
    const sheetName = await sheetNameOf(ctx, r);
    const kind = String(a.rule ?? "").toLowerCase().replace(/[\s-]/g, "_");
    let cf: Excel.ConditionalFormat;
    let what: string;
    switch (kind) {
      case "cell_value": {
        const op = normalizeOperator(a.operator ?? "greater_than");
        if (!op) throw new ExcelUserError(`Unknown operator "${String(a.operator)}".`);
        const f1 = cfFormula(a.value1);
        if (!f1) throw new ExcelUserError("value1 is required for a cell-value rule.");
        cf = r.conditionalFormats.add(Excel.ConditionalFormatType.cellValue);
        cf.cellValue.rule = { formula1: f1, formula2: op === "Between" || op === "NotBetween" ? cfFormula(a.value2) : undefined, operator: op };
        applyCfFormat(cf.cellValue.format, a);
        what = `cells ${op} ${f1}`;
        break;
      }
      case "color_scale": {
        cf = r.conditionalFormats.add(Excel.ConditionalFormatType.colorScale);
        const mid = typeof a.mid_color === "string" && a.mid_color ? a.mid_color : null;
        cf.colorScale.criteria = {
          minimum: { type: Excel.ConditionalFormatColorCriterionType.lowestValue, formula: undefined, color: String(a.min_color || "#F8696B") },
          ...(mid ? { midpoint: { type: Excel.ConditionalFormatColorCriterionType.percentile, formula: "50", color: mid } } : {}),
          maximum: { type: Excel.ConditionalFormatColorCriterionType.highestValue, formula: undefined, color: String(a.max_color || "#63BE7B") },
        };
        what = "colour scale";
        break;
      }
      case "data_bar":
        cf = r.conditionalFormats.add(Excel.ConditionalFormatType.dataBar);
        cf.dataBar.positiveFormat.fillColor = String(a.bar_color || "#638EC6");
        what = "data bars";
        break;
      case "top_bottom": {
        const rank = Math.max(1, Math.round(Number(a.rank) || 10));
        const bottom = a.bottom === true;
        const pct = a.percent === true;
        cf = r.conditionalFormats.add(Excel.ConditionalFormatType.topBottom);
        cf.topBottom.rule = { rank, type: `${bottom ? "Bottom" : "Top"}${pct ? "Percent" : "Items"}` as Excel.ConditionalTopBottomCriterionType };
        applyCfFormat(cf.topBottom.format, a);
        what = `${bottom ? "bottom" : "top"} ${rank}${pct ? "%" : ""}`;
        break;
      }
      case "text_contains": {
        const text = String(a.text ?? "").trim();
        if (!text) throw new ExcelUserError("text is required for a text rule.");
        cf = r.conditionalFormats.add(Excel.ConditionalFormatType.containsText);
        cf.textComparison.rule = { operator: Excel.ConditionalTextOperator.contains, text };
        applyCfFormat(cf.textComparison.format, a);
        what = `text contains "${text}"`;
        break;
      }
      case "formula": {
        const f = String(a.formula ?? "").trim();
        if (!f) throw new ExcelUserError("formula is required for a formula rule.");
        cf = r.conditionalFormats.add(Excel.ConditionalFormatType.custom);
        cf.custom.rule.formula = f.startsWith("=") ? f : `=${f}`;
        applyCfFormat(cf.custom.format, a);
        what = `formula ${f}`;
        break;
      }
      default:
        throw new ExcelUserError(`Unknown rule "${a.rule}". Use cell_value, color_scale, data_bar, top_bottom, text_contains or formula.`);
    }
    cf.load("id");
    await ctx.sync();
    setSnapshot({ kind: "addedConditionalFormat", sheetName, address: r.address, id: cf.id, label: `conditional format ${stripSheet(r.address)}` });
    return `${what} on ${stripSheet(r.address)}`;
  });
}

/* ---------------------------------------------------------------- data validation */

export interface ValidationArgs {
  type: string;
  source?: unknown; // list: "A,B,C" or "=Lists!A2:A20"
  operator?: unknown;
  min?: unknown;
  max?: unknown;
  formula?: unknown;
  error_message?: unknown;
  clear?: boolean;
}

export async function setDataValidation(address: string, a: ValidationArgs, sheet?: string): Promise<string> {
  return run(async (ctx) => {
    const r = resolveRange(ctx, address, sheet);
    r.load("address");
    const dv = r.dataValidation;
    dv.load("type,rule");
    const sheetName = await sheetNameOf(ctx, r);
    let previous: Excel.DataValidationRule | null = null;
    try {
      previous = dv.type === Excel.DataValidationType.none || String(dv.type) === "Inconsistent" ? null : dv.rule;
    } catch {
      previous = null;
    }
    const type = String(a.type ?? "").toLowerCase().replace(/[\s-]/g, "_");
    let rule: Excel.DataValidationRule | null = null;
    if (!a.clear && type !== "clear") {
      const op = normalizeOperator(a.operator ?? (a.min !== undefined && a.max !== undefined ? "between" : a.min !== undefined ? ">=" : "<="));
      const basic = (): Excel.BasicDataValidation => {
        if (!op) throw new ExcelUserError(`Unknown operator "${String(a.operator)}".`);
        const f1 = cfFormula(op === "LessThanOrEqual" || op === "LessThan" ? a.max ?? a.min : a.min ?? a.max);
        if (!f1) throw new ExcelUserError("Give min and/or max.");
        // Validation spells the inclusive operators "…OrEqualTo" (conditional formats don't).
        const vop = (op === "GreaterThanOrEqual" ? "GreaterThanOrEqualTo" : op === "LessThanOrEqual" ? "LessThanOrEqualTo" : op) as Excel.DataValidationOperator;
        return { operator: vop, formula1: f1, formula2: op === "Between" || op === "NotBetween" ? cfFormula(a.max) : undefined };
      };
      switch (type) {
        case "list": {
          const src = Array.isArray(a.source) ? a.source.map(String).join(",") : String(a.source ?? "").trim();
          if (!src) throw new ExcelUserError("A list needs source values (\"A,B,C\") or a range (\"=Lists!A2:A20\").");
          rule = { list: { inCellDropDown: true, source: src } };
          break;
        }
        case "whole_number":
        case "whole":
        case "integer":
          rule = { wholeNumber: basic() };
          break;
        case "decimal":
        case "number":
          rule = { decimal: basic() };
          break;
        case "date": {
          // Dates go in as ISO text ("2026-01-31") or a formula ("=TODAY()"), not as quoted strings.
          const b = basic();
          const raw = (v: unknown) => (v === undefined || v === null || v === "" ? undefined : String(v).trim());
          const lo = raw(a.min);
          const hi = raw(a.max);
          const first = b.operator === "LessThanOrEqualTo" || b.operator === "LessThan" ? hi ?? lo : lo ?? hi;
          rule = { date: { operator: b.operator, formula1: first ?? "", formula2: b.formula2 === undefined ? undefined : hi } };
          break;
        }
        case "text_length":
          rule = { textLength: basic() };
          break;
        case "custom": {
          const f = String(a.formula ?? "").trim();
          if (!f) throw new ExcelUserError("A custom rule needs a formula.");
          rule = { custom: { formula: f.startsWith("=") ? f : `=${f}` } };
          break;
        }
        default:
          throw new ExcelUserError(`Unknown validation type "${a.type}". Use list, whole_number, decimal, date, text_length, custom or clear.`);
      }
    }
    dv.clear();
    if (rule) {
      dv.rule = rule;
      if (typeof a.error_message === "string" && a.error_message.trim()) dv.errorAlert = { message: a.error_message.trim(), showAlert: true, style: Excel.DataValidationAlertStyle.stop, title: "Invalid value" };
    }
    await ctx.sync();
    setSnapshot({ kind: "dataValidation", sheetName, address: r.address, rule: previous, label: `validation ${stripSheet(r.address)}` });
    return rule ? `${Object.keys(rule)[0]} validation on ${stripSheet(r.address)}` : `validation cleared on ${stripSheet(r.address)}`;
  });
}

/* ---------------------------------------------------------------- named ranges */

export async function addNamedRange(name: string, address: string, comment?: string, sheet?: string): Promise<string> {
  const n = name.trim();
  if (!validName(n)) throw new ExcelUserError(`"${name}" is not a valid Excel name (start with a letter, no spaces, not a cell address).`);
  return run(async (ctx) => {
    const r = resolveRange(ctx, address, sheet);
    r.load("address");
    const existing = ctx.workbook.names.getItemOrNullObject(n);
    await ctx.sync();
    if (!existing.isNullObject) throw new ExcelUserError(`The name "${n}" already exists.`);
    ctx.workbook.names.add(n, r, comment);
    await ctx.sync();
    setSnapshot({ kind: "addedName", name: n, label: `name ${n}` });
    invalidateOverview();
    return `${n} → ${r.address}`;
  });
}

/* ---------------------------------------------------------------- PivotTable */

export interface PivotArgs {
  source: string;
  rows: unknown[];
  columns?: unknown[];
  values: Array<{ field: unknown; summarize?: unknown }>;
  filters?: unknown[];
  sheetName?: string;
}

export async function createPivotTable(a: PivotArgs, sheet?: string): Promise<string> {
  if (!a.values?.length) throw new ExcelUserError("A PivotTable needs at least one value field.");
  // Read headers first so bad field names fail before a sheet is created.
  const headers = await run(async (ctx) => {
    const src = resolveRange(ctx, a.source, sheet);
    const head = src.getRow(0);
    head.load("values");
    src.load("rowCount");
    await ctx.sync();
    if (src.rowCount < 2) throw new ExcelUserError("The source range needs a header row and data.");
    return head.values[0].map((h) => String(h ?? ""));
  });
  const field = (f: unknown) => {
    const h = matchHeader(f, headers);
    if (!h) throw new ExcelUserError(`Field "${String(f)}" is not a header in ${a.source}. Headers: ${headers.filter(Boolean).join(", ")}`);
    return h;
  };
  const rows = (a.rows ?? []).map(field);
  const cols = (a.columns ?? []).map(field);
  const filters = (a.filters ?? []).map(field);
  const vals = a.values.map((v) => ({ field: field(v.field), how: normalizeAggregation(v.summarize) }));

  const sheetName = await addSheet(a.sheetName || "Pivot", true);
  return run(async (ctx) => {
    const src = resolveRange(ctx, a.source, sheet);
    const ws = ctx.workbook.worksheets.getItem(sheetName);
    const name = `ZeePivot${Date.now().toString(36)}`;
    const pt = ws.pivotTables.add(name, src, ws.getRange("A3"));
    rows.forEach((h) => pt.rowHierarchies.add(pt.hierarchies.getItem(h)));
    cols.forEach((h) => pt.columnHierarchies.add(pt.hierarchies.getItem(h)));
    filters.forEach((h) => pt.filterHierarchies.add(pt.hierarchies.getItem(h)));
    vals.forEach((v) => {
      const dh = pt.dataHierarchies.add(pt.hierarchies.getItem(v.field));
      dh.summarizeBy = v.how as Excel.AggregationFunction;
    });
    await ctx.sync();
    invalidateOverview();
    return `PivotTable on sheet "${sheetName}" (rows: ${rows.join(", ") || "–"}; columns: ${cols.join(", ") || "–"}; values: ${vals.map((v) => `${v.how} of ${v.field}`).join(", ")})`;
  });
}
