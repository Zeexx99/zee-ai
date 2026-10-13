/**
 * Data tools: duplicates / variants / outliers, gap filler, import fixer, sheet map, hard-coded numbers →
 * named inputs, agent dry-run previews and big-data summaries. Every change keeps an undo snapshot.
 */
import { ExcelUserError, addSheet, prepareFormatSnapshot, resolveRange, run, setSnapshot, type Snapshot } from "./excelService";
import { a1, columnLetter, quoteSheet, stripSheet } from "../utils/a1";
import { findDuplicateRows, findOutliers, findSpellingVariants, type DuplicateGroup, type Outlier, type VariantGroup } from "../utils/dupes";
import { fillGaps, type GapOptions, type GapResult } from "../utils/gaps";
import { fixCell, planColumn, type ColumnFixPlan } from "../utils/importFix";
import { buildSheetMap, sheetMapDigest, type SheetMap } from "../utils/sheetMap";
import { collectConstants, isValidName, replaceLiteral, suggestName, knownConstant, type ConstantUse } from "../utils/constants";
import { profileColumns, profilesForPrompt } from "../utils/columnRules";
import type { HeaderedSelection } from "./toolActions";
import { serialLabel } from "../utils/dailyReport";

const UNDO_CELLS = 50_000;

/* ================================================================ 18. duplicates, spelling variants, outliers */

export interface DupeScan {
  sel: HeaderedSelection;
  duplicates: DuplicateGroup[];
  variants: Array<{ col: number; header: string; groups: VariantGroup[] }>;
  outliers: Array<{ col: number; header: string; items: Outlier[] }>;
}

export function scanDupes(sel: HeaderedSelection): DupeScan {
  const profiles = profileColumns(sel.headers, sel.rows, sel.texts);
  const variants = profiles
    .filter((p) => p.kind === "text" || p.kind === "mixed")
    .map((p) => ({ col: p.index, header: p.name, groups: findSpellingVariants(sel.rows, p.index) }))
    .filter((v) => v.groups.length);
  const outliers = profiles
    .filter((p) => p.kind === "number")
    .map((p) => ({ col: p.index, header: p.name, items: findOutliers(sel.rows, p.index) }))
    .filter((o) => o.items.length);
  return { sel, duplicates: findDuplicateRows(sel.rows), variants, outliers };
}

export interface DupeChoice {
  highlight: boolean;
  unify: boolean;
  deleteDuplicates: boolean;
}

export interface DupeResult {
  highlighted: number;
  unified: number;
  deleted: number;
  undoable: boolean;
}

/** Applies the chosen actions. Highlight → unify → delete (each a separate undo step). */
export async function applyDupes(scan: DupeScan, c: DupeChoice): Promise<DupeResult> {
  const b = scan.sel.block;
  const sheet = b.sheetName;
  const cols = scan.sel.headers.length;
  const dataTop = b.rowIndex + 1; // 0-based sheet row of the first data row
  const res: DupeResult = { highlighted: 0, unified: 0, deleted: 0, undoable: true };

  if (c.highlight) {
    await run(async (ctx) => {
      const ws = ctx.workbook.worksheets.getItem(sheet);
      const area = ws.getRangeByIndexes(dataTop, b.columnIndex, scan.sel.rows.length, cols);
      area.load("rowCount,columnCount");
      await ctx.sync();
      const snap = prepareFormatSnapshot(area, sheet, "highlight duplicates & outliers");
      await ctx.sync();
      if (snap) setSnapshot(snap());
      else res.undoable = false;
      let n = 0;
      const paint = (r: number, col: number, width: number, color: string) => {
        ws.getRangeByIndexes(dataTop + r, b.columnIndex + col, 1, width).format.fill.color = color;
        n++;
      };
      scan.duplicates.forEach((g) => g.dupes.forEach((r) => paint(r, 0, cols, "#FFF2CC")));
      scan.variants.forEach((v) => v.groups.forEach((g) => g.variants.slice(1).forEach((x) => x.rows.forEach((r) => paint(r, v.col, 1, "#DDEBF7")))));
      scan.outliers.forEach((o) => o.items.forEach((x) => paint(x.row, o.col, 1, "#F8CBAD")));
      await ctx.sync();
      res.highlighted = n;
    });
  }

  if (c.unify && scan.variants.length) {
    await run(async (ctx) => {
      const ws = ctx.workbook.worksheets.getItem(sheet);
      const steps: Snapshot[] = [];
      const ranges = scan.variants.map((v) => {
        const r = ws.getRangeByIndexes(dataTop, b.columnIndex + v.col, scan.sel.rows.length, 1);
        r.load("address,formulas,numberFormat");
        return { v, r };
      });
      await ctx.sync();
      for (const { v, r } of ranges) {
        steps.push({ kind: "cells", sheetName: sheet, address: r.address, formulas: r.formulas, numberFormat: r.numberFormat, label: "" });
        const next = r.formulas.map((row) => row.slice());
        v.groups.forEach((g) =>
          g.variants.slice(1).forEach((x) =>
            x.rows.forEach((i) => {
              if (typeof next[i][0] === "string" && !(next[i][0] as string).startsWith("=")) {
                next[i][0] = g.canonical;
                res.unified++;
              }
            })
          )
        );
        r.formulas = next;
      }
      await ctx.sync();
      setSnapshot({ kind: "multi", steps, label: `unify spelling (${res.unified} cells)` });
    });
  }

  if (c.deleteDuplicates && scan.duplicates.length) {
    const rows = scan.duplicates.flatMap((g) => g.dupes).sort((x, y) => y - x);
    await run(async (ctx) => {
      const ws = ctx.workbook.worksheets.getItem(sheet);
      const all = ws.getRangeByIndexes(dataTop, b.columnIndex, scan.sel.rows.length, cols);
      all.load("address,formulas,numberFormat");
      await ctx.sync();
      const keep = scan.sel.rows.length * cols <= UNDO_CELLS;
      if (!keep) res.undoable = false;
      // Bottom-up so earlier row numbers stay valid. Only the table's columns shift up.
      let q = 0;
      for (const r of rows) {
        ws.getRangeByIndexes(dataTop + r, b.columnIndex, 1, cols).delete(Excel.DeleteShiftDirection.up);
        if (++q % 300 === 0) await ctx.sync();
      }
      await ctx.sync();
      if (keep) setSnapshot({ kind: "cells", sheetName: sheet, address: all.address, formulas: all.formulas, numberFormat: all.numberFormat, label: `delete ${rows.length} duplicate rows` });
      res.deleted = rows.length;
    });
  }
  return res;
}

/* ================================================================ 19. gap filler */

export interface GapRunResult {
  sheet: string;
  result: GapResult;
}

export async function runGapFill(sel: HeaderedSelection, o: GapOptions): Promise<GapRunResult> {
  const result = fillGaps(sel.rows, o);
  if (!result.rows.length) throw new ExcelUserError("No rows with a valid date were found in the date column.");
  const name = await addSheet(`Filled ${sel.block.sheetName}`);
  const headers = [...sel.headers, "Zee: filled"];
  const out = result.rows.map((r, i) => {
    const note = result.added[i] ? "added row" : result.filledCols[i].length ? `filled: ${result.filledCols[i].map((j) => sel.headers[j]).join(", ")}` : "";
    return [...r.map((v) => (v === null || v === undefined ? "" : v)), note];
  });
  const width = headers.length;
  const CHUNK = 4000;
  await run(async (ctx) => {
    const ws = ctx.workbook.worksheets.getItem(name);
    const h = ws.getRangeByIndexes(0, 0, 1, width);
    h.values = [headers];
    h.format.font.bold = true;
    h.format.fill.color = "#EDEBFB";
    for (let s = 0; s < out.length; s += CHUNK) {
      const part = out.slice(s, s + CHUNK);
      ws.getRangeByIndexes(1 + s, 0, part.length, width).values = part as (string | number | boolean)[][];
      await ctx.sync();
    }
    ws.getRangeByIndexes(1, o.dateCol, out.length, 1).numberFormat = out.map(() => ["dd-mmm-yyyy"]);
    // Mark added rows (blue) and filled cells (yellow) – at most 3,000 marks to keep it quick.
    let marks = 0;
    for (let i = 0; i < out.length && marks < 3000; i++) {
      if (result.added[i]) {
        ws.getRangeByIndexes(1 + i, 0, 1, width).format.fill.color = "#DDEBF7";
        marks++;
      }
      for (const j of result.filledCols[i]) {
        if (marks >= 3000) break;
        ws.getRangeByIndexes(1 + i, j, 1, 1).format.fill.color = "#FFF2CC";
        marks++;
      }
      if (marks % 500 === 0) await ctx.sync();
    }
    ws.freezePanes.freezeRows(1);
    ws.getRangeByIndexes(0, 0, out.length + 1, width).getEntireColumn().format.autofitColumns();
    ws.getRange("A1").select();
    await ctx.sync();
  });
  return { sheet: name, result };
}

/* ================================================================ 21. import fixer */

export interface ImportFixPreview {
  sel: HeaderedSelection;
  plans: ColumnFixPlan[];
  /** Cell formulas (to skip formula cells) and number formats as read for the preview. */
  formulas: unknown[][];
  numberFormats: string[][];
  changes: number;
  byKind: Record<"number" | "date" | "text", number>;
  otherUnits: number;
  samples: Array<{ address: string; before: string; after: string }>;
}

export async function previewImportFix(sel: HeaderedSelection): Promise<ImportFixPreview> {
  const b = sel.block;
  const rows = sel.rows.length;
  const cols = sel.headers.length;
  if (rows * cols > 200_000) throw new ExcelUserError("Select up to 200,000 cells at a time.");
  const { formulas, numberFormats } = await run(async (ctx) => {
    const r = ctx.workbook.worksheets.getItem(b.sheetName).getRangeByIndexes(b.rowIndex + 1, b.columnIndex, rows, cols);
    r.load("formulas,numberFormat");
    await ctx.sync();
    return { formulas: r.formulas, numberFormats: r.numberFormat as string[][] };
  });
  const plans = sel.headers.map((h, j) =>
    planColumn(
      j,
      h,
      sel.rows.map((r, i) => (typeof formulas[i][j] === "string" && (formulas[i][j] as string).startsWith("=") ? "" : r[j])),
      sel.texts.map((t) => t[j])
    )
  );
  const byKind = { number: 0, date: 0, text: 0 };
  let changes = 0;
  let otherUnits = 0;
  const samples: ImportFixPreview["samples"] = [];
  for (let i = 0; i < rows; i++)
    for (let j = 0; j < cols; j++) {
      const f = formulas[i][j];
      if (typeof f === "string" && f.startsWith("=")) continue;
      const fx = fixCell(sel.rows[i][j], plans[j]);
      if (fx.otherUnit) otherUnits++;
      if (!fx.changed) continue;
      changes++;
      byKind[plans[j].kind]++;
      if (samples.length < 6) samples.push({ address: a1(b.rowIndex + 1 + i, b.columnIndex + j), before: String(sel.rows[i][j]), after: plans[j].kind === "date" && typeof fx.value === "number" ? serialLabel(fx.value) : String(fx.value) });
    }
  return { sel, plans, formulas, numberFormats, changes, byKind, otherUnits, samples };
}

export async function applyImportFix(p: ImportFixPreview): Promise<{ changed: number; headers: number; undoable: boolean }> {
  const b = p.sel.block;
  const rows = p.sel.rows.length;
  const cols = p.sel.headers.length;
  let changed = 0;
  const headerChanges = p.plans.filter((x) => x.newHeader);
  const undoable = (rows + 1) * cols <= UNDO_CELLS;
  await run(async (ctx) => {
    const ws = ctx.workbook.worksheets.getItem(b.sheetName);
    const all = ws.getRangeByIndexes(b.rowIndex, b.columnIndex, rows + 1, cols);
    all.load("address,formulas,numberFormat");
    await ctx.sync();
    if (undoable) setSnapshot({ kind: "cells", sheetName: b.sheetName, address: all.address, formulas: all.formulas, numberFormat: all.numberFormat, label: `fix imported data ${stripSheet(all.address)}` });
    const header = all.formulas[0].slice();
    headerChanges.forEach((x) => (header[x.index] = x.newHeader as string));
    const nextF = p.formulas.map((r) => r.slice());
    const nextN = p.numberFormats.map((r) => r.slice());
    for (let i = 0; i < rows; i++)
      for (let j = 0; j < cols; j++) {
        const f = p.formulas[i][j];
        if (typeof f === "string" && f.startsWith("=")) continue;
        const fx = fixCell(p.sel.rows[i][j], p.plans[j]);
        if (!fx.changed) continue;
        nextF[i][j] = fx.value;
        if (fx.numberFormat) nextN[i][j] = fx.numberFormat;
        changed++;
      }
    ws.getRangeByIndexes(b.rowIndex, b.columnIndex, 1, cols).formulas = [header];
    const CHUNK = Math.max(1, Math.floor(20_000 / cols));
    for (let s = 0; s < rows; s += CHUNK) {
      const n = Math.min(CHUNK, rows - s);
      const r = ws.getRangeByIndexes(b.rowIndex + 1 + s, b.columnIndex, n, cols);
      r.numberFormat = nextN.slice(s, s + n);
      r.formulas = nextF.slice(s, s + n) as (string | number | boolean)[][];
      await ctx.sync();
    }
    all.getEntireColumn().format.autofitColumns();
    await ctx.sync();
  });
  return { changed, headers: headerChanges.length, undoable };
}

/* ================================================================ 28. explain this sheet */

export interface SheetMapResult {
  sheet: string;
  address: string;
  map: SheetMap;
  digest: string;
}

export async function runSheetMap(): Promise<SheetMapResult> {
  const MAX = 150_000;
  const data = await run(async (ctx) => {
    const ws = ctx.workbook.worksheets.getActiveWorksheet();
    ws.load("name");
    const used = ws.getUsedRangeOrNullObject(true);
    used.load("address,rowIndex,columnIndex,rowCount,columnCount");
    await ctx.sync();
    if (used.isNullObject) throw new ExcelUserError("This sheet is empty.");
    if (used.rowCount * used.columnCount > MAX) throw new ExcelUserError(`This sheet has ${(used.rowCount * used.columnCount).toLocaleString()} cells – the map handles up to ${MAX.toLocaleString()}.`);
    const formulas: unknown[][] = [];
    const formulasR1C1: unknown[][] = [];
    const chunk = Math.max(1, Math.floor(15_000 / used.columnCount));
    for (let s = 0; s < used.rowCount; s += chunk) {
      const r = ws.getRangeByIndexes(used.rowIndex + s, used.columnIndex, Math.min(chunk, used.rowCount - s), used.columnCount);
      r.load("formulas,formulasR1C1");
      await ctx.sync();
      formulas.push(...r.formulas);
      formulasR1C1.push(...r.formulasR1C1);
    }
    return { sheet: ws.name, address: used.address, block: { rowIndex: used.rowIndex, columnIndex: used.columnIndex, formulas, formulasR1C1 } };
  });
  const map = buildSheetMap(data.block);
  return { sheet: data.sheet, address: data.address, map, digest: sheetMapDigest(data.sheet, stripSheet(data.address), map) };
}

/** Writes the map as a "Guide" sheet (new sheet → undo deletes it). */
export async function writeSheetGuide(m: SheetMapResult): Promise<string> {
  const name = await addSheet(`Guide ${m.sheet}`);
  const roleLabel = { input: "Input (typed data)", calculated: "Calculated", mixed: "Mixed – formulas and typed values", label: "Labels / text", empty: "Empty" };
  const columns = [["Column", "Header", "Role", "Typed values", "Formulas"], ...m.map.columns.map((c) => [c.column, c.header, roleLabel[c.role], c.values, c.formulas])];
  const groups = [
    ["Column", "Header", "Rows", "Count", "Formula pattern (example)", "Depends on"],
    ...m.map.groups.map((g) => [g.column, g.header, `${g.firstRow}–${g.lastRow}`, g.count, `'${g.example}`, g.dependsOn.join(", ")]),
  ];
  await run(async (ctx) => {
    const ws = ctx.workbook.worksheets.getItem(name);
    ws.getRange("A1").values = [[`Sheet guide – ${m.sheet} (${stripSheet(m.address)})`]];
    ws.getRange("A1").format.font.bold = true;
    ws.getRange("A1").format.font.size = 14;
    const c = ws.getRangeByIndexes(2, 0, columns.length, 5);
    c.values = columns as (string | number)[][];
    c.getRow(0).format.font.bold = true;
    c.getRow(0).format.fill.color = "#EDEBFB";
    const top = 2 + columns.length + 1;
    const g = ws.getRangeByIndexes(top, 0, groups.length, 6);
    g.values = groups as (string | number)[][];
    g.getRow(0).format.font.bold = true;
    g.getRow(0).format.fill.color = "#EDEBFB";
    if (m.map.otherSheets.length || m.map.names.length)
      ws.getRangeByIndexes(top + groups.length + 1, 0, 2, 1).values = [[`Reads other sheets: ${m.map.otherSheets.join(", ") || "none"}`], [`Uses names: ${m.map.names.join(", ") || "none"}`]];
    ws.getRange("A:F").format.autofitColumns();
    ws.getRange("E:E").format.columnWidth = 320;
    await ctx.sync();
  });
  return name;
}

/* ================================================================ 29. hard-coded numbers → named inputs */

export interface ConstantsScan {
  sheet: string;
  address: string;
  uses: ConstantUse[];
  suggestions: Record<string, { name: string; note: string }>;
}

export async function scanConstants(sheetName?: string): Promise<ConstantsScan> {
  const data = await run(async (ctx) => {
    const ws = sheetName ? ctx.workbook.worksheets.getItem(sheetName) : ctx.workbook.worksheets.getActiveWorksheet();
    ws.load("name");
    const used = ws.getUsedRangeOrNullObject(true);
    used.load("address,rowIndex,columnIndex,rowCount,columnCount");
    const names = ctx.workbook.names;
    names.load("items/name");
    await ctx.sync();
    if (used.isNullObject) throw new ExcelUserError("This sheet is empty.");
    if (used.rowCount * used.columnCount > UNDO_CELLS)
      throw new ExcelUserError(`This sheet has ${(used.rowCount * used.columnCount).toLocaleString()} cells – up to ${UNDO_CELLS.toLocaleString()} can be converted with undo.`);
    used.load("formulas");
    await ctx.sync();
    return { sheet: ws.name, address: used.address, rowIndex: used.rowIndex, columnIndex: used.columnIndex, formulas: used.formulas, names: names.items.map((n) => n.name.toLowerCase()) };
  });
  const first = data.formulas[0] ?? [];
  const headerish = first.filter((v) => typeof v === "string" && v && !v.startsWith("=")).length >= Math.ceil(first.length / 2);
  const headers = first.map((v) => (headerish && typeof v === "string" ? v : ""));
  const uses = collectConstants({ rowIndex: data.rowIndex, columnIndex: data.columnIndex, formulas: data.formulas, headers });
  const taken = new Set(data.names);
  const suggestions: ConstantsScan["suggestions"] = {};
  for (const u of uses) {
    const name = suggestName(u.value, u.header, taken);
    taken.add(name.toLowerCase());
    suggestions[u.value] = { name, note: knownConstant(u.value)?.note ?? (u.header ? `used in "${u.header}"` : "") };
  }
  return { sheet: data.sheet, address: data.address, uses, suggestions };
}

export interface ConstantChoice {
  value: string;
  name: string;
  note: string;
}

/** Creates/extends the "Inputs" sheet, defines the names and rewrites the formulas. One undo step. */
export async function applyConstants(scan: ConstantsScan, chosen: ConstantChoice[]): Promise<{ names: number; formulas: number; inputsSheet: string }> {
  if (!chosen.length) throw new ExcelUserError("Pick at least one number.");
  const bad = chosen.find((c) => !isValidName(c.name));
  if (bad) throw new ExcelUserError(`"${bad.name}" isn't a valid Excel name (letters, digits, _ and .; must not look like a cell address such as BBL1).`);
  const dupName = chosen.find((c, i) => chosen.findIndex((x) => x.name.toLowerCase() === c.name.toLowerCase()) !== i);
  if (dupName) throw new ExcelUserError(`The name "${dupName.name}" is used twice.`);
  const INPUTS = "Inputs";
  let changedFormulas = 0;
  await run(async (ctx) => {
    const ws = ctx.workbook.worksheets.getItem(scan.sheet);
    const used = ws.getUsedRange(true);
    used.load("address,formulas,numberFormat");
    const names = ctx.workbook.names;
    names.load("items/name");
    let inputs = ctx.workbook.worksheets.getItemOrNullObject(INPUTS);
    inputs.load("name");
    await ctx.sync();
    if (used.address !== scan.address) throw new ExcelUserError("The sheet changed since the scan – run it again.");
    const existing = new Set(names.items.map((n) => n.name.toLowerCase()));
    const clash = chosen.find((c) => existing.has(c.name.toLowerCase()));
    if (clash) throw new ExcelUserError(`A name "${clash.name}" already exists in this workbook – choose another.`);

    const steps: Snapshot[] = [];
    let startRow = 1;
    if (inputs.isNullObject) {
      inputs = ctx.workbook.worksheets.add(INPUTS);
      inputs.getRange("A1:D1").values = [["Name", "Value", "Note", "Used in"]];
      inputs.getRange("A1:D1").format.font.bold = true;
      inputs.getRange("A1:D1").format.fill.color = "#EDEBFB";
      await ctx.sync();
      steps.push({ kind: "addedSheet", sheetName: INPUTS, label: "" });
    } else {
      const u = inputs.getUsedRangeOrNullObject(true);
      u.load("rowIndex,rowCount");
      await ctx.sync();
      startRow = u.isNullObject ? 1 : u.rowIndex + u.rowCount + 1;
      const area = inputs.getRangeByIndexes(startRow, 0, chosen.length, 4);
      area.load("address,formulas");
      await ctx.sync();
      steps.push({ kind: "cells", sheetName: INPUTS, address: area.address, formulas: area.formulas, label: "" });
    }
    const usesBy = new Map(scan.uses.map((u) => [u.value, u]));
    inputs.getRangeByIndexes(startRow, 0, chosen.length, 4).values = chosen.map((c) => {
      const u = usesBy.get(c.value);
      return [c.name, Number(c.value), c.note, `${scan.sheet}: ${u ? u.count : 0} formula${u && u.count === 1 ? "" : "s"}`];
    });
    chosen.forEach((c, i) => {
      ctx.workbook.names.add(c.name, `=${quoteSheet(INPUTS)}!$B$${startRow + i + 1}`, c.note || `Zee AI input (was ${c.value} in formulas)`);
      steps.push({ kind: "addedName", name: c.name, label: "" });
    });
    // Rewrite the formulas that use the chosen numbers.
    const next = used.formulas.map((row) =>
      row.map((f) => {
        if (typeof f !== "string" || !f.startsWith("=")) return f;
        let out = f;
        for (const c of chosen) out = replaceLiteral(out, c.value, c.name).formula;
        if (out !== f) changedFormulas++;
        return out;
      })
    );
    steps.push({ kind: "cells", sheetName: scan.sheet, address: used.address, formulas: used.formulas, numberFormat: used.numberFormat, label: "" });
    used.formulas = next as (string | number | boolean)[][];
    inputs.getRange("A:D").format.autofitColumns();
    await ctx.sync();
    setSnapshot({ kind: "multi", steps, label: `named inputs (${chosen.length})` });
  });
  return { names: chosen.length, formulas: changedFormulas, inputsSheet: INPUTS };
}

/* ================================================================ 30. agent dry run */

export interface DryRunCell {
  before: string;
  after?: string;
}

export interface DryRunPreview {
  title: string;
  /** First rows/columns of the affected cells, with real row numbers and column letters. */
  columns: string[];
  rows: Array<{ row: number; cells: DryRunCell[] }>;
  more: string;
}

const show = (v: unknown, text?: string) => {
  const t = text !== undefined && text !== "" ? text : v === null || v === undefined ? "" : String(v);
  return t.length > 28 ? `${t.slice(0, 26)}…` : t;
};

/** What a write tool would change, cell by cell (first 8 rows × 6 columns). null when there's nothing to show. */
export async function dryRunPreview(name: string, args: Record<string, unknown>): Promise<DryRunPreview | null> {
  const sheet = typeof args.sheet === "string" && args.sheet ? args.sheet : undefined;
  try {
    if (name === "set_cells" && Array.isArray(args.values) && typeof args.address === "string") {
      const vals = (args.values as unknown[]).map((r) => (Array.isArray(r) ? r : [r]));
      if (!vals.length) return null;
      const width = Math.max(1, ...vals.map((r) => r.length));
      return await run(async (ctx) => {
        const start = resolveRange(ctx, args.address as string, sheet).getCell(0, 0);
        start.load("rowIndex,columnIndex");
        const ws = start.worksheet;
        await ctx.sync();
        const R = Math.min(vals.length, 8);
        const C = Math.min(width, 6);
        const r = ws.getRangeByIndexes(start.rowIndex, start.columnIndex, R, C);
        r.load("values,text,formulas");
        await ctx.sync();
        const changed = vals.reduce((n, row) => n + row.length, 0);
        return {
          title: "Cells before → after",
          columns: Array.from({ length: C }, (_, j) => columnLetter(start.columnIndex + j)),
          rows: Array.from({ length: R }, (_, i) => ({
            row: start.rowIndex + i + 1,
            cells: Array.from({ length: C }, (_, j) => {
              const f = r.formulas[i][j];
              const before = typeof f === "string" && f.startsWith("=") ? f : show(r.values[i][j], r.text[i][j]);
              const nv = vals[i]?.[j];
              return { before, after: nv === undefined ? undefined : show(nv) };
            }),
          })),
          more: vals.length > R || width > C ? `Showing ${R} × ${C} of ${vals.length} × ${width} cells (${changed.toLocaleString()} in total).` : "",
        };
      });
    }
    if ((name === "clear_range" || name === "delete_rows_columns") && typeof args.address === "string") {
      return await run(async (ctx) => {
        const target = resolveRange(ctx, args.address as string, sheet);
        const ws = target.worksheet;
        const used = ws.getUsedRangeOrNullObject(true);
        used.load("address");
        await ctx.sync();
        if (used.isNullObject) return null;
        const inter = target.getIntersectionOrNullObject(used);
        inter.load("rowIndex,columnIndex,rowCount,columnCount");
        await ctx.sync();
        if (inter.isNullObject) return { title: "Nothing in this range has data", columns: [], rows: [], more: "" };
        const R = Math.min(inter.rowCount, 6);
        const C = Math.min(inter.columnCount, 6);
        const r = ws.getRangeByIndexes(inter.rowIndex, inter.columnIndex, R, C);
        r.load("values,text,formulas");
        await ctx.sync();
        return {
          title: name === "clear_range" ? (args.what === "formats" ? "Formatting removed from (values stay)" : "Content that will be cleared") : "Content in the rows/columns that will be deleted",
          columns: Array.from({ length: C }, (_, j) => columnLetter(inter.columnIndex + j)),
          rows: Array.from({ length: R }, (_, i) => ({
            row: inter.rowIndex + i + 1,
            cells: Array.from({ length: C }, (_, j) => {
              const f = r.formulas[i][j];
              return { before: typeof f === "string" && f.startsWith("=") ? f : show(r.values[i][j], r.text[i][j]) };
            }),
          })),
          more: inter.rowCount > R || inter.columnCount > C ? `Showing ${R} × ${C} of ${inter.rowCount.toLocaleString()} × ${inter.columnCount} cells with data.` : "",
        };
      });
    }
  } catch {
    return null; // a preview must never block the approval
  }
  return null;
}

/* ================================================================ 33. big-data mode */

export interface BigDataSummary {
  address: string;
  rows: number;
  cols: number;
  text: string;
  sampled: number;
}

/**
 * Reads a large block in chunks and returns column statistics over ALL rows plus a spread sample
 * (first, last and evenly spaced rows) – instead of only the first rows.
 */
export async function summarizeLargeRange(address: string | null, sheet?: string, sampleRows = 50, maxCells = 3_000_000): Promise<BigDataSummary> {
  return run(async (ctx) => {
    let target: Excel.Range;
    if (address) target = resolveRange(ctx, address, sheet);
    else target = ctx.workbook.getSelectedRange();
    const ws = target.worksheet;
    ws.load("name");
    const used = ws.getUsedRangeOrNullObject(true);
    used.load("address");
    await ctx.sync();
    if (used.isNullObject) throw new ExcelUserError("The range is empty.");
    const r = target.getIntersectionOrNullObject(used);
    r.load("address,rowIndex,columnIndex,rowCount,columnCount");
    await ctx.sync();
    if (r.isNullObject) throw new ExcelUserError("The range is empty.");
    const cols = Math.min(r.columnCount, 60);
    if (r.rowCount * cols > maxCells) throw new ExcelUserError(`Too large to summarise (${(r.rowCount * cols).toLocaleString()} cells; max ${maxCells.toLocaleString()}).`);
    const values: unknown[][] = [];
    const chunk = Math.max(1, Math.floor(20_000 / cols));
    for (let s = 0; s < r.rowCount; s += chunk) {
      const part = ws.getRangeByIndexes(r.rowIndex + s, r.columnIndex, Math.min(chunk, r.rowCount - s), cols);
      part.load("values");
      await ctx.sync();
      values.push(...part.values);
    }
    // Sample: first 15, last 15 and evenly spaced rows in between (sorted, unique).
    const n = values.length;
    const pick = new Set<number>();
    for (let i = 0; i < Math.min(15, n); i++) pick.add(i);
    for (let i = Math.max(0, n - 15); i < n; i++) pick.add(i);
    const middle = Math.max(0, sampleRows - pick.size);
    for (let k = 1; k <= middle; k++) pick.add(Math.floor((k * n) / (middle + 1)));
    const sample = Array.from(pick).sort((a, b) => a - b);
    const texts = ws.getRangeByIndexes(r.rowIndex, r.columnIndex, 2, cols);
    texts.load("text,numberFormat");
    await ctx.sync();
    const first = values[0] ?? [];
    const header = first.filter((v) => typeof v === "string" && v.trim()).length >= Math.ceil(cols / 2);
    const headers = first.map((v, j) => (header ? String(v ?? "").trim() : "") || columnLetter(r.columnIndex + j));
    const body = header ? values.slice(1) : values;
    const isDate = (j: number) => /[dmy]{2}|yyyy|mmm/i.test(String((texts.numberFormat as string[][])[header ? 1 : 0]?.[j] ?? ""));
    const profiles = profileColumns(headers, body);
    const iso = (v: number | null) => (v === null ? "–" : new Date(Date.UTC(1899, 11, 30) + v * 86_400_000).toISOString().slice(0, 10));
    const stats = profilesForPrompt(profiles)
      .split("\n")
      .map((line, j) => (isDate(j) && profiles[j]?.kind === "number" ? `- "${profiles[j].name}": dates from ${iso(profiles[j].min)} to ${iso(profiles[j].max)}, n=${profiles[j].count}` : line))
      .join("\n");
    const letters = Array.from({ length: cols }, (_, j) => columnLetter(r.columnIndex + j));
    const cell = (v: unknown, j: number) => {
      if (typeof v === "number" && isDate(j)) return iso(v);
      const t = v === null || v === undefined ? "" : String(v);
      return (t.length > 40 ? `${t.slice(0, 38)}…` : t).replace(/\|/g, "\\|");
    };
    const table = [`| # | ${letters.join(" | ")} |`, `|${["#", ...letters].map(() => "---").join("|")}|`, ...sample.map((i) => `| ${r.rowIndex + i + 1} | ${values[i].map(cell).join(" | ")} |`)].join("\n");
    const text = `[Excel selection – big-data mode]
Sheet "${ws.name}", range ${stripSheet(r.address)} (${r.rowCount.toLocaleString()} rows × ${r.columnCount} columns${r.columnCount > cols ? `, first ${cols} columns summarised` : ""}).
Too large to send whole, so here are statistics over ALL ${body.length.toLocaleString()} data rows${header ? ` (headers in row ${r.rowIndex + 1})` : ""}, then ${sample.length} sample rows (first, last and evenly spaced). Base answers on the statistics; say when an exact answer needs every row (the user can then use a formula or the agent).
Column statistics:
${stats}

Sample rows:
${table}`;
    return { address: r.address, rows: r.rowCount, cols: r.columnCount, text, sampled: sample.length };
  });
}

/** Row count of the current selection (clipped to the used range) – cheap check before big-data mode. */
export async function selectionRowCount(): Promise<number> {
  return run(async (ctx) => {
    const sel = ctx.workbook.getSelectedRange();
    const used = sel.worksheet.getUsedRangeOrNullObject(true);
    used.load("address");
    await ctx.sync();
    if (used.isNullObject) return 0;
    const i = sel.getIntersectionOrNullObject(used);
    i.load("rowCount");
    await ctx.sync();
    return i.isNullObject ? 0 : i.rowCount;
  });
}

/** For the approval card label. */
export const quoteForDisplay = (s: string) => quoteSheet(s);
