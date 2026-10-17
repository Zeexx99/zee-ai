/**
 * Report-style tools: daily production report (hand-over to Word), ESP failure log + analysis,
 * formula audit and AI-suggested data validation. Results go to NEW sheets (undo deletes the sheet)
 * or use the undoable helpers in workbookOps.
 */
import { ExcelUserError, addSheet, readBlockAt, readSelectionBlock, run, type SelectionBlock } from "./excelService";
import { addConditionalFormat, setDataValidation } from "./workbookOps";
import { a1, columnLetter, quoteSheet, stripSheet } from "../utils/a1";
import { buildDailySummary, serialLabel, serialToIso, summaryTable, toMarkdownTable, toSerial, type DailyMapping, type DailySummary } from "../utils/dailyReport";
import { ESP_CAUSES, ESP_COMPONENTS, ESP_LOG_HEADERS, ESP_STATUSES, analyzeEspLog, type EspFailureReport, type EspLogMapping, type GroupStat } from "../utils/espFailures";
import { AUDIT_LABELS, auditFormulas, type AuditIssue } from "../utils/formulaAudit";
import { findViolations, profileColumns, type ColumnProfile, type ColumnRule } from "../utils/columnRules";
import { hasHeaderRow } from "./petroActions";
import { saveDailyReportPackage, type DailyReportPackage } from "../qol/sharedPackage";
import type { ProgressOptions } from "./excelService";

type Row = Array<string | number>;

/** Writes rows (values or formulas) at a cell of a sheet created by this module; bold/filled header optional. */
async function put(sheet: string, topLeft: string, rows: Row[], header = false): Promise<{ address: string; rowIndex: number; columnIndex: number; rows: number; cols: number }> {
  const width = Math.max(1, ...rows.map((r) => r.length));
  const data = rows.map((r) => [...r, ...Array<string | number>(width - r.length).fill("")]);
  return run(async (ctx) => {
    const ws = ctx.workbook.worksheets.getItem(sheet);
    const start = ws.getRange(topLeft);
    start.load("rowIndex,columnIndex");
    await ctx.sync();
    const target = ws.getRangeByIndexes(start.rowIndex, start.columnIndex, data.length, width);
    target.formulas = data;
    if (header) {
      target.getRow(0).format.font.bold = true;
      target.getRow(0).format.fill.color = "#EDEBFB";
    }
    target.load("address");
    await ctx.sync();
    return { address: target.address, rowIndex: start.rowIndex, columnIndex: start.columnIndex, rows: data.length, cols: width };
  });
}

/** Adds a chart on a sheet and returns a PNG of it (null where the host can't render images). */
async function chartWithImage(sheet: string, dataAddress: string, type: string, title: string, anchor: string, axis?: { x?: string; y?: string }): Promise<string | null> {
  return run(async (ctx) => {
    const ws = ctx.workbook.worksheets.getItem(sheet);
    const chart = ws.charts.add(type as Excel.ChartType, ws.getRange(stripSheet(dataAddress)), Excel.ChartSeriesBy.columns);
    chart.title.text = title;
    chart.legend.visible = false;
    chart.setPosition(anchor);
    chart.width = 520;
    chart.height = 300;
    if (axis?.x) {
      chart.axes.categoryAxis.title.text = axis.x;
      chart.axes.categoryAxis.title.visible = true;
    }
    if (axis?.y) {
      chart.axes.valueAxis.title.text = axis.y;
      chart.axes.valueAxis.title.visible = true;
    }
    await ctx.sync();
    try {
      const img = chart.getImage(1040);
      await ctx.sync();
      return img.value || null;
    } catch {
      return null;
    }
  });
}

async function finishSheet(sheet: string) {
  await run(async (ctx) => {
    const ws = ctx.workbook.worksheets.getItem(sheet);
    ws.getRange("A1").format.font.bold = true;
    ws.getRange("A1").format.font.size = 14;
    const used = ws.getUsedRangeOrNullObject(true);
    await ctx.sync();
    if (!used.isNullObject) used.getEntireColumn().format.autofitColumns();
    ws.getRange("A1").select();
    await ctx.sync();
  });
}

/* ================================================================ selection with header */

/** Tables read in chunks, so this is a memory guard rather than an Office limit (≈ 150k rows × 10 columns). */
export const BIG_TABLE_CELLS = 1_500_000;

export interface HeaderedSelection {
  block: SelectionBlock;
  headers: string[];
  rows: unknown[][];
  texts: string[][];
}

export async function readHeaderedSelection(what: string): Promise<HeaderedSelection> {
  return headered(await readSelectionBlock(BIG_TABLE_CELLS), what);
}

/** Same for an address (agent tools). */
export async function readHeaderedAt(address: string, sheet: string | undefined, what: string): Promise<HeaderedSelection> {
  return headered(await readBlockAt(address, sheet, BIG_TABLE_CELLS), what);
}

function headered(block: SelectionBlock, what: string): HeaderedSelection {
  if (block.values.length < 2) throw new ExcelUserError(`Select the ${what} including its header row (at least one data row).`);
  if (!hasHeaderRow(block.values)) throw new ExcelUserError(`The first selected row should be the column headers (e.g. Date, Well, Oil…). Select the ${what} from its header row.`);
  return {
    block,
    headers: block.values[0].map((h, j) => String(h ?? "").trim() || `Column ${columnLetter(block.columnIndex + j)}`),
    rows: block.values.slice(1),
    texts: block.text.slice(1),
  };
}

/* ================================================================ 1. daily production report */

export interface DailyReportResult {
  sheet: string;
  summary: DailySummary;
  pkg: DailyReportPackage;
  packageText: string;
  hasChart: boolean;
}

export async function runDailyReport(sel: HeaderedSelection, m: DailyMapping, reportDate?: number): Promise<DailyReportResult> {
  let summary: DailySummary;
  try {
    summary = buildDailySummary(sel.rows, m, reportDate);
  } catch (e) {
    throw new ExcelUserError((e as Error).message);
  }
  if (!summary.wells.length) throw new ExcelUserError("No well has data on the report date.");
  const h = sel.headers;
  const table = summaryTable(summary, { oil: h[m.oil], water: m.water !== undefined ? h[m.water] : undefined, gas: m.gas !== undefined ? h[m.gas] : undefined });
  const label = serialLabel(summary.reportDate);

  const sheet = await addSheet(`Daily ${serialToIso(summary.reportDate)}`);
  await put(sheet, "A1", [
    [`Daily production report – ${label}`],
    [`Source: ${sel.block.address} · ${summary.wells.length} wells${summary.previousDate !== null ? ` · compared with each well's previous reading (latest ${serialLabel(summary.previousDate)})` : ""}`],
  ]);
  const t = await put(sheet, "A4", table, true);
  await run(async (ctx) => {
    const ws = ctx.workbook.worksheets.getItem(sheet);
    const last = ws.getRangeByIndexes(t.rowIndex + t.rows - 1, t.columnIndex, 1, t.cols);
    last.format.font.bold = true;
    last.format.borders.getItem(Excel.BorderIndex.edgeTop).style = Excel.BorderLineStyle.continuous;
    await ctx.sync();
  });
  const hlRow = t.rowIndex + t.rows + 1;
  await put(sheet, a1(hlRow, 0), [["Highlights"], ...(summary.highlights.length ? summary.highlights : ["No notable changes."]).map((x) => [`• ${x}`])]);
  await run(async (ctx) => {
    ctx.workbook.worksheets.getItem(sheet).getRange(a1(hlRow, 0)).format.font.bold = true;
    await ctx.sync();
  });

  // Oil by well: columns A (well) and B (oil), without the totals row.
  const oilRange = `A${t.rowIndex + 1}:B${t.rowIndex + t.rows - 1}`;
  const png = await chartWithImage(sheet, oilRange, "ColumnClustered", `Oil by well – ${label}`, `${columnLetter(t.cols + 1)}4`, { y: h[m.oil] });
  await finishSheet(sheet);

  const pkg: DailyReportPackage = {
    zeeDailyReport: 1,
    createdAt: new Date().toISOString(),
    sheet,
    reportDate: serialToIso(summary.reportDate),
    previousDate: summary.previousDate !== null ? serialToIso(summary.previousDate) : null,
    dateLabel: label,
    table,
    highlights: summary.highlights,
    wells: summary.wells.length,
    missingWells: summary.missingWells,
    ...(png ? { chartPng: png } : {}),
  };
  const packageText = saveDailyReportPackage(pkg);
  return { sheet, summary, pkg, packageText, hasChart: !!png };
}

/** The dates found in the date column (latest first) – for the report-date picker. */
export function datesIn(sel: HeaderedSelection, dateCol: number): number[] {
  const set = new Set<number>();
  for (const r of sel.rows) {
    const d = toSerial(r[dateCol]);
    if (d !== null) set.add(d);
  }
  return Array.from(set).sort((a, b) => b - a);
}

/* ================================================================ 3. ESP failure log */

export async function createEspFailureLog(): Promise<string> {
  const sheet = await addSheet("ESP Failure Log");
  const N = 300;
  await run(async (ctx) => {
    const ws = ctx.workbook.worksheets.getItem(sheet);
    ws.getRange("A1:I1").values = [ESP_LOG_HEADERS];
    ws.getRange(`F2:F${N + 1}`).formulas = Array.from({ length: N }, (_, i) => [`=IF(C${i + 2}="","",IF(D${i + 2}="",TODAY(),D${i + 2})-C${i + 2})`]);
    ws.getRange(`C2:D${N + 1}`).numberFormat = Array.from({ length: N }, () => ["dd-mmm-yyyy", "dd-mmm-yyyy"]);
    ws.getRange(`F2:F${N + 1}`).numberFormat = Array.from({ length: N }, () => ["0"]);
    const list = (col: string, values: string[]) => {
      ws.getRange(`${col}2:${col}${N + 1}`).dataValidation.rule = { list: { inCellDropDown: true, source: values.join(",") } };
    };
    list("E", ESP_STATUSES);
    list("G", ESP_COMPONENTS);
    list("H", ESP_CAUSES);
    ws.getRange(`C2:D${N + 1}`).dataValidation.rule = { date: { operator: Excel.DataValidationOperator.greaterThan, formula1: "1990-01-01" } };
    const table = ws.tables.add(`A1:I${N + 1}`, true);
    table.style = "TableStyleMedium2";
    ws.freezePanes.freezeRows(1);
    ws.getRange("A1:I1").format.columnWidth = 110;
    ws.getRange("I1").format.columnWidth = 220;
    ws.getRange("A2").select();
    await ctx.sync();
  });
  return sheet;
}

export interface EspAnalysisResult {
  sheet: string;
  report: EspFailureReport;
  /** Markdown digest for the AI summary. */
  digest: string;
}

const groupRows = (title: string, g: GroupStat[], withInstalls: boolean): Row[] => [
  withInstalls ? [title, "Installs", "Failures", "MTBF (days)", "Avg run life of failures (days)"] : [title, "Failures", "Avg run life (days)"],
  ...g.slice(0, 40).map((x): Row => (withInstalls ? [x.name, x.installs, x.failures, x.mtbf ?? "", x.avgFailedRunLife ?? ""] : [x.name, x.failures, x.avgFailedRunLife ?? ""])),
];

export async function runEspFailureAnalysis(sel: HeaderedSelection, m: EspLogMapping): Promise<EspAnalysisResult> {
  const today = toSerial(new Date().toISOString().slice(0, 10)) as number;
  const report = analyzeEspLog(sel.rows, m, today);
  if (!report.installs) throw new ExcelUserError("No rows with a well name and a valid install date were found.");
  const sheet = await addSheet("ESP Failure Analysis");
  const kpi: Row[] = [
    ["Metric", "Value"],
    ["ESP installs (runs)", report.installs],
    ["Failures", report.failures],
    ["Still running", report.running],
    ["Total run days", report.totalRunDays],
    ["MTBF (days) = run days / failures", report.mtbf ?? "no failures"],
    ["Average run life of failed units (days)", report.avgFailedRunLife ?? ""],
    ["Median run life of failed units (days)", report.medianFailedRunLife ?? ""],
    ["Wells with 2+ failures", report.repeatWells.join(", ") || "none"],
  ];
  await put(sheet, "A1", [["ESP failure analysis"], [`Source: ${sel.block.address} · as of ${serialLabel(today)}${report.skipped ? ` · ${report.skipped} rows skipped (no well / install date)` : ""}`]]);
  let row = 3;
  const k = await put(sheet, a1(row, 0), kpi, true);
  row += k.rows + 1;

  const blocks: Array<{ rows: Row[]; chart?: { type: string; title: string; x: string } }> = [
    { rows: groupRows("Failure cause", report.byCause, false), chart: { type: "BarClustered", title: "Failures by cause", x: "Cause" } },
    { rows: groupRows("Failed component", report.byComponent, false) },
    { rows: groupRows("ESP model", report.byModel, true) },
    { rows: groupRows("Well", report.byWell, true) },
    { rows: [["Month", "Failures", "Installs"], ...report.monthly.map((x): Row => [x.month, x.failures, x.installs])], chart: { type: "ColumnClustered", title: "Failures per month", x: "Month" } },
  ];
  let chartRow = 3;
  for (const b of blocks) {
    if (b.rows.length < 2) continue;
    const w = await put(sheet, a1(row, 0), b.rows, true);
    if (b.chart) {
      await chartWithImage(sheet, `A${w.rowIndex + 1}:B${w.rowIndex + w.rows}`, b.chart.type, b.chart.title, `H${chartRow + 1}`, { x: b.chart.x, y: "Failures" });
      chartRow += 17;
    }
    row += w.rows + 1;
  }
  await finishSheet(sheet);

  const md = (title: string, rows: Row[]) => `**${title}**\n${toMarkdownTable(rows)}`;
  const digest = [
    md("Key figures", kpi),
    md("By failure cause", groupRows("Cause", report.byCause, false)),
    md("By failed component", groupRows("Component", report.byComponent, false)),
    md("By ESP model", groupRows("Model", report.byModel, true).slice(0, 16)),
    md("Wells with most failures", groupRows("Well", report.byWell.filter((w) => w.failures > 0), true).slice(0, 16)),
    md("Monthly trend (last 24 months)", [["Month", "Failures", "Installs"], ...report.monthly.slice(-24).map((x): Row => [x.month, x.failures, x.installs])]),
  ].join("\n\n");
  return { sheet, report, digest };
}

export const ESP_FAILURE_PROMPT = `Act as a senior artificial-lift (ESP) reliability engineer. Below are statistics computed from an ESP failure log (run life in days; MTBF = total run days of all units / failures).
Write a concise failure analysis in Markdown:
1. **Summary** – MTBF, failure count and what stands out (2–3 sentences).
2. **Top failure modes** – the main causes / components with numbers, and what usually drives them.
3. **Problem wells and models** – repeat failures, models with short run life.
4. **Trend** – is the failure rate improving or worsening? Mention seasonality only if the data shows it.
5. **Recommended actions** – specific, prioritised (e.g. scale inhibition, gas handling, VSD settings, cable/handling procedures, design review), and what extra data to log.
Use only the numbers given; say so where data is thin.`;

/* ================================================================ 8. formula audit */

export interface AuditResult {
  sheet: string;
  source: string;
  issues: AuditIssue[];
  counts: Record<string, number>;
  formulaCells: number;
}

export async function runFormulaAudit(opts: ProgressOptions = {}, sheetName?: string): Promise<AuditResult> {
  const MAX = 300_000;
  const data = await run(async (ctx) => {
    const ws = sheetName ? ctx.workbook.worksheets.getItem(sheetName) : ctx.workbook.worksheets.getActiveWorksheet();
    ws.load("name");
    const used = ws.getUsedRangeOrNullObject(true);
    used.load("address,rowIndex,columnIndex,rowCount,columnCount");
    await ctx.sync();
    if (used.isNullObject) throw new ExcelUserError("This sheet is empty.");
    const total = used.rowCount * used.columnCount;
    if (total > MAX) throw new ExcelUserError(`This sheet has ${total.toLocaleString()} cells – the audit handles up to ${MAX.toLocaleString()}.`);
    const block = { rowIndex: used.rowIndex, columnIndex: used.columnIndex, formulas: [] as unknown[][], formulasR1C1: [] as unknown[][], values: [] as unknown[][], valueTypes: [] as string[][] };
    const chunk = Math.max(1, Math.floor(15_000 / used.columnCount));
    for (let start = 0; start < used.rowCount; start += chunk) {
      if (opts.signal?.aborted) throw new ExcelUserError("Cancelled.");
      const n = Math.min(chunk, used.rowCount - start);
      opts.onProgress?.(start / used.rowCount, `Reading rows ${(start + 1).toLocaleString()}–${(start + n).toLocaleString()} of ${used.rowCount.toLocaleString()}…`);
      const r = ws.getRangeByIndexes(used.rowIndex + start, used.columnIndex, n, used.columnCount);
      r.load("formulas,formulasR1C1,values,valueTypes");
      await ctx.sync();
      block.formulas.push(...r.formulas);
      block.formulasR1C1.push(...r.formulasR1C1);
      block.values.push(...r.values);
      block.valueTypes.push(...(r.valueTypes as unknown as string[][]));
    }
    return { sheet: ws.name, address: used.address, block };
  });
  opts.onProgress?.(1, "Checking formulas…");
  const issues = auditFormulas(data.block);
  const formulaCells = data.block.formulas.reduce((n, r) => n + r.filter((f) => typeof f === "string" && f.startsWith("=")).length, 0);
  const counts: Record<string, number> = {};
  issues.forEach((i) => (counts[i.kind] = (counts[i.kind] ?? 0) + 1));

  const sheet = await addSheet("Formula Audit");
  await put(sheet, "A1", [
    [`Formula audit – ${data.sheet}`],
    [`${data.address} · ${formulaCells.toLocaleString()} formulas · ${issues.length.toLocaleString()} findings${issues.length >= 2000 ? " (first 2,000)" : ""} · click a cell address to jump there`],
  ]);
  const summary: Row[] = [["Finding", "Count"], ...Object.entries(AUDIT_LABELS).map(([k, label]): Row => [label, counts[k] ?? 0])];
  const s = await put(sheet, "A4", summary, true);
  const link = (i: AuditIssue) => `=HYPERLINK("#${quoteSheet(data.sheet).replace(/"/g, '""')}!${i.address}","${i.address}")`;
  const rows: Row[] = [
    ["Severity", "Cell", "Finding", "Detail", "Formula / value"],
    ...issues.map((i): Row => [i.severity, link(i), AUDIT_LABELS[i.kind], i.detail, i.formula.startsWith("=") ? `'${i.formula}` : i.formula]),
  ];
  if (rows.length === 1) rows.push(["–", "", "No problems found", "", ""]);
  await put(sheet, a1(s.rowIndex + s.rows + 1, 0), rows, true);
  await run(async (ctx) => {
    const ws = ctx.workbook.worksheets.getItem(sheet);
    ws.getRange("D:D").format.columnWidth = 380;
    ws.getRange("D:D").format.wrapText = true;
    await ctx.sync();
  });
  await finishSheet(sheet);
  return { sheet, source: data.sheet, issues, counts, formulaCells };
}

/* ================================================================ 10. data validation from AI */

export interface RulesTarget {
  sel: HeaderedSelection;
  profiles: ColumnProfile[];
}

export async function profileSelectionForRules(): Promise<RulesTarget> {
  const sel = await readHeaderedSelection("data table");
  return { sel, profiles: profileColumns(sel.headers, sel.rows, sel.texts) };
}

export interface RuleApplyResult {
  column: string;
  violations: number;
  /** First few bad cells, e.g. "C14". */
  examples: string[];
  highlighted: boolean;
  validated: boolean;
  note?: string;
}

const xlText = (s: string) => `"${s.replace(/"/g, '""')}"`;

/** Counts violations and (optionally) highlights them and adds Excel validation. Each change is undoable. */
export async function applyColumnRules(t: RulesTarget, rules: ColumnRule[], opts: { highlight: boolean; validate: boolean }): Promise<RuleApplyResult[]> {
  const b = t.sel.block;
  const out: RuleApplyResult[] = [];
  const firstRow = b.rowIndex + 1; // 0-based sheet row of the first data row
  const lastRow = b.rowIndex + t.sel.rows.length;
  for (const rule of rules) {
    const p = t.profiles.find((x) => x.name === rule.column);
    if (!p || rule.type === "none") continue;
    const col = b.columnIndex + p.index;
    const letter = columnLetter(col);
    const address = `${quoteSheet(b.sheetName)}!${letter}${firstRow + 1}:${letter}${lastRow + 1}`;
    const bad = findViolations(t.sel.rows, p.index, rule);
    const res: RuleApplyResult = { column: rule.column, violations: bad.length, examples: bad.slice(0, 5).map((i) => a1(firstRow + i, col)), highlighted: false, validated: false };
    const ref = `${letter}${firstRow + 1}`;
    if (opts.highlight) {
      let formula = "";
      if (rule.type === "number") {
        const tests = [rule.min !== null && rule.min !== undefined ? `${ref}<${rule.min}` : "", rule.max !== null && rule.max !== undefined ? `${ref}>${rule.max}` : ""].filter(Boolean);
        formula = `=AND(${ref}<>"",OR(NOT(ISNUMBER(${ref})),${tests.join(",")}))`;
      } else {
        formula = `=AND(${ref}<>"",NOT(OR(${(rule.allowed ?? []).map((a) => `TRIM(${ref})=${xlText(a)}`).join(",")})))`;
      }
      if (formula.length <= 250) {
        await addConditionalFormat(address, { rule: "formula", formula, fill_color: "#FFC7CE", font_color: "#9C0006" });
        res.highlighted = true;
      } else res.note = "Too many allowed values to highlight with a rule.";
    }
    if (opts.validate) {
      if (rule.type === "number") {
        await setDataValidation(address, { type: "decimal", min: rule.min ?? undefined, max: rule.max ?? undefined, error_message: `${rule.column}: ${rule.reason || "value outside the allowed range"}` });
        res.validated = true;
      } else {
        const values = (rule.allowed ?? []).filter((a) => !a.includes(","));
        if (values.length) {
          await setDataValidation(address, { type: "list", source: values.join(","), error_message: `${rule.column}: pick a value from the list.` });
          res.validated = true;
        }
      }
    }
    out.push(res);
  }
  return out;
}
