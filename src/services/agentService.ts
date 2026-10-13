/**
 * Agent mode: Gemini function-calling tools that act on the workbook.
 * Read tools run immediately; write tools are shown to the user for approval first
 * (unless auto-approve is on) and every write is recorded on the undo stack.
 */
import { groundedSearch, type FunctionDeclaration } from "./geminiService";
import type { ZeeSettings } from "./settingsService";
import {
  addSheet,
  createChart,
  createTableAt,
  formatRangeAt,
  formatSelectionForPrompt,
  getSelectionContext,
  getWorkbookOverview,
  readBlockAt,
  readRangeText,
  readRangeValues,
  setCells,
  type ChartKind,
} from "./excelService";
import {
  addConditionalFormat,
  addNamedRange,
  clearRange,
  copyRange,
  createPivotTable,
  deleteLines,
  filterRange,
  findCells,
  freezePanes,
  insertLines,
  replaceText,
  setDataValidation,
  sortRange,
  type CfArgs,
} from "./workbookOps";
import { runProductionQc } from "./petroActions";
import { readHeaderedAt, runEspFailureAnalysis, runFormulaAudit } from "./toolActions";
import { guessEspLogMapping, type EspLogMapping } from "../utils/espFailures";
import { fitArps, prepareDca, arpsTimeToRate, arpsCum } from "../utils/petro";
import { convertUnit, normalizeUnit, UNIT_HELP } from "../utils/units";
import { analyzeEsp, parseCurve } from "../utils/espAnalysis";
import { a1, quoteSheet } from "../utils/a1";
import { summarizeLargeRange } from "./dataTools";

export const AGENT_SYSTEM_PROMPT = `You are Zee AI, an expert Microsoft Excel agent running inside an Excel task pane. You can inspect and change the user's workbook with tools. The user is a petroleum engineer; production, well-test and ESP data are common.

HOW TO WORK
- If the user only asks a question, answer it; use read tools only when you need data.
- Before changing anything, read what you need (get_selection, read_range, get_workbook_overview, find_cells). Never guess addresses.
- Prefer live formulas over hard-coded results so the sheet stays correct when data changes. Use English function names and comma separators (e.g. =XLOOKUP(A2,Sheet2!A:A,Sheet2!B:B)). For unit conversions in cells you may use =ZEE.UNITS(value,"from","to").
- In set_cells every value is a string; strings starting with "=" are formulas. Write whole blocks in one call, not cell by cell.
- Don't overwrite existing data unless the user asked for it. Put new results in empty cells/columns (e.g. next to the data) or on a new sheet.
- Use the dedicated tools instead of workarounds: insert_rows_columns / delete_rows_columns, sort_range, filter_range, conditional_format, data_validation, create_pivot_table, freeze_panes, find_cells / replace_text, copy_range, clear_range, add_named_range.
- Petroleum: fit_decline_curve for Arps decline; production_qc for data-quality checks (writes a report sheet); esp_analysis for ESP performance (TDH, affinity laws, operating range, flags); esp_failure_analysis for run life / MTBF / failure causes from an ESP failure log; convert_units for unit columns.
- For very large ranges (tens of thousands of rows) use profile_range for statistics over every row plus a sample, instead of read_range.
- formula_audit checks a sheet's formulas (errors, #REF!, inconsistent formulas, hard-coded numbers) and writes a report sheet.
- For charts, pass a contiguous data range whose first row holds the series names.
- If the user attached images or PDFs (pump curves, well-test reports, scanned tables), read them carefully. When digitising a curve, sample enough points to follow its shape and say the values are read from the image (approximate).
- Keep changes minimal and focused. The user approves each batch of changes, and can undo them.
- When finished, reply with a short summary: what you changed (with addresses) or the answer. Use markdown, keep it brief.`;

const STR = { type: "STRING" };
const NUM = { type: "NUMBER" };
const BOOL = { type: "BOOLEAN" };
const SHEET = { type: "STRING", description: "Sheet name if not in the address; defaults to the active sheet." };

const BASE_TOOLS: FunctionDeclaration[] = [
  { name: "get_workbook_overview", description: "List all sheets with used ranges and header rows, plus tables and named ranges.", parameters: { type: "OBJECT", properties: {} } },
  { name: "get_selection", description: "Return the user's currently selected cells (values and formulas) with their addresses.", parameters: { type: "OBJECT", properties: {} } },
  {
    name: "read_range",
    description: "Read cell values/formulas from a range such as 'A1:D20', 'Sheet2!B:B' or a table column. Large ranges are truncated.",
    parameters: { type: "OBJECT", properties: { address: STR, sheet: SHEET }, required: ["address"] },
  },
  {
    name: "find_cells",
    description: "Find cells containing text (or a number) on one sheet or all sheets. Returns matching addresses (max 50).",
    parameters: {
      type: "OBJECT",
      properties: { text: STR, sheet: { type: "STRING", description: "Limit to this sheet; omit for all sheets." }, match_case: BOOL, whole_cell: { type: "BOOLEAN", description: "Match the entire cell content only." } },
      required: ["text"],
    },
  },
  {
    name: "set_cells",
    description: "Write values and/or formulas as a 2D block starting at the top-left cell of 'address'. Strings starting with '=' become formulas.",
    parameters: {
      type: "OBJECT",
      properties: {
        address: { type: "STRING", description: "Top-left cell (or range) to write to, e.g. 'E2' or 'Summary!A1'." },
        values: { type: "ARRAY", description: "Rows of cell values.", items: { type: "ARRAY", items: STR } },
        sheet: SHEET,
      },
      required: ["address", "values"],
    },
  },
  {
    name: "format_range",
    description: "Apply formatting to a range. Undoable.",
    parameters: {
      type: "OBJECT",
      properties: {
        address: STR,
        sheet: SHEET,
        bold: BOOL,
        fill_color: { type: "STRING", description: "Hex colour like #FFF2CC" },
        font_color: { type: "STRING", description: "Hex colour" },
        number_format: { type: "STRING", description: "Excel number format, e.g. 0.00, #,##0, yyyy-mm-dd, 0%" },
        borders: BOOL,
        autofit: BOOL,
        wrap: BOOL,
      },
      required: ["address"],
    },
  },
  {
    name: "insert_rows_columns",
    description: "Insert whole rows (above) or columns (to the left of) the given address. E.g. address '5:5' + kind rows inserts one row above row 5; 'C1' + kind columns + count 2 inserts two columns before C.",
    parameters: { type: "OBJECT", properties: { address: STR, kind: { type: "STRING", enum: ["rows", "columns"] }, count: NUM, sheet: SHEET }, required: ["address", "kind"] },
  },
  {
    name: "delete_rows_columns",
    description: "Delete the whole rows or columns that the address covers, e.g. '7:9' rows or 'D:E' columns. Undoable (contents are restored; formatting of deleted lines is not).",
    parameters: { type: "OBJECT", properties: { address: STR, kind: { type: "STRING", enum: ["rows", "columns"] }, sheet: SHEET }, required: ["address", "kind"] },
  },
  {
    name: "clear_range",
    description: "Clear contents, formats or both from a range.",
    parameters: { type: "OBJECT", properties: { address: STR, what: { type: "STRING", enum: ["contents", "formats", "all"] }, sheet: SHEET }, required: ["address", "what"] },
  },
  {
    name: "copy_range",
    description: "Copy a range to another place (same or other sheet). Relative formula references adjust like Excel copy/paste.",
    parameters: {
      type: "OBJECT",
      properties: {
        source: { type: "STRING", description: "e.g. 'Data!A1:D50'" },
        destination: { type: "STRING", description: "Top-left cell of the destination, e.g. 'Summary!A1'" },
        mode: { type: "STRING", enum: ["all", "values", "formulas", "formats"] },
        sheet: SHEET,
      },
      required: ["source", "destination"],
    },
  },
  {
    name: "sort_range",
    description: "Sort a range by one or more columns. Columns can be header names, sheet column letters, or 1-based positions within the range.",
    parameters: {
      type: "OBJECT",
      properties: {
        address: STR,
        by: { type: "ARRAY", items: { type: "OBJECT", properties: { column: STR, ascending: BOOL }, required: ["column"] } },
        has_headers: { type: "BOOLEAN", description: "First row is a header (default true)." },
        sheet: SHEET,
      },
      required: ["address", "by"],
    },
  },
  {
    name: "filter_range",
    description:
      "Apply an AutoFilter to a range on one column: a list of values to show, a condition like '>100' or '<>0' (optionally a second condition with AND/OR), or top/bottom N. Set clear=true to remove the sheet's filter.",
    parameters: {
      type: "OBJECT",
      properties: {
        address: { type: "STRING", description: "The whole data range including the header row." },
        column: { type: "STRING", description: "Header name, column letter or 1-based position." },
        values: { type: "ARRAY", items: STR },
        condition: STR,
        condition2: STR,
        operator: { type: "STRING", enum: ["and", "or"] },
        top: NUM,
        bottom: NUM,
        clear: BOOL,
        sheet: SHEET,
      },
      required: ["address"],
    },
  },
  {
    name: "replace_text",
    description: "Find and replace text in cells on one sheet (or all sheets if sheet is omitted).",
    parameters: { type: "OBJECT", properties: { find: STR, replace: STR, sheet: STR, match_case: BOOL, whole_cell: BOOL }, required: ["find", "replace"] },
  },
  {
    name: "freeze_panes",
    description: "Freeze the top N rows and/or left N columns of a sheet (0 and 0 = unfreeze).",
    parameters: { type: "OBJECT", properties: { rows: NUM, columns: NUM, sheet: STR }, required: ["rows", "columns"] },
  },
  {
    name: "conditional_format",
    description:
      "Add a conditional format. rule: cell_value (operator + value1 [+ value2]), color_scale (min_color/mid_color/max_color), data_bar (bar_color), top_bottom (rank, bottom, percent), text_contains (text), formula (formula relative to the top-left cell, e.g. '=$C2<100'). Default highlight is light red fill + dark red text.",
    parameters: {
      type: "OBJECT",
      properties: {
        address: STR,
        rule: { type: "STRING", enum: ["cell_value", "color_scale", "data_bar", "top_bottom", "text_contains", "formula"] },
        operator: { type: "STRING", enum: ["greater_than", "less_than", "between", "not_between", "equal_to", "not_equal_to", "greater_than_or_equal", "less_than_or_equal"] },
        value1: STR,
        value2: STR,
        text: STR,
        formula: STR,
        rank: NUM,
        bottom: BOOL,
        percent: BOOL,
        fill_color: STR,
        font_color: STR,
        bold: BOOL,
        min_color: STR,
        mid_color: STR,
        max_color: STR,
        bar_color: STR,
        sheet: SHEET,
      },
      required: ["address", "rule"],
    },
  },
  {
    name: "data_validation",
    description:
      "Restrict what can be typed in cells. type: list (source = 'A,B,C' or '=Lists!A2:A20' → drop-down), whole_number / decimal / date / text_length (operator + min/max), custom (formula), or clear.",
    parameters: {
      type: "OBJECT",
      properties: {
        address: STR,
        type: { type: "STRING", enum: ["list", "whole_number", "decimal", "date", "text_length", "custom", "clear"] },
        source: STR,
        operator: { type: "STRING", enum: ["between", "not_between", "equal_to", "not_equal_to", "greater_than", "less_than", "greater_than_or_equal", "less_than_or_equal"] },
        min: STR,
        max: STR,
        formula: STR,
        error_message: STR,
        sheet: SHEET,
      },
      required: ["address", "type"],
    },
  },
  {
    name: "add_named_range",
    description: "Create a workbook-level named range (e.g. OilRate → Data!C2:C500).",
    parameters: { type: "OBJECT", properties: { name: STR, address: STR, comment: STR, sheet: SHEET }, required: ["name", "address"] },
  },
  {
    name: "create_pivot_table",
    description: "Create a PivotTable on a new sheet from a source range with a header row. Fields are header names.",
    parameters: {
      type: "OBJECT",
      properties: {
        source: { type: "STRING", description: "Source data including headers, e.g. 'Data!A1:H5000' or a table name range." },
        rows: { type: "ARRAY", items: STR },
        columns: { type: "ARRAY", items: STR },
        values: {
          type: "ARRAY",
          items: { type: "OBJECT", properties: { field: STR, summarize: { type: "STRING", enum: ["sum", "average", "count", "max", "min", "stdev"] } }, required: ["field"] },
        },
        filters: { type: "ARRAY", items: STR },
        sheet_name: { type: "STRING", description: "Name for the new sheet (default 'Pivot')." },
        sheet: SHEET,
      },
      required: ["source", "rows", "values"],
    },
  },
  {
    name: "create_chart",
    description: "Create a chart from a contiguous data range (first row = series names, first column = categories/x).",
    parameters: {
      type: "OBJECT",
      properties: {
        data_address: STR,
        chart_type: { type: "STRING", enum: ["line", "column", "bar", "scatter", "scatterLines", "pie", "area"] },
        title: STR,
        sheet: SHEET,
      },
      required: ["data_address", "chart_type"],
    },
  },
  {
    name: "create_table",
    description: "Convert a range into a formatted Excel table.",
    parameters: { type: "OBJECT", properties: { address: STR, has_headers: BOOL, sheet: SHEET }, required: ["address"] },
  },
  { name: "add_sheet", description: "Add a new worksheet (name made unique automatically). Returns the final name.", parameters: { type: "OBJECT", properties: { name: STR }, required: ["name"] } },
  {
    name: "fit_decline_curve",
    description:
      "Fit Arps decline curves (exponential, harmonic, hyperbolic) to production data. Time can be Excel dates (days) or elapsed time. Returns parameters, R², and optional time-to-limit/EUR.",
    parameters: {
      type: "OBJECT",
      properties: {
        time_address: { type: "STRING", description: "Range with dates/times (no header)" },
        rate_address: { type: "STRING", description: "Range with rates, same length (no header)" },
        sheet: SHEET,
        economic_limit: { type: "NUMBER", description: "Optional economic-limit rate for EUR" },
      },
      required: ["time_address", "rate_address"],
    },
  },
  {
    name: "esp_failure_analysis",
    description:
      "Analyze an ESP failure / run-life log (one row per ESP run: well, install date, pull/fail date, optional status, model, failed component, cause). Computes run life, MTBF, failures by cause/component/model/well and a monthly trend; writes an 'ESP Failure Analysis' sheet with charts and returns the key numbers.",
    parameters: { type: "OBJECT", properties: { address: { type: "STRING", description: "The log including its header row." }, sheet: SHEET }, required: ["address"] },
  },
  {
    name: "profile_range",
    description:
      "Summarize a LARGE range without reading every cell: per-column statistics over all rows (count, blanks, min/max/mean, date span, top text values) plus ~50 sample rows (first, last, evenly spaced). Use for ranges too big for read_range.",
    parameters: { type: "OBJECT", properties: { address: { type: "STRING", description: "Range including the header row, e.g. 'A1:H250000' or 'A:H'." }, sheet: SHEET }, required: ["address"] },
  },
  {
    name: "formula_audit",
    description:
      "Audit every formula on a sheet: error values (#DIV/0!, #N/A…), broken #REF! references, formulas that differ from the rest of their column, typed values in formula columns and hard-coded numbers. Writes a 'Formula Audit' sheet (clickable cell links) and returns counts plus the first findings.",
    parameters: { type: "OBJECT", properties: { sheet: { type: "STRING", description: "Sheet to audit; defaults to the active sheet." } } },
  },
  {
    name: "production_qc",
    description:
      "Quality-check production data (missing, negative, zero, outlier and stuck-gauge values; date gaps and duplicates when the first column is dates). Writes a 'QC Report' sheet and returns the issue count.",
    parameters: { type: "OBJECT", properties: { address: { type: "STRING", description: "Data range including the header row." }, sheet: SHEET }, required: ["address"] },
  },
  {
    name: "convert_units",
    description: `Convert a column of numbers between units and write the results (header '<name> (<to unit>)'). Units: ${UNIT_HELP}.`,
    parameters: {
      type: "OBJECT",
      properties: {
        source: { type: "STRING", description: "One column of values, optionally with its header, e.g. 'C1:C200'." },
        from_unit: STR,
        to_unit: STR,
        destination: { type: "STRING", description: "Top-left cell for the output; defaults to the first empty column right of the data." },
        sheet: SHEET,
      },
      required: ["source", "from_unit", "to_unit"],
    },
  },
  {
    name: "esp_analysis",
    description:
      "Calculated ESP performance check over rows of well-test data: TDH, hydraulic HP, affinity-law correction to base frequency, head vs catalog curve (if given), operating range, motor load and flags. Each *_address is one column of the same length (no headers).",
    parameters: {
      type: "OBJECT",
      properties: {
        rate_address: { type: "STRING", description: "Liquid rate, bbl/d" },
        pip_address: { type: "STRING", description: "Pump intake pressure, psi" },
        pd_address: { type: "STRING", description: "Pump discharge pressure, psi" },
        hz_address: { type: "STRING", description: "Running frequency, Hz" },
        amps_address: { type: "STRING", description: "Optional motor current, A" },
        water_cut_address: { type: "STRING", description: "Optional water cut (% or fraction)" },
        sheet: SHEET,
        base_hz: { type: "NUMBER", description: "Catalog curve frequency (default 60)" },
        sg: { type: "NUMBER", description: "Fluid SG when no water cut (default 1.0)" },
        api: { type: "NUMBER", description: "Oil API, with water cut gives mixture SG" },
        stages: NUM,
        curve_address: { type: "STRING", description: "Optional pump curve: 2 columns, rate (bbl/d) and head at base_hz" },
        min_rate: NUM,
        max_rate: NUM,
        motor_amps: { type: "NUMBER", description: "Motor nameplate current" },
        min_pip: { type: "NUMBER", description: "Flag intake pressure below this (e.g. bubble point)" },
      },
      required: ["rate_address", "pip_address", "pd_address", "hz_address"],
    },
  },
];

const WEB_SEARCH_TOOL: FunctionDeclaration = {
  name: "web_search",
  description: "Search the web (Google) for current facts: prices, standards, equipment specs, news. Returns an answer with source links.",
  parameters: { type: "OBJECT", properties: { query: STR }, required: ["query"] },
};

/** Tools offered to the model; web_search only when enabled in Settings. */
export function agentTools(s: Pick<ZeeSettings, "webSearch">): FunctionDeclaration[] {
  return s.webSearch ? [...BASE_TOOLS, WEB_SEARCH_TOOL] : BASE_TOOLS;
}

/** @deprecated Use agentTools(settings). Kept for older imports. */
export const AGENT_TOOLS = BASE_TOOLS;

export const WRITE_TOOLS = new Set([
  "set_cells",
  "format_range",
  "create_chart",
  "create_table",
  "add_sheet",
  "insert_rows_columns",
  "delete_rows_columns",
  "clear_range",
  "copy_range",
  "sort_range",
  "filter_range",
  "replace_text",
  "freeze_panes",
  "conditional_format",
  "data_validation",
  "add_named_range",
  "create_pivot_table",
  "production_qc",
  "convert_units",
  "esp_failure_analysis",
  "formula_audit",
]);

export interface ToolCall {
  name: string;
  args: Record<string, unknown>;
  id?: string;
}

const s = (v: unknown) => (v === undefined || v === null || v === "" ? undefined : String(v));
const n = (v: unknown) => (typeof v === "number" && Number.isFinite(v) ? v : typeof v === "string" && v.trim() !== "" && Number.isFinite(Number(v)) ? Number(v) : undefined);

/** Human-readable description of a tool call (for the approval card and the action log). */
export function describeCall(c: ToolCall): string {
  const a = c.args;
  const where = (addr: unknown, sheet: unknown) => (sheet ? `${sheet}!${addr}` : String(addr));
  switch (c.name) {
    case "get_workbook_overview":
      return "Look at the workbook structure";
    case "get_selection":
      return "Read the selected cells";
    case "read_range":
      return `Read ${where(a.address, a.sheet)}`;
    case "profile_range":
      return `Summarize ${where(a.address, a.sheet)} (statistics + sample)`;
    case "find_cells":
      return `Find "${a.text}"${a.sheet ? ` on ${a.sheet}` : ""}`;
    case "set_cells": {
      const rows = Array.isArray(a.values) ? (a.values as unknown[][]) : [];
      const cols = rows.reduce((m, r) => Math.max(m, Array.isArray(r) ? r.length : 0), 0);
      const formulas = rows.flat().filter((v) => typeof v === "string" && v.startsWith("=")).length;
      return `Write ${rows.length}×${cols} cells at ${where(a.address, a.sheet)}${formulas ? ` (${formulas} formulas)` : ""}`;
    }
    case "format_range": {
      const bits = [a.bold ? "bold" : "", a.fill_color ? `fill ${a.fill_color}` : "", a.number_format ? `format ${a.number_format}` : "", a.borders ? "borders" : "", a.autofit ? "autofit" : ""].filter(Boolean);
      return `Format ${where(a.address, a.sheet)}${bits.length ? ` (${bits.join(", ")})` : ""}`;
    }
    case "insert_rows_columns":
      return `Insert ${a.count ?? ""} ${a.kind} at ${where(a.address, a.sheet)}`.replace("  ", " ");
    case "delete_rows_columns":
      return `Delete ${a.kind} ${where(a.address, a.sheet)}`;
    case "clear_range":
      return `Clear ${a.what ?? "contents"} of ${where(a.address, a.sheet)}`;
    case "copy_range":
      return `Copy ${a.source} to ${a.destination}${a.mode && a.mode !== "all" ? ` (${a.mode})` : ""}`;
    case "sort_range": {
      const by = Array.isArray(a.by) ? (a.by as Array<{ column?: unknown; ascending?: unknown }>).map((b) => `${b.column}${b.ascending === false ? " ↓" : " ↑"}`).join(", ") : "";
      return `Sort ${where(a.address, a.sheet)} by ${by}`;
    }
    case "filter_range":
      if (a.clear) return `Remove the filter on ${a.sheet ?? "the sheet"}`;
      return `Filter ${where(a.address, a.sheet)} on ${a.column}: ${Array.isArray(a.values) ? (a.values as unknown[]).join(", ") : a.top ? `top ${a.top}` : a.bottom ? `bottom ${a.bottom}` : `${a.condition ?? ""}${a.condition2 ? ` ${a.operator ?? "and"} ${a.condition2}` : ""}`}`;
    case "replace_text":
      return `Replace "${a.find}" with "${a.replace}"${a.sheet ? ` on ${a.sheet}` : " on all sheets"}`;
    case "freeze_panes":
      return Number(a.rows) || Number(a.columns) ? `Freeze ${a.rows ?? 0} row(s), ${a.columns ?? 0} column(s)` : "Unfreeze panes";
    case "conditional_format":
      return `Conditional format (${a.rule}${a.operator ? ` ${a.operator} ${a.value1 ?? ""}${a.value2 ? `–${a.value2}` : ""}` : ""}) on ${where(a.address, a.sheet)}`;
    case "data_validation":
      return `${a.type === "clear" ? "Clear validation" : `Validation (${a.type}${a.source ? `: ${a.source}` : ""})`} on ${where(a.address, a.sheet)}`;
    case "add_named_range":
      return `Name ${where(a.address, a.sheet)} as "${a.name}"`;
    case "create_pivot_table": {
      const vals = Array.isArray(a.values) ? (a.values as Array<{ field?: unknown; summarize?: unknown }>).map((v) => `${v.summarize ?? "sum"} of ${v.field}`).join(", ") : "";
      return `PivotTable from ${a.source}: rows ${Array.isArray(a.rows) ? a.rows.join(", ") : ""}; ${vals}`;
    }
    case "create_chart":
      return `Create ${a.chart_type} chart${a.title ? ` "${a.title}"` : ""} from ${where(a.data_address, a.sheet)}`;
    case "create_table":
      return `Make ${where(a.address, a.sheet)} a table`;
    case "add_sheet":
      return `Add sheet "${a.name}"`;
    case "fit_decline_curve":
      return `Fit decline curve (time ${a.time_address}, rate ${a.rate_address})`;
    case "production_qc":
      return `Production QC of ${where(a.address, a.sheet)} → new "QC Report" sheet`;
    case "esp_failure_analysis":
      return `ESP failure analysis of ${where(a.address, a.sheet)} → new "ESP Failure Analysis" sheet`;
    case "formula_audit":
      return `Formula audit of ${a.sheet ? `sheet "${a.sheet}"` : "the active sheet"} → new "Formula Audit" sheet`;
    case "convert_units":
      return `Convert ${where(a.source, a.sheet)} from ${a.from_unit} to ${a.to_unit}${a.destination ? ` → ${a.destination}` : ""}`;
    case "esp_analysis":
      return `ESP analysis (rate ${a.rate_address}, PIP ${a.pip_address}, Pd ${a.pd_address}, Hz ${a.hz_address})`;
    case "web_search":
      return `Search the web: "${a.query}"`;
    default:
      return c.name;
  }
}

/** Preview of the values a set_cells call would write (first rows), for the approval card. */
export function previewCall(c: ToolCall): string[][] | null {
  if (c.name !== "set_cells" || !Array.isArray(c.args.values)) return null;
  return (c.args.values as unknown[][]).slice(0, 6).map((r) => (Array.isArray(r) ? r.slice(0, 6).map((v) => String(v ?? "")) : []));
}

const firstColumn = (rows: unknown[][]) => rows.map((r) => r[0]);

/** Convert one column (first cell may be a header). Pure – unit-tested. */
export function convertColumn(values: unknown[], from: string, to: string): { out: Array<number | string>; converted: number; skipped: number; header: string | null } {
  normalizeUnit(from);
  normalizeUnit(to);
  const isNum = (v: unknown) => typeof v === "number" || (typeof v === "string" && v.trim() !== "" && Number.isFinite(Number(v.replace(/,/g, ""))));
  const header = values.length && !isNum(values[0]) && String(values[0] ?? "").trim() !== "" ? String(values[0]).trim() : null;
  const body = header === null ? values : values.slice(1);
  let converted = 0;
  let skipped = 0;
  const out = body.map((v) => {
    if (v === "" || v === null || v === undefined) return "";
    if (!isNum(v)) {
      skipped++;
      return "";
    }
    converted++;
    const x = convertUnit(typeof v === "number" ? v : Number(String(v).replace(/,/g, "")), from, to);
    return +x.toPrecision(10);
  });
  return { out: header === null ? out : [`${header} (${to})`, ...out], converted, skipped, header };
}

export interface ToolContext {
  settings: ZeeSettings;
  signal?: AbortSignal;
}

/** Execute one tool call against the workbook. Never throws: errors are returned to the model. */
export async function executeTool(c: ToolCall, ctx?: ToolContext): Promise<Record<string, unknown>> {
  const a = c.args;
  try {
    switch (c.name) {
      case "get_workbook_overview":
        return { overview: await getWorkbookOverview(6000, 0) };
      case "get_selection":
        return { selection: formatSelectionForPrompt(await getSelectionContext(2000)) };
      case "read_range":
        return { range: await readRangeText(String(a.address), s(a.sheet)) };
      case "profile_range": {
        const big = await summarizeLargeRange(String(a.address), s(a.sheet));
        return { summary: big.text };
      }
      case "find_cells": {
        const hits = await findCells(String(a.text ?? ""), { sheet: s(a.sheet), matchCase: a.match_case === true, wholeCell: a.whole_cell === true });
        return hits.length ? { matches: hits } : { matches: [], note: "No cells found." };
      }
      case "set_cells": {
        const rows = (Array.isArray(a.values) ? a.values : []) as unknown[][];
        const written = await setCells(String(a.address), rows.map((r) => (Array.isArray(r) ? r : [r])), s(a.sheet), "agent: write cells");
        return { ok: true, written };
      }
      case "format_range": {
        const addr = await formatRangeAt(
          String(a.address),
          {
            bold: typeof a.bold === "boolean" ? a.bold : undefined,
            fillColor: s(a.fill_color),
            fontColor: s(a.font_color),
            numberFormat: s(a.number_format),
            borders: a.borders === true,
            autofit: a.autofit === true,
            wrap: typeof a.wrap === "boolean" ? a.wrap : undefined,
          },
          s(a.sheet),
          true
        );
        return { ok: true, formatted: addr };
      }
      case "insert_rows_columns":
        return { ok: true, inserted: await insertLines(String(a.address), a.kind === "columns" ? "columns" : "rows", n(a.count), s(a.sheet)) };
      case "delete_rows_columns":
        return { ok: true, deleted: await deleteLines(String(a.address), a.kind === "columns" ? "columns" : "rows", s(a.sheet)) };
      case "clear_range": {
        const what = a.what === "formats" || a.what === "all" ? a.what : "contents";
        return { ok: true, cleared: await clearRange(String(a.address), what, s(a.sheet)) };
      }
      case "copy_range": {
        const mode = (["all", "values", "formulas", "formats"].includes(String(a.mode)) ? a.mode : "all") as "all" | "values" | "formulas" | "formats";
        return { ok: true, pasted: await copyRange(String(a.source), String(a.destination), mode, s(a.sheet)) };
      }
      case "sort_range": {
        const by = (Array.isArray(a.by) ? a.by : []) as Array<{ column: unknown; ascending?: boolean }>;
        return { ok: true, sorted: await sortRange(String(a.address), by, a.has_headers !== false, s(a.sheet)) };
      }
      case "filter_range":
        return { ok: true, result: await filterRange(String(a.address ?? "A1"), a.column, { ...a, clear: a.clear === true }, s(a.sheet)) };
      case "replace_text": {
        const count = await replaceText(String(a.find ?? ""), String(a.replace ?? ""), { sheet: s(a.sheet), matchCase: a.match_case === true, wholeCell: a.whole_cell === true });
        return { ok: true, replaced: count };
      }
      case "freeze_panes":
        return { ok: true, result: await freezePanes(n(a.rows) ?? 0, n(a.columns) ?? 0, s(a.sheet)) };
      case "conditional_format":
        return { ok: true, added: await addConditionalFormat(String(a.address), a as unknown as CfArgs, s(a.sheet)) };
      case "data_validation":
        return { ok: true, result: await setDataValidation(String(a.address), { ...a, type: String(a.type ?? "") }, s(a.sheet)) };
      case "add_named_range":
        return { ok: true, name: await addNamedRange(String(a.name ?? ""), String(a.address), s(a.comment), s(a.sheet)) };
      case "create_pivot_table":
        return {
          ok: true,
          pivot: await createPivotTable(
            {
              source: String(a.source),
              rows: Array.isArray(a.rows) ? a.rows : [],
              columns: Array.isArray(a.columns) ? a.columns : [],
              values: (Array.isArray(a.values) ? a.values : []) as Array<{ field: unknown; summarize?: unknown }>,
              filters: Array.isArray(a.filters) ? a.filters : [],
              sheetName: s(a.sheet_name),
            },
            s(a.sheet)
          ),
        };
      case "create_chart":
        return { ok: true, chart: await createChart(String(a.data_address), (s(a.chart_type) ?? "line") as ChartKind, s(a.title), s(a.sheet)) };
      case "create_table":
        return { ok: true, table: await createTableAt(String(a.address), a.has_headers !== false, s(a.sheet)) };
      case "add_sheet":
        return { ok: true, sheet: await addSheet(String(a.name ?? "Sheet"), false) };
      case "fit_decline_curve": {
        const tVals = firstColumn(await readRangeValues(String(a.time_address), s(a.sheet)));
        const qVals = firstColumn(await readRangeValues(String(a.rate_address), s(a.sheet)));
        const prep = prepareDca(tVals, qVals);
        const fits = fitArps(prep);
        const limit = n(a.economic_limit);
        return {
          points_used: prep.t.length,
          points_dropped: prep.dropped,
          time_origin: prep.t0,
          note: "di is per time unit of the input (per day if dates). t measured from the first point.",
          fits: fits.map((f) => {
            const tl = limit ? arpsTimeToRate(f, limit) : undefined;
            return {
              model: f.model,
              qi: +f.qi.toPrecision(6),
              di: +f.di.toPrecision(6),
              b: +f.b.toFixed(3),
              r2: +f.r2.toFixed(4),
              rmse: +f.rmse.toPrecision(4),
              ...(limit ? { time_to_limit: Number.isFinite(tl) ? +(tl as number).toFixed(1) : "never", eur_from_t0: Number.isFinite(tl) ? +arpsCum(f, tl as number).toPrecision(6) : "unbounded" } : {}),
            };
          }),
        };
      }
      case "production_qc": {
        const r = await runProductionQc({ address: String(a.address), sheet: s(a.sheet) });
        return { ok: true, report_sheet: r.sheet, issues: r.issues, columns_checked: r.columns };
      }
      case "esp_failure_analysis": {
        const sel = await readHeaderedAt(String(a.address), s(a.sheet), "failure log");
        const g = guessEspLogMapping(sel.headers);
        if (g.well === undefined || g.install === undefined || g.end === undefined)
          return { error: `Couldn't identify the well, install-date and pull/fail-date columns from the headers (${sel.headers.join(", ")}). Ask the user to use Petroleum → ESP failure log → Analyze, which lets them pick the columns.` };
        const r = await runEspFailureAnalysis(sel, g as EspLogMapping);
        const rep = r.report;
        return {
          ok: true,
          report_sheet: r.sheet,
          installs: rep.installs,
          failures: rep.failures,
          running: rep.running,
          mtbf_days: rep.mtbf,
          avg_failed_run_life_days: rep.avgFailedRunLife,
          top_causes: rep.byCause.slice(0, 5).map((c) => `${c.name}: ${c.failures}`),
          repeat_failure_wells: rep.repeatWells,
        };
      }
      case "formula_audit": {
        const r = await runFormulaAudit({}, s(a.sheet));
        return { ok: true, report_sheet: r.sheet, formulas: r.formulaCells, counts: r.counts, first_findings: r.issues.slice(0, 15).map((i) => `${i.address}: ${i.detail}`) };
      }
      case "convert_units": {
        const blk = await readBlockAt(String(a.source), s(a.sheet), 200_000);
        if ((blk.values[0]?.length ?? 0) !== 1) return { error: "convert_units takes a single column." };
        const conv = convertColumn(firstColumn(blk.values), String(a.from_unit), String(a.to_unit));
        let dest = s(a.destination);
        if (!dest) {
          // First empty column right of the data on the same rows.
          const used = await readBlockAt(`${a1(blk.rowIndex, blk.columnIndex)}:${a1(blk.rowIndex, blk.columnIndex + 200)}`, blk.sheetName, 10_000).catch(() => null);
          const width = used ? used.columnIndex - blk.columnIndex + (used.values[0]?.length ?? 1) : 1;
          dest = `${quoteSheet(blk.sheetName)}!${a1(blk.rowIndex, blk.columnIndex + Math.max(1, width))}`;
        }
        const written = await setCells(dest, conv.out.map((v) => [v]), undefined, `convert ${a.from_unit}→${a.to_unit}`);
        return { ok: true, written, converted: conv.converted, skipped_non_numeric: conv.skipped };
      }
      case "esp_analysis": {
        const col = async (addr: unknown) => (s(addr) ? firstColumn(await readRangeValues(String(addr), s(a.sheet))) : undefined);
        const [rate, pip, pd, hz, amps, wc] = await Promise.all([a.rate_address, a.pip_address, a.pd_address, a.hz_address, a.amps_address, a.water_cut_address].map(col));
        const len = Math.max(rate?.length ?? 0, pip?.length ?? 0, pd?.length ?? 0, hz?.length ?? 0);
        const curve = s(a.curve_address) ? parseCurve(await readRangeValues(String(a.curve_address), s(a.sheet))) : undefined;
        const inputs = Array.from({ length: len }, (_, i) => ({ rate: rate?.[i], pip: pip?.[i], pd: pd?.[i], hz: hz?.[i], amps: amps?.[i], waterCut: wc?.[i] }));
        const { rows, summary } = analyzeEsp(inputs, {
          baseHz: n(a.base_hz) ?? 60,
          sg: n(a.sg) ?? 1,
          api: n(a.api),
          stages: n(a.stages),
          curve: curve && curve.length >= 2 ? curve : undefined,
          minRate: n(a.min_rate),
          maxRate: n(a.max_rate),
          motorAmps: n(a.motor_amps),
          minPip: n(a.min_pip),
          degradationPct: 10,
        });
        const flagged = rows
          .map((r, i) => ({ row: i + 1, tdh_ft: r.tdh, head_vs_curve_pct: r.headDeviationPct, motor_load_pct: r.motorLoadPct, range: r.range, flags: r.flags }))
          .filter((r) => r.flags.length)
          .slice(0, 30);
        return { summary, flagged_rows: flagged, note: "Rows are 1-based positions within the given ranges. TDH from (Pd − PIP)/(0.433·SG)." };
      }
      case "web_search": {
        if (!ctx?.settings.webSearch) return { error: "Web search is turned off in Settings." };
        const r = await groundedSearch({ settings: ctx.settings, query: String(a.query ?? ""), signal: ctx.signal });
        return { answer: r.text, sources: r.sources.slice(0, 8) };
      }
      default:
        return { error: `Unknown tool ${c.name}` };
    }
  } catch (e) {
    return { error: (e as Error).message ?? String(e) };
  }
}
