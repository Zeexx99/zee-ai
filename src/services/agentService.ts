/**
 * Agent mode: Gemini function-calling tools that act on the workbook.
 * Read tools run immediately; write tools are shown to the user for approval first
 * (unless auto-approve is on) and every write is recorded on the undo stack.
 */
import type { FunctionDeclaration } from "./geminiService";
import {
  addSheet,
  createChart,
  createTableAt,
  formatRangeAt,
  formatSelectionForPrompt,
  getSelectionContext,
  getWorkbookOverview,
  readRangeText,
  readRangeValues,
  setCells,
  type ChartKind,
} from "./excelService";
import { fitArps, prepareDca, arpsTimeToRate, arpsCum } from "../utils/petro";

export const AGENT_SYSTEM_PROMPT = `You are Zee AI, an expert Microsoft Excel agent running inside an Excel task pane. You can inspect and change the user's workbook with tools.

HOW TO WORK
- If the user only asks a question, answer it; use read tools only when you need data.
- Before changing anything, read what you need (get_selection, read_range, get_workbook_overview). Never guess addresses.
- Prefer live formulas over hard-coded results so the sheet stays correct when data changes. Use English function names and comma separators (e.g. =XLOOKUP(A2,Sheet2!A:A,Sheet2!B:B)).
- In set_cells every value is a string; strings starting with "=" are formulas. Write whole blocks in one call, not cell by cell.
- Don't overwrite existing data unless the user asked for it. Put new results in empty cells/columns (e.g. next to the data) or on a new sheet.
- For charts, pass a contiguous data range whose first row holds the series names.
- Keep changes minimal and focused. The user approves each batch of changes, and can undo them.
- For decline-curve requests use fit_decline_curve, then write the results/forecast if the user wants them.
- When finished, reply with a short summary: what you changed (with addresses) or the answer. Use markdown, keep it brief.`;

const STR = { type: "STRING" };
const BOOL = { type: "BOOLEAN" };

export const AGENT_TOOLS: FunctionDeclaration[] = [
  { name: "get_workbook_overview", description: "List all sheets with used ranges and header rows, plus tables and named ranges.", parameters: { type: "OBJECT", properties: {} } },
  { name: "get_selection", description: "Return the user's currently selected cells (values and formulas) with their addresses.", parameters: { type: "OBJECT", properties: {} } },
  {
    name: "read_range",
    description: "Read cell values/formulas from a range such as 'A1:D20', 'Sheet2!B:B' or a table column. Large ranges are truncated.",
    parameters: { type: "OBJECT", properties: { address: STR, sheet: { type: "STRING", description: "Sheet name if not in the address; defaults to the active sheet." } }, required: ["address"] },
  },
  {
    name: "set_cells",
    description: "Write values and/or formulas as a 2D block starting at the top-left cell of 'address'. Strings starting with '=' become formulas.",
    parameters: {
      type: "OBJECT",
      properties: {
        address: { type: "STRING", description: "Top-left cell (or range) to write to, e.g. 'E2' or 'Summary!A1'." },
        values: { type: "ARRAY", description: "Rows of cell values.", items: { type: "ARRAY", items: STR } },
        sheet: STR,
      },
      required: ["address", "values"],
    },
  },
  {
    name: "format_range",
    description: "Apply formatting to a range.",
    parameters: {
      type: "OBJECT",
      properties: {
        address: STR,
        sheet: STR,
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
    name: "create_chart",
    description: "Create a chart from a contiguous data range (first row = series names, first column = categories/x).",
    parameters: {
      type: "OBJECT",
      properties: {
        data_address: STR,
        chart_type: { type: "STRING", enum: ["line", "column", "bar", "scatter", "scatterLines", "pie", "area"] },
        title: STR,
        sheet: STR,
      },
      required: ["data_address", "chart_type"],
    },
  },
  {
    name: "create_table",
    description: "Convert a range into a formatted Excel table.",
    parameters: { type: "OBJECT", properties: { address: STR, has_headers: BOOL, sheet: STR }, required: ["address"] },
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
        sheet: STR,
        economic_limit: { type: "NUMBER", description: "Optional economic-limit rate for EUR" },
      },
      required: ["time_address", "rate_address"],
    },
  },
];

export const WRITE_TOOLS = new Set(["set_cells", "format_range", "create_chart", "create_table", "add_sheet"]);

export interface ToolCall {
  name: string;
  args: Record<string, unknown>;
  id?: string;
}

const s = (v: unknown) => (v === undefined || v === null ? undefined : String(v));

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
    case "create_chart":
      return `Create ${a.chart_type} chart${a.title ? ` "${a.title}"` : ""} from ${where(a.data_address, a.sheet)}`;
    case "create_table":
      return `Make ${where(a.address, a.sheet)} a table`;
    case "add_sheet":
      return `Add sheet "${a.name}"`;
    case "fit_decline_curve":
      return `Fit decline curve (time ${a.time_address}, rate ${a.rate_address})`;
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

/** Execute one tool call against the workbook. Never throws: errors are returned to the model. */
export async function executeTool(c: ToolCall): Promise<Record<string, unknown>> {
  const a = c.args;
  try {
    switch (c.name) {
      case "get_workbook_overview":
        return { overview: await getWorkbookOverview(6000, 0) };
      case "get_selection":
        return { selection: formatSelectionForPrompt(await getSelectionContext(2000)) };
      case "read_range":
        return { range: await readRangeText(String(a.address), s(a.sheet)) };
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
          s(a.sheet)
        );
        return { ok: true, formatted: addr, note: "Formatting changes are not covered by Zee AI undo." };
      }
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
        const limit = typeof a.economic_limit === "number" ? a.economic_limit : undefined;
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
      default:
        return { error: `Unknown tool ${c.name}` };
    }
  } catch (e) {
    return { error: (e as Error).message ?? String(e) };
  }
}
