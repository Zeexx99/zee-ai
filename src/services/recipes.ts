/**
 * Recipes: record a sequence of Zee AI actions once, replay it next month on new data with one click.
 * Pure part (types, labels, storage, sharing, column mapping) – unit tested. The runner is recipeRunner.ts.
 *
 * Where a step runs is stored as a Target:
 *  - "data"     – the recipe's data sheet. At replay the user picks it (default: the active sheet), so a
 *                 recipe recorded on "Oct production" runs on "Nov production" with the same layout.
 *  - "step:k"   – the sheet that step k created (e.g. resample → chart the resampled table).
 *  - "fixed"    – a sheet that stays the same (e.g. a lookup / master sheet), by name.
 * Columns and the first row are kept; rows run to the end of the data so more or fewer rows are fine.
 * Columns that tools pick by name (date, well, …) are stored by header text, not by position.
 */
import type { CleanOptions } from "../utils/clean";
import type { GapMethod } from "../utils/gaps";
import type { Agg, Period, PeriodFill } from "../utils/resample";
import type { DupeChoice } from "./dataTools";
import type { AggregateFn, ChartKind } from "./excelService";
import type { NodalInput } from "../utils/nodal";
import type { DateMatch, MergeHow } from "../utils/merge";

export type TargetSheet = { kind: "data" } | { kind: "step"; index: number } | { kind: "fixed"; name: string };

export interface Target {
  sheet: TargetSheet;
  /** 0-based first row and column span of the selection when recorded. */
  startRow: number;
  startCol: number;
  endCol: number;
  /** Single-cell selection (e.g. "insert at the active cell"): replay selects just that cell. */
  single?: boolean;
}

export type StepAction =
  | { kind: "clean"; options: CleanOptions }
  | { kind: "removeEmpty"; lines: "rows" | "columns" }
  | { kind: "formatTable" }
  | { kind: "autofit" }
  | { kind: "fixImport" }
  | { kind: "fillGaps"; dateHeader: string; groupHeader?: string; method: GapMethod; maxGapSteps: number }
  | { kind: "resample"; dateHeader: string; groupHeader?: string; period: Period; aggs: Record<string, Agg>; fill: PeriodFill; countDays: boolean }
  | { kind: "dupes"; choice: DupeChoice }
  | { kind: "productionQc" }
  | { kind: "formulaAudit" }
  | { kind: "aggregate"; fn: AggregateFn }
  | { kind: "units"; id: string; header: boolean }
  | { kind: "chart"; chartKind: ChartKind; title: string }
  | { kind: "merge"; sheets: Array<{ sheet: TargetSheet; wellHeader: string; dateHeader?: string }>; how: MergeHow; dateMatch: DateMatch; maxAgeDays: number }
  | { kind: "nodal"; input: NodalInput }
  | { kind: "prompt"; text: string; agent: boolean };

export type RecipeStep = StepAction & { target?: Target; enabled?: boolean };

export interface Recipe {
  id: string;
  name: string;
  created: number;
  steps: RecipeStep[];
}

/** Steps that read the selection (they get a Target when recorded). */
export const SELECTION_STEPS = new Set<StepAction["kind"]>(["clean", "removeEmpty", "formatTable", "autofit", "fixImport", "fillGaps", "resample", "dupes", "productionQc", "formulaAudit", "aggregate", "units", "chart", "prompt"]);

const PERIOD_WORD: Record<Period, string> = { day: "daily", week: "weekly", month: "monthly", quarter: "quarterly", year: "yearly" };

export function describeStep(s: StepAction): string {
  switch (s.kind) {
    case "clean":
      return "Clean data";
    case "removeEmpty":
      return `Remove empty ${s.lines}`;
    case "formatTable":
      return "Format as table";
    case "autofit":
      return "Auto-fit columns";
    case "fixImport":
      return "Fix imported data";
    case "fillGaps":
      return `Fill gaps (${s.dateHeader}${s.groupHeader ? ` per ${s.groupHeader}` : ""}, ${s.method})`;
    case "resample":
      return `Resample ${PERIOD_WORD[s.period]} (${s.dateHeader}${s.groupHeader ? ` per ${s.groupHeader}` : ""})`;
    case "dupes":
      return `Duplicates & outliers (${[s.choice.highlight && "highlight", s.choice.unify && "unify spellings", s.choice.deleteDuplicates && "delete duplicates"].filter(Boolean).join(", ") || "scan"})`;
    case "productionQc":
      return "Production data QC";
    case "formulaAudit":
      return "Formula audit";
    case "aggregate":
      return `Insert ${s.fn}`;
    case "units":
      return `Convert units (${s.id})`;
    case "chart":
      return `Chart (${s.chartKind}${s.title ? `: ${s.title}` : ""})`;
    case "merge":
      return `Merge ${s.sheets.length} sheets (${s.dateMatch === "asof" ? "latest on/before" : "same date"})`;
    case "nodal":
      return `Nodal analysis (Pr ${s.input.ipr.pr} psia)`;
    case "prompt":
      return `${s.agent ? "Agent" : "Ask"}: “${s.text.length > 60 ? `${s.text.slice(0, 57)}…` : s.text}”`;
  }
}

export function describeTarget(t: Target | undefined, stepLabels: string[] = []): string {
  if (!t) return "";
  const where = t.sheet.kind === "data" ? "data sheet" : t.sheet.kind === "fixed" ? `"${t.sheet.name}"` : `result of step ${t.sheet.index + 1}${stepLabels[t.sheet.index] ? ` (${stepLabels[t.sheet.index]})` : ""}`;
  return `on ${where}`;
}

/* ---------------------------------------------------------------- recording helpers */

export interface RecordState {
  steps: RecipeStep[];
  /** Data sheet = the first sheet a selection step ran on. */
  dataSheet: string | null;
  /** Sheets created by recorded steps → step index. */
  outputs: Record<string, number>;
}

export const newRecording = (): RecordState => ({ steps: [], dataSheet: null, outputs: {} });

export interface SelectionInfo {
  sheet: string;
  rowIndex: number;
  columnIndex: number;
  rowCount: number;
  columnCount: number;
}

export function sheetRef(r: RecordState, sheet: string): TargetSheet {
  if (r.outputs[sheet] !== undefined) return { kind: "step", index: r.outputs[sheet] };
  if (r.dataSheet === null || r.dataSheet === sheet) return { kind: "data" };
  return { kind: "fixed", name: sheet };
}

/** Add a step to a recording. Returns the new state (pure). */
export function recordStep(r: RecordState, action: StepAction, sel: SelectionInfo | null, outputSheet?: string): RecordState {
  let dataSheet = r.dataSheet;
  let target: Target | undefined;
  if (sel && SELECTION_STEPS.has(action.kind)) {
    if (dataSheet === null && r.outputs[sel.sheet] === undefined) dataSheet = sel.sheet;
    target = {
      sheet: sheetRef({ ...r, dataSheet }, sel.sheet),
      startRow: sel.rowIndex,
      startCol: sel.columnIndex,
      endCol: sel.columnIndex + Math.max(1, sel.columnCount) - 1,
      ...(sel.rowCount === 1 && sel.columnCount === 1 ? { single: true } : {}),
    };
  }
  const steps = [...r.steps, { ...action, ...(target ? { target } : {}) }];
  const outputs = outputSheet ? { ...r.outputs, [outputSheet]: steps.length - 1 } : r.outputs;
  return { steps, dataSheet, outputs };
}

/** Record a merge: the first picked sheet becomes the data sheet when there isn't one yet. */
export function recordMerge(
  r: RecordState,
  picks: Array<{ sheet: string; wellHeader: string; dateHeader?: string }>,
  o: { how: MergeHow; dateMatch: DateMatch; maxAgeDays: number },
  outputSheet?: string
): RecordState {
  let dataSheet = r.dataSheet;
  const sheets = picks.map((p) => {
    if (dataSheet === null && r.outputs[p.sheet] === undefined) dataSheet = p.sheet;
    return { sheet: sheetRef({ ...r, dataSheet }, p.sheet), wellHeader: p.wellHeader, ...(p.dateHeader ? { dateHeader: p.dateHeader } : {}) };
  });
  const action: StepAction = {
    kind: "merge",
    sheets,
    how: o.how,
    dateMatch: o.dateMatch,
    maxAgeDays: o.maxAgeDays,
  };
  const steps = [...r.steps, action];
  return { steps, dataSheet, outputs: outputSheet ? { ...r.outputs, [outputSheet]: steps.length - 1 } : r.outputs };
}

/* ---------------------------------------------------------------- column mapping at replay */

export function findColumn(headers: string[], name: string | undefined): number {
  if (!name) return -1;
  const norm = (s: string) => s.trim().toLowerCase().replace(/\s+/g, " ");
  const want = norm(name);
  let i = headers.findIndex((h) => norm(h) === want);
  if (i < 0) i = headers.findIndex((h) => norm(h).replace(/[^a-z0-9]/g, "") === want.replace(/[^a-z0-9]/g, ""));
  return i;
}

export function requireColumn(headers: string[], name: string, what: string): number {
  const i = findColumn(headers, name);
  if (i < 0) throw new Error(`Column "${name}" (${what}) wasn't found. Headers here: ${headers.slice(0, 12).join(", ")}${headers.length > 12 ? "…" : ""}`);
  return i;
}

/** Resample aggregations by header → by column index (unknown headers fall back to the default rule). */
export function aggsByIndex(headers: string[], aggs: Record<string, Agg>): Record<number, Agg> {
  const out: Record<number, Agg> = {};
  for (const [h, a] of Object.entries(aggs)) {
    const i = findColumn(headers, h);
    if (i >= 0) out[i] = a;
  }
  return out;
}

/* ---------------------------------------------------------------- storage & sharing */

const KEY = "zee.recipes";

export function loadRecipes(): Recipe[] {
  try {
    const raw = localStorage.getItem(KEY);
    const list = raw ? (JSON.parse(raw) as Recipe[]) : [];
    return Array.isArray(list) ? list.filter(isRecipe) : [];
  } catch {
    return [];
  }
}

export function saveRecipes(list: Recipe[]): boolean {
  try {
    localStorage.setItem(KEY, JSON.stringify(list));
    return true;
  } catch {
    return false;
  }
}

export function isRecipe(x: unknown): x is Recipe {
  const r = x as Recipe;
  return !!r && typeof r.name === "string" && Array.isArray(r.steps) && r.steps.every((s) => !!s && typeof (s as { kind?: unknown }).kind === "string");
}

const SHARE_PREFIX = "ZEE-RECIPE v1\n";

/** Text to paste to a colleague (Teams, email). */
export function recipeToText(r: Recipe): string {
  return SHARE_PREFIX + JSON.stringify({ name: r.name, steps: r.steps });
}

export function recipeFromText(text: string): Recipe {
  const t = text.trim();
  const body = t.startsWith(SHARE_PREFIX.trim()) ? t.slice(SHARE_PREFIX.trim().length) : t;
  let parsed: unknown;
  try {
    parsed = JSON.parse(body.trim());
  } catch {
    throw new Error("That isn't a Zee AI recipe – copy it with Recipes → Share.");
  }
  const p = parsed as { name?: unknown; steps?: unknown };
  const r = { id: newId(), name: String(p.name ?? "Imported recipe"), created: Date.now(), steps: p.steps } as Recipe;
  if (!isRecipe(r) || !r.steps.length) throw new Error("The recipe has no steps.");
  const known = new Set(["clean", "removeEmpty", "formatTable", "autofit", "fixImport", "fillGaps", "resample", "dupes", "productionQc", "formulaAudit", "aggregate", "units", "chart", "merge", "nodal", "prompt"]);
  const bad = r.steps.filter((s) => !known.has(s.kind));
  if (bad.length) throw new Error(`Unknown step(s): ${bad.map((s) => s.kind).join(", ")} – update Zee AI on this PC.`);
  return r;
}

export function newId(): string {
  return `${Date.now().toString(36)}${Math.random().toString(36).slice(2, 7)}`;
}

/** Display list: does the recipe need a data sheet (any step targets "data")? */
export function usesDataSheet(r: Recipe): boolean {
  return r.steps.some((s) => s.enabled !== false && (s.target?.sheet.kind === "data" || (s.kind === "merge" && s.sheets.some((x) => x.sheet.kind === "data"))));
}
