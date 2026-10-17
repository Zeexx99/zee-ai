/**
 * Replays a recipe step by step. Runs without the usual confirmation dialogs (the user confirmed the whole
 * recipe), every change still lands on the undo / change-history list. Stops at the first failing step.
 */
import {
  ExcelUserError,
  applyEmptyLines,
  autofitAllColumns,
  cleanSelection,
  createChart,
  formatSelectionAsTable,
  getSelectedAddress,
  insertAggregate,
  insertConvertedColumns,
  scanEmptyLines,
  selectBlock,
} from "./excelService";
import { applyDupes, applyImportFix, previewImportFix, runGapFill, scanDupes } from "./dataTools";
import { readHeaderedSelection, runFormulaAudit } from "./toolActions";
import { runProductionQc } from "./petroActions";
import { listSheetTables, runMerge, runResample } from "./dataMenu";
import { runNodalSheet } from "./nodalActions";
import { UNIT_CONVERSIONS } from "../utils/petro";
import { aggsByIndex, describeStep, findColumn, requireColumn, type Recipe, type RecipeStep, type TargetSheet } from "./recipes";

export interface RunnerDeps {
  /** Send a chat / agent prompt and wait for the answer (agent approvals still apply). */
  sendPrompt: (text: string, agent: boolean) => Promise<void>;
  onStep: (index: number, total: number, label: string) => void;
  signal?: AbortSignal;
}

export interface RunReport {
  done: number;
  total: number;
  outputs: string[];
  failed?: { index: number; label: string; error: string };
  cancelled?: boolean;
}

export async function runRecipe(recipe: Recipe, dataSheet: string, deps: RunnerDeps): Promise<RunReport> {
  const steps = recipe.steps.map((s, i) => ({ s, i })).filter(({ s }) => s.enabled !== false);
  const outputs: Record<number, string> = {};
  const report: RunReport = { done: 0, total: steps.length, outputs: [] };
  const sheetOf = (t: TargetSheet): string => {
    if (t.kind === "data") return dataSheet;
    if (t.kind === "fixed") return t.name;
    const out = outputs[t.index];
    if (!out) throw new ExcelUserError(`Step ${t.index + 1} didn't create a sheet in this run (it may be switched off).`);
    return out;
  };

  for (let k = 0; k < steps.length; k++) {
    const { s, i } = steps[k];
    const label = describeStep(s);
    if (deps.signal?.aborted) {
      report.cancelled = true;
      return report;
    }
    deps.onStep(k, steps.length, label);
    try {
      if (s.target) {
        const t = s.target;
        await selectBlock(sheetOf(t.sheet), t.startRow, t.startCol, t.endCol - t.startCol + 1, t.single ? 1 : undefined);
      }
      const out = await runStep(s, sheetOf, deps);
      if (out) {
        outputs[i] = out;
        report.outputs.push(out);
      }
      report.done++;
    } catch (e) {
      report.failed = { index: i, label, error: (e as Error).message ?? String(e) };
      return report;
    }
  }
  return report;
}

async function runStep(s: RecipeStep, sheetOf: (t: TargetSheet) => string, deps: RunnerDeps): Promise<string | undefined> {
  switch (s.kind) {
    case "clean":
      await cleanSelection(s.options, { signal: deps.signal });
      return;
    case "removeEmpty": {
      const plan = await scanEmptyLines(s.lines, { signal: deps.signal });
      if (plan.hits.length) await applyEmptyLines(plan);
      return;
    }
    case "formatTable":
      await formatSelectionAsTable();
      return;
    case "autofit":
      await autofitAllColumns();
      return;
    case "fixImport": {
      const p = await previewImportFix(await readHeaderedSelection("imported table"));
      if (p.changes || p.plans.some((x) => x.newHeader)) await applyImportFix(p);
      return;
    }
    case "fillGaps": {
      const sel = await readHeaderedSelection("time series");
      const dateCol = requireColumn(sel.headers, s.dateHeader, "date");
      const groupCol = s.groupHeader ? requireColumn(sel.headers, s.groupHeader, "well") : undefined;
      return (await runGapFill(sel, { dateCol, groupCol, method: s.method, maxGapSteps: s.maxGapSteps })).sheet;
    }
    case "resample": {
      const sel = await readHeaderedSelection("time series");
      const dateCol = requireColumn(sel.headers, s.dateHeader, "date");
      const groupCol = s.groupHeader ? requireColumn(sel.headers, s.groupHeader, "well") : undefined;
      return (await runResample(sel, { dateCol, groupCol, period: s.period, aggs: aggsByIndex(sel.headers, s.aggs), fill: s.fill, countDays: s.countDays })).sheet;
    }
    case "dupes": {
      const scan = scanDupes(await readHeaderedSelection("table"));
      await applyDupes(scan, s.choice);
      return;
    }
    case "productionQc":
      return (await runProductionQc()).sheet;
    case "formulaAudit":
      return (await runFormulaAudit({ signal: deps.signal })).sheet;
    case "aggregate":
      await insertAggregate(s.fn, false);
      return;
    case "units": {
      const conv = UNIT_CONVERSIONS.find((u) => u.id === s.id);
      if (!conv) throw new ExcelUserError(`Unknown conversion "${s.id}".`);
      await insertConvertedColumns(conv, s.header);
      return;
    }
    case "chart":
      await createChart(await getSelectedAddress(), s.chartKind, s.title || undefined);
      return;
    case "merge": {
      const tables = await listSheetTables();
      const picks = s.sheets.map((m) => {
        const name = sheetOf(m.sheet);
        const t = tables.find((x) => x.name === name);
        if (!t) throw new ExcelUserError(`Sheet "${name}" is empty or missing.`);
        return { sheet: name, wellCol: requireColumn(t.headers, m.wellHeader, "well"), dateCol: m.dateHeader ? findColumn(t.headers, m.dateHeader) : -1 };
      });
      return (await runMerge(picks, { how: s.how, dateMatch: s.dateMatch, maxAgeDays: s.maxAgeDays })).sheet;
    }
    case "nodal":
      return (await runNodalSheet(s.input)).sheet;
    case "prompt":
      await deps.sendPrompt(s.text, s.agent);
      return;
  }
}
