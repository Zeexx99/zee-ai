/**
 * Which Zee AI features this Excel can run. Older Excel builds (e.g. Excel 2016/2019 perpetual) lack newer
 * Office.js API sets; instead of failing with an error, the matching buttons are greyed out with a note and
 * the agent isn't offered tools it can't use. The pure part takes an isSupported() probe so it's testable.
 */

export type FeatureId = "core" | "charts" | "axes" | "formatUndo" | "condFormat" | "validation" | "pivot" | "filterFind" | "copy" | "functions";

export interface FeatureReq {
  id: FeatureId;
  label: string;
  set: "ExcelApi" | "CustomFunctionsRuntime";
  version: string;
}

export const FEATURES: FeatureReq[] = [
  { id: "core", label: "Reading & writing cells, Clean, Data, Petroleum tools, reports", set: "ExcelApi", version: "1.7" },
  { id: "charts", label: "Charts (Draw, Chart from selection, ESP / decline / nodal charts)", set: "ExcelApi", version: "1.7" },
  { id: "axes", label: "Secondary axis in Draw", set: "ExcelApi", version: "1.8" },
  { id: "formatUndo", label: "Undo of formatting changes", set: "ExcelApi", version: "1.9" },
  { id: "condFormat", label: "Conditional formatting (agent)", set: "ExcelApi", version: "1.6" },
  { id: "validation", label: "Data validation (Suggest data validation, agent)", set: "ExcelApi", version: "1.8" },
  { id: "pivot", label: "PivotTables (agent)", set: "ExcelApi", version: "1.8" },
  { id: "filterFind", label: "Filter, find & replace (agent)", set: "ExcelApi", version: "1.9" },
  { id: "copy", label: "Copy ranges (agent)", set: "ExcelApi", version: "1.9" },
  { id: "functions", label: "=ZEE.* worksheet functions and the Functions list", set: "CustomFunctionsRuntime", version: "1.1" },
];

/** Agent tools that need more than the core API set. */
export const TOOL_FEATURE: Record<string, FeatureId> = {
  create_chart: "charts",
  conditional_format: "condFormat",
  data_validation: "validation",
  create_pivot_table: "pivot",
  filter_range: "filterFind",
  find_cells: "filterFind",
  replace_text: "filterFind",
  copy_range: "copy",
  esp_failure_analysis: "charts",
  nodal_analysis: "charts",
};

export type IsSupported = (set: string, version: string) => boolean;

export interface FeatureStatus extends FeatureReq {
  ok: boolean;
}

export function evaluateFeatures(isSupported: IsSupported): FeatureStatus[] {
  return FEATURES.map((f) => ({ ...f, ok: safe(() => isSupported(f.set, f.version)) }));
}

/** Highest ExcelApi 1.x this Excel reports (e.g. "1.17"), or null. */
export function highestExcelApi(isSupported: IsSupported, upTo = 25): string | null {
  let best: string | null = null;
  for (let v = 1; v <= upTo; v++) if (safe(() => isSupported("ExcelApi", `1.${v}`))) best = `1.${v}`;
  return best;
}

function safe(f: () => boolean): boolean {
  try {
    return !!f();
  } catch {
    return false;
  }
}

/* ---------------------------------------------------------------- live probe (cached) */

let cache: Map<FeatureId, boolean> | null = null;

function officeProbe(): IsSupported | null {
  try {
    const req = (globalThis as { Office?: typeof Office }).Office?.context?.requirements;
    return req ? (set, version) => req.isSetSupported(set, version) : null;
  } catch {
    return null;
  }
}

/** Is a feature available here? Outside Excel (browser preview) everything reports true. */
export function canUse(id: FeatureId): boolean {
  const probe = officeProbe();
  if (!probe) return true;
  if (!cache) cache = new Map(evaluateFeatures(probe).map((f) => [f.id, f.ok]));
  return cache.get(id) ?? true;
}

export function canUseTool(name: string): boolean {
  const f = TOOL_FEATURE[name];
  return (f ? canUse(f) : true) && canUse("core");
}

export function featureNote(id: FeatureId): string {
  const f = FEATURES.find((x) => x.id === id);
  return f ? `Needs a newer Excel (${f.set} ${f.version})` : "Not available in this Excel";
}

export function liveFeatureReport(): { highest: string | null; features: FeatureStatus[] } | null {
  const probe = officeProbe();
  if (!probe) return null;
  return { highest: highestExcelApi(probe), features: evaluateFeatures(probe) };
}

/** @internal tests */
export function resetFeatureCache() {
  cache = null;
}
