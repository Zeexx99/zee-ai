/**
 * One list of every toolbar tool, used by the Ctrl+K command search and by pinned toolbar buttons.
 * `searchCommands` is pure and unit-tested.
 */
import type { FeatureId } from "../qol/requirements";
import { AGGREGATE_LABELS, type AggregateFn } from "../services/excelService";
import type { QuickActionHandlers } from "./QuickActions";

export type CommandGroup = "Math" | "Clean" | "Analyze" | "Petroleum" | "Draw" | "Data" | "Zee AI";

export interface Command {
  /** Stable id – used to remember pins. */
  id: string;
  label: string;
  group: CommandGroup;
  /** Short label for a pinned toolbar button. */
  short?: string;
  hint?: string;
  /** Extra words the search should match (abbreviations, synonyms). */
  keywords?: string;
  /** Greyed out when this Excel lacks the API set. */
  feature?: FeatureId;
  run: () => void;
}

const AGG: AggregateFn[] = ["SUM", "AVERAGE", "MIN", "MAX", "COUNT", "COUNTA", "UNIQUE"];

/** The toolbar's tools as commands (same handlers as the menus). */
export function toolCommands(p: QuickActionHandlers): Command[] {
  const list: Command[] = [
    ...AGG.map((fn): Command => ({ id: `math.${fn}`, label: `${AGGREGATE_LABELS[fn]} below selection`, short: fn === "AVERAGE" ? "AVG" : fn, group: "Math", keywords: `${fn} aggregate total`, run: () => p.onAggregate(fn) })),

    { id: "clean.clean", label: "Clean data…", short: "Clean", group: "Clean", keywords: "trim spaces text numbers dates case", run: p.onClean },
    { id: "clean.extract", label: "Split / Extract…", short: "Split", group: "Clean", keywords: "split extract column first name", run: p.onExtract },
    { id: "clean.dupes", label: "Find duplicates & outliers…", short: "Dupes", group: "Clean", keywords: "duplicate outlier anomaly", run: p.onFindDupes },
    { id: "clean.gaps", label: "Fill gaps in time series…", short: "Gaps", group: "Clean", keywords: "missing dates interpolate", run: p.onFillGaps },
    { id: "clean.fixImport", label: "Fix imported data…", short: "Fix import", group: "Clean", keywords: "arabic digits mixed dates units", run: p.onFixImport },
    { id: "clean.emptyCols", label: "Remove empty columns", group: "Clean", keywords: "blank null delete", run: p.onRemoveEmptyColumns },
    { id: "clean.emptyRows", label: "Remove empty rows", group: "Clean", keywords: "blank null delete", run: p.onRemoveEmptyRows },
    { id: "clean.formatTable", label: "Format as table", short: "Table", group: "Clean", run: p.onFormatTable },
    { id: "clean.autofit", label: "Auto-fit all columns", short: "Autofit", group: "Clean", keywords: "width", run: p.onAutofit },

    { id: "analyze.explainFormula", label: "Explain formula in active cell", short: "Explain", group: "Analyze", run: p.onExplainFormula },
    { id: "analyze.generateFormula", label: "Generate formula…", short: "Formula", group: "Analyze", run: p.onGenerateFormula },
    { id: "analyze.summarize", label: "Summarize selection", short: "Summary", group: "Analyze", keywords: "stats statistics", run: p.onAnalyze },
    { id: "analyze.chart", label: "Chart from selection…", short: "Chart", group: "Analyze", feature: "charts", keywords: "graph plot", run: p.onChart },
    { id: "analyze.audit", label: "Formula audit (this sheet)", short: "Audit", group: "Analyze", keywords: "errors broken references check", run: p.onFormulaAudit },
    { id: "analyze.rules", label: "Suggest data validation…", short: "Validation", group: "Analyze", feature: "validation", keywords: "rules", run: p.onSuggestRules },
    { id: "analyze.explainSheet", label: "Explain this sheet", group: "Analyze", keywords: "map formulas", run: p.onExplainSheet },
    { id: "analyze.sheetGuide", label: "Add a sheet guide tab", group: "Analyze", run: p.onSheetGuide },
    { id: "analyze.namedInputs", label: "Move hard-coded numbers to inputs…", group: "Analyze", keywords: "constants named ranges", run: p.onNamedInputs },

    { id: "petro.daily", label: "Daily production report…", short: "Daily report", group: "Petroleum", hint: "→ Word", keywords: "production report word", run: p.onDailyReport },
    { id: "petro.dca", label: "Decline curve analysis…", short: "DCA", group: "Petroleum", feature: "charts", keywords: "dca arps decline forecast eur", run: p.onDeclineCurve },
    { id: "petro.qc", label: "Production data QC", short: "QC", group: "Petroleum", keywords: "quality check production", run: p.onProductionQc },
    { id: "petro.units", label: "Unit conversion…", short: "Units", group: "Petroleum", keywords: "bbl m3 psi bar convert", run: p.onUnits },
    { id: "petro.esp", label: "ESP analysis…", short: "ESP", group: "Petroleum", feature: "charts", keywords: "esp pump electrical submersible", run: p.onEspAnalysis },
    { id: "petro.espCheck", label: "ESP performance check", group: "Petroleum", keywords: "esp pump", run: p.onEspCheck },
    { id: "petro.espLog", label: "Create an ESP failure log sheet", group: "Petroleum", keywords: "esp failure template", run: p.onEspLogTemplate },
    { id: "petro.espLogAnalyze", label: "Analyze ESP failure log…", short: "MTBF", group: "Petroleum", feature: "charts", keywords: "esp failure mtbf run life", run: p.onEspLogAnalyze },
    { id: "petro.nodal", label: "Nodal analysis…", short: "Nodal", group: "Petroleum", feature: "charts", keywords: "ipr vlp beggs brill well performance", run: p.onNodal },

    { id: "draw.plot", label: "Plot columns (X vs several Y)…", short: "Plot", group: "Draw", feature: "charts", keywords: "chart graph draw", run: p.onDraw },

    { id: "data.import", label: "Import file…", short: "Import", group: "Data", hint: "CSV · LAS · SCADA", keywords: "csv las scada load", run: p.onImportFile },
    { id: "data.merge", label: "Merge well data…", short: "Merge", group: "Data", keywords: "join combine wells dates", run: p.onMerge },
    { id: "data.resample", label: "Resample time series…", short: "Resample", group: "Data", hint: "day → month", keywords: "daily monthly aggregate", run: p.onResample },
  ];
  return list;
}

const norm = (s: string) => s.toLowerCase().normalize("NFKD").replace(/[^\p{L}\p{N}\s]/gu, " ");

/**
 * Every query word must appear in the label, group or keywords. Label matches rank first,
 * then label-starts-with, then the original order. Empty query → all commands.
 */
export function searchCommands<T extends Pick<Command, "label" | "group" | "keywords" | "hint">>(cmds: T[], query: string): T[] {
  const words = norm(query).split(/\s+/).filter(Boolean);
  if (!words.length) return cmds;
  return cmds
    .map((c, i) => {
      const label = norm(c.label);
      const hay = `${label} ${norm(c.group)} ${norm(c.keywords ?? "")} ${norm(c.hint ?? "")}`;
      if (!words.every((w) => hay.includes(w))) return null;
      const inLabel = words.filter((w) => label.includes(w)).length;
      const starts = label.startsWith(words[0]) || label.split(/\s+/).some((t) => t.startsWith(words[0])) ? 1 : 0;
      return { c, i, score: inLabel * 10 + starts * 5 };
    })
    .filter((x): x is { c: T; i: number; score: number } => !!x)
    .sort((a, b) => b.score - a.score || a.i - b.i)
    .map((x) => x.c);
}

/** Adds or removes a pin, keeping at most `max` (oldest pin drops). Pure. */
export function togglePin(pins: string[], id: string, max: number): string[] {
  if (pins.includes(id)) return pins.filter((p) => p !== id);
  const next = [...pins, id];
  return next.length > max ? next.slice(next.length - max) : next;
}
