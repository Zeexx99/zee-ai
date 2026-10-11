/**
 * Catalog of every =ZEE.* worksheet function, used by the task pane's "Functions" drop list.
 * Pure (no Office / React imports) so formula building is unit-testable.
 */

export type FnCategory = "AI – text" | "AI – analysis" | "Petroleum" | "Units";

/**
 * input  – one cell per row, taken from a column of the selection (or typed: a cell like B2, or a value)
 * range  – the whole selected data block
 * text / number / choice – typed once in the dialog and used for every row
 */
export type ArgKind = "input" | "range" | "text" | "number" | "choice";

export interface FnArg {
  name: string;
  label: string;
  kind: ArgKind;
  placeholder?: string;
  default?: string;
  choices?: string[];
  optional?: boolean;
  /** Wrap the cell reference when applying to a selection, e.g. "FORMULATEXT({})". */
  wrap?: string;
}

/** perRow = one formula per selected row · range = one formula for the whole block · single = active cell only */
export type FnMode = "perRow" | "range" | "single";

export interface FnSpec {
  id: string;
  category: FnCategory;
  brief: string;
  syntax: string;
  example: string;
  ai: boolean;
  mode: FnMode;
  args: FnArg[];
}

const num = (name: string, label: string, optional = false): FnArg => ({ name, label, kind: "input", optional, placeholder: "cell (B2) or number" });

export const FUNCTION_CATALOG: FnSpec[] = [
  /* ---------------------------------------------------------------- AI – text */
  {
    id: "ASK",
    category: "AI – text",
    brief: "Ask anything about a cell; short answer in the cell",
    syntax: "=ZEE.ASK(prompt, [cells…])",
    example: '=ZEE.ASK("Is this failure electrical or mechanical?", B2)',
    ai: true,
    mode: "perRow",
    args: [
      { name: "prompt", label: "Prompt", kind: "text", placeholder: "e.g. Is this failure electrical or mechanical?" },
      { name: "context", label: "Cell", kind: "input", optional: true, placeholder: "column of the selection, or B2" },
    ],
  },
  {
    id: "EXTRACT",
    category: "AI – text",
    brief: "Pull one piece of information out of messy text",
    syntax: "=ZEE.EXTRACT(text, what)",
    example: '=ZEE.EXTRACT(B2, "pump model")',
    ai: true,
    mode: "perRow",
    args: [
      { name: "text", label: "Text", kind: "input" },
      { name: "what", label: "What to extract", kind: "text", placeholder: "e.g. pump model, date, well name" },
    ],
  },
  {
    id: "CLASSIFY",
    category: "AI – text",
    brief: "Put text into exactly one of your categories",
    syntax: "=ZEE.CLASSIFY(text, categories)",
    example: '=ZEE.CLASSIFY(B2, "Electrical, Mechanical, Reservoir, Other")',
    ai: true,
    mode: "perRow",
    args: [
      { name: "text", label: "Text", kind: "input" },
      { name: "categories", label: "Categories (comma separated)", kind: "text", placeholder: "Electrical, Mechanical, Reservoir, Other" },
    ],
  },
  {
    id: "CLEAN",
    category: "AI – text",
    brief: "Fix spelling, spacing, casing and format of a value",
    syntax: "=ZEE.CLEAN(text, [style])",
    example: '=ZEE.CLEAN(B2, "Title Case company names")',
    ai: true,
    mode: "perRow",
    args: [
      { name: "text", label: "Text", kind: "input" },
      { name: "style", label: "Style (optional)", kind: "text", optional: true, placeholder: "e.g. dates as YYYY-MM-DD" },
    ],
  },
  {
    id: "TRANSLATE",
    category: "AI – text",
    brief: "Translate text (e.g. Arabic ↔ English field reports)",
    syntax: "=ZEE.TRANSLATE(text, [language])",
    example: '=ZEE.TRANSLATE(B2, "English")',
    ai: true,
    mode: "perRow",
    args: [
      { name: "text", label: "Text", kind: "input" },
      { name: "language", label: "To language", kind: "text", optional: true, default: "English", placeholder: "English" },
    ],
  },
  {
    id: "SUMMARIZE",
    category: "AI – text",
    brief: "Summarize a block of cells in a few words",
    syntax: "=ZEE.SUMMARIZE(range, [max_words])",
    example: "=ZEE.SUMMARIZE(A2:D50, 30)",
    ai: true,
    mode: "range",
    args: [
      { name: "range", label: "Range", kind: "range" },
      { name: "max_words", label: "Max words", kind: "number", optional: true, default: "30" },
    ],
  },
  /* ---------------------------------------------------------------- AI – analysis */
  {
    id: "FLAG",
    category: "AI – analysis",
    brief: "Mark rows that break a rule you describe (spills a column)",
    syntax: "=ZEE.FLAG(range, rule)",
    example: '=ZEE.FLAG(A2:D200, "motor amps jump >20% or intake pressure drops")',
    ai: true,
    mode: "range",
    args: [
      { name: "range", label: "Range", kind: "range" },
      { name: "rule", label: "Rule in words", kind: "text", placeholder: "e.g. current spikes more than 20%" },
    ],
  },
  {
    id: "FORMULA",
    category: "AI – analysis",
    brief: "Write an Excel formula from a description (as text)",
    syntax: "=ZEE.FORMULA(description, [sample])",
    example: '=ZEE.FORMULA("average of column C where column B is ESP")',
    ai: true,
    mode: "single",
    args: [
      { name: "description", label: "What should the formula do?", kind: "text" },
      { name: "sample", label: "Sample data (optional)", kind: "input", optional: true, placeholder: "e.g. A1:D5" },
    ],
  },
  {
    id: "EXPLAIN",
    category: "AI – analysis",
    brief: "Explain a formula in plain language",
    syntax: "=ZEE.EXPLAIN(formula_text)",
    example: "=ZEE.EXPLAIN(FORMULATEXT(B2))",
    ai: true,
    mode: "perRow",
    args: [{ name: "formula", label: "Formula cell", kind: "input", wrap: "FORMULATEXT({})", placeholder: "column of formulas, or B2" }],
  },
  /* ---------------------------------------------------------------- Units (no AI) */
  {
    id: "UNITS",
    category: "Units",
    brief: "Convert any oilfield unit: psi, bar, bbl/d, m3/d, °F, API…",
    syntax: "=ZEE.UNITS(value, from, to)",
    example: '=ZEE.UNITS(B2, "bbl/d", "m3/d")',
    ai: false,
    mode: "perRow",
    args: [
      num("value", "Value"),
      { name: "from", label: "From unit", kind: "text", placeholder: "e.g. psi, bbl/d, F, API" },
      { name: "to", label: "To unit", kind: "text", placeholder: "e.g. bar, m3/d, C, SG" },
    ],
  },
  {
    id: "CONVERT",
    category: "Units",
    brief: "Quick preset conversions (psi_bar, bbl_m3, f_c…)",
    syntax: "=ZEE.CONVERT(value, conversion)",
    example: '=ZEE.CONVERT(B2, "psi_bar")',
    ai: false,
    mode: "perRow",
    args: [
      num("value", "Value"),
      {
        name: "conversion",
        label: "Conversion",
        kind: "choice",
        choices: ["bbl_m3", "m3_bbl", "scf_m3", "mscf_m3", "mmscf_mm3", "m3_scf", "psi_bar", "bar_psi", "psi_kpa", "kpa_psi", "f_c", "c_f", "ft_m", "m_ft", "api_sg", "sg_api", "psift_sg", "hp_kw", "kw_hp"],
      },
    ],
  },
  /* ---------------------------------------------------------------- Petroleum (no AI) */
  {
    id: "VOGEL",
    category: "Petroleum",
    brief: "Vogel IPR: rate at a flowing pressure from one test",
    syntax: "=ZEE.VOGEL(Pr, Pwf, q_test, Pwf_test)",
    example: "=ZEE.VOGEL(3000, 1500, 800, 2000)",
    ai: false,
    mode: "perRow",
    args: [num("Pr", "Reservoir pressure (psia)"), num("Pwf", "Flowing pressure (psia)"), num("q_test", "Test rate (bbl/d)"), num("Pwf_test", "Test Pwf (psia)")],
  },
  {
    id: "VOGEL_QMAX",
    category: "Petroleum",
    brief: "Vogel AOF (rate at Pwf = 0)",
    syntax: "=ZEE.VOGEL_QMAX(Pr, q_test, Pwf_test)",
    example: "=ZEE.VOGEL_QMAX(3000, 800, 2000)",
    ai: false,
    mode: "perRow",
    args: [num("Pr", "Reservoir pressure (psia)"), num("q_test", "Test rate (bbl/d)"), num("Pwf_test", "Test Pwf (psia)")],
  },
  {
    id: "PI",
    category: "Petroleum",
    brief: "Productivity index J = q / (Pr − Pwf)",
    syntax: "=ZEE.PI(q, Pr, Pwf)",
    example: "=ZEE.PI(800, 3000, 2000)",
    ai: false,
    mode: "perRow",
    args: [num("q", "Rate (bbl/d)"), num("Pr", "Reservoir pressure (psia)"), num("Pwf", "Flowing pressure (psia)")],
  },
  {
    id: "RS",
    category: "Petroleum",
    brief: "Standing solution GOR (scf/STB)",
    syntax: "=ZEE.RS(p, T, API, gas_SG)",
    example: "=ZEE.RS(2000, 180, 35, 0.75)",
    ai: false,
    mode: "perRow",
    args: [num("p", "Pressure (psia)"), num("T", "Temperature (°F)"), num("API", "Oil API"), num("gas_SG", "Gas SG")],
  },
  {
    id: "PB",
    category: "Petroleum",
    brief: "Standing bubble-point pressure (psia)",
    syntax: "=ZEE.PB(Rs, T, API, gas_SG)",
    example: "=ZEE.PB(500, 180, 35, 0.75)",
    ai: false,
    mode: "perRow",
    args: [num("Rs", "Solution GOR (scf/STB)"), num("T", "Temperature (°F)"), num("API", "Oil API"), num("gas_SG", "Gas SG")],
  },
  {
    id: "BO",
    category: "Petroleum",
    brief: "Standing oil FVF Bo (rb/STB)",
    syntax: "=ZEE.BO(p, T, API, gas_SG)",
    example: "=ZEE.BO(2000, 180, 35, 0.75)",
    ai: false,
    mode: "perRow",
    args: [num("p", "Pressure (psia)"), num("T", "Temperature (°F)"), num("API", "Oil API"), num("gas_SG", "Gas SG")],
  },
  {
    id: "ARPS",
    category: "Petroleum",
    brief: "Arps decline rate at time t",
    syntax: "=ZEE.ARPS(qi, Di, b, t)",
    example: "=ZEE.ARPS(1000, 0.05, 0.5, 12)",
    ai: false,
    mode: "perRow",
    args: [num("qi", "Initial rate"), num("Di", "Initial decline (1/period)"), num("b", "b exponent"), num("t", "Time")],
  },
  {
    id: "ARPS_CUM",
    category: "Petroleum",
    brief: "Arps cumulative production to time t",
    syntax: "=ZEE.ARPS_CUM(qi, Di, b, t)",
    example: "=ZEE.ARPS_CUM(1000, 0.05, 0.5, 12)",
    ai: false,
    mode: "perRow",
    args: [num("qi", "Initial rate"), num("Di", "Initial decline (1/period)"), num("b", "b exponent"), num("t", "Time")],
  },
  {
    id: "TDH",
    category: "Petroleum",
    brief: "Pump total dynamic head (ft) from pressures",
    syntax: "=ZEE.TDH(P_discharge, P_intake, SG)",
    example: "=ZEE.TDH(3500, 800, 1.02)",
    ai: false,
    mode: "perRow",
    args: [num("P_discharge", "Discharge pressure (psi)"), num("P_intake", "Intake pressure (psi)"), num("SG", "Fluid SG")],
  },
  {
    id: "HHP",
    category: "Petroleum",
    brief: "Hydraulic horsepower from rate and ΔP",
    syntax: "=ZEE.HHP(q, dP)",
    example: "=ZEE.HHP(2000, 2700)",
    ai: false,
    mode: "perRow",
    args: [num("q", "Rate (bbl/d)"), num("dP", "Pressure rise (psi)")],
  },
  {
    id: "AFFINITY",
    category: "Petroleum",
    brief: "Pump affinity laws: scale Q, H or P with frequency",
    syntax: "=ZEE.AFFINITY(value, f_from, f_to, kind)",
    example: '=ZEE.AFFINITY(B2, 60, 50, "H")',
    ai: false,
    mode: "perRow",
    args: [num("value", "Value at f_from"), num("f_from", "From Hz"), num("f_to", "To Hz"), { name: "kind", label: "Kind", kind: "choice", choices: ["Q", "H", "P"], default: "Q" }],
  },
  {
    id: "SG_MIX",
    category: "Petroleum",
    brief: "Liquid mixture SG from water cut and API",
    syntax: "=ZEE.SG_MIX(water_cut, API, [water_SG])",
    example: "=ZEE.SG_MIX(0.85, 32, 1.05)",
    ai: false,
    mode: "perRow",
    args: [num("water_cut", "Water cut (0–1)"), num("API", "Oil API"), { ...num("water_SG", "Water SG (optional)", true), kind: "number" }],
  },
];

export const CATEGORY_ORDER: FnCategory[] = ["AI – text", "AI – analysis", "Units", "Petroleum"];

export function findFunction(id: string): FnSpec | undefined {
  return FUNCTION_CATALOG.find((f) => f.id === id);
}

/* ---------------------------------------------------------------- formula building */

const CELL = String.raw`\$?[A-Za-z]{1,3}\$?\d{1,7}`;
const SHEET = String.raw`(?:'(?:[^']|'')+'|[A-Za-z_][\w.]*)!`;
const REF_RE = new RegExp(String.raw`^(?:${SHEET})?(?:${CELL}(?::${CELL})?|\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}|\$?\d{1,7}:\$?\d{1,7})$`);

export function isCellRef(s: string): boolean {
  return REF_RE.test(s.trim());
}

export function quoteText(s: string): string {
  return `"${s.replace(/"/g, '""')}"`;
}

/** Turn what the user typed into an Excel argument. */
export function toArg(raw: string | undefined, kind: ArgKind): string {
  const s = (raw ?? "").trim();
  if (s === "") return "";
  if (s.startsWith("=")) return s.slice(1).trim(); // explicit expression: =B2*2
  if (/^".*"$/.test(s) && s.length >= 2) return s; // already quoted
  if (isCellRef(s)) return s;
  if ((kind === "number" || kind === "input" || kind === "range") && /^-?\d+(\.\d+)?(e-?\d+)?$/i.test(s)) return s;
  return quoteText(s);
}

/** Does the user still need to type something before this function can run? */
export function missingArgs(spec: FnSpec, values: Record<string, string>, fromSelection: boolean, selectionColumns = 0): FnArg[] {
  let inputIndex = 0;
  return spec.args.filter((a) => {
    const v = (values[a.name] ?? "").trim();
    if (a.kind === "input" || a.kind === "range") {
      if (v) return false; // typed values/refs don't use up a selection column
      const idx = a.kind === "input" ? inputIndex++ : -1;
      if (a.optional) return false;
      if (!fromSelection) return true;
      if (a.kind === "range") return false;
      return idx >= Math.max(1, selectionColumns);
    }
    if (v || a.default) return false;
    return !a.optional;
  });
}

function joinArgs(args: string[]): string {
  // Drop trailing empty optional args; keep empty middle ones as "".
  const out = [...args];
  while (out.length && out[out.length - 1] === "") out.pop();
  return out.map((a) => (a === "" ? '""' : a)).join(", ");
}

/**
 * Build one formula.
 * @param values  what the user typed per arg name
 * @param cells   for "input" args not typed by the user: references in order (the 1st untyped input gets cells[0], …)
 * @param range   for "range" args not typed by the user
 */
export function buildFormula(spec: FnSpec, values: Record<string, string>, cells: string[] = [], range?: string): string {
  let inputIndex = 0;
  const args = spec.args.map((a) => {
    const typed = (values[a.name] ?? "").trim();
    if (a.kind === "input") {
      if (typed) return toArg(typed, a.kind);
      const ref = cells[inputIndex++];
      if (!ref) return "";
      return a.wrap ? a.wrap.replace("{}", ref) : ref;
    }
    if (a.kind === "range") return typed ? toArg(typed, a.kind) : range ?? "";
    return toArg(typed || a.default || "", a.kind);
  });
  return `=ZEE.${spec.id}(${joinArgs(args)})`;
}

/** Count of "input" args (each maps to one column of the selection). */
export function inputCount(spec: FnSpec): number {
  return spec.args.filter((a) => a.kind === "input").length;
}

/** Can a click apply this function immediately (nothing to type)? */
export function canApplyDirectly(spec: FnSpec, selectionColumns: number): boolean {
  if (spec.mode === "single") return false;
  return missingArgs(spec, {}, true, selectionColumns).length === 0;
}

/** Simple header guess: first row is all text and some later cell is a number, or the first row has no blanks and differs in type. */
export function guessHeader(values: unknown[][]): boolean {
  if (values.length < 2) return false;
  const first = values[0];
  if (!first.every((v) => typeof v === "string" && v.trim() !== "" && !/^-?\d+(\.\d+)?$/.test(v.trim()))) return false;
  const rest = values.slice(1);
  if (rest.some((r) => r.some((v) => typeof v === "number"))) return true;
  // All-text data: treat short, distinct labels as a header.
  return first.every((v) => String(v).length <= 30) && new Set(first.map(String)).size === first.length && first.length > 1;
}
