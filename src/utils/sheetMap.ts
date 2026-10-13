/**
 * "Explain this sheet": a map of how a sheet calculates – input columns, formula patterns per column
 * (grouped by their R1C1 shape) and what each one depends on. Pure – unit tested.
 */
import { columnLetter } from "./a1";

export interface MapBlock {
  rowIndex: number;
  columnIndex: number;
  formulas: unknown[][];
  formulasR1C1: unknown[][];
}

export interface FormulaGroup {
  column: string;
  header: string;
  /** Example formula in A1 style (from the first row of the group). */
  example: string;
  firstRow: number;
  lastRow: number;
  count: number;
  /** Columns on this sheet, other sheets and names it reads. */
  dependsOn: string[];
}

export interface ColumnRole {
  column: string;
  header: string;
  role: "input" | "calculated" | "mixed" | "label" | "empty";
  values: number;
  formulas: number;
}

export interface SheetMap {
  headerRow: number | null;
  columns: ColumnRole[];
  groups: FormulaGroup[];
  otherSheets: string[];
  names: string[];
  formulaCount: number;
}

const isFormula = (f: unknown): f is string => typeof f === "string" && f.startsWith("=");

const FUNCTIONS = /\b([A-Z][A-Z0-9.]*)\s*\(/g;

/** References in a formula: same-sheet columns, other sheets and defined names. */
export function formulaRefs(formula: string): { columns: string[]; sheets: string[]; names: string[] } {
  const noStrings = formula.replace(/"([^"]|"")*"/g, " ");
  const sheets = new Set<string>();
  const sheetRe = /(?:'((?:[^']|'')+)'|([A-Za-z_][\w.]*))!/g;
  let m: RegExpExecArray | null;
  while ((m = sheetRe.exec(noStrings))) sheets.add((m[1] ?? m[2]).replace(/''/g, "'"));
  const local = noStrings.replace(/(?:'(?:[^']|'')+'|[A-Za-z_][\w.]*)!\$?[A-Za-z]{1,3}\$?\d*(?::\$?[A-Za-z]{1,3}\$?\d*)?/g, " ");
  const columns = new Set<string>();
  const refRe = /\$?\b([A-Za-z]{1,3})\$?\d+\b|\$?\b([A-Za-z]{1,3}):\$?([A-Za-z]{1,3})\b/g;
  while ((m = refRe.exec(local))) {
    if (m[1]) columns.add(m[1].toUpperCase());
    else {
      columns.add(m[2].toUpperCase());
      columns.add(m[3].toUpperCase());
    }
  }
  const funcs = new Set<string>();
  let f: RegExpExecArray | null;
  while ((f = FUNCTIONS.exec(local.toUpperCase()))) funcs.add(f[1]);
  const names = new Set<string>();
  const cleaned = local.replace(/\$?\b[A-Za-z]{1,3}\$?\d+\b/g, " ").replace(/\b[A-Za-z_][\w.]*\s*\(/g, "(");
  const nameRe = /\b([A-Za-z_][A-Za-z0-9_.]{2,})\b/g;
  while ((m = nameRe.exec(cleaned))) {
    const n = m[1];
    if (/^(TRUE|FALSE)$/i.test(n) || funcs.has(n.toUpperCase())) continue;
    names.add(n);
  }
  return { columns: Array.from(columns), sheets: Array.from(sheets), names: Array.from(names) };
}

export function buildSheetMap(b: MapBlock): SheetMap {
  const rows = b.formulas.length;
  const cols = b.formulas[0]?.length ?? 0;
  const first = b.formulas[0] ?? [];
  const textCells = first.filter((v) => typeof v === "string" && v.trim() !== "" && !isFormula(v)).length;
  const headerRow = rows > 1 && textCells >= Math.max(1, Math.ceil(cols / 2)) ? 0 : null;
  const headerOf = (j: number) => (headerRow !== null ? String(first[j] ?? "").trim() : "");
  const start = headerRow === null ? 0 : 1;

  const columns: ColumnRole[] = [];
  const groups: FormulaGroup[] = [];
  const otherSheets = new Set<string>();
  const names = new Set<string>();
  let formulaCount = 0;

  for (let j = 0; j < cols; j++) {
    const letter = columnLetter(b.columnIndex + j);
    let values = 0;
    let formulas = 0;
    let labels = 0;
    const byShape = new Map<string, FormulaGroup>();
    for (let i = start; i < rows; i++) {
      const f = b.formulas[i][j];
      if (isFormula(f)) {
        formulas++;
        const shape = String(b.formulasR1C1[i]?.[j] ?? f);
        const sheetRow = b.rowIndex + i + 1;
        const g = byShape.get(shape);
        if (g) {
          g.count++;
          g.lastRow = sheetRow;
        } else {
          const refs = formulaRefs(f);
          refs.sheets.forEach((s) => otherSheets.add(s));
          refs.names.forEach((n) => names.add(n));
          const deps = [...refs.columns.filter((c) => c !== letter).map((c) => `column ${c}`), ...refs.sheets.map((s) => `sheet "${s}"`), ...refs.names.map((n) => `name ${n}`)];
          byShape.set(shape, { column: letter, header: headerOf(j), example: f, firstRow: sheetRow, lastRow: sheetRow, count: 1, dependsOn: Array.from(new Set(deps)) });
        }
      } else if (typeof f === "number") values++;
      else if (typeof f === "string" && f.trim()) labels++;
    }
    formulaCount += formulas;
    const role: ColumnRole["role"] = formulas && values ? "mixed" : formulas ? "calculated" : values ? "input" : labels ? "label" : "empty";
    columns.push({ column: letter, header: headerOf(j), role, values, formulas });
    groups.push(...Array.from(byShape.values()).sort((a, c) => c.count - a.count).slice(0, 6));
  }
  return { headerRow: headerRow === null ? null : b.rowIndex + 1, columns, groups, otherSheets: Array.from(otherSheets), names: Array.from(names), formulaCount };
}

/** Compact text for the AI explanation. */
export function sheetMapDigest(sheet: string, address: string, m: SheetMap): string {
  const col = (c: ColumnRole) => `${c.column}${c.header ? ` "${c.header}"` : ""}`;
  const by = (r: ColumnRole["role"]) => m.columns.filter((c) => c.role === r).map(col);
  const lines = [
    `Sheet "${sheet}", used range ${address}, ${m.formulaCount} formulas${m.headerRow ? `, headers in row ${m.headerRow}` : ""}.`,
    `Input columns (typed numbers): ${by("input").join(", ") || "none"}`,
    `Label / text columns: ${by("label").join(", ") || "none"}`,
    `Calculated columns: ${by("calculated").join(", ") || "none"}`,
    `Mixed (formulas and typed values): ${by("mixed").join(", ") || "none"}`,
    m.otherSheets.length ? `Reads other sheets: ${m.otherSheets.join(", ")}` : "",
    m.names.length ? `Uses defined names: ${m.names.join(", ")}` : "",
  ].filter(Boolean);
  const patterns = m.groups.slice(0, 80).map((g) => `- ${g.column}${g.header ? ` "${g.header}"` : ""} rows ${g.firstRow}–${g.lastRow} (${g.count}×): ${g.example}${g.dependsOn.length ? ` ← ${g.dependsOn.join(", ")}` : ""}`);
  return [...lines, "", "Formula patterns (column, rows, count, example, depends on):", ...(patterns.length ? patterns : ["- none"])].join("\n");
}

export const EXPLAIN_SHEET_PROMPT = `Explain how this Excel sheet works, for an engineer who didn't build it. Use the map below (computed from the real formulas). Write Markdown:
1. **Purpose** – what the sheet calculates, in 1–2 sentences (infer from headers and formulas).
2. **Inputs** – which columns/cells are typed data or parameters.
3. **Calculation flow** – step by step, which columns feed which, with the key formulas explained in plain words (units where obvious, e.g. bbl/d, psi).
4. **Outputs** – the main results.
5. **Risks & suggestions** – mixed columns, hard-coded numbers, links to other sheets, anything fragile.
Don't invent columns or formulas that aren't in the map.`;
