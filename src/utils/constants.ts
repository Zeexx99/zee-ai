/**
 * Hard-coded numbers → named inputs: collect numeric literals in a sheet's formulas, suggest names,
 * validate them, and rewrite formulas to use the names. Pure – unit tested.
 */
import { a1 } from "./a1";
import { numericLiterals } from "./formulaAudit";

export interface ConstantUse {
  value: string;
  count: number;
  /** A1 addresses of the cells using it. */
  cells: string[];
  example: string;
  /** Header of the column where it's used most (for naming). */
  header: string;
}

/** Well-known oilfield / unit constants. */
const KNOWN: Record<string, { name: string; note: string }> = {
  "5.615": { name: "FT3_PER_BBL", note: "cubic feet per barrel" },
  "5.6146": { name: "FT3_PER_BBL", note: "cubic feet per barrel" },
  "0.1781": { name: "BBL_PER_FT3", note: "barrels per cubic foot" },
  "6.2898": { name: "BBL_PER_M3", note: "barrels per cubic metre" },
  "6.29": { name: "BBL_PER_M3", note: "barrels per cubic metre" },
  "0.159": { name: "M3_PER_BBL", note: "cubic metres per barrel" },
  "0.158987": { name: "M3_PER_BBL", note: "cubic metres per barrel" },
  "42": { name: "GAL_PER_BBL", note: "US gallons per barrel" },
  "0.433": { name: "PSI_PER_FT_WATER", note: "fresh-water gradient, psi/ft" },
  "0.4335": { name: "PSI_PER_FT_WATER", note: "fresh-water gradient, psi/ft" },
  "0.052": { name: "PSI_PER_FT_PER_PPG", note: "psi/ft per ppg mud weight" },
  "14.7": { name: "ATM_PSI", note: "atmospheric pressure, psi" },
  "14.696": { name: "ATM_PSI", note: "atmospheric pressure, psi" },
  "0.3048": { name: "M_PER_FT", note: "metres per foot" },
  "3.28084": { name: "FT_PER_M", note: "feet per metre" },
  "3.2808": { name: "FT_PER_M", note: "feet per metre" },
  "14.5038": { name: "PSI_PER_BAR", note: "psi per bar" },
  "6.89476": { name: "KPA_PER_PSI", note: "kPa per psi" },
  "141.5": { name: "API_A", note: "API gravity constant 141.5" },
  "131.5": { name: "API_B", note: "API gravity constant 131.5" },
  "1.8": { name: "F_PER_C", note: "°F per °C" },
  "459.67": { name: "RANKINE_OFFSET", note: "°F to °R" },
  "273.15": { name: "KELVIN_OFFSET", note: "°C to K" },
  "0.7457": { name: "KW_PER_HP", note: "kW per horsepower" },
  "1.341": { name: "HP_PER_KW", note: "horsepower per kW" },
  "3960": { name: "HP_HEAD_CONST", note: "hydraulic HP = Q·H·SG / 3960" },
  "135770": { name: "HHP_CONST_BBL", note: "HHP = Q[bbl/d]·H[ft]·SG / 135,770" },
  "1000": { name: "THOUSAND", note: "unit scaling" },
  "35.3147": { name: "FT3_PER_M3", note: "cubic feet per cubic metre" },
  "0.0283168": { name: "M3_PER_FT3", note: "cubic metres per cubic foot" },
};

export function knownConstant(value: string): { name: string; note: string } | null {
  const n = Number(value);
  for (const [k, v] of Object.entries(KNOWN)) if (Number(k) === n) return v;
  return null;
}

/** Valid Excel name: starts with a letter or _, letters/digits/_/. only, not a cell reference, not reserved. */
export function isValidName(name: string): boolean {
  if (!/^[A-Za-z_][A-Za-z0-9_.]{0,254}$/.test(name)) return false;
  if (/^[A-Za-z]{1,3}\d+$/.test(name)) return false; // looks like A1 / ABC12
  if (/^[RrCc]$|^[Rr]\d*[Cc]\d*$/.test(name)) return false; // R1C1 style
  if (/^(TRUE|FALSE)$/i.test(name)) return false;
  return true;
}

export function suggestName(value: string, header: string, taken: Set<string>): string {
  const known = knownConstant(value)?.name;
  let base =
    known ??
    (header
      ? `K_${header
          .replace(/\(.*?\)/g, "")
          .trim()
          .replace(/[^A-Za-z0-9]+/g, "_")
          .replace(/^_+|_+$/g, "")
          .slice(0, 30)
          .toUpperCase()}`
      : `K_${value.replace(/\./g, "_").replace(/-/g, "M")}`);
  if (!isValidName(base)) base = `K_${base.replace(/[^A-Za-z0-9_]/g, "_")}`;
  let name = base;
  for (let i = 2; taken.has(name.toLowerCase()) || !isValidName(name); i++) name = `${base}_${i}`;
  return name;
}

export interface ConstBlock {
  rowIndex: number;
  columnIndex: number;
  formulas: unknown[][];
  headers: string[];
}

/** Numeric literals used in formulas, most used first (counts exclude ROUND(…, 2)-style counts). */
export function collectConstants(b: ConstBlock, max = 60): ConstantUse[] {
  const map = new Map<string, { cells: string[]; example: string; headers: Map<string, number> }>();
  b.formulas.forEach((row, i) =>
    row.forEach((f, j) => {
      if (typeof f !== "string" || !f.startsWith("=")) return;
      for (const lit of new Set(numericLiterals(f))) {
        const key = String(Number(lit));
        const e = map.get(key) ?? { cells: [], example: f, headers: new Map<string, number>() };
        map.set(key, e);
        e.cells.push(a1(b.rowIndex + i, b.columnIndex + j));
        const h = b.headers[j] ?? "";
        if (h) e.headers.set(h, (e.headers.get(h) ?? 0) + 1);
      }
    })
  );
  return Array.from(map.entries())
    .map(([value, e]) => ({
      value,
      count: e.cells.length,
      cells: e.cells,
      example: e.example,
      header: Array.from(e.headers.entries()).sort((a, c) => c[1] - a[1])[0]?.[0] ?? "",
    }))
    .sort((a, c) => c.count - a.count || Number(a.value) - Number(c.value))
    .slice(0, max);
}

/**
 * Replaces standalone occurrences of the number `literal` in a formula with `name`, skipping strings,
 * sheet names, cell references and function names. Returns the new formula and how many were replaced.
 */
export function replaceLiteral(formula: string, literal: string, name: string): { formula: string; n: number } {
  const target = Number(literal);
  let out = "";
  let n = 0;
  let i = 0;
  const f = formula;
  while (i < f.length) {
    const c = f[i];
    if (c === '"') {
      let j = i + 1;
      while (j < f.length && !(f[j] === '"' && f[j + 1] !== '"')) j += f[j] === '"' ? 2 : 1;
      out += f.slice(i, j + 1);
      i = j + 1;
      continue;
    }
    if (c === "'") {
      let j = i + 1;
      while (j < f.length && !(f[j] === "'" && f[j + 1] !== "'")) j += f[j] === "'" ? 2 : 1;
      out += f.slice(i, j + 1);
      i = j + 1;
      continue;
    }
    // A word (function, name, sheet, cell reference like $AB$12): copy it whole.
    const word = f.slice(i).match(/^\$?[A-Za-z_][A-Za-z0-9_.]*\$?\d*/);
    if (word && /[A-Za-z_$]/.test(c)) {
      out += word[0];
      i += word[0].length;
      continue;
    }
    const num = f.slice(i).match(/^\d+(?:\.\d+)?(?:[eE][-+]?\d+)?/);
    if (num) {
      const prev = f[i - 1] ?? "";
      const after = f[i + num[0].length] ?? "";
      const standalone = !/[A-Za-z0-9_.$:]/.test(prev) && !/[A-Za-z0-9_.(:]/.test(after);
      if (standalone && Number(num[0]) === target) {
        out += name;
        n++;
      } else out += num[0];
      i += num[0].length;
      continue;
    }
    out += c;
    i++;
  }
  return { formula: out, n };
}
