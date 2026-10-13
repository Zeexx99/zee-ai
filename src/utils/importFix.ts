/**
 * Fixer for messy imported data: Arabic digits, numbers stored as text (thousands separators, %,
 * (negatives)), units stuck to numbers ("1250 bbl") and mixed date formats. Decided per column so a
 * column ends up with one type. Pure – unit tested.
 */

export type DateOrder = "DMY" | "MDY";

export function latinDigits(s: string): string {
  return s
    .replace(/[٠-٩]/g, (d) => String(d.charCodeAt(0) - 0x0660))
    .replace(/[۰-۹]/g, (d) => String(d.charCodeAt(0) - 0x06f0))
    .replace(/٫/g, ".")
    .replace(/[٬،]/g, ",")
    .replace(/[   ]/g, " ")
    .replace(/−/g, "-");
}

export interface LooseNumber {
  value: number;
  unit: string;
}

/** "1,250.5 bbl/d" → 1250.5 + "bbl/d"; "(12)" → -12; "35 %" → 35 + "%"; "1,25" → 1.25. null when not a number. */
export function parseLooseNumber(raw: string): LooseNumber | null {
  const s = latinDigits(raw).trim();
  const m = s.match(/^([+-]?)\(?\s*([+-]?)(\d[\d,\s]*(?:\.\d+)?|\.\d+)\s*\)?\s*(-?)\s*([A-Za-z°%µ/³²][A-Za-z0-9°%µ/³²._ ]{0,14})?$/);
  if (!m) return null;
  let digits = m[3].replace(/\s/g, "");
  const comma = digits.match(/,(\d+)$/);
  if (comma && !digits.includes(".") && comma[1].length !== 3 && (digits.match(/,/g) ?? []).length === 1) digits = digits.replace(",", ".");
  else digits = digits.replace(/,/g, "");
  let value = Number(digits);
  if (!Number.isFinite(value)) return null;
  const negative = m[1] === "-" || m[2] === "-" || m[4] === "-" || (s.startsWith("(") && s.includes(")"));
  if (negative) value = -value;
  return { value, unit: (m[5] ?? "").trim() };
}

const MONTHS: Record<string, number> = { jan: 1, feb: 2, mar: 3, apr: 4, may: 5, jun: 6, jul: 7, aug: 8, sep: 9, sept: 9, oct: 10, nov: 11, dec: 12 };
const EPOCH = Date.UTC(1899, 11, 30);
const monthOf = (name: string): number | undefined => MONTHS[name.slice(0, 4).toLowerCase()] ?? MONTHS[name.slice(0, 3).toLowerCase()];

const serialOf = (y: number, m: number, d: number) => {
  if (y < 100) y += y < 70 ? 2000 : 1900;
  if (m < 1 || m > 12 || d < 1 || d > 31) return null;
  const t = Date.UTC(y, m - 1, d);
  const back = new Date(t);
  if (back.getUTCMonth() !== m - 1) return null;
  return Math.round((t - EPOCH) / 86_400_000);
};

/** Date text in common formats → Excel serial (time of day dropped). null when not a date. */
export function parseLooseDate(raw: string, order: DateOrder): number | null {
  const s = latinDigits(raw).trim().replace(/\s+\d{1,2}:\d{2}(:\d{2})?(\s*[AaPp][Mm])?$/, "");
  let m = s.match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})$/);
  if (m) return serialOf(+m[1], +m[2], +m[3]);
  m = s.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{2}|\d{4})$/);
  if (m) {
    const [a, b] = [+m[1], +m[2]];
    const dmy = order === "DMY" ? !(a <= 12 && b > 12) : a > 12;
    return dmy ? serialOf(+m[3], b, a) : serialOf(+m[3], a, b);
  }
  m = s.match(/^(\d{1,2})[-\s/.]([A-Za-z]{3,9})[-\s/.,]*(\d{2}|\d{4})$/);
  if (m) {
    const mo = monthOf(m[2]);
    if (mo) return serialOf(+m[3], mo, +m[1]);
  }
  m = s.match(/^([A-Za-z]{3,9})[-\s/.]+(\d{1,2})(?:st|nd|rd|th)?[,\s]+(\d{4})$/);
  if (m) {
    const mo = monthOf(m[1]);
    if (mo) return serialOf(+m[3], mo, +m[2]);
  }
  return null;
}

/** Day-first unless a value proves month-first (second part > 12). Egypt and most of the world use day-first. */
export function detectDateOrder(texts: string[]): DateOrder {
  let dmy = 0;
  let mdy = 0;
  for (const t of texts) {
    const m = latinDigits(t).trim().match(/^(\d{1,2})[-/.](\d{1,2})[-/.]\d{2,4}/);
    if (!m) continue;
    if (+m[1] > 12) dmy++;
    else if (+m[2] > 12) mdy++;
  }
  return mdy > dmy ? "MDY" : "DMY";
}

export type ColumnKind = "number" | "date" | "text";

export interface ColumnFixPlan {
  index: number;
  header: string;
  kind: ColumnKind;
  /** Main unit found in the column ("" when none). */
  unit: string;
  /** New header when the unit is moved into it. */
  newHeader?: string;
  order: DateOrder;
}

const blank = (v: unknown) => v === null || v === undefined || (typeof v === "string" && v.trim() === "");

export function planColumn(index: number, header: string, values: unknown[], texts: string[]): ColumnFixPlan {
  const nonBlank = values.map((v, i) => ({ v, t: texts[i] ?? String(v ?? "") })).filter((x) => !blank(x.v));
  const order = detectDateOrder(nonBlank.filter((x) => typeof x.v === "string").map((x) => x.v as string));
  let nums = 0;
  let dates = 0;
  const units = new Map<string, number>();
  for (const x of nonBlank) {
    if (typeof x.v === "number") {
      // A real number formatted as a date counts as a date.
      if (/[/\-.]|[A-Za-z]{3}/.test(x.t) && x.v > 3000) dates++;
      else nums++;
      continue;
    }
    const s = String(x.v);
    if (parseLooseDate(s, order) !== null) {
      dates++;
      continue;
    }
    const n = parseLooseNumber(s);
    if (n) {
      nums++;
      if (n.unit) units.set(n.unit, (units.get(n.unit) ?? 0) + 1);
    }
  }
  const total = nonBlank.length || 1;
  const kind: ColumnKind = dates >= total * 0.7 && dates >= nums ? "date" : nums >= total * 0.7 ? "number" : "text";
  const unit = kind === "number" && units.size ? Array.from(units.entries()).sort((a, b) => b[1] - a[1])[0][0] : "";
  const hasUnitInHeader = unit && header.toLowerCase().includes(unit.toLowerCase());
  return { index, header, kind, unit, order, ...(unit && !hasUnitInHeader && header.trim() ? { newHeader: `${header.trim()} (${unit})` } : {}) };
}

export interface CellFix {
  value: unknown;
  changed: boolean;
  numberFormat?: string;
  /** The cell had a different unit than the column's main unit – left unchanged. */
  otherUnit?: string;
}

export function fixCell(v: unknown, plan: ColumnFixPlan): CellFix {
  if (blank(v) || typeof v !== "string") {
    if (typeof v === "string" && v !== "" && !v.trim()) return { value: "", changed: true };
    return { value: v, changed: false };
  }
  const s = v;
  if (plan.kind === "date") {
    const d = parseLooseDate(s, plan.order);
    return d !== null ? { value: d, changed: true, numberFormat: "dd-mmm-yyyy" } : { value: v, changed: false };
  }
  if (plan.kind === "number") {
    const n = parseLooseNumber(s);
    if (!n) return { value: v, changed: false };
    if (n.unit && plan.unit && n.unit.toLowerCase() !== plan.unit.toLowerCase()) return { value: v, changed: false, otherUnit: n.unit };
    return { value: n.value, changed: true };
  }
  const t = latinDigits(s).replace(/\s+/g, " ").trim();
  return t !== s ? { value: t, changed: true } : { value: v, changed: false };
}
