/**
 * "Data validation from AI": column profiles sent to Gemini, the rules it suggests, and a local check
 * that counts the values breaking each rule. Pure functions (no Office) – unit tested.
 */

export interface ColumnProfile {
  index: number;
  name: string;
  kind: "number" | "text" | "date" | "empty" | "mixed";
  count: number;
  blanks: number;
  min: number | null;
  max: number | null;
  mean: number | null;
  /** 1st / 99th percentile (numbers). */
  p01: number | null;
  p99: number | null;
  /** Most common text values (text columns), up to 12. */
  top: Array<{ value: string; n: number }>;
  distinct: number;
}

export interface ColumnRule {
  column: string;
  type: "number" | "list" | "none";
  min?: number | null;
  max?: number | null;
  allowed?: string[];
  reason: string;
}

const numberOf = (v: unknown): number | null => {
  if (typeof v === "number" && Number.isFinite(v)) return v;
  if (typeof v === "string" && /^\s*-?[\d,]*\.?\d+([eE][-+]?\d+)?\s*%?\s*$/.test(v)) {
    const n = Number(v.replace(/[,%\s]/g, ""));
    return Number.isFinite(n) ? n : null;
  }
  return null;
};

const blank = (v: unknown) => v === null || v === undefined || (typeof v === "string" && v.trim() === "");

function pct(sorted: number[], p: number): number | null {
  if (!sorted.length) return null;
  const k = Math.min(sorted.length - 1, Math.max(0, Math.round((sorted.length - 1) * p)));
  return sorted[k];
}

/** `rows` = data rows (no header); `texts` (optional, the displayed text) tells date columns apart from numbers. */
export function profileColumns(headers: string[], rows: unknown[][], texts?: string[][]): ColumnProfile[] {
  return headers.map((name, j) => {
    const vals = rows.map((r) => r[j]);
    const nonBlank = vals.filter((v) => !blank(v));
    const nums = nonBlank.map(numberOf).filter((n): n is number => n !== null);
    const words = nonBlank.filter((v) => numberOf(v) === null).map((v) => String(v).trim());
    const isDate = !!texts && nums.length > 0 && rows.every((r, i) => typeof r[j] !== "number" || (r[j] as number) > 3000 && /[/\-.]|[A-Za-z]{3}/.test(texts[i]?.[j] ?? ""));
    const kind: ColumnProfile["kind"] = !nonBlank.length ? "empty" : isDate ? "date" : nums.length >= nonBlank.length * 0.9 ? "number" : words.length >= nonBlank.length * 0.9 ? "text" : "mixed";
    const sorted = [...nums].sort((a, b) => a - b);
    const freq = new Map<string, number>();
    words.forEach((t) => freq.set(t, (freq.get(t) ?? 0) + 1));
    return {
      index: j,
      name: String(name ?? "").trim() || `Column ${j + 1}`,
      kind,
      count: nonBlank.length,
      blanks: vals.length - nonBlank.length,
      min: sorted.length ? sorted[0] : null,
      max: sorted.length ? sorted[sorted.length - 1] : null,
      mean: sorted.length ? sorted.reduce((a, b) => a + b, 0) / sorted.length : null,
      p01: pct(sorted, 0.01),
      p99: pct(sorted, 0.99),
      top: Array.from(freq.entries())
        .sort((a, b) => b[1] - a[1])
        .slice(0, 12)
        .map(([value, n]) => ({ value, n })),
      distinct: new Set(nonBlank.map((v) => String(v).trim())).size,
    };
  });
}

/** Compact text for the prompt. */
export function profilesForPrompt(ps: ColumnProfile[]): string {
  const f = (n: number | null) => (n === null ? "–" : +n.toPrecision(6));
  return ps
    .map((p) =>
      p.kind === "number"
        ? `- "${p.name}": numbers, n=${p.count}, blanks=${p.blanks}, min=${f(p.min)}, p1=${f(p.p01)}, mean=${f(p.mean)}, p99=${f(p.p99)}, max=${f(p.max)}`
        : p.kind === "text" || p.kind === "mixed"
          ? `- "${p.name}": ${p.kind}, n=${p.count}, distinct=${p.distinct}, most common: ${p.top.map((t) => `${t.value} (${t.n})`).join(", ")}`
          : `- "${p.name}": ${p.kind}`
    )
    .join("\n");
}

/** Cleans the model's rules: known columns only, sensible numbers, short lists. */
export function sanitizeRules(raw: unknown, profiles: ColumnProfile[]): ColumnRule[] {
  const list = Array.isArray(raw) ? raw : [];
  const out: ColumnRule[] = [];
  for (const r of list as Array<Partial<ColumnRule>>) {
    const p = profiles.find((x) => x.name.toLowerCase() === String(r?.column ?? "").trim().toLowerCase());
    if (!p || out.some((o) => o.column === p.name)) continue;
    const reason = typeof r.reason === "string" ? r.reason.slice(0, 200) : "";
    if (r.type === "number") {
      const min = typeof r.min === "number" && Number.isFinite(r.min) ? r.min : null;
      const max = typeof r.max === "number" && Number.isFinite(r.max) ? r.max : null;
      if (min === null && max === null) continue;
      out.push({ column: p.name, type: "number", min, max: min !== null && max !== null && max < min ? min : max, reason });
    } else if (r.type === "list" && Array.isArray(r.allowed)) {
      const allowed = Array.from(new Set(r.allowed.map((a) => String(a).trim()).filter(Boolean))).slice(0, 40);
      if (allowed.length) out.push({ column: p.name, type: "list", allowed, reason });
    }
  }
  return out;
}

/** 0-based row indexes (within `rows`) whose value breaks the rule. Blanks never break a rule. */
export function findViolations(rows: unknown[][], col: number, rule: ColumnRule): number[] {
  const bad: number[] = [];
  rows.forEach((r, i) => {
    const v = r[col];
    if (blank(v)) return;
    if (rule.type === "number") {
      const n = numberOf(v);
      if (n === null || (rule.min !== null && rule.min !== undefined && n < rule.min) || (rule.max !== null && rule.max !== undefined && n > rule.max)) bad.push(i);
    } else if (rule.type === "list") {
      const t = String(v).trim().toLowerCase();
      if (!(rule.allowed ?? []).some((a) => a.toLowerCase() === t)) bad.push(i);
    }
  });
  return bad;
}

export const RULES_SCHEMA: Record<string, unknown> = {
  type: "OBJECT",
  properties: {
    rules: {
      type: "ARRAY",
      items: {
        type: "OBJECT",
        properties: {
          column: { type: "STRING" },
          type: { type: "STRING" },
          min: { type: "NUMBER" },
          max: { type: "NUMBER" },
          allowed: { type: "ARRAY", items: { type: "STRING" } },
          reason: { type: "STRING" },
        },
        required: ["column", "type", "reason"],
      },
    },
  },
  required: ["rules"],
};

export const RULES_SYSTEM_PROMPT = `You are Zee AI, a data-quality engineer for oil & gas production, well-test and ESP data in Excel.
For each column you get a statistical profile. Propose a validation rule per column that a careful engineer would accept:
- "number" with physical/engineering limits (min and/or max). Use domain knowledge, e.g. water cut 0–100 (%) or 0–1 (fraction – decide from the data), rates ≥ 0, frequency 30–70 Hz, API gravity 5–70, pressures ≥ 0 and below a plausible maximum for the unit, temperatures in a plausible range, GOR ≥ 0. Prefer physical limits over the observed min/max; never set limits that most existing values break.
- "list" for categorical columns with a small fixed vocabulary (status, well type, field…): the allowed values exactly as they should be spelled; leave out obvious typos so they get flagged.
- "none" for IDs, free text, dates and columns where no rule makes sense.
"reason": one short sentence (units, why). Use the exact column names given.`;
