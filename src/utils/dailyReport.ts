/**
 * Daily production report: per-well rates on the report date vs the previous date, totals and
 * highlights. Pure functions (no Office) – unit tested.
 */

export interface DailyMapping {
  date: number;
  well: number;
  oil: number;
  water?: number;
  gas?: number;
}

export interface WellDay {
  well: string;
  oil: number | null;
  water: number | null;
  gas: number | null;
  /** Oil + water. */
  liquid: number | null;
  /** Water cut, %. */
  waterCut: number | null;
  /** Gas-oil ratio in the gas unit per oil unit (e.g. scf/bbl when gas is in scf/d, Mscf/bbl for Mscf/d). */
  gor: number | null;
  oilPrev: number | null;
  /** Oil change vs the previous date, %. */
  oilChangePct: number | null;
  waterCutPrev: number | null;
}

export interface DailySummary {
  /** Excel date serials. */
  reportDate: number;
  previousDate: number | null;
  wells: WellDay[];
  totals: { oil: number; water: number; gas: number; liquid: number; waterCut: number | null; oilPrev: number | null; oilChangePct: number | null };
  /** Wells that have data on earlier dates but none on the report date. */
  missingWells: string[];
  highlights: string[];
}

const num = (v: unknown): number | null => {
  if (typeof v === "number" && Number.isFinite(v)) return v;
  if (typeof v === "string" && v.trim() !== "") {
    const n = Number(v.replace(/,/g, "").trim());
    return Number.isFinite(n) ? n : null;
  }
  return null;
};

/** Excel serial (days since 1899-12-30) from a serial number or a date text. */
export function toSerial(v: unknown): number | null {
  if (typeof v === "number" && Number.isFinite(v)) return v > 1000 && v < 200000 ? Math.floor(v) : null;
  if (typeof v === "string" && v.trim()) {
    const t = Date.parse(v.trim());
    if (Number.isFinite(t)) {
      const d = new Date(t);
      return Math.round((Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()) - Date.UTC(1899, 11, 30)) / 86_400_000);
    }
  }
  return null;
}

export function serialToIso(serial: number): string {
  return new Date(Date.UTC(1899, 11, 30) + serial * 86_400_000).toISOString().slice(0, 10);
}

/** "13 Oct 2026" */
export function serialLabel(serial: number): string {
  const d = new Date(Date.UTC(1899, 11, 30) + serial * 86_400_000);
  return `${d.getUTCDate()} ${["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"][d.getUTCMonth()]} ${d.getUTCFullYear()}`;
}

/** Guesses the column roles from header names. -1 / undefined when not found. */
export function guessDailyMapping(headers: string[]): { date: number; well: number; oil: number; water?: number; gas?: number } {
  const find = (re: RegExp, not?: RegExp) => headers.findIndex((h) => re.test(h) && !(not && not.test(h)));
  const water = find(/\b(water|wtr|bwpd)\b/i, /cut|wc\b|%/i);
  const gas = find(/\b(gas|mscf|mmscf|scf)\b/i, /ratio|gor|lift|inj/i);
  return {
    date: find(/\b(date|day|time)\b/i),
    well: find(/\b(well|wellname|well name|uwi|string|completion)\b/i),
    oil: find(/\b(oil|bopd|net)\b/i, /cut|ratio|%/i),
    water: water >= 0 ? water : undefined,
    gas: gas >= 0 ? gas : undefined,
  };
}

const r1 = (n: number) => Math.round(n * 10) / 10;

/**
 * Builds the summary. Rows are data rows (no header). When several rows exist for the same well and
 * date their rates are added (e.g. several strings / tests). `reportDate` defaults to the latest date.
 */
export function buildDailySummary(rows: unknown[][], m: DailyMapping, reportDate?: number): DailySummary {
  type Acc = { oil: number | null; water: number | null; gas: number | null };
  const byWell = new Map<string, Map<number, Acc>>();
  const dates = new Set<number>();
  const add = (a: number | null, b: number | null) => (b === null ? a : (a ?? 0) + b);
  for (const r of rows) {
    const d = toSerial(r[m.date]);
    const w = String(r[m.well] ?? "").trim();
    if (d === null || !w) continue;
    dates.add(d);
    const days = byWell.get(w) ?? new Map<number, Acc>();
    byWell.set(w, days);
    const cur = days.get(d) ?? { oil: null, water: null, gas: null };
    cur.oil = add(cur.oil, num(r[m.oil]));
    if (m.water !== undefined) cur.water = add(cur.water, num(r[m.water]));
    if (m.gas !== undefined) cur.gas = add(cur.gas, num(r[m.gas]));
    days.set(d, cur);
  }
  const sorted = Array.from(dates).sort((a, b) => a - b);
  if (!sorted.length) throw new Error("No rows with both a date and a well name were found.");
  const rep = reportDate !== undefined && dates.has(reportDate) ? reportDate : sorted[sorted.length - 1];
  const prevDates = sorted.filter((d) => d < rep);
  const previousDate = prevDates.length ? prevDates[prevDates.length - 1] : null;

  const wells: WellDay[] = [];
  const missingWells: string[] = [];
  for (const [well, days] of Array.from(byWell.entries()).sort((a, b) => a[0].localeCompare(b[0], undefined, { numeric: true }))) {
    const today = days.get(rep);
    // "Previous" = this well's latest earlier day (wells aren't always tested daily).
    const earlier = Array.from(days.keys()).filter((d) => d < rep).sort((a, b) => b - a)[0];
    const prev = earlier !== undefined ? days.get(earlier) : undefined;
    if (!today) {
      if (prev) missingWells.push(well);
      continue;
    }
    const liquid = today.oil === null && today.water === null ? null : (today.oil ?? 0) + (today.water ?? 0);
    const wc = (o: number | null, w: number | null) => (w !== null && o !== null && o + w > 0 ? (w / (o + w)) * 100 : null);
    wells.push({
      well,
      oil: today.oil,
      water: today.water,
      gas: today.gas,
      liquid,
      waterCut: wc(today.oil, today.water),
      gor: today.gas !== null && today.oil ? today.gas / today.oil : null,
      oilPrev: prev?.oil ?? null,
      oilChangePct: prev?.oil && today.oil !== null ? ((today.oil - prev.oil) / prev.oil) * 100 : null,
      waterCutPrev: prev ? wc(prev.oil, prev.water) : null,
    });
  }

  const sum = (k: "oil" | "water" | "gas" | "oilPrev") => wells.reduce((s, w) => s + (w[k] ?? 0), 0);
  const oil = sum("oil");
  const water = sum("water");
  const oilPrevComparable = wells.filter((w) => w.oilPrev !== null);
  const oilPrevTotal = oilPrevComparable.length ? sum("oilPrev") : null;
  const oilNowComparable = oilPrevComparable.reduce((s, w) => s + (w.oil ?? 0), 0);
  const totals = {
    oil,
    water,
    gas: sum("gas"),
    liquid: oil + water,
    waterCut: oil + water > 0 && m.water !== undefined ? (water / (oil + water)) * 100 : null,
    oilPrev: oilPrevTotal,
    oilChangePct: oilPrevTotal ? ((oilNowComparable - oilPrevTotal) / oilPrevTotal) * 100 : null,
  };

  const highlights: string[] = [];
  if (totals.oilChangePct !== null && Math.abs(totals.oilChangePct) >= 2)
    highlights.push(`Total oil ${totals.oilChangePct > 0 ? "up" : "down"} ${r1(Math.abs(totals.oilChangePct))}% vs the previous reading of the same wells.`);
  for (const w of wells) {
    if (w.oilChangePct !== null && w.oilChangePct <= -20) highlights.push(`${w.well}: oil down ${r1(-w.oilChangePct)}% (${r1(w.oilPrev ?? 0)} → ${r1(w.oil ?? 0)}).`);
    if (w.oilChangePct !== null && w.oilChangePct >= 20) highlights.push(`${w.well}: oil up ${r1(w.oilChangePct)}% (${r1(w.oilPrev ?? 0)} → ${r1(w.oil ?? 0)}).`);
    if (w.waterCut !== null && w.waterCutPrev !== null && w.waterCut - w.waterCutPrev >= 5)
      highlights.push(`${w.well}: water cut up ${r1(w.waterCut - w.waterCutPrev)} points to ${r1(w.waterCut)}%.`);
    if (w.oil === 0) highlights.push(`${w.well}: no oil on the report date (shut in or down?).`);
  }
  if (missingWells.length) highlights.push(`No data on the report date for: ${missingWells.slice(0, 15).join(", ")}${missingWells.length > 15 ? "…" : ""}.`);
  return { reportDate: rep, previousDate, wells, totals, missingWells, highlights };
}

/** Table rows for the report sheet / Word (header first). */
export function summaryTable(s: DailySummary, names: { oil: string; water?: string; gas?: string }): Array<Array<string | number>> {
  const hasW = names.water !== undefined;
  const hasG = names.gas !== undefined;
  const header = ["Well", names.oil, ...(hasW ? [names.water as string, "Water cut %"] : []), ...(hasG ? [names.gas as string, "GOR"] : []), "Prev. oil", "Oil change %"];
  const cell = (v: number | null, digits = 1) => (v === null ? "" : Math.round(v * 10 ** digits) / 10 ** digits);
  const rows = s.wells.map((w) => [
    w.well,
    cell(w.oil),
    ...(hasW ? [cell(w.water), cell(w.waterCut)] : []),
    ...(hasG ? [cell(w.gas), cell(w.gor, 2)] : []),
    cell(w.oilPrev),
    cell(w.oilChangePct),
  ]);
  const t = s.totals;
  const total = ["Total", cell(t.oil), ...(hasW ? [cell(t.water), cell(t.waterCut)] : []), ...(hasG ? [cell(t.gas), t.oil ? cell(t.gas / t.oil, 2) : ""] : []), cell(t.oilPrev), cell(t.oilChangePct)];
  return [header, ...rows, total];
}

/** GitHub-Markdown table from rows (header first). */
export function toMarkdownTable(rows: Array<Array<string | number>>): string {
  if (!rows.length) return "";
  const esc = (v: string | number) => String(v).replace(/\|/g, "\\|");
  const [h, ...body] = rows;
  return [`| ${h.map(esc).join(" | ")} |`, `| ${h.map(() => "---").join(" | ")} |`, ...body.map((r) => `| ${r.map(esc).join(" | ")} |`)].join("\n");
}
