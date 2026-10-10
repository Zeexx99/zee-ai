/**
 * Petroleum quick actions built on the pure math in utils/petro.ts and the Excel helpers.
 * Results go to a NEW sheet (never overwriting data) and are undoable (delete sheet).
 */
import { addSheet, createChart, formatRangeAt, readSelectionBlock, writeBlock, ExcelUserError, type SelectionBlock } from "./excelService";
import { arpsCum, arpsRate, arpsTimeToRate, fitArps, prepareDca, qcAnalyze, type ArpsFit } from "../utils/petro";
import { a1, quoteSheet } from "../utils/a1";
import type { Cell } from "../utils/markdown";

const isNumeric = (v: unknown) => typeof v === "number" || (typeof v === "string" && v.trim() !== "" && Number.isFinite(Number(v.replace(/,/g, ""))));

/** First row is a header if most of its cells are non-numeric text. */
export function hasHeaderRow(values: unknown[][]): boolean {
  const first = values[0] ?? [];
  const text = first.filter((v) => typeof v === "string" && v.trim() !== "" && !isNumeric(v)).length;
  return text >= Math.ceil(first.length / 2);
}

/** Column looks like Excel dates: numeric serials in a plausible range and displayed with / - or month names. */
function looksLikeDates(values: unknown[], texts: string[]): boolean {
  const sample = values.slice(0, 20).map((v, i) => [v, texts[i]] as const).filter(([v]) => typeof v === "number");
  if (!sample.length) return false;
  return sample.every(([v, t]) => (v as number) > 3000 && (v as number) < 120000 && /[/\-.]|[A-Za-z]{3}/.test(String(t)));
}

export interface DcaOptions {
  rateColumn: number; // index within selection
  timeColumn: number;
  forecastPeriods: number; // extra periods beyond last data point (in typical step units)
  economicLimit?: number;
}

export interface DcaResult {
  sheet: string;
  best: ArpsFit;
  fits: ArpsFit[];
  pointsUsed: number;
  dropped: number;
}

/** Read headers of the current selection (for the DCA dialog). */
export async function getSelectionHeaders(): Promise<{ headers: string[]; hasHeader: boolean; rows: number }> {
  const blk = await readSelectionBlock(200_000);
  const hasHeader = hasHeaderRow(blk.values);
  const headers = (blk.values[0] ?? []).map((h, j) => (hasHeader ? String(h).trim() || `Column ${j + 1}` : `Column ${j + 1}`));
  return { headers, hasHeader, rows: blk.values.length - (hasHeader ? 1 : 0) };
}

export async function runDeclineCurve(o: DcaOptions): Promise<DcaResult> {
  const blk: SelectionBlock = await readSelectionBlock(200_000);
  if ((blk.values[0]?.length ?? 0) < 2) throw new ExcelUserError("Select at least two columns: time/date and rate.");
  const header = hasHeaderRow(blk.values);
  const body = header ? blk.values.slice(1) : blk.values;
  const bodyText = header ? blk.text.slice(1) : blk.text;
  const rawT = body.map((r) => r[o.timeColumn]);
  const rawQ = body.map((r) => r[o.rateColumn]);
  const dates = looksLikeDates(rawT, bodyText.map((r) => r[o.timeColumn]));
  const rateName = header ? String(blk.values[0][o.rateColumn]).trim() || "Rate" : "Rate";
  const timeName = header ? String(blk.values[0][o.timeColumn]).trim() || (dates ? "Date" : "Time") : dates ? "Date" : "Time";

  const prep = prepareDca(rawT, rawQ);
  const fits = fitArps(prep);
  const best = fits[0];

  // Forecast grid: data points, then extend by forecastPeriods typical steps.
  const steps = prep.t.slice(1).map((t, i) => t - prep.t[i]).filter((d) => d > 0).sort((a, b) => a - b);
  const step = steps.length ? steps[Math.floor(steps.length / 2)] : 1;
  const lastT = prep.t[prep.t.length - 1];
  const unit = dates ? "day" : "time unit";

  const sheet = await addSheet(`DCA ${rateName}`);
  const params: Cell[][] = [
    ["Decline curve analysis (Arps)", "", "", "", "", "", "", ""],
    [`Source: ${blk.address}`, "", "", "", "", "", "", ""],
    ["Model", "qi", `Di (1/${unit})`, "b", "R²", "RMSE", o.economicLimit ? `Time to ${o.economicLimit} (${unit}s)` : "", o.economicLimit ? "EUR from first point" : ""],
    ...fits.map((f, i): Cell[] => {
      const tl = o.economicLimit ? arpsTimeToRate(f, o.economicLimit) : NaN;
      return [
        `${f.model}${i === 0 ? " (best)" : ""}`,
        +f.qi.toPrecision(6),
        +f.di.toPrecision(6),
        +f.b.toFixed(3),
        +f.r2.toFixed(4),
        +f.rmse.toPrecision(4),
        o.economicLimit ? (Number.isFinite(tl) ? +tl.toFixed(1) : "never") : "",
        o.economicLimit ? (Number.isFinite(tl) ? +arpsCum(f, tl).toPrecision(6) : "unbounded") : "",
      ];
    }),
    [`Effective annual decline (best): ${dates ? `${((1 - arpsRate(best, 365) / best.qi) * 100).toFixed(1)}%` : "n/a (time not in days)"}`, "", "", "", "", "", "", ""],
  ];
  await writeBlock(sheet, "A1", params);
  await formatRangeAt("A1", { bold: true }, sheet);
  await formatRangeAt("A3:H3", { bold: true, fillColor: "#EDEBFB" }, sheet);

  const table: Cell[][] = [[timeName, `${rateName} (actual)`, `${best.model} fit`, "Cumulative (fit)"]];
  prep.t.forEach((t, i) => table.push([t + prep.t0, prep.q[i], +arpsRate(best, t).toPrecision(6), +arpsCum(best, t).toPrecision(6)]));
  for (let k = 1; k <= o.forecastPeriods; k++) {
    const t = lastT + k * step;
    table.push([t + prep.t0, "", +arpsRate(best, t).toPrecision(6), +arpsCum(best, t).toPrecision(6)]);
  }
  const startRow = params.length + 2; // leave a blank row
  const topLeft = a1(startRow, 0);
  const tableAddr = await writeBlock(sheet, topLeft, table, { header: true });
  if (dates) await formatRangeAt(`A${startRow + 2}:A${startRow + table.length}`, { numberFormat: "yyyy-mm-dd" }, sheet);

  const chartRange = `${quoteSheet(sheet)}!A${startRow + 1}:C${startRow + table.length}`;
  await createChart(chartRange, "scatterLines", `${rateName}: actual vs ${best.model} decline`).catch(() => undefined);
  void tableAddr;
  return { sheet, best, fits, pointsUsed: prep.t.length, dropped: prep.dropped };
}

export interface QcResult {
  sheet: string;
  issues: number;
  columns: number;
}

export async function runProductionQc(): Promise<QcResult> {
  const blk = await readSelectionBlock(200_000);
  const header = hasHeaderRow(blk.values);
  const values = header ? blk.values : [blk.values[0].map((_, j) => `Column ${j + 1}`), ...blk.values];
  const texts = header ? blk.text.slice(1) : blk.text;
  const dataRows = values.slice(1);
  const dateCol = looksLikeDates(dataRows.map((r) => r[0]), texts.map((r) => r[0])) ? 0 : -1;
  const firstDataRow = blk.rowIndex + (header ? 2 : 1); // 1-based sheet row
  const rep = qcAnalyze(values, dateCol, firstDataRow);

  const sheet = await addSheet("QC Report");
  const summary: Cell[][] = [
    ["Production data QC", "", "", "", "", "", "", "", "", "", ""],
    [`Source: ${blk.address} · ${dataRows.length} rows${dateCol === 0 ? " · date column: " + values[0][0] : ""}`, "", "", "", "", "", "", "", "", "", ""],
    ["Column", "Rows", "Missing", "Non-numeric", "Negative", "Zero", "Outliers", "Stuck runs", "Min", "Max", "Mean"],
    ...rep.columns.map((c): Cell[] => [c.name, c.count, c.missing, c.nonNumeric, c.negatives, c.zeros, c.outliers, c.flatRuns, c.min ?? "", c.max ?? "", c.mean === null ? "" : +c.mean.toPrecision(6)]),
  ];
  await writeBlock(sheet, "A1", summary);
  await formatRangeAt("A1", { bold: true }, sheet);
  await formatRangeAt("A3:K3", { bold: true, fillColor: "#EDEBFB" }, sheet);

  const issues: Cell[][] = [["Column", "Issue"]];
  rep.columns.forEach((c) => c.issues.forEach((i) => issues.push([c.name, i])));
  rep.dateGaps.forEach((g) => issues.push([String(values[0][0]), g]));
  rep.duplicateDates.forEach((g) => issues.push([String(values[0][0]), g]));
  if (issues.length === 1) issues.push(["—", "No issues found."]);
  const issuesRow = summary.length + 2;
  await writeBlock(sheet, a1(issuesRow, 0), issues, { header: true });
  return { sheet, issues: issues.length - 1, columns: rep.columns.length };
}

export const ESP_PROMPT = `Act as a senior artificial-lift (ESP) engineer. Review the ESP data in the selection and give a concise diagnosis.

Do the following where the data allows (state which columns you used):
1. Compute pump differential pressure and total dynamic head: TDH (ft) = (P_discharge − P_intake) [psi] / (0.433 × SG). Use SG from the data or assume water-cut-weighted SG and say so.
2. Compare actual head/rate with the expected pump curve at the running frequency (affinity laws: Q ∝ f, H ∝ f², BHP ∝ f³) if a catalog curve or rated point is present; estimate head degradation (pump wear).
3. Check whether the pump runs inside its recommended operating range (downthrust vs upthrust).
4. Flag gas interference / gas lock (PIP near or below bubble point, amp fluctuation, low efficiency), scale/plugging, high motor temperature, current imbalance, underload/overload trips.
5. Summarise as: **Status**, **Key findings** (bullets with numbers), **Likely causes**, **Recommended actions** (e.g. frequency change, choke, gas handling, workover).
If key data is missing, list exactly which measurements are needed. Put any computed table in a markdown table so it can be inserted.`;
