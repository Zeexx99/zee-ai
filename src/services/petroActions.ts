/**
 * Petroleum quick actions built on the pure math in utils/petro.ts and the Excel helpers.
 * Results go to a NEW sheet (never overwriting data) and are undoable (delete sheet).
 */
import {
  addSheet,
  createChart,
  createScatterSeriesChart,
  formatRangeAt,
  readBlockAt,
  readRangeValues,
  readSelectionBlock,
  readSourceColumns,
  writeBlock,
  ExcelUserError,
  type PlotSource,
  type SelectionBlock,
} from "./excelService";
import { analyzeEsp, parseCurve, type EspOptions } from "../utils/espAnalysis";
import { arpsCum, arpsRate, arpsTimeToRate, fitArps, fitArpsByGroup, prepareDca, qcAnalyze, type ArpsFit } from "../utils/petro";
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
  /** Optional well / group column: fit every well separately and write a summary table. */
  wellColumn?: number;
}

export interface DcaByWellResult {
  sheet: string;
  wells: number;
  fitted: number;
  failed: number;
}

/** Decline-curve fit for every well in the selection; one summary row per well. */
export async function runDeclineByWell(o: DcaOptions): Promise<DcaByWellResult> {
  if (o.wellColumn === undefined) throw new ExcelUserError("Choose the well column.");
  const blk: SelectionBlock = await readSelectionBlock(200_000);
  const header = hasHeaderRow(blk.values);
  const body = header ? blk.values.slice(1) : blk.values;
  const bodyText = header ? blk.text.slice(1) : blk.text;
  const rawT = body.map((r) => r[o.timeColumn]);
  const dates = looksLikeDates(rawT, bodyText.map((r) => r[o.timeColumn]));
  const fits = fitArpsByGroup(
    body.map((r) => r[o.wellColumn as number]),
    rawT,
    body.map((r) => r[o.rateColumn])
  );
  if (!fits.length) throw new ExcelUserError("The well column is empty.");
  const unit = dates ? "day" : "time unit";
  const rateName = header ? String(blk.values[0][o.rateColumn]).trim() || "Rate" : "Rate";
  const lim = o.economicLimit;

  const sheet = await addSheet(`DCA by well`);
  const top: Cell[][] = [
    [`Decline curve analysis by well – ${rateName}`],
    [`Source: ${blk.address} · ${fits.length} wells · Di per ${unit}${lim ? ` · economic limit ${lim}` : ""}`],
  ];
  await writeBlock(sheet, "A1", top);
  await formatRangeAt("A1", { bold: true }, sheet);
  const table: Cell[][] = [
    [
      "Well",
      "Points",
      "Skipped",
      "Best model",
      "qi",
      `Di (1/${unit})`,
      "b",
      "R²",
      dates ? "Annual decline (%)" : "",
      dates ? "First date" : "First time",
      lim ? `Time to limit (${unit}s)` : "",
      lim ? "EUR from first point" : "",
      "Note",
    ],
    ...fits.map((g): Cell[] => {
      const f = g.best;
      if (!f) return [g.group, g.points, g.dropped, "", "", "", "", "", "", g.t0 || "", "", "", g.error ?? "not fitted"];
      const tl = lim ? arpsTimeToRate(f, lim) : NaN;
      return [
        g.group,
        g.points,
        g.dropped,
        f.model,
        +f.qi.toPrecision(6),
        +f.di.toPrecision(5),
        +f.b.toFixed(3),
        +f.r2.toFixed(4),
        dates ? +((1 - arpsRate(f, 365) / f.qi) * 100).toFixed(1) : "",
        g.t0,
        lim ? (Number.isFinite(tl) ? +tl.toFixed(1) : "never") : "",
        lim ? (Number.isFinite(tl) ? +arpsCum(f, tl).toPrecision(6) : "unbounded") : "",
        f.r2 < 0.7 ? "low R² – check data / pick a period" : "",
      ];
    }),
  ];
  await writeBlock(sheet, "A4", table, { header: true });
  if (dates) await formatRangeAt(`J5:J${4 + fits.length}`, { numberFormat: "yyyy-mm-dd" }, sheet);
  await formatRangeAt(`A4:M4`, { autofit: true }, sheet);
  const fitted = fits.filter((g) => g.best).length;
  return { sheet, wells: fits.length, fitted, failed: fits.length - fitted };
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

/** QC the selection, or any range when `source` is given (agent mode). */
export async function runProductionQc(source?: { address: string; sheet?: string }): Promise<QcResult> {
  const blk = source ? await readBlockAt(source.address, source.sheet, 200_000) : await readSelectionBlock(200_000);
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

/* ================================================================ ESP analysis (calculated, no AI) */

export interface EspColumns {
  rate: number;
  pip: number;
  pd: number;
  hz: number;
  amps?: number;
  waterCut?: number;
  /** Date / time / well-test id column used as the first report column and the trend chart X. */
  x?: number;
}

export interface EspChoice {
  source: PlotSource;
  columns: EspColumns;
  options: Omit<EspOptions, "curve">;
  /** Optional pump curve range: two columns, rate (bbl/d) and head at the base frequency, e.g. "Curves!A2:B40". */
  curveAddress?: string;
}

export interface EspResult {
  sheet: string;
  status: string;
  rows: number;
  flagged: number;
}

export async function runEspAnalysis(c: EspChoice): Promise<EspResult> {
  const curve = c.curveAddress?.trim() ? parseCurve(await readRangeValues(c.curveAddress.trim())) : undefined;
  if (c.curveAddress?.trim() && (!curve || curve.length < 2)) throw new ExcelUserError(`No pump curve found in ${c.curveAddress}: it needs two columns – rate and head – with at least 2 numeric rows.`);
  const opts: EspOptions = { ...c.options, curve };

  const k = c.columns;
  const order = [k.rate, k.pip, k.pd, k.hz, k.amps ?? -1, k.waterCut ?? -1, k.x ?? -1];
  const wanted = order.filter((i) => i >= 0);
  const data = await readSourceColumns(c.source, wanted);
  const col = (i: number | undefined) => (i === undefined || i < 0 ? null : data[wanted.indexOf(i)]);
  const n = data[0]?.length ?? 0;
  const [rate, pip, pd, hz, amps, wc, xs] = [k.rate, k.pip, k.pd, k.hz, k.amps, k.waterCut, k.x].map(col);
  const inputs = Array.from({ length: n }, (_, i) => ({
    rate: rate![i],
    pip: pip![i],
    pd: pd![i],
    hz: hz![i],
    amps: amps?.[i],
    waterCut: wc?.[i],
  }));
  const { rows, summary } = analyzeEsp(inputs, opts);
  if (!summary.validRows) throw new ExcelUserError("No row had numeric rate, intake and discharge pressure and frequency. Check the column choices.");

  const nameOf = (i: number | undefined) => (i === undefined ? "" : c.source.columns.find((cc) => cc.index === i)?.name ?? "");
  const headUnit = opts.stages && opts.stages > 1 ? "ft/stage" : "ft";
  const base = `${opts.baseHz} Hz`;
  const sheet = await addSheet("ESP Analysis");

  const S = summary;
  const top: Cell[][] = [
    ["ESP performance analysis (calculated)", ""],
    [`Source: ${c.source.sheetName}!${c.source.address}`, ""],
    ["Status", S.status],
    ["Rows analysed", `${S.validRows} of ${S.rows}`],
    ["In range / downthrust / upthrust", curve || (opts.minRate !== undefined && opts.maxRate !== undefined) ? `${S.inRange} / ${S.downthrust} / ${S.upthrust}` : "n/a (no curve or range given)"],
    ["Head vs curve – average / latest (%)", S.avgHeadDeviationPct === null ? "n/a (no curve)" : `${S.avgHeadDeviationPct} / ${S.latestHeadDeviationPct}`],
    ["Average motor load (%)", S.avgMotorLoadPct ?? "n/a (no amps / nameplate)"],
    ["Rows with flags", S.flaggedRows],
    ...Object.entries(S.flagCounts).map(([f, cnt]): Cell[] => [`  ${f}`, cnt]),
    [`Assumptions: SG ${opts.api !== undefined && k.waterCut !== undefined ? `from water cut + ${opts.api} °API` : opts.sg}; base ${base}; ${opts.stages ? `${opts.stages} stages` : "curve head = total pump head"}; head-loss flag at −${opts.degradationPct} %.`, ""],
  ];
  await writeBlock(sheet, "A1", top);
  await formatRangeAt("A1", { bold: true }, sheet);
  await formatRangeAt(`A3:A${top.length - 1}`, { bold: true }, sheet);

  const headerRow = top.length + 2; // 1-based
  const table: Cell[][] = [
    [
      nameOf(k.x) || "Row",
      "Rate (bbl/d)",
      "Freq (Hz)",
      "PIP (psi)",
      "Pdis (psi)",
      "ΔP (psi)",
      "SG",
      "TDH (ft)",
      `Rate @ ${base}`,
      `Head @ ${base} (${headUnit})`,
      `Curve head (${headUnit})`,
      "Head vs curve (%)",
      "HHP",
      "Motor load (%)",
      "Range",
      "Flags",
    ],
    ...rows.map((r, i): Cell[] => [
      xs ? ((xs[i] as Cell) ?? "") : i + 1,
      r.rate ?? "",
      r.hz ?? "",
      r.pip ?? "",
      r.pd ?? "",
      r.dp ?? "",
      r.sg ?? "",
      r.tdh ?? "",
      r.rateBase ?? "",
      r.headBase ?? "",
      r.headCurve ?? "",
      r.headDeviationPct ?? "",
      r.hhp ?? "",
      r.motorLoadPct ?? "",
      r.range,
      r.flags.join("; "),
    ]),
  ];
  await writeBlock(sheet, `A${headerRow}`, table, { header: true });
  const lastRow = headerRow + rows.length;
  const xText = xs ? xs.find((v) => v !== "" && v !== null) : undefined;
  if (typeof xText === "number" && xText > 3000 && xText < 120000 && /date|day|time/i.test(nameOf(k.x))) {
    await formatRangeAt(`A${headerRow + 1}:A${lastRow}`, { numberFormat: "yyyy-mm-dd" }, sheet);
  }

  // Pump curve table + charts to the right of the table.
  const anchorCol = "R";
  if (curve) {
    const curveTop = headerRow;
    await writeBlock(sheet, `${anchorCol}${curveTop}`, [[`Curve rate @ ${base}`, `Curve head (${headUnit})`], ...curve.map((p): Cell[] => [p.q, p.h])], { header: true });
    const cEnd = curveTop + curve.length;
    await createScatterSeriesChart(
      sheet,
      "Operating points on the pump curve",
      [
        { name: `Catalog curve @ ${base}`, x: `R${curveTop + 1}:R${cEnd}`, y: `S${curveTop + 1}:S${cEnd}`, style: "smooth" },
        { name: "Measured (corrected)", x: `I${headerRow + 1}:I${lastRow}`, y: `J${headerRow + 1}:J${lastRow}`, style: "points" },
      ],
      "U2",
      { x: `Rate @ ${base} (bbl/d)`, y: `Head (${headUnit})` }
    ).catch(() => undefined);
    await createScatterSeriesChart(
      sheet,
      "Head vs catalog curve (%) – wear trend",
      [{ name: "Head vs curve (%)", x: `A${headerRow + 1}:A${lastRow}`, y: `L${headerRow + 1}:L${lastRow}`, style: "line" }],
      "U24",
      { x: nameOf(k.x) || "Row", y: "%" }
    ).catch(() => undefined);
  } else {
    await createScatterSeriesChart(
      sheet,
      "TDH trend",
      [{ name: "TDH (ft)", x: `A${headerRow + 1}:A${lastRow}`, y: `H${headerRow + 1}:H${lastRow}`, style: "line" }],
      "R2",
      { x: nameOf(k.x) || "Row", y: "TDH (ft)" }
    ).catch(() => undefined);
  }
  return { sheet, status: S.status, rows: S.validRows, flagged: S.flaggedRows };
}
