/**
 * Nodal analysis → a new "Nodal" sheet: inputs, IPR + VLP table, operating point(s) and a chart.
 */
import { addSheet, createScatterSeriesChart, formatRangeAt, writeBlock, type ScatterSeriesSpec } from "./excelService";
import { runNodal, SENSITIVITY_LABELS, type NodalInput, type NodalResult } from "../utils/nodal";
import { a1 } from "../utils/a1";
import type { Cell } from "../utils/markdown";

const LAST_KEY = "zee.nodal.last";

export function loadLastNodal(): NodalInput | null {
  try {
    const raw = localStorage.getItem(LAST_KEY);
    return raw ? (JSON.parse(raw) as NodalInput) : null;
  } catch {
    return null;
  }
}
function saveLastNodal(i: NodalInput) {
  try {
    localStorage.setItem(LAST_KEY, JSON.stringify(i));
  } catch {
    /* storage blocked – not important */
  }
}

const r1 = (v: number | null | undefined) => (v === null || v === undefined || !Number.isFinite(v) ? "" : Math.round(v * 10) / 10);

export interface NodalSheetResult {
  sheet: string;
  result: NodalResult;
}

export async function runNodalSheet(input: NodalInput): Promise<NodalSheetResult> {
  const result = runNodal(input);
  saveLastNodal(input);
  const sheet = await addSheet("Nodal Analysis");
  const { ipr, cases } = result;
  const w = input.well;
  const ip = input.ipr;

  const top: Cell[][] = [
    ["Nodal analysis – node at bottom-hole (calculated)", ""],
    ["Inflow (IPR)", ipr.label],
    ["Reservoir pressure (psia)", ip.pr],
    ...(ip.qTest !== undefined && ip.pwfTest !== undefined ? [["Well test", `${ip.qTest} STB/d @ ${ip.pwfTest} psia`] as Cell[]] : []),
    ...(ip.model === "composite" && ip.pb !== undefined ? [["Bubble point (psia)", ip.pb] as Cell[]] : []),
    ["Outflow (VLP)", "Beggs & Brill + Payne correction, vertical well, 50 ft steps"],
    ["Depth – mid perfs (ft)", w.depth],
    ["Tubing ID (in)", w.tubingId],
    ["Wellhead pressure (psia)", w.whp],
    ["GOR (scf/STB)", w.gor],
    ["Water cut (%)", +(w.waterCut * 100).toFixed(1)],
    ["Oil API / gas SG / water SG", `${w.api} / ${w.gasSg} / ${w.waterSg}`],
    ["Temperature wellhead → bottom (°F)", `${w.tWh} → ${w.tBh}`],
    ["", ""],
    ["Case", "Liquid rate (STB/d)", "Oil rate (STB/d)", "Pwf (psia)", "Drawdown (psi)", "Note"],
    ...cases.map((c): Cell[] =>
      c.op
        ? [c.label, r1(c.op.q), r1(c.op.oil), r1(c.op.pwf), r1(c.op.drawdown), c.op.stable ? "" : "Left of the tubing-curve minimum – flow may be unstable (heading / loading)."]
        : [c.label, "", "", "", "", c.note ?? "No operating point"]
    ),
  ];
  await writeBlock(sheet, "A1", top);
  const opHeader = top.length - cases.length; // 1-based row of the "Case" header
  await formatRangeAt("A1", { bold: true }, sheet);
  await formatRangeAt(`A2:A${opHeader - 2}`, { bold: true }, sheet);
  await formatRangeAt(`A${opHeader}:F${opHeader}`, { bold: true, fillColor: "#EDEBFB" }, sheet);

  // Curve table: q | IPR Pwf | VLP per case.
  const tRow = top.length + 2; // 0-based row index of its header
  const header: Cell[] = ["Liquid rate (STB/d)", "IPR Pwf (psia)", ...cases.map((c) => `VLP ${c.label} (psia)`)];
  const rows: Cell[][] = result.rates.map((q, k) => [q, r1(cases[0].curve[k].pwfIpr), ...cases.map((c) => r1(c.curve[k].pwfVlp))]);
  await writeBlock(sheet, a1(tRow, 0), [header, ...rows], { header: true });

  // Operating points as their own little x/y block (one series each).
  const opCol = header.length + 1;
  const opRows: Cell[][] = [["Operating point", "q", "Pwf"], ...cases.map((c): Cell[] => [c.label, c.op ? r1(c.op.q) : "", c.op ? r1(c.op.pwf) : ""])];
  await writeBlock(sheet, a1(tRow, opCol), opRows, { header: true });

  const first = tRow + 1;
  const last = tRow + rows.length;
  const col = (j: number) => `${a1(first, j)}:${a1(last, j)}`;
  const series: ScatterSeriesSpec[] = [
    { name: "IPR", x: col(0), y: col(1), style: "smooth" },
    ...cases.map((c, k): ScatterSeriesSpec => ({ name: `VLP ${c.label}`, x: col(0), y: col(2 + k), style: "smooth" })),
    ...cases
      .map((c, k) => ({ c, k }))
      .filter(({ c }) => c.op)
      .map(({ c, k }): ScatterSeriesSpec => ({ name: `Op. point ${c.label}`, x: a1(tRow + 1 + k, opCol + 1), y: a1(tRow + 1 + k, opCol + 2), style: "points" })),
  ];
  await createScatterSeriesChart(sheet, "Nodal analysis – IPR vs VLP", series, a1(0, 7), { x: "Liquid rate (STB/d)", y: "Pwf (psia)" });
  return { sheet, result };
}

export function nodalSummary(r: NodalResult): string {
  return r.cases.map((c) => (c.op ? `${c.label}: ${Math.round(c.op.q)} STB/d at Pwf ${Math.round(c.op.pwf)} psia` : `${c.label}: ${c.note ?? "no operating point"}`)).join(" · ");
}

export { SENSITIVITY_LABELS };
