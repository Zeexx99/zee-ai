/**
 * ESP failure log analysis: run life, MTBF, failures by cause / component / model / well and a
 * monthly trend. Pure functions (no Office) – unit tested.
 */
import { serialToIso, toSerial } from "./dailyReport";

export const ESP_LOG_HEADERS = ["Well", "ESP Model", "Install Date", "Pull / Fail Date", "Status", "Run Life (days)", "Failed Component", "Failure Cause", "Notes"];
export const ESP_STATUSES = ["Running", "Failed", "Pulled – no failure"];
export const ESP_COMPONENTS = ["Pump", "Motor", "Seal / Protector", "Cable / MLE", "Intake / Gas separator", "Sensor", "VSD / Surface", "Tubing", "Other"];
export const ESP_CAUSES = [
  "Electrical (insulation / short)",
  "Mechanical wear",
  "Scale",
  "Sand / Solids",
  "Gas lock / Gas interference",
  "Corrosion",
  "High temperature",
  "Overload",
  "Underload",
  "Manufacturing defect",
  "Installation / Handling",
  "Unknown",
];

export interface EspLogMapping {
  well: number;
  install: number;
  end: number;
  status?: number;
  model?: number;
  component?: number;
  cause?: number;
}

export function guessEspLogMapping(headers: string[]): Partial<EspLogMapping> {
  const f = (re: RegExp, not?: RegExp) => {
    const i = headers.findIndex((h) => re.test(h) && !(not && not.test(h)));
    return i >= 0 ? i : undefined;
  };
  return {
    well: f(/\bwell\b|uwi|string/i),
    install: f(/install|run ?in|start/i),
    end: f(/pull|fail|end|stop|out/i, /cause|component|reason|mode/i),
    status: f(/status|state/i),
    model: f(/model|type|pump size|series/i),
    component: f(/component|part|item/i),
    cause: f(/cause|reason|\bmode\b|root/i),
  };
}

export interface EspRun {
  well: string;
  model: string;
  install: number;
  end: number | null;
  runLife: number;
  failed: boolean;
  component: string;
  cause: string;
}

export interface GroupStat {
  name: string;
  installs: number;
  failures: number;
  /** Average run life of the failed units (days). */
  avgFailedRunLife: number | null;
  /** Total run days / failures. */
  mtbf: number | null;
}

export interface EspFailureReport {
  asOf: number;
  runs: EspRun[];
  installs: number;
  failures: number;
  running: number;
  totalRunDays: number;
  /** Mean time between failures (days) = total run days of all units / failures. */
  mtbf: number | null;
  avgFailedRunLife: number | null;
  medianFailedRunLife: number | null;
  byCause: GroupStat[];
  byComponent: GroupStat[];
  byModel: GroupStat[];
  byWell: GroupStat[];
  /** Failures per month ("2026-03"). */
  monthly: Array<{ month: string; failures: number; installs: number }>;
  /** Wells with 2+ failures. */
  repeatWells: string[];
  skipped: number;
}

const txt = (v: unknown) => String(v ?? "").trim();

const isFailed = (status: string, cause: string, end: number | null) => {
  const s = status.toLowerCase();
  if (/running|run\b|on ?line|producing|active/.test(s)) return false;
  if (/no ?fail|pulled.*(no|without)|upgrade|resize|workover only/.test(s)) return false;
  if (/fail|down|dead|trip|grounded/.test(s)) return true;
  // No status: a pull date with a cause counts as a failure.
  return end !== null && cause !== "" && !/^(n\/?a|none|-)$/i.test(cause);
};

function group(runs: EspRun[], key: (r: EspRun) => string): GroupStat[] {
  const m = new Map<string, EspRun[]>();
  for (const r of runs) {
    const k = key(r) || "(blank)";
    m.set(k, [...(m.get(k) ?? []), r]);
  }
  return Array.from(m.entries())
    .map(([name, rs]) => {
      const failed = rs.filter((r) => r.failed);
      const days = rs.reduce((s, r) => s + r.runLife, 0);
      return {
        name,
        installs: rs.length,
        failures: failed.length,
        avgFailedRunLife: failed.length ? Math.round(failed.reduce((s, r) => s + r.runLife, 0) / failed.length) : null,
        mtbf: failed.length ? Math.round(days / failed.length) : null,
      };
    })
    .sort((a, b) => b.failures - a.failures || b.installs - a.installs || a.name.localeCompare(b.name));
}

/** `rows` = data rows (no header). `asOf` = Excel serial used as the end of running units (today by default). */
export function analyzeEspLog(rows: unknown[][], m: EspLogMapping, asOf: number): EspFailureReport {
  const runs: EspRun[] = [];
  let skipped = 0;
  for (const r of rows) {
    const well = txt(r[m.well]);
    const install = toSerial(r[m.install]);
    if (!well || install === null) {
      if (r.some((v) => txt(v) !== "")) skipped++;
      continue;
    }
    const end = toSerial(r[m.end]);
    const status = m.status !== undefined ? txt(r[m.status]) : "";
    const cause = m.cause !== undefined ? txt(r[m.cause]) : "";
    const failed = isFailed(status, cause, end);
    const stop = end ?? asOf;
    if (stop < install) {
      skipped++;
      continue;
    }
    runs.push({
      well,
      model: m.model !== undefined ? txt(r[m.model]) : "",
      install,
      end,
      runLife: stop - install,
      failed,
      component: m.component !== undefined ? txt(r[m.component]) : "",
      cause: failed ? cause || "Unknown" : cause,
    });
  }
  const failedRuns = runs.filter((r) => r.failed);
  const totalRunDays = runs.reduce((s, r) => s + r.runLife, 0);
  const lives = failedRuns.map((r) => r.runLife).sort((a, b) => a - b);
  const median = lives.length ? (lives.length % 2 ? lives[(lives.length - 1) / 2] : (lives[lives.length / 2 - 1] + lives[lives.length / 2]) / 2) : null;

  const months = new Map<string, { failures: number; installs: number }>();
  const bump = (serial: number, k: "failures" | "installs") => {
    const mo = serialToIso(serial).slice(0, 7);
    const cur = months.get(mo) ?? { failures: 0, installs: 0 };
    cur[k]++;
    months.set(mo, cur);
  };
  runs.forEach((r) => {
    bump(r.install, "installs");
    if (r.failed && r.end !== null) bump(r.end, "failures");
  });
  const byWell = group(runs, (r) => r.well);
  return {
    asOf,
    runs,
    installs: runs.length,
    failures: failedRuns.length,
    running: runs.filter((r) => !r.failed && r.end === null).length,
    totalRunDays,
    mtbf: failedRuns.length ? Math.round(totalRunDays / failedRuns.length) : null,
    avgFailedRunLife: lives.length ? Math.round(lives.reduce((a, b) => a + b, 0) / lives.length) : null,
    medianFailedRunLife: median === null ? null : Math.round(median),
    byCause: group(failedRuns, (r) => r.cause),
    byComponent: group(failedRuns, (r) => r.component),
    byModel: group(runs, (r) => r.model),
    byWell,
    monthly: Array.from(months.entries())
      .sort((a, b) => a[0].localeCompare(b[0]))
      .map(([month, v]) => ({ month, ...v })),
    repeatWells: byWell.filter((w) => w.failures >= 2).map((w) => w.name),
    skipped,
  };
}
