import { useEffect, useState } from "react";
import { Button, Checkbox, Dialog, DialogActions, DialogBody, DialogContent, DialogSurface, DialogTitle, Dropdown, Field, Input, Option } from "@fluentui/react-components";
import { guessDailyMapping, serialLabel, type DailyMapping } from "../utils/dailyReport";
import { guessEspLogMapping, type EspLogMapping } from "../utils/espFailures";
import type { ColumnRule } from "../utils/columnRules";

const NONE = -1;

function ColumnPick({ label, headers, value, onChange, optional, hint }: { label: string; headers: string[]; value: number; onChange: (v: number) => void; optional?: boolean; hint?: string }) {
  return (
    <Field label={label} hint={hint}>
      <Dropdown
        value={value >= 0 ? headers[value] ?? "" : optional ? "(none)" : "Choose…"}
        selectedOptions={[String(value)]}
        onOptionSelect={(_, d) => d.optionValue !== undefined && onChange(Number(d.optionValue))}
      >
        {optional && <Option value={String(NONE)}>(none)</Option>}
        {headers.map((h, i) => (
          <Option key={i} value={String(i)}>
            {h}
          </Option>
        ))}
      </Dropdown>
    </Field>
  );
}

/* ------------------------------------------------------------ daily production report */

export interface DailyReportTarget {
  headers: string[];
  /** Excel date serials found in the guessed date column (latest first). */
  datesFor: (dateCol: number) => number[];
}

export function DailyReportDialog({ target, onClose, onRun }: { target: DailyReportTarget | null; onClose: () => void; onRun: (m: DailyMapping, reportDate?: number) => void }) {
  const headers = target?.headers ?? [];
  const [m, setM] = useState({ date: NONE, well: NONE, oil: NONE, water: NONE, gas: NONE });
  const [date, setDate] = useState<number | null>(null);
  useEffect(() => {
    if (!target) return;
    const g = guessDailyMapping(target.headers);
    setM({ date: g.date, well: g.well, oil: g.oil, water: g.water ?? NONE, gas: g.gas ?? NONE });
    setDate(null);
  }, [target]);
  const dates = target && m.date >= 0 ? target.datesFor(m.date).slice(0, 90) : [];
  const set = (k: keyof typeof m) => (v: number) => setM((p) => ({ ...p, [k]: v }));
  const used = [m.date, m.well, m.oil, m.water, m.gas].filter((v) => v >= 0);
  const valid = m.date >= 0 && m.well >= 0 && m.oil >= 0 && new Set(used).size === used.length;
  return (
    <Dialog open={!!target} onOpenChange={(_, d) => !d.open && onClose()}>
      <DialogSurface className="zee-dialog">
        <DialogBody>
          <DialogTitle>Daily production report</DialogTitle>
          <DialogContent className="zee-form">
            <p className="zee-muted">
              Builds a report sheet (per-well rates, change vs the previous reading, water cut, GOR, totals, highlights and a chart) and hands it to Zee AI for Word: there,
              Templates → <b>Daily report from Excel</b> writes the report text around it.
            </p>
            <ColumnPick label="Date column" headers={headers} value={m.date} onChange={set("date")} />
            <ColumnPick label="Well column" headers={headers} value={m.well} onChange={set("well")} />
            <ColumnPick label="Oil rate" headers={headers} value={m.oil} onChange={set("oil")} />
            <ColumnPick label="Water rate (optional)" headers={headers} value={m.water} onChange={set("water")} optional />
            <ColumnPick label="Gas rate (optional)" headers={headers} value={m.gas} onChange={set("gas")} optional />
            <Field label="Report date" hint={dates.length ? `${dates.length} dates in the data` : "Pick the date column first"}>
              <Dropdown
                disabled={!dates.length}
                value={date !== null ? serialLabel(date) : dates.length ? `Latest – ${serialLabel(dates[0])}` : ""}
                selectedOptions={[String(date ?? "latest")]}
                onOptionSelect={(_, d) => setDate(d.optionValue === "latest" || d.optionValue === undefined ? null : Number(d.optionValue))}
              >
                <Option value="latest" text={dates.length ? `Latest – ${serialLabel(dates[0])}` : "Latest"}>
                  {dates.length ? `Latest – ${serialLabel(dates[0])}` : "Latest"}
                </Option>
                {dates.map((d) => (
                  <Option key={d} value={String(d)} text={serialLabel(d)}>
                    {serialLabel(d)}
                  </Option>
                ))}
              </Dropdown>
            </Field>
          </DialogContent>
          <DialogActions>
            <Button appearance="primary" disabled={!valid} onClick={() => onRun({ date: m.date, well: m.well, oil: m.oil, water: m.water >= 0 ? m.water : undefined, gas: m.gas >= 0 ? m.gas : undefined }, date ?? undefined)}>
              Build report
            </Button>
            <Button onClick={onClose}>Cancel</Button>
          </DialogActions>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
}

/* ------------------------------------------------------------ ESP failure log */

export function EspLogDialog({ headers, onClose, onRun }: { headers: string[] | null; onClose: () => void; onRun: (m: EspLogMapping) => void }) {
  const [m, setM] = useState({ well: NONE, install: NONE, end: NONE, status: NONE, model: NONE, component: NONE, cause: NONE });
  useEffect(() => {
    if (!headers) return;
    const g = guessEspLogMapping(headers);
    setM({ well: g.well ?? NONE, install: g.install ?? NONE, end: g.end ?? NONE, status: g.status ?? NONE, model: g.model ?? NONE, component: g.component ?? NONE, cause: g.cause ?? NONE });
  }, [headers]);
  const h = headers ?? [];
  const set = (k: keyof typeof m) => (v: number) => setM((p) => ({ ...p, [k]: v }));
  const opt = (v: number) => (v >= 0 ? v : undefined);
  const valid = m.well >= 0 && m.install >= 0 && m.end >= 0 && m.install !== m.end;
  return (
    <Dialog open={!!headers} onOpenChange={(_, d) => !d.open && onClose()}>
      <DialogSurface className="zee-dialog">
        <DialogBody>
          <DialogTitle>Analyze ESP failure log</DialogTitle>
          <DialogContent className="zee-form">
            <p className="zee-muted">Run life, MTBF, failures by cause / component / model / well and a monthly trend on a new sheet, then an AI summary in the chat.</p>
            <ColumnPick label="Well" headers={h} value={m.well} onChange={set("well")} />
            <ColumnPick label="Install date" headers={h} value={m.install} onChange={set("install")} />
            <ColumnPick label="Pull / fail date" headers={h} value={m.end} onChange={set("end")} hint="Empty = still running (run life counted to today)." />
            <ColumnPick label="Status (optional)" headers={h} value={m.status} onChange={set("status")} optional hint="Running / Failed / Pulled – no failure. Without it, a pull date + cause counts as a failure." />
            <ColumnPick label="ESP model (optional)" headers={h} value={m.model} onChange={set("model")} optional />
            <ColumnPick label="Failed component (optional)" headers={h} value={m.component} onChange={set("component")} optional />
            <ColumnPick label="Failure cause (optional)" headers={h} value={m.cause} onChange={set("cause")} optional />
          </DialogContent>
          <DialogActions>
            <Button
              appearance="primary"
              disabled={!valid}
              onClick={() => onRun({ well: m.well, install: m.install, end: m.end, status: opt(m.status), model: opt(m.model), component: opt(m.component), cause: opt(m.cause) })}
            >
              Analyze
            </Button>
            <Button onClick={onClose}>Cancel</Button>
          </DialogActions>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
}

/* ------------------------------------------------------------ AI data validation */

interface EditableRule extends ColumnRule {
  on: boolean;
  text: string;
}

export function RulesDialog({ rules, onClose, onApply }: { rules: ColumnRule[] | null; onClose: () => void; onApply: (rules: ColumnRule[], opts: { highlight: boolean; validate: boolean }) => void }) {
  const [list, setList] = useState<EditableRule[]>([]);
  const [highlight, setHighlight] = useState(true);
  const [validate, setValidate] = useState(false);
  useEffect(() => {
    if (rules) setList(rules.map((r) => ({ ...r, on: true, text: (r.allowed ?? []).join(", ") })));
  }, [rules]);
  const patch = (i: number, p: Partial<EditableRule>) => setList((l) => l.map((r, k) => (k === i ? { ...r, ...p } : r)));
  const num = (v: string) => (v.trim() === "" || !Number.isFinite(Number(v)) ? null : Number(v));
  const chosen = list
    .filter((r) => r.on)
    .map(({ on: _on, text, ...r }): ColumnRule => (r.type === "list" ? { ...r, allowed: text.split(",").map((x) => x.trim()).filter(Boolean) } : r))
    .filter((r) => (r.type === "number" ? r.min !== null || r.max !== null : (r.allowed ?? []).length > 0));
  return (
    <Dialog open={!!rules} onOpenChange={(_, d) => !d.open && onClose()}>
      <DialogSurface className="zee-dialog">
        <DialogBody>
          <DialogTitle>Suggested validation rules</DialogTitle>
          <DialogContent className="zee-form">
            {!list.length && <p>Zee AI found no columns that need a rule.</p>}
            {list.map((r, i) => (
              <div key={r.column} className="zee-rule">
                <Checkbox checked={r.on} onChange={(_, d) => patch(i, { on: !!d.checked })} label={<b>{r.column}</b>} />
                {r.type === "number" ? (
                  <div className="zee-row">
                    <Input size="small" type="number" className="zee-rule-num" value={r.min === null || r.min === undefined ? "" : String(r.min)} placeholder="min" onChange={(_, d) => patch(i, { min: num(d.value) })} />
                    <span>to</span>
                    <Input size="small" type="number" className="zee-rule-num" value={r.max === null || r.max === undefined ? "" : String(r.max)} placeholder="max" onChange={(_, d) => patch(i, { max: num(d.value) })} />
                  </div>
                ) : (
                  <Input size="small" value={r.text} onChange={(_, d) => patch(i, { text: d.value })} placeholder="Allowed values, comma separated" />
                )}
                {r.reason && <span className="zee-muted">{r.reason}</span>}
              </div>
            ))}
            <Checkbox checked={highlight} onChange={(_, d) => setHighlight(!!d.checked)} label="Highlight values that break a rule (red fill, updates live)" />
            <Checkbox checked={validate} onChange={(_, d) => setValidate(!!d.checked)} label="Also add Excel data validation, so new entries are checked when typed" />
            <p className="zee-muted">Each rule is a separate undo step (↶ or the history list).</p>
          </DialogContent>
          <DialogActions>
            <Button appearance="primary" disabled={!chosen.length} onClick={() => onApply(chosen, { highlight, validate })}>
              Check {chosen.length} column{chosen.length === 1 ? "" : "s"}
            </Button>
            <Button onClick={onClose}>Cancel</Button>
          </DialogActions>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
}
