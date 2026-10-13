import { useEffect, useState } from "react";
import {
  Button,
  Checkbox,
  Dialog,
  DialogActions,
  DialogBody,
  DialogContent,
  DialogSurface,
  DialogTitle,
  Dropdown,
  Field,
  Input,
  Option,
  OptionGroup,
  Radio,
  RadioGroup,
  Switch,
  Textarea,
} from "@fluentui/react-components";
import { DEFAULT_CLEAN_OPTIONS, type CaseMode, type CleanOptions, type DateOrder } from "../utils/clean";
import { EXTRACT_LABELS, type ExtractKind } from "../utils/extract";
import { UNIT_CONVERSIONS } from "../utils/petro";
import type { ChartKind, PlotSource } from "../services/excelService";
import { CHART_GROUPS, CHART_TYPES, chartTypeInfo, checkYCount, supportsLogScale, supportsSecondaryAxis } from "../utils/chartTypes";

/* ------------------------------------------------------------ confirm */

export interface ConfirmState {
  title: string;
  message: string;
  confirmLabel?: string;
}

export function ConfirmDialog({ state, onResult }: { state: ConfirmState | null; onResult: (ok: boolean) => void }) {
  return (
    <Dialog open={!!state} onOpenChange={(_, d) => !d.open && onResult(false)}>
      <DialogSurface>
        <DialogBody>
          <DialogTitle>{state?.title}</DialogTitle>
          <DialogContent style={{ whiteSpace: "pre-line" }}>{state?.message}</DialogContent>
          <DialogActions>
            <Button appearance="secondary" onClick={() => onResult(false)}>
              Cancel
            </Button>
            <Button appearance="primary" onClick={() => onResult(true)}>
              {state?.confirmLabel ?? "OK"}
            </Button>
          </DialogActions>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
}

/* ------------------------------------------------------------ clean */

const CASES: Record<CaseMode, string> = {
  none: "Keep as is",
  upper: "UPPER CASE",
  lower: "lower case",
  proper: "Proper Case",
  sentence: "Sentence case",
};
const DATES: Record<DateOrder, string> = {
  off: "Don't convert dates",
  MDY: "Text dates are M/D/Y (US)",
  DMY: "Text dates are D/M/Y",
  YMD: "Only ISO dates (Y-M-D)",
};

export function CleanDialog({ open, onClose, onRun }: { open: boolean; onClose: () => void; onRun: (o: CleanOptions) => void }) {
  const [o, setO] = useState<CleanOptions>(DEFAULT_CLEAN_OPTIONS);
  return (
    <Dialog open={open} onOpenChange={(_, d) => !d.open && onClose()}>
      <DialogSurface>
        <DialogBody>
          <DialogTitle>Clean data</DialogTitle>
          <DialogContent className="zee-form">
            <p className="zee-muted">Cleans the selected cells in place. Formulas are never changed. You can undo with the ↶ button.</p>
            <Switch checked={o.trimWhitespace} onChange={(_, d) => setO({ ...o, trimWhitespace: d.checked })} label="Trim and collapse extra spaces" />
            <Switch checked={o.removeNonPrintable} onChange={(_, d) => setO({ ...o, removeNonPrintable: d.checked })} label="Remove hidden / non-printable characters" />
            <Switch checked={o.convertNumbers} onChange={(_, d) => setO({ ...o, convertNumbers: d.checked })} label='Convert numbers stored as text ("1,250" → 1250)' />
            <Field label="Capitalization">
              <Dropdown value={CASES[o.caseMode]} selectedOptions={[o.caseMode]} onOptionSelect={(_, d) => setO({ ...o, caseMode: d.optionValue as CaseMode })}>
                {Object.entries(CASES).map(([k, v]) => (
                  <Option key={k} value={k}>
                    {v}
                  </Option>
                ))}
              </Dropdown>
            </Field>
            <Field label="Fix date formats (text → real dates, yyyy-mm-dd)">
              <Dropdown value={DATES[o.dateOrder]} selectedOptions={[o.dateOrder]} onOptionSelect={(_, d) => setO({ ...o, dateOrder: d.optionValue as DateOrder })}>
                {Object.entries(DATES).map(([k, v]) => (
                  <Option key={k} value={k}>
                    {v}
                  </Option>
                ))}
              </Dropdown>
            </Field>
          </DialogContent>
          <DialogActions>
            <Button appearance="secondary" onClick={onClose}>
              Cancel
            </Button>
            <Button appearance="primary" onClick={() => onRun(o)}>
              Clean selection
            </Button>
          </DialogActions>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
}

/* ------------------------------------------------------------ extract */

export type ExtractChoice = { kind: ExtractKind; header: boolean } | { kind: "ai"; instruction: string; header: boolean };

export function ExtractDialog({ open, onClose, onRun }: { open: boolean; onClose: () => void; onRun: (c: ExtractChoice) => void }) {
  const [kind, setKind] = useState<ExtractKind | "ai">("firstName");
  const [instruction, setInstruction] = useState("");
  const [header, setHeader] = useState(true);
  return (
    <Dialog open={open} onOpenChange={(_, d) => !d.open && onClose()}>
      <DialogSurface>
        <DialogBody>
          <DialogTitle>Split / Extract</DialogTitle>
          <DialogContent className="zee-form">
            <p className="zee-muted">Select one column. Results go into a new column inserted to its right.</p>
            <RadioGroup value={kind} onChange={(_, d) => setKind(d.value as ExtractKind | "ai")}>
              {(Object.keys(EXTRACT_LABELS) as ExtractKind[]).map((k) => (
                <Radio key={k} value={k} label={EXTRACT_LABELS[k]} />
              ))}
              <Radio value="ai" label="Custom – describe it (uses Gemini)" />
            </RadioGroup>
            {kind === "ai" && (
              <Field label="What should be extracted or generated for each row?" hint="e.g. “the city”, “sentiment: positive/neutral/negative”, “translate to Arabic”">
                <Textarea value={instruction} onChange={(_, d) => setInstruction(d.value)} rows={3} />
              </Field>
            )}
            <Checkbox checked={header} onChange={(_, d) => setHeader(!!d.checked)} label="First selected row is a header" />
          </DialogContent>
          <DialogActions>
            <Button appearance="secondary" onClick={onClose}>
              Cancel
            </Button>
            <Button
              appearance="primary"
              disabled={kind === "ai" && !instruction.trim()}
              onClick={() => onRun(kind === "ai" ? { kind, instruction: instruction.trim(), header } : { kind, header })}
            >
              Extract
            </Button>
          </DialogActions>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
}

/* ------------------------------------------------------------ generate formula */

export function GenerateFormulaDialog({ open, onClose, onRun }: { open: boolean; onClose: () => void; onRun: (desc: string) => void }) {
  const [desc, setDesc] = useState("");
  useEffect(() => {
    if (open) setDesc("");
  }, [open]);
  return (
    <Dialog open={open} onOpenChange={(_, d) => !d.open && onClose()}>
      <DialogSurface>
        <DialogBody>
          <DialogTitle>Generate formula</DialogTitle>
          <DialogContent className="zee-form">
            <p className="zee-muted">Click the cell where the formula should go first. Zee AI reads your column headers for context.</p>
            <Field label="Describe what the formula should do">
              <Textarea
                value={desc}
                onChange={(_, d) => setDesc(d.value)}
                rows={3}
                placeholder="Look up the salary for the employee ID in A2 from the Staff sheet"
                onKeyDown={(e) => {
                  if (e.key === "Enter" && !e.shiftKey && desc.trim()) {
                    e.preventDefault();
                    onRun(desc.trim());
                  }
                }}
              />
            </Field>
          </DialogContent>
          <DialogActions>
            <Button appearance="secondary" onClick={onClose}>
              Cancel
            </Button>
            <Button appearance="primary" disabled={!desc.trim()} onClick={() => onRun(desc.trim())}>
              Generate
            </Button>
          </DialogActions>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
}

/* ------------------------------------------------------------ decline curve */

export interface DcaChoice {
  timeColumn: number;
  rateColumn: number;
  forecastPeriods: number;
  economicLimit?: number;
  wellColumn?: number;
}

export function DeclineCurveDialog({ open, headers, onClose, onRun }: { open: boolean; headers: string[]; onClose: () => void; onRun: (c: DcaChoice) => void }) {
  const [timeCol, setTimeCol] = useState(0);
  const [rateCol, setRateCol] = useState(1);
  const [periods, setPeriods] = useState("60");
  const [limit, setLimit] = useState("");
  const [wellCol, setWellCol] = useState(-1);
  useEffect(() => {
    if (open) {
      setTimeCol(0);
      setRateCol(Math.min(1, Math.max(0, headers.length - 1)));
      setWellCol(headers.findIndex((h) => /\b(well|wellname|well name|uwi|api ?no|string|completion)\b/i.test(h)));
    }
  }, [open, headers.length]);
  const pick = (value: number, set: (n: number) => void, label: string) => (
    <Field label={label}>
      <Dropdown value={headers[value] ?? ""} selectedOptions={[String(value)]} onOptionSelect={(_, d) => set(Number(d.optionValue))}>
        {headers.map((h, i) => (
          <Option key={i} value={String(i)}>
            {h}
          </Option>
        ))}
      </Dropdown>
    </Field>
  );
  return (
    <Dialog open={open} onOpenChange={(_, d) => !d.open && onClose()}>
      <DialogSurface>
        <DialogBody>
          <DialogTitle>Decline curve analysis</DialogTitle>
          <DialogContent className="zee-form">
            <p className="zee-muted">Fits exponential, harmonic and hyperbolic Arps curves to the selected data and writes a report sheet with parameters, a forecast and a chart.</p>
            {pick(timeCol, setTimeCol, "Time / date column")}
            {pick(rateCol, setRateCol, "Rate column")}
            <Field label="Well column (optional)" hint="Pick it to fit every well separately and get one summary row per well.">
              <Dropdown
                value={wellCol >= 0 ? headers[wellCol] ?? "" : "(single well – no grouping)"}
                selectedOptions={[String(wellCol)]}
                onOptionSelect={(_, d) => d.optionValue !== undefined && setWellCol(Number(d.optionValue))}
              >
                <Option value="-1">(single well – no grouping)</Option>
                {headers.map((h, i) => (
                  <Option key={i} value={String(i)}>
                    {h}
                  </Option>
                ))}
              </Dropdown>
            </Field>
            <Field label="Forecast periods after the last data point" hint="In the data's typical step (e.g. days or months).">
              <Input type="number" value={periods} onChange={(_, d) => setPeriods(d.value)} />
            </Field>
            <Field label="Economic limit rate (optional)" hint="Adds time-to-limit and EUR.">
              <Input type="number" value={limit} onChange={(_, d) => setLimit(d.value)} placeholder="e.g. 10" />
            </Field>
          </DialogContent>
          <DialogActions>
            <Button appearance="secondary" onClick={onClose}>
              Cancel
            </Button>
            <Button
              appearance="primary"
              disabled={timeCol === rateCol || wellCol === timeCol || wellCol === rateCol}
              onClick={() =>
                onRun({
                  timeColumn: timeCol,
                  rateColumn: rateCol,
                  forecastPeriods: Math.max(0, Math.min(5000, Math.round(Number(periods) || 0))),
                  economicLimit: Number(limit) > 0 ? Number(limit) : undefined,
                  wellColumn: wellCol >= 0 ? wellCol : undefined,
                })
              }
            >
              {wellCol >= 0 ? "Fit every well" : "Fit curves"}
            </Button>
          </DialogActions>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
}

/* ------------------------------------------------------------ units */

export function UnitDialog({ open, onClose, onRun }: { open: boolean; onClose: () => void; onRun: (id: string, header: boolean) => void }) {
  const [id, setId] = useState(UNIT_CONVERSIONS[0].id);
  const [header, setHeader] = useState(true);
  const groups = Array.from(new Set(UNIT_CONVERSIONS.map((u) => u.group)));
  const current = UNIT_CONVERSIONS.find((u) => u.id === id);
  return (
    <Dialog open={open} onOpenChange={(_, d) => !d.open && onClose()}>
      <DialogSurface>
        <DialogBody>
          <DialogTitle>Unit conversion</DialogTitle>
          <DialogContent className="zee-form">
            <p className="zee-muted">Converts the selected column(s) into new columns inserted to the right. The original data is not changed.</p>
            <Field label="Conversion">
              <Dropdown value={current?.label ?? ""} selectedOptions={[id]} onOptionSelect={(_, d) => d.optionValue && setId(d.optionValue)}>
                {groups.map((g) => (
                  <OptionGroup key={g} label={g}>
                    {UNIT_CONVERSIONS.filter((u) => u.group === g).map((u) => (
                      <Option key={u.id} value={u.id}>
                        {u.label}
                      </Option>
                    ))}
                  </OptionGroup>
                ))}
              </Dropdown>
            </Field>
            <Checkbox checked={header} onChange={(_, d) => setHeader(!!d.checked)} label="First selected row is a header" />
          </DialogContent>
          <DialogActions>
            <Button appearance="secondary" onClick={onClose}>
              Cancel
            </Button>
            <Button appearance="primary" onClick={() => onRun(id, header)}>
              Convert
            </Button>
          </DialogActions>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
}

/* ------------------------------------------------------------ chart */

const CHART_LABELS: Record<ChartKind, string> = {
  line: "Line",
  column: "Column",
  bar: "Bar",
  scatter: "Scatter (points)",
  scatterLines: "Scatter with lines (x = first column)",
  area: "Area",
  pie: "Pie",
};

export function ChartDialog({ open, onClose, onRun }: { open: boolean; onClose: () => void; onRun: (kind: ChartKind, title: string) => void }) {
  const [kind, setKind] = useState<ChartKind>("line");
  const [title, setTitle] = useState("");
  return (
    <Dialog open={open} onOpenChange={(_, d) => !d.open && onClose()}>
      <DialogSurface>
        <DialogBody>
          <DialogTitle>Chart from selection</DialogTitle>
          <DialogContent className="zee-form">
            <p className="zee-muted">Select a block whose first row holds the series names (and first column the x values / categories).</p>
            <Field label="Chart type">
              <Dropdown value={CHART_LABELS[kind]} selectedOptions={[kind]} onOptionSelect={(_, d) => d.optionValue && setKind(d.optionValue as ChartKind)}>
                {(Object.keys(CHART_LABELS) as ChartKind[]).map((k) => (
                  <Option key={k} value={k}>
                    {CHART_LABELS[k]}
                  </Option>
                ))}
              </Dropdown>
            </Field>
            <Field label="Title (optional)">
              <Input value={title} onChange={(_, d) => setTitle(d.value)} />
            </Field>
          </DialogContent>
          <DialogActions>
            <Button appearance="secondary" onClick={onClose}>
              Cancel
            </Button>
            <Button appearance="primary" onClick={() => onRun(kind, title.trim())}>
              Create chart
            </Button>
          </DialogActions>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
}

/* ------------------------------------------------------------ draw: plot columns */

export interface DrawChoice {
  xIndex: number;
  yIndexes: number[];
  /** Excel.ChartType value. */
  chartType: string;
  title: string;
  secondaryYIndexes: number[];
  logLeft: boolean;
  logRight: boolean;
}

const LAST_TYPE_KEY = "zeeai.draw.lastType";
function lastChartType(): string | null {
  try {
    const v = localStorage.getItem(LAST_TYPE_KEY);
    return v && CHART_TYPES.some((t) => t.id === v) ? v : null;
  } catch {
    return null;
  }
}
function rememberChartType(id: string) {
  try {
    localStorage.setItem(LAST_TYPE_KEY, id);
  } catch {
    /* ignore */
  }
}

export function DrawDialog({ source, onClose, onRun }: { source: PlotSource | null; onClose: () => void; onRun: (c: DrawChoice) => void }) {
  const cols = source?.columns ?? [];
  const [xIndex, setXIndex] = useState(0);
  const [yIndexes, setYIndexes] = useState<number[]>([]);
  const [chartType, setChartType] = useState("Line");
  const [title, setTitle] = useState("");
  const [rightAxis, setRightAxis] = useState<number[]>([]);
  const [logLeft, setLogLeft] = useState(false);
  const [logRight, setLogRight] = useState(false);

  // New data source → sensible defaults: first column on X, the next numeric columns (up to 4) on Y.
  useEffect(() => {
    if (!source) return;
    const x = source.columns[0]?.index ?? 0;
    setXIndex(x);
    setYIndexes(source.columns.filter((c) => c.index !== x && c.numeric).slice(0, 4).map((c) => c.index));
    setChartType(lastChartType() ?? (source.columns[0]?.numeric ? "XYScatterLinesNoMarkers" : "Line"));
    setTitle("");
    setRightAxis([]);
    setLogLeft(false);
    setLogRight(false);
  }, [source]);

  // Types that take a fixed number of Y columns (pie, pareto, bubble…): drop the extras automatically.
  useEffect(() => {
    const max = chartTypeInfo(chartType).maxY;
    if (max !== undefined) setYIndexes((prev) => (prev.length > max ? prev.slice(0, max) : prev));
  }, [chartType, source]);

  const xCol = cols.find((c) => c.index === xIndex);
  const info = chartTypeInfo(chartType);
  const yProblem = checkYCount(info, yIndexes.length);
  const yChoices = cols.filter((c) => c.index !== xIndex);
  const canRight = supportsSecondaryAxis(info) && yIndexes.length > 1;
  const canLog = supportsLogScale(info);
  const activeRight = canRight ? rightAxis.filter((i) => yIndexes.includes(i)) : [];
  const allOnRight = activeRight.length > 0 && activeRight.length === yIndexes.length;
  const toggle = (idx: number, on: boolean) =>
    setYIndexes((prev) => (on ? [...prev, idx].sort((a, b) => a - b) : prev.filter((i) => i !== idx)));
  const toggleRight = (idx: number, on: boolean) => setRightAxis((prev) => (on ? [...prev, idx] : prev.filter((i) => i !== idx)));
  const pickX = (idx: number) => {
    setXIndex(idx);
    setYIndexes((prev) => prev.filter((i) => i !== idx));
  };

  return (
    <Dialog open={!!source} onOpenChange={(_, d) => !d.open && onClose()}>
      <DialogSurface>
        <DialogBody>
          <DialogTitle>Draw – plot columns</DialogTitle>
          <DialogContent className="zee-form">
            <p className="zee-muted">
              Data from {source?.sheetName}!{source?.address}
              {source?.hasHeader ? " (first row used as names)" : ""}. Select a range first to limit the rows.
            </p>

            <section className="zee-draw-section">
              <h4>1 · X axis</h4>
              <Dropdown
                aria-label="X axis column"
                value={xCol ? `${xCol.letter} – ${xCol.name}` : ""}
                selectedOptions={[String(xIndex)]}
                onOptionSelect={(_, d) => d.optionValue && pickX(Number(d.optionValue))}
              >
                {cols.map((c) => (
                  <Option key={c.index} value={String(c.index)} text={`${c.letter} – ${c.name}`}>
                    {c.letter} – {c.name}
                  </Option>
                ))}
              </Dropdown>
            </section>

            <section className="zee-draw-section">
              <h4>
                2 · Y axis data <span className="zee-muted">({yIndexes.length} selected)</span>
              </h4>
              <div className="zee-draw-list">
                {yChoices.map((c) => (
                  <div key={c.index} className="zee-draw-row">
                    <Checkbox
                      checked={yIndexes.includes(c.index)}
                      onChange={(_, d) => toggle(c.index, !!d.checked)}
                      label={
                        <span className="zee-draw-label">
                          <span className="zee-draw-letter">{c.letter}</span>
                          <span className="zee-draw-name">{c.name}</span>
                          {!c.numeric && <span className="zee-muted"> (text)</span>}
                        </span>
                      }
                    />
                    {canRight && yIndexes.includes(c.index) && (
                      <Checkbox
                        className="zee-draw-right"
                        size="medium"
                        checked={rightAxis.includes(c.index)}
                        onChange={(_, d) => toggleRight(c.index, !!d.checked)}
                        label="Right axis"
                      />
                    )}
                  </div>
                ))}
              </div>
              <div className="zee-row">
                <Button size="small" appearance="subtle" onClick={() => setYIndexes(yChoices.filter((c) => c.numeric).map((c) => c.index))}>
                  All numeric
                </Button>
                <Button size="small" appearance="subtle" onClick={() => setYIndexes([])}>
                  Clear
                </Button>
              </div>
            </section>

            <Field
              label={`Chart type (${CHART_TYPES.length} available)`}
              hint={info.hint ?? (info.group === "Pie & doughnut" || info.group === "Radar" ? undefined : "X goes on the horizontal axis; each ticked Y column is one series.")}
              validationState={yProblem && yIndexes.length > 0 ? "warning" : "none"}
              validationMessage={yProblem && yIndexes.length > 0 ? yProblem : undefined}
            >
              <Dropdown
                value={`${info.group} › ${info.label}`}
                selectedOptions={[chartType]}
                onOptionSelect={(_, d) => d.optionValue && setChartType(d.optionValue)}
                listbox={{ style: { maxHeight: "320px" } }}
              >
                {CHART_GROUPS.map((g) => (
                  <OptionGroup key={g} label={g}>
                    {CHART_TYPES.filter((t) => t.group === g).map((t) => (
                      <Option key={t.id} value={t.id} text={`${g} › ${t.label}`}>
                        {t.label}
                      </Option>
                    ))}
                  </OptionGroup>
                ))}
              </Dropdown>
            </Field>
            {canLog && (
              <div className="zee-row">
                <Checkbox checked={logLeft} onChange={(_, d) => setLogLeft(!!d.checked)} label="Log scale – left Y axis" />
                {activeRight.length > 0 && !allOnRight && (
                  <Checkbox checked={logRight} onChange={(_, d) => setLogRight(!!d.checked)} label="Log scale – right Y axis" />
                )}
              </div>
            )}
            {allOnRight && <p className="zee-muted">Keep at least one series on the left axis – otherwise everything stays on one axis.</p>}
            <Field label="Title (optional)" hint="Leave empty for “Y names vs X name”.">
              <Input value={title} onChange={(_, d) => setTitle(d.value)} />
            </Field>
          </DialogContent>
          <DialogActions>
            <Button appearance="secondary" onClick={onClose}>
              Cancel
            </Button>
            <Button
              appearance="primary"
              disabled={!!yProblem}
              onClick={() => {
                rememberChartType(chartType);
                onRun({ xIndex, yIndexes, chartType, title: title.trim(), secondaryYIndexes: activeRight, logLeft: canLog && logLeft, logRight: canLog && logRight });
              }}
            >
              Draw chart
            </Button>
          </DialogActions>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
}

/* ------------------------------------------------------------ ESP analysis (calculated) */

export interface EspDialogChoice {
  columns: { rate: number; pip: number; pd: number; hz: number; amps?: number; waterCut?: number; x?: number };
  options: {
    baseHz: number;
    sg: number;
    api?: number;
    stages?: number;
    minRate?: number;
    maxRate?: number;
    motorAmps?: number;
    minPip?: number;
    degradationPct: number;
  };
  curveAddress?: string;
}

type EspNumKey = "baseHz" | "sg" | "api" | "stages" | "minRate" | "maxRate" | "motorAmps" | "minPip" | "degradationPct";
const ESP_FORM_KEY = "zeeai.esp.form.v1";
const ESP_DEFAULTS: Record<EspNumKey, string> = { baseHz: "60", sg: "1.0", api: "", stages: "", minRate: "", maxRate: "", motorAmps: "", minPip: "", degradationPct: "15" };

const ESP_GUESS: Record<"rate" | "pip" | "pd" | "hz" | "amps" | "waterCut" | "x", RegExp> = {
  rate: /\b(rate|q|bfpd|blpd|bpd|liquid|gross|total ?fluid)\b/i,
  pip: /(intake|\bpip\b|suction)/i,
  pd: /(discharge|\bpdis|\bpd\b)/i,
  hz: /(freq|\bhz\b)/i,
  amps: /(amp|current|\bmotor ?a\b)/i,
  waterCut: /(water ?cut|\bwc\b|bs&?w)/i,
  x: /(date|time|day|test)/i,
};

const NONE = "-1";

export function EspDialog({
  source,
  onClose,
  onRun,
  onPickSelection,
}: {
  source: PlotSource | null;
  onClose: () => void;
  onRun: (c: EspDialogChoice) => void;
  onPickSelection: () => Promise<string | null>;
}) {
  const cols = source?.columns ?? [];
  const [map, setMap] = useState<Record<keyof typeof ESP_GUESS, number>>({ rate: -1, pip: -1, pd: -1, hz: -1, amps: -1, waterCut: -1, x: -1 });
  const [form, setForm] = useState<Record<EspNumKey, string>>(ESP_DEFAULTS);
  const [curve, setCurve] = useState("");

  useEffect(() => {
    if (!source) return;
    const used = new Set<number>();
    const guess = {} as Record<keyof typeof ESP_GUESS, number>;
    // Specific columns first so "rate" doesn't grab "intake pressure" etc.
    (["pip", "pd", "hz", "amps", "waterCut", "x", "rate"] as const).forEach((k) => {
      const hit = source.columns.find((c) => !used.has(c.index) && ESP_GUESS[k].test(c.name));
      guess[k] = hit ? hit.index : -1;
      if (hit) used.add(hit.index);
    });
    setMap(guess);
    try {
      const saved = JSON.parse(localStorage.getItem(ESP_FORM_KEY) ?? "null") as { form?: Record<EspNumKey, string>; curve?: string } | null;
      setForm({ ...ESP_DEFAULTS, ...(saved?.form ?? {}) });
      setCurve(saved?.curve ?? "");
    } catch {
      setForm(ESP_DEFAULTS);
    }
  }, [source]);

  const colPick = (key: keyof typeof ESP_GUESS, label: string, required: boolean) => {
    const cur = cols.find((c) => c.index === map[key]);
    return (
      <Field label={label} required={required} size="small">
        <Dropdown
          size="small"
          value={cur ? `${cur.letter} – ${cur.name}` : required ? "" : "(none)"}
          selectedOptions={[String(map[key])]}
          onOptionSelect={(_, d) => d.optionValue !== undefined && setMap((m) => ({ ...m, [key]: Number(d.optionValue) }))}
          placeholder="Choose a column"
        >
          {!required && (
            <Option value={NONE} text="(none)">
              (none)
            </Option>
          )}
          {cols.map((c) => (
            <Option key={c.index} value={String(c.index)} text={`${c.letter} – ${c.name}`}>
              {c.letter} – {c.name}
            </Option>
          ))}
        </Dropdown>
      </Field>
    );
  };
  const numField = (key: EspNumKey, label: string, hint?: string) => (
    <Field label={label} hint={hint} size="small">
      <Input size="small" type="number" value={form[key]} onChange={(_, d) => setForm((f) => ({ ...f, [key]: d.value }))} />
    </Field>
  );
  const opt = (k: EspNumKey) => (form[k].trim() === "" || !Number.isFinite(Number(form[k])) ? undefined : Number(form[k]));
  const required = [map.rate, map.pip, map.pd, map.hz];
  const missing = required.some((v) => v < 0);
  const duplicate = new Set(required).size !== required.length;

  const run = () => {
    try {
      localStorage.setItem(ESP_FORM_KEY, JSON.stringify({ form, curve }));
    } catch {
      /* ignore */
    }
    const o = (v: number) => (v >= 0 ? v : undefined);
    onRun({
      columns: { rate: map.rate, pip: map.pip, pd: map.pd, hz: map.hz, amps: o(map.amps), waterCut: o(map.waterCut), x: o(map.x) },
      options: {
        baseHz: opt("baseHz") ?? 60,
        sg: opt("sg") ?? 1,
        api: opt("api"),
        stages: opt("stages"),
        minRate: opt("minRate"),
        maxRate: opt("maxRate"),
        motorAmps: opt("motorAmps"),
        minPip: opt("minPip"),
        degradationPct: opt("degradationPct") ?? 15,
      },
      curveAddress: curve.trim() || undefined,
    });
  };

  return (
    <Dialog open={!!source} onOpenChange={(_, d) => !d.open && onClose()}>
      <DialogSurface>
        <DialogBody>
          <DialogTitle>ESP analysis (calculated)</DialogTitle>
          <DialogContent className="zee-form">
            <p className="zee-muted">
              Calculates TDH, hydraulic HP, motor load, affinity-law correction to the catalog frequency, head vs the pump curve and the operating range for every row of{" "}
              {source?.sheetName}!{source?.address}. Results go to a new sheet with charts.
            </p>
            <section className="zee-draw-section">
              <h4>1 · Measured data columns</h4>
              <div className="zee-grid2">
                {colPick("rate", "Liquid rate (bbl/d)", true)}
                {colPick("hz", "Frequency (Hz)", true)}
                {colPick("pip", "Intake pressure PIP (psi)", true)}
                {colPick("pd", "Discharge pressure (psi)", true)}
                {colPick("amps", "Motor current (A)", false)}
                {colPick("waterCut", "Water cut (% or 0–1)", false)}
                {colPick("x", "Date / test (for trends)", false)}
              </div>
              {duplicate && !missing && <p className="zee-muted">Each required measurement needs its own column.</p>}
            </section>
            <section className="zee-draw-section">
              <h4>2 · Pump & fluid</h4>
              <div className="zee-grid2">
                {numField("baseHz", "Catalog frequency (Hz)")}
                {numField("stages", "Stages", "Leave empty if curve head is total head")}
                {numField("sg", "Fluid SG", "Used when no water cut + API")}
                {numField("api", "Oil API", "With water cut → SG per row")}
                {numField("motorAmps", "Motor nameplate amps")}
                {numField("minPip", "Min PIP / bubble point (psi)")}
                {numField("minRate", "Min recommended rate @ catalog Hz")}
                {numField("maxRate", "Max recommended rate @ catalog Hz")}
                {numField("degradationPct", "Head-loss flag at (%)")}
              </div>
              <Field
                size="small"
                label="Pump curve range (optional)"
                hint="Two columns at the catalog frequency: rate (bbl/d) and head per stage (or total head). Select it in Excel, then click Use selection."
              >
                <div className="zee-row">
                  <Input size="small" className="zee-grow" value={curve} onChange={(_, d) => setCurve(d.value)} placeholder="e.g. Curves!A2:B60" />
                  <Button
                    size="small"
                    onClick={async () => {
                      const a = await onPickSelection();
                      if (a) setCurve(a);
                    }}
                  >
                    Use selection
                  </Button>
                </div>
              </Field>
            </section>
          </DialogContent>
          <DialogActions>
            <Button appearance="secondary" onClick={onClose}>
              Cancel
            </Button>
            <Button appearance="primary" disabled={missing || duplicate} onClick={run}>
              Analyse
            </Button>
          </DialogActions>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
}
