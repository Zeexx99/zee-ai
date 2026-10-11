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
import type { ChartKind, PlotKind, PlotSource } from "../services/excelService";

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
          <DialogContent>{state?.message}</DialogContent>
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
}

export function DeclineCurveDialog({ open, headers, onClose, onRun }: { open: boolean; headers: string[]; onClose: () => void; onRun: (c: DcaChoice) => void }) {
  const [timeCol, setTimeCol] = useState(0);
  const [rateCol, setRateCol] = useState(1);
  const [periods, setPeriods] = useState("60");
  const [limit, setLimit] = useState("");
  useEffect(() => {
    if (open) {
      setTimeCol(0);
      setRateCol(Math.min(1, Math.max(0, headers.length - 1)));
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
              disabled={timeCol === rateCol}
              onClick={() =>
                onRun({
                  timeColumn: timeCol,
                  rateColumn: rateCol,
                  forecastPeriods: Math.max(0, Math.min(5000, Math.round(Number(periods) || 0))),
                  economicLimit: Number(limit) > 0 ? Number(limit) : undefined,
                })
              }
            >
              Fit curves
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
  kind: PlotKind;
  title: string;
}

const PLOT_LABELS: Record<PlotKind, string> = {
  line: "Line",
  scatterLines: "Scatter with lines (numeric / date X)",
  scatter: "Scatter (points only)",
  column: "Column",
};

export function DrawDialog({ source, onClose, onRun }: { source: PlotSource | null; onClose: () => void; onRun: (c: DrawChoice) => void }) {
  const cols = source?.columns ?? [];
  const [xIndex, setXIndex] = useState(0);
  const [yIndexes, setYIndexes] = useState<number[]>([]);
  const [kind, setKind] = useState<PlotKind>("line");
  const [title, setTitle] = useState("");

  // New data source → sensible defaults: first column on X, the next numeric columns (up to 4) on Y.
  useEffect(() => {
    if (!source) return;
    const x = source.columns[0]?.index ?? 0;
    setXIndex(x);
    setYIndexes(source.columns.filter((c) => c.index !== x && c.numeric).slice(0, 4).map((c) => c.index));
    setKind(source.columns[0]?.numeric ? "scatterLines" : "line");
    setTitle("");
  }, [source]);

  const xCol = cols.find((c) => c.index === xIndex);
  const yChoices = cols.filter((c) => c.index !== xIndex);
  const toggle = (idx: number, on: boolean) =>
    setYIndexes((prev) => (on ? [...prev, idx].sort((a, b) => a - b) : prev.filter((i) => i !== idx)));
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
                  <Checkbox
                    key={c.index}
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

            <Field label="Chart type">
              <Dropdown value={PLOT_LABELS[kind]} selectedOptions={[kind]} onOptionSelect={(_, d) => d.optionValue && setKind(d.optionValue as PlotKind)}>
                {(Object.keys(PLOT_LABELS) as PlotKind[]).map((k) => (
                  <Option key={k} value={k}>
                    {PLOT_LABELS[k]}
                  </Option>
                ))}
              </Dropdown>
            </Field>
            <Field label="Title (optional)" hint="Leave empty for “Y names vs X name”.">
              <Input value={title} onChange={(_, d) => setTitle(d.value)} />
            </Field>
          </DialogContent>
          <DialogActions>
            <Button appearance="secondary" onClick={onClose}>
              Cancel
            </Button>
            <Button appearance="primary" disabled={!yIndexes.length} onClick={() => onRun({ xIndex, yIndexes, kind, title: title.trim() })}>
              Draw chart
            </Button>
          </DialogActions>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
}
