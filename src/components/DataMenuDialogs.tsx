import { useEffect, useMemo, useState } from "react";
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
  Radio,
  RadioGroup,
  Spinner,
} from "@fluentui/react-components";
import { IMPORT_ACCEPT, IMPORT_MAX_BYTES, parseFile, pivotLong, type ParsedFile } from "../utils/fileImport";
import { guessKeys, type DateMatch, type MergeHow } from "../utils/merge";
import { AGG_LABELS, FILL_LABELS, PERIOD_LABELS, guessAgg, type Agg, type Period, type PeriodFill, type ResampleOptions } from "../utils/resample";
import type { ImportTarget, MergePick, SheetInfo } from "../services/dataMenu";
import type { MergeOptions } from "../utils/merge";

const NONE = -1;

function ColumnPick({ headers, value, onChange, optional, noneLabel = "(none)" }: { headers: string[]; value: number; onChange: (v: number) => void; optional?: boolean; noneLabel?: string }) {
  return (
    <Dropdown
      value={value >= 0 ? headers[value] ?? "" : optional ? noneLabel : "Choose…"}
      selectedOptions={[String(value)]}
      onOptionSelect={(_, d) => d.optionValue !== undefined && onChange(Number(d.optionValue))}
    >
      {optional && <Option value={String(NONE)}>{noneLabel}</Option>}
      {headers.map((h, i) => (
        <Option key={i} value={String(i)}>
          {h}
        </Option>
      ))}
    </Dropdown>
  );
}

/* ================================================================ import file */

type Encoding = "utf-8" | "windows-1256" | "windows-1252";
const ENCODINGS: Array<{ key: Encoding; label: string }> = [
  { key: "utf-8", label: "UTF-8 (most files)" },
  { key: "windows-1256", label: "Arabic (Windows-1256)" },
  { key: "windows-1252", label: "Western (Windows-1252)" },
];
const DELIMS: Array<{ key: string; label: string }> = [
  { key: "auto", label: "Detect" },
  { key: ",", label: "Comma ," },
  { key: ";", label: "Semicolon ;" },
  { key: "\t", label: "Tab" },
  { key: "|", label: "Pipe |" },
];

export interface ImportChoice {
  fileName: string;
  parsed: ParsedFile;
  target: ImportTarget;
}

export function ImportDialog({ open, onClose, onRun }: { open: boolean; onClose: () => void; onRun: (c: ImportChoice) => void }) {
  const [file, setFile] = useState<File | null>(null);
  const [encoding, setEncoding] = useState<Encoding>("utf-8");
  const [delim, setDelim] = useState("auto");
  const [order, setOrder] = useState<"DMY" | "MDY">("DMY");
  const [skip, setSkip] = useState("");
  const [pivot, setPivot] = useState(true);
  const [roundMin, setRoundMin] = useState("0");
  const [target, setTarget] = useState<ImportTarget>("newSheet");
  const [parsed, setParsed] = useState<ParsedFile | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [busy, setBusy] = useState(false);

  useEffect(() => {
    if (!open) {
      setFile(null);
      setParsed(null);
      setError(null);
    }
  }, [open]);

  useEffect(() => {
    if (!file) return;
    let cancelled = false;
    setBusy(true);
    setError(null);
    (async () => {
      try {
        if (file.size > IMPORT_MAX_BYTES) throw new Error(`The file is ${(file.size / 1e6).toFixed(0)} MB – the limit is ${IMPORT_MAX_BYTES / 1e6} MB. Split it, or filter it first.`);
        const buf = await file.arrayBuffer();
        let text: string;
        try {
          text = new TextDecoder(encoding).decode(buf);
        } catch {
          text = new TextDecoder("utf-8").decode(buf);
        }
        // Give the spinner a frame before parsing a big file.
        await new Promise((r) => setTimeout(r, 0));
        const skipRows = skip.trim() === "" ? undefined : Math.max(0, Math.floor(Number(skip)) || 0);
        const p = parseFile(file.name, text, { delimiter: delim === "auto" ? undefined : delim, dateOrder: order, skipRows });
        if (!cancelled) setParsed(p);
      } catch (e) {
        if (!cancelled) {
          setParsed(null);
          setError((e as Error).message ?? String(e));
        }
      } finally {
        if (!cancelled) setBusy(false);
      }
    })();
    return () => {
      cancelled = true;
    };
  }, [file, encoding, delim, order, skip]);

  const final = useMemo(() => (parsed && parsed.long && pivot ? pivotLong(parsed, Number(roundMin) || 0) : parsed), [parsed, pivot, roundMin]);
  if (!open) return null;
  const preview = final ? final.rows.slice(0, 5) : [];
  const cols = final ? Math.min(final.headers.length, 8) : 0;
  const show = (v: unknown, j: number) => {
    if (final?.dateCols.includes(j) && typeof v === "number") {
      const d = new Date(Date.UTC(1899, 11, 30) + v * 86_400_000);
      return d.toISOString().replace("T", " ").slice(0, v % 1 ? 16 : 10);
    }
    return String(v ?? "");
  };

  return (
    <Dialog open onOpenChange={(_, d) => !d.open && onClose()}>
      <DialogSurface className="zee-dialog">
        <DialogBody>
          <DialogTitle>Import file</DialogTitle>
          <DialogContent className="zee-form">
            <p className="zee-muted">CSV / TXT / TSV, LAS 2.0 well logs and SCADA or historian exports. Numbers, dates and Arabic digits are converted on the way in.</p>
            <Field label="File">
              <input type="file" accept={IMPORT_ACCEPT} onChange={(e) => setFile(e.target.files?.[0] ?? null)} />
            </Field>
            {file && parsed?.kind !== "las" && (
              <div className="zee-grid2">
                <Field label="Encoding">
                  <Dropdown value={ENCODINGS.find((x) => x.key === encoding)?.label} selectedOptions={[encoding]} onOptionSelect={(_, d) => d.optionValue && setEncoding(d.optionValue as Encoding)}>
                    {ENCODINGS.map((x) => (
                      <Option key={x.key} value={x.key}>
                        {x.label}
                      </Option>
                    ))}
                  </Dropdown>
                </Field>
                <Field label="Separator">
                  <Dropdown value={DELIMS.find((x) => x.key === delim)?.label} selectedOptions={[delim]} onOptionSelect={(_, d) => d.optionValue && setDelim(d.optionValue)}>
                    {DELIMS.map((x) => (
                      <Option key={x.key} value={x.key}>
                        {x.label}
                      </Option>
                    ))}
                  </Dropdown>
                </Field>
                <Field label="Dates like 03/04/2026 are">
                  <Dropdown value={order === "DMY" ? "Day first (3 Apr)" : "Month first (Mar 4)"} selectedOptions={[order]} onOptionSelect={(_, d) => d.optionValue && setOrder(d.optionValue as "DMY" | "MDY")}>
                    <Option value="DMY">Day first (3 Apr)</Option>
                    <Option value="MDY">Month first (Mar 4)</Option>
                  </Dropdown>
                </Field>
                <Field label="Lines above the header" hint="Empty = detect">
                  <Input type="number" min={0} value={skip} onChange={(_, d) => setSkip(d.value)} />
                </Field>
              </div>
            )}
            {busy && <Spinner size="tiny" label="Reading the file…" />}
            {error && <p className="zee-error-text">{error}</p>}
            {final && !busy && (
              <>
                <p className="zee-muted">
                  <b>{final.kind === "las" ? "LAS well log" : final.kind === "scada" ? "SCADA / historian export" : "Delimited text"}</b> · {final.rows.length.toLocaleString()} rows × {final.headers.length} columns
                  {final.dateCols.length ? ` · ${final.dateCols.length} date column${final.dateCols.length === 1 ? "" : "s"}` : ""}
                </p>
                {final.notes.map((n, i) => (
                  <p key={i} className="zee-muted">
                    {n}
                  </p>
                ))}
                {parsed?.long && (
                  <div className="zee-row">
                    <Checkbox checked={pivot} onChange={(_, d) => setPivot(!!d.checked)} label="One column per tag (pivot the Tag / Time / Value list)" />
                    {pivot && (
                      <Field label="Round times to (min)" orientation="horizontal">
                        <Input type="number" min={0} value={roundMin} onChange={(_, d) => setRoundMin(d.value)} style={{ width: 70 }} />
                      </Field>
                    )}
                  </div>
                )}
                <div className="zee-preview">
                  <table>
                    <thead>
                      <tr>
                        {final.headers.slice(0, cols).map((h, j) => (
                          <th key={j}>{h}</th>
                        ))}
                        {final.headers.length > cols && <th>…</th>}
                      </tr>
                    </thead>
                    <tbody>
                      {preview.map((r, i) => (
                        <tr key={i}>
                          {r.slice(0, cols).map((v, j) => (
                            <td key={j}>{show(v, j)}</td>
                          ))}
                          {final.headers.length > cols && <td />}
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
                <Field label="Put it">
                  <RadioGroup value={target} onChange={(_, d) => setTarget(d.value as ImportTarget)} layout="horizontal">
                    <Radio value="newSheet" label="On a new sheet" />
                    <Radio value="activeCell" label="At the active cell" />
                  </RadioGroup>
                </Field>
              </>
            )}
          </DialogContent>
          <DialogActions>
            <Button appearance="primary" disabled={!final || busy || !final.rows.length} onClick={() => file && final && onRun({ fileName: file.name, parsed: final, target })}>
              Import {final ? `${final.rows.length.toLocaleString()} rows` : ""}
            </Button>
            <Button onClick={onClose}>Cancel</Button>
          </DialogActions>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
}

/* ================================================================ merge well data */

interface PickState {
  on: boolean;
  wellCol: number;
  dateCol: number;
}

export function MergeDialog({ sheets, onClose, onRun }: { sheets: SheetInfo[] | null; onClose: () => void; onRun: (picks: MergePick[], o: MergeOptions) => void }) {
  const [state, setState] = useState<Record<string, PickState>>({});
  const [order, setOrder] = useState<string[]>([]);
  const [how, setHow] = useState<MergeHow>("outer");
  const [dateMatch, setDateMatch] = useState<DateMatch>("exact");
  const [maxAge, setMaxAge] = useState("0");
  useEffect(() => {
    if (!sheets) return;
    setState(Object.fromEntries(sheets.map((s) => [s.name, { on: false, ...guessKeys(s.headers) }])));
    setOrder([]);
  }, [sheets]);
  if (!sheets) return null;

  const toggle = (name: string, on: boolean) => {
    setState((p) => ({ ...p, [name]: { ...p[name], on } }));
    setOrder((p) => (on ? [...p.filter((x) => x !== name), name] : p.filter((x) => x !== name)));
  };
  const setCol = (name: string, k: "wellCol" | "dateCol", v: number) => setState((p) => ({ ...p, [name]: { ...p[name], [k]: v } }));
  const picked = order.filter((n) => state[n]?.on);
  const firstHasDate = picked.length ? (state[picked[0]]?.dateCol ?? NONE) >= 0 : false;
  const anyDate = picked.some((n) => (state[n]?.dateCol ?? NONE) >= 0);
  const needsFirstDate = anyDate && (how === "left" || dateMatch === "asof") && !firstHasDate;
  const valid = picked.length >= 2 && picked.every((n) => state[n].wellCol >= 0) && !needsFirstDate;
  const info = (n: string) => sheets.find((s) => s.name === n)!;

  return (
    <Dialog open onOpenChange={(_, d) => !d.open && onClose()}>
      <DialogSurface className="zee-dialog">
        <DialogBody>
          <DialogTitle>Merge well data</DialogTitle>
          <DialogContent className="zee-form">
            <p className="zee-muted">Joins sheets by Well (and Date) into one table on a new sheet. Well names match even when written differently ("W-01", "w 01"). Your sheets aren't changed.</p>
            <Field label="Sheets to merge" hint="Tick in order – the first ticked sheet is the main one.">
              <div className="zee-merge-list">
                {sheets.map((s) => (
                  <Checkbox
                    key={s.name}
                    checked={!!state[s.name]?.on}
                    onChange={(_, d) => toggle(s.name, !!d.checked)}
                    label={`${picked.indexOf(s.name) >= 0 ? `${picked.indexOf(s.name) + 1}. ` : ""}${s.name} (${s.rows.toLocaleString()} rows)`}
                  />
                ))}
              </div>
            </Field>
            {picked.map((n) => (
              <div key={n} className="zee-finding">
                <b>{n}</b>
                <div className="zee-grid2">
                  <Field label="Well column">
                    <ColumnPick headers={info(n).headers} value={state[n].wellCol} onChange={(v) => setCol(n, "wellCol", v)} />
                  </Field>
                  <Field label="Date column">
                    <ColumnPick headers={info(n).headers} value={state[n].dateCol} onChange={(v) => setCol(n, "dateCol", v)} optional noneLabel="(no date – one row per well)" />
                  </Field>
                </div>
              </div>
            ))}
            {picked.length >= 2 && (
              <>
                <Field label="Dates must">
                  <RadioGroup value={dateMatch} onChange={(_, d) => setDateMatch(d.value as DateMatch)}>
                    <Radio value="exact" label="Match exactly (same day)" />
                    <Radio value="asof" label={`Use the latest row on or before each date of "${picked[0]}"`} />
                  </RadioGroup>
                </Field>
                {dateMatch === "asof" ? (
                  <Field label="Ignore matches older than (days)" hint="e.g. 90 so a well test older than 3 months isn't used. 0 = no limit.">
                    <Input type="number" min={0} value={maxAge} onChange={(_, d) => setMaxAge(d.value)} />
                  </Field>
                ) : (
                  <Field label="Rows">
                    <RadioGroup value={how} onChange={(_, d) => setHow(d.value as MergeHow)}>
                      <Radio value="outer" label="Every well/date found in any sheet" />
                      <Radio value="left" label={`Only the wells/dates of "${picked[0]}"`} />
                    </RadioGroup>
                  </Field>
                )}
              </>
            )}
            {needsFirstDate && <p className="zee-error-text">Put a sheet with a date column first (e.g. daily production) – or pick its date column.</p>}
          </DialogContent>
          <DialogActions>
            <Button
              appearance="primary"
              disabled={!valid}
              onClick={() =>
                onRun(
                  picked.map((n) => ({ sheet: n, wellCol: state[n].wellCol, dateCol: state[n].dateCol })),
                  { how: dateMatch === "asof" ? "left" : how, dateMatch, maxAgeDays: Math.max(0, Number(maxAge) || 0) }
                )
              }
            >
              Merge {picked.length >= 2 ? `${picked.length} sheets` : ""}
            </Button>
            <Button onClick={onClose}>Cancel</Button>
          </DialogActions>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
}

/* ================================================================ resample */

export interface ResampleTarget {
  headers: string[];
  /** A sample of rows (to suggest an aggregation per column). */
  sample: unknown[][];
}

export function ResampleDialog({ target, onClose, onRun }: { target: ResampleTarget | null; onClose: () => void; onRun: (o: ResampleOptions) => void }) {
  const [dateCol, setDateCol] = useState(NONE);
  const [groupCol, setGroupCol] = useState(NONE);
  const [period, setPeriod] = useState<Period>("month");
  const [fill, setFill] = useState<PeriodFill>("none");
  const [countDays, setCountDays] = useState(true);
  const [aggs, setAggs] = useState<Record<number, Agg>>({});
  useEffect(() => {
    if (!target) return;
    const h = target.headers;
    setDateCol(h.findIndex((x) => /date|day|time|تاريخ/i.test(x)));
    setGroupCol(h.findIndex((x) => /well|بئر/i.test(x)));
    setAggs(Object.fromEntries(h.map((x, j) => [j, guessAgg(x, target.sample.map((r) => r[j]))])));
  }, [target]);
  if (!target) return null;
  const h = target.headers;
  const valueCols = h.map((_, j) => j).filter((j) => j !== dateCol && j !== groupCol);
  const valid = dateCol >= 0 && dateCol !== groupCol && valueCols.some((j) => aggs[j] !== "skip");
  return (
    <Dialog open onOpenChange={(_, d) => !d.open && onClose()}>
      <DialogSurface className="zee-dialog">
        <DialogBody>
          <DialogTitle>Resample time series</DialogTitle>
          <DialogContent className="zee-form">
            <p className="zee-muted">Rolls rows up to days, weeks, months, quarters or years (per well). Each column gets its own rule – rates are averaged, volumes and hours summed, cumulatives take the last value. The result goes to a new sheet.</p>
            <div className="zee-grid2">
              <Field label="Date column">
                <ColumnPick headers={h} value={dateCol} onChange={setDateCol} />
              </Field>
              <Field label="Well column (optional)">
                <ColumnPick headers={h} value={groupCol} onChange={setGroupCol} optional />
              </Field>
              <Field label="Period">
                <Dropdown value={PERIOD_LABELS[period]} selectedOptions={[period]} onOptionSelect={(_, d) => d.optionValue && setPeriod(d.optionValue as Period)}>
                  {(Object.keys(PERIOD_LABELS) as Period[]).map((p) => (
                    <Option key={p} value={p}>
                      {PERIOD_LABELS[p]}
                    </Option>
                  ))}
                </Dropdown>
              </Field>
              <Field label="Missing periods">
                <Dropdown value={FILL_LABELS[fill]} selectedOptions={[fill]} onOptionSelect={(_, d) => d.optionValue && setFill(d.optionValue as PeriodFill)}>
                  {(Object.keys(FILL_LABELS) as PeriodFill[]).map((p) => (
                    <Option key={p} value={p}>
                      {FILL_LABELS[p]}
                    </Option>
                  ))}
                </Dropdown>
              </Field>
            </div>
            <Field label="Rule per column">
              <div className="zee-agg-list">
                {valueCols.map((j) => (
                  <div key={j} className="zee-row">
                    <span className="zee-grow" title={h[j]}>
                      {h[j]}
                    </span>
                    <Dropdown className="zee-agg-pick" value={AGG_LABELS[aggs[j] ?? "avg"]} selectedOptions={[aggs[j] ?? "avg"]} onOptionSelect={(_, d) => d.optionValue && setAggs((p) => ({ ...p, [j]: d.optionValue as Agg }))}>
                      {(Object.keys(AGG_LABELS) as Agg[]).map((a) => (
                        <Option key={a} value={a}>
                          {AGG_LABELS[a]}
                        </Option>
                      ))}
                    </Dropdown>
                  </div>
                ))}
              </div>
            </Field>
            <Checkbox checked={countDays} onChange={(_, d) => setCountDays(!!d.checked)} label='Add a "Days with data" column' />
          </DialogContent>
          <DialogActions>
            <Button appearance="primary" disabled={!valid} onClick={() => onRun({ dateCol, groupCol: groupCol >= 0 ? groupCol : undefined, period, aggs, fill, countDays })}>
              Resample
            </Button>
            <Button onClick={onClose}>Cancel</Button>
          </DialogActions>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
}
