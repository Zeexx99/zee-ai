import { useEffect, useState } from "react";
import { Button, Checkbox, Dialog, DialogActions, DialogBody, DialogContent, DialogSurface, DialogTitle, Dropdown, Field, Input, Option } from "@fluentui/react-components";
import type { ConstantChoice, ConstantsScan, DupeChoice, DupeScan } from "../services/dataTools";
import type { GapMethod, GapOptions } from "../utils/gaps";
import { isValidName, knownConstant } from "../utils/constants";

const NONE = -1;

/* ------------------------------------------------------------ 18. duplicates & outliers */

export function DupesDialog({ scan, onClose, onApply }: { scan: DupeScan | null; onClose: () => void; onApply: (c: DupeChoice) => void }) {
  const [c, setC] = useState<DupeChoice>({ highlight: true, unify: false, deleteDuplicates: false });
  useEffect(() => {
    if (scan) setC({ highlight: true, unify: false, deleteDuplicates: false });
  }, [scan]);
  if (!scan) return null;
  const firstRow = scan.sel.block.rowIndex + 2; // sheet row of data row 0
  const dupRows = scan.duplicates.reduce((n, g) => n + g.dupes.length, 0);
  const variantCells = scan.variants.reduce((n, v) => n + v.groups.reduce((m, g) => m + g.variants.slice(1).reduce((k, x) => k + x.n, 0), 0), 0);
  const outlierCount = scan.outliers.reduce((n, o) => n + o.items.length, 0);
  const nothing = !dupRows && !variantCells && !outlierCount;
  const any = c.highlight || (c.unify && variantCells > 0) || (c.deleteDuplicates && dupRows > 0);
  return (
    <Dialog open onOpenChange={(_, d) => !d.open && onClose()}>
      <DialogSurface className="zee-dialog">
        <DialogBody>
          <DialogTitle>Duplicates & outliers</DialogTitle>
          <DialogContent className="zee-form">
            <p className="zee-muted">
              {scan.sel.rows.length.toLocaleString()} rows × {scan.sel.headers.length} columns checked in {scan.sel.block.address}.
            </p>
            {nothing && <p>No duplicate rows, spelling variants or outliers found. 👍</p>}

            {dupRows > 0 && (
              <div className="zee-finding">
                <b className="zee-swatch zee-swatch-dup">
                  {dupRows} duplicate row{dupRows === 1 ? "" : "s"}
                </b>
                <ul>
                  {scan.duplicates.slice(0, 4).map((g) => (
                    <li key={g.keep}>
                      Row {firstRow + g.keep} repeated in row{g.dupes.length === 1 ? "" : "s"} {g.dupes.slice(0, 6).map((r) => firstRow + r).join(", ")}
                      {g.dupes.length > 6 ? "…" : ""}
                    </li>
                  ))}
                  {scan.duplicates.length > 4 && <li>… and {scan.duplicates.length - 4} more groups</li>}
                </ul>
              </div>
            )}

            {variantCells > 0 && (
              <div className="zee-finding">
                <b className="zee-swatch zee-swatch-var">
                  {variantCells} cell{variantCells === 1 ? "" : "s"} spelled differently
                </b>
                <ul>
                  {scan.variants.slice(0, 4).flatMap((v) =>
                    v.groups.slice(0, 3).map((g) => (
                      <li key={`${v.col}-${g.canonical}`}>
                        {v.header}: {g.variants.map((x) => `"${x.value}" (${x.n})`).join(", ")} → <b>{g.canonical}</b>
                      </li>
                    ))
                  )}
                </ul>
              </div>
            )}

            {outlierCount > 0 && (
              <div className="zee-finding">
                <b className="zee-swatch zee-swatch-out">
                  {outlierCount} outlier{outlierCount === 1 ? "" : "s"}
                </b>
                <ul>
                  {scan.outliers.slice(0, 5).map((o) => (
                    <li key={o.col}>
                      {o.header}:{" "}
                      {o.items
                        .slice(0, 4)
                        .map((x) => `${x.value.toLocaleString()} (row ${firstRow + x.row}, ${x.direction})`)
                        .join(", ")}
                      {o.items.length > 4 ? "…" : ""}
                    </li>
                  ))}
                </ul>
                <span className="zee-muted">Outliers are only flagged – check them, they may be real (e.g. a well test or shut-in).</span>
              </div>
            )}

            {!nothing && (
              <>
                <Checkbox checked={c.highlight} onChange={(_, d) => setC((p) => ({ ...p, highlight: !!d.checked }))} label="Highlight them (yellow = duplicate, blue = spelling, orange = outlier)" />
                <Checkbox
                  disabled={!variantCells}
                  checked={c.unify && variantCells > 0}
                  onChange={(_, d) => setC((p) => ({ ...p, unify: !!d.checked }))}
                  label="Unify spellings to the most common one"
                />
                <Checkbox
                  disabled={!dupRows}
                  checked={c.deleteDuplicates && dupRows > 0}
                  onChange={(_, d) => setC((p) => ({ ...p, deleteDuplicates: !!d.checked }))}
                  label={`Delete the ${dupRows} duplicate row${dupRows === 1 ? "" : "s"} (first one is kept)`}
                />
                <p className="zee-muted">Each action is a separate undo step (↶ or the history list).</p>
              </>
            )}
          </DialogContent>
          <DialogActions>
            {!nothing && (
              <Button appearance="primary" disabled={!any} onClick={() => onApply(c)}>
                Apply
              </Button>
            )}
            <Button onClick={onClose}>{nothing ? "Close" : "Cancel"}</Button>
          </DialogActions>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
}

/* ------------------------------------------------------------ 19. gap filler */

const METHODS: Array<{ key: GapMethod; label: string; hint: string }> = [
  { key: "linear", label: "Straight line between neighbours", hint: "Rates move smoothly from the reading before to the one after." },
  { key: "previous", label: "Repeat the previous value", hint: "Carries the last reading forward (also copies text columns)." },
  { key: "blank", label: "Add the missing dates only", hint: "Rows are added with empty values – nothing is invented." },
];

export function GapDialog({ headers, onClose, onRun }: { headers: string[] | null; onClose: () => void; onRun: (o: GapOptions) => void }) {
  const [dateCol, setDateCol] = useState(NONE);
  const [groupCol, setGroupCol] = useState(NONE);
  const [method, setMethod] = useState<GapMethod>("linear");
  const [maxGap, setMaxGap] = useState("31");
  useEffect(() => {
    if (!headers) return;
    setDateCol(headers.findIndex((h) => /date|day|time|تاريخ/i.test(h)));
    setGroupCol(headers.findIndex((h) => /well|بئر/i.test(h)));
  }, [headers]);
  if (!headers) return null;
  const m = METHODS.find((x) => x.key === method) ?? METHODS[0];
  const steps = Number(maxGap);
  const valid = dateCol >= 0 && dateCol !== groupCol && Number.isFinite(steps) && steps >= 0;
  const pick = (value: number, set: (v: number) => void, optional: boolean) => (
    <Dropdown
      value={value >= 0 ? headers[value] ?? "" : optional ? "(none)" : "Choose…"}
      selectedOptions={[String(value)]}
      onOptionSelect={(_, d) => d.optionValue !== undefined && set(Number(d.optionValue))}
    >
      {optional && <Option value={String(NONE)}>(none)</Option>}
      {headers.map((h, i) => (
        <Option key={i} value={String(i)}>
          {h}
        </Option>
      ))}
    </Dropdown>
  );
  return (
    <Dialog open onOpenChange={(_, d) => !d.open && onClose()}>
      <DialogSurface className="zee-dialog">
        <DialogBody>
          <DialogTitle>Fill gaps in time series</DialogTitle>
          <DialogContent className="zee-form">
            <p className="zee-muted">Adds the missing dates (daily, weekly or monthly – detected from the data) and fills empty numbers. The result goes to a new sheet; your data isn't touched. Added rows are blue, filled cells yellow.</p>
            <Field label="Date column">{pick(dateCol, setDateCol, false)}</Field>
            <Field label="Well column (optional)" hint="Gaps are filled per well.">
              {pick(groupCol, setGroupCol, true)}
            </Field>
            <Field label="Fill method" hint={m.hint}>
              <Dropdown value={m.label} selectedOptions={[method]} onOptionSelect={(_, d) => d.optionValue && setMethod(d.optionValue as GapMethod)}>
                {METHODS.map((x) => (
                  <Option key={x.key} value={x.key}>
                    {x.label}
                  </Option>
                ))}
              </Dropdown>
            </Field>
            <Field label="Longest gap to fill (steps)" hint="Longer gaps (e.g. a well shut in for months) are left alone. 0 = no limit.">
              <Input type="number" min={0} value={maxGap} onChange={(_, d) => setMaxGap(d.value)} />
            </Field>
          </DialogContent>
          <DialogActions>
            <Button appearance="primary" disabled={!valid} onClick={() => onRun({ dateCol, groupCol: groupCol >= 0 ? groupCol : undefined, method, maxGapSteps: Math.floor(steps) })}>
              Fill gaps
            </Button>
            <Button onClick={onClose}>Cancel</Button>
          </DialogActions>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
}

/* ------------------------------------------------------------ 29. hard-coded numbers → named inputs */

interface ConstRow extends ConstantChoice {
  on: boolean;
  count: number;
  example: string;
}

export function ConstantsDialog({ scan, onClose, onApply }: { scan: ConstantsScan | null; onClose: () => void; onApply: (chosen: ConstantChoice[]) => void }) {
  const [list, setList] = useState<ConstRow[]>([]);
  useEffect(() => {
    if (!scan) return;
    setList(
      scan.uses.map((u) => ({
        value: u.value,
        name: scan.suggestions[u.value]?.name ?? "",
        note: scan.suggestions[u.value]?.note ?? "",
        on: !!knownConstant(u.value) || u.count >= 3,
        count: u.count,
        example: u.example,
      }))
    );
  }, [scan]);
  if (!scan) return null;
  const patch = (i: number, p: Partial<ConstRow>) => setList((l) => l.map((r, k) => (k === i ? { ...r, ...p } : r)));
  const chosen = list.filter((r) => r.on);
  const names = chosen.map((r) => r.name.toLowerCase());
  const badName = (r: ConstRow) => r.on && (!isValidName(r.name) || names.filter((n) => n === r.name.toLowerCase()).length > 1);
  const invalid = chosen.some(badName);
  return (
    <Dialog open onOpenChange={(_, d) => !d.open && onClose()}>
      <DialogSurface className="zee-dialog zee-dialog-wide">
        <DialogBody>
          <DialogTitle>Move hard-coded numbers to inputs</DialogTitle>
          <DialogContent className="zee-form">
            {!list.length ? (
              <p>No hard-coded numbers found in the formulas on "{scan.sheet}". 👍</p>
            ) : (
              <>
                <p className="zee-muted">
                  Ticked numbers go to an <b>Inputs</b> sheet as named cells, and the formulas on "{scan.sheet}" are rewritten to use the names (e.g. <code>=B2*5.615</code> → <code>=B2*FT3_PER_BBL</code>). Change
                  a value once on Inputs and every formula follows. One undo step.
                </p>
                <div className="zee-const-list">
                  {list.map((r, i) => (
                    <div key={r.value} className="zee-const">
                      <Checkbox checked={r.on} onChange={(_, d) => patch(i, { on: !!d.checked })} label={<b className="zee-const-val">{Number(r.value).toLocaleString(undefined, { maximumFractionDigits: 10 })}</b>} />
                      <span className="zee-muted">
                        {r.count}× · <code title={r.example}>{r.example.length > 40 ? `${r.example.slice(0, 38)}…` : r.example}</code>
                      </span>
                      {r.on && (
                        <div className="zee-row">
                          <Input size="small" value={r.name} onChange={(_, d) => patch(i, { name: d.value.trim() })} aria-invalid={badName(r)} className={badName(r) ? "zee-invalid" : undefined} aria-label="Name" />
                          <Input size="small" value={r.note} onChange={(_, d) => patch(i, { note: d.value })} placeholder="Note" aria-label="Note" />
                        </div>
                      )}
                    </div>
                  ))}
                </div>
                {invalid && <p className="zee-error-text">Fix the red names: letters, digits, _ and . only, not like a cell address (e.g. BBL1), and each used once.</p>}
              </>
            )}
          </DialogContent>
          <DialogActions>
            {!!list.length && (
              <Button appearance="primary" disabled={!chosen.length || invalid} onClick={() => onApply(chosen.map(({ value, name, note }) => ({ value, name, note })))}>
                Create {chosen.length} input{chosen.length === 1 ? "" : "s"}
              </Button>
            )}
            <Button onClick={onClose}>{list.length ? "Cancel" : "Close"}</Button>
          </DialogActions>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
}
