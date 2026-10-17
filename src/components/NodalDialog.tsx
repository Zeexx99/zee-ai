import { useEffect, useState } from "react";
import { Button, Combobox, Dialog, DialogActions, DialogBody, DialogContent, DialogSurface, DialogTitle, Dropdown, Field, Input, Option } from "@fluentui/react-components";
import { SENSITIVITY_LABELS, type IprModel, type NodalInput, type SensitivityKind } from "../utils/nodal";

type Form = Record<
  "pr" | "pi" | "qTest" | "pwfTest" | "pb" | "depth" | "tubingId" | "whp" | "gor" | "waterCut" | "api" | "gasSg" | "waterSg" | "tWh" | "tBh" | "sensValues",
  string
>;

const DEFAULTS: Form = {
  pr: "3000",
  pi: "",
  qTest: "1000",
  pwfTest: "2000",
  pb: "2200",
  depth: "8000",
  tubingId: "2.441",
  whp: "150",
  gor: "500",
  waterCut: "50",
  api: "35",
  gasSg: "0.7",
  waterSg: "1.05",
  tWh: "100",
  tBh: "200",
  sensValues: "",
};

const IPR_MODELS: Array<{ key: IprModel; label: string; hint: string }> = [
  { key: "composite", label: "Composite (PI above Pb, Vogel below)", hint: "Best for most oil wells: straight line above the bubble point, Vogel below. Uses the test (or PI) and Pb." },
  { key: "vogel", label: "Vogel (from one test)", hint: "Saturated reservoir (Pr ≤ Pb). Needs one test: rate and flowing pressure." },
  { key: "pi", label: "Straight-line PI", hint: "Undersaturated / high water cut. Type a PI, or give one test and it is calculated." },
];

const TUBING = ["1.995", "2.441", "2.992", "3.476", "3.958", "4.892"];

const SENS_HINT: Record<SensitivityKind, string> = {
  none: "",
  whp: "e.g. 100, 200, 300",
  tubingId: "e.g. 1.995, 2.992, 3.958",
  waterCut: "e.g. 30, 60, 90",
  gor: "e.g. 300, 800, 1500",
};

function fromInput(i: NodalInput | null): { form: Form; model: IprModel; sens: SensitivityKind } {
  if (!i) return { form: DEFAULTS, model: "composite", sens: "none" };
  const s = (v: number | undefined) => (v === undefined || !Number.isFinite(v) ? "" : String(v));
  return {
    model: i.ipr.model,
    sens: i.sensitivity?.kind ?? "none",
    form: {
      pr: s(i.ipr.pr),
      pi: s(i.ipr.pi),
      qTest: s(i.ipr.qTest),
      pwfTest: s(i.ipr.pwfTest),
      pb: s(i.ipr.pb),
      depth: s(i.well.depth),
      tubingId: s(i.well.tubingId),
      whp: s(i.well.whp),
      gor: s(i.well.gor),
      waterCut: s(i.well.waterCut * 100),
      api: s(i.well.api),
      gasSg: s(i.well.gasSg),
      waterSg: s(i.well.waterSg),
      tWh: s(i.well.tWh),
      tBh: s(i.well.tBh),
      sensValues: i.sensitivity?.values.join(", ") ?? "",
    },
  };
}

export function NodalDialog({ open, last, onClose, onRun }: { open: boolean; last: NodalInput | null; onClose: () => void; onRun: (i: NodalInput) => void }) {
  const [form, setForm] = useState<Form>(DEFAULTS);
  const [model, setModel] = useState<IprModel>("composite");
  const [sens, setSens] = useState<SensitivityKind>("none");
  useEffect(() => {
    if (!open) return;
    const f = fromInput(last);
    setForm(f.form);
    setModel(f.model);
    setSens(f.sens);
  }, [open, last]);
  if (!open) return null;

  const num = (k: keyof Form) => (form[k].trim() === "" ? undefined : Number(form[k]));
  const set = (k: keyof Form) => (_: unknown, d: { value: string }) => setForm((p) => ({ ...p, [k]: d.value }));
  const box = (k: keyof Form, label: string, hint?: string) => (
    <Field label={label} hint={hint} validationState={form[k].trim() !== "" && !Number.isFinite(Number(form[k])) ? "error" : "none"}>
      <Input value={form[k]} onChange={set(k)} inputMode="decimal" />
    </Field>
  );
  const m = IPR_MODELS.find((x) => x.key === model) ?? IPR_MODELS[0];
  const values = form.sensValues
    .split(/[,;\s]+/)
    .map((x) => Number(x))
    .filter((x) => Number.isFinite(x));

  const required: Array<keyof Form> = ["pr", "depth", "tubingId", "whp", "gor", "waterCut", "api", "gasSg", "waterSg", "tWh", "tBh"];
  const hasTest = num("qTest") !== undefined && num("pwfTest") !== undefined;
  const okIpr = model === "vogel" ? hasTest : model === "pi" ? hasTest || (num("pi") ?? 0) > 0 : (hasTest || (num("pi") ?? 0) > 0) && num("pb") !== undefined;
  const valid = required.every((k) => Number.isFinite(num(k) ?? NaN)) && okIpr && (sens === "none" || values.length > 0);

  const run = () =>
    onRun({
      ipr: { model, pr: num("pr")!, pi: num("pi"), qTest: num("qTest"), pwfTest: num("pwfTest"), pb: model === "composite" ? num("pb") : undefined },
      well: {
        depth: num("depth")!,
        tubingId: num("tubingId")!,
        whp: num("whp")!,
        gor: num("gor")!,
        waterCut: num("waterCut")! / 100,
        api: num("api")!,
        gasSg: num("gasSg")!,
        waterSg: num("waterSg")!,
        tWh: num("tWh")!,
        tBh: num("tBh")!,
      },
      sensitivity: sens === "none" ? undefined : { kind: sens, values: values.slice(0, 6) },
    });

  return (
    <Dialog open onOpenChange={(_, d) => !d.open && onClose()}>
      <DialogSurface className="zee-dialog">
        <DialogBody>
          <DialogTitle>Nodal analysis</DialogTitle>
          <DialogContent className="zee-form">
            <p className="zee-muted">Finds where the well's inflow (IPR) meets the tubing curve (VLP). Results, curves and a chart go to a new sheet. Pressures are absolute (psig + 14.7).</p>
            <h4 className="zee-form-section">Inflow</h4>
            <Field label="IPR model" hint={m.hint}>
              <Dropdown value={m.label} selectedOptions={[model]} onOptionSelect={(_, d) => d.optionValue && setModel(d.optionValue as IprModel)}>
                {IPR_MODELS.map((x) => (
                  <Option key={x.key} value={x.key}>
                    {x.label}
                  </Option>
                ))}
              </Dropdown>
            </Field>
            <div className="zee-grid2">
              {box("pr", "Reservoir pressure (psia)")}
              {model !== "vogel" ? box("pi", "PI (STB/d/psi)", "Leave empty to use the test") : <span />}
              {box("qTest", "Test liquid rate (STB/d)")}
              {box("pwfTest", "Test Pwf (psia)")}
              {model === "composite" && box("pb", "Bubble point (psia)")}
            </div>
            <h4 className="zee-form-section">Well & fluids</h4>
            <div className="zee-grid2">
              {box("depth", "Mid-perf depth TVD (ft)")}
              <Field label="Tubing ID (in)" hint="2 3/8″ = 1.995 · 2 7/8″ = 2.441 · 3 1/2″ = 2.992 · 4 1/2″ = 3.958">
                <Combobox
                  freeform
                  value={form.tubingId}
                  selectedOptions={TUBING.includes(form.tubingId) ? [form.tubingId] : []}
                  onOptionSelect={(_, d) => d.optionValue && setForm((p) => ({ ...p, tubingId: d.optionValue! }))}
                  onChange={(e) => setForm((p) => ({ ...p, tubingId: e.target.value }))}
                >
                  {TUBING.map((t) => (
                    <Option key={t} value={t}>
                      {t}
                    </Option>
                  ))}
                </Combobox>
              </Field>
              {box("whp", "Wellhead pressure (psia)")}
              {box("gor", "Producing GOR (scf/STB)")}
              {box("waterCut", "Water cut (%)")}
              {box("api", "Oil API")}
              {box("gasSg", "Gas SG (air = 1)")}
              {box("waterSg", "Water SG")}
              {box("tWh", "Wellhead temp (°F)")}
              {box("tBh", "Bottom-hole temp (°F)")}
            </div>
            <h4 className="zee-form-section">What if…</h4>
            <Field label="Sensitivity (optional)" hint="Adds one tubing curve and operating point per value (up to 6).">
              <Dropdown value={SENSITIVITY_LABELS[sens]} selectedOptions={[sens]} onOptionSelect={(_, d) => d.optionValue && setSens(d.optionValue as SensitivityKind)}>
                {(Object.keys(SENSITIVITY_LABELS) as SensitivityKind[]).map((k) => (
                  <Option key={k} value={k}>
                    {SENSITIVITY_LABELS[k]}
                  </Option>
                ))}
              </Dropdown>
            </Field>
            {sens !== "none" && (
              <Field label="Values" hint={SENS_HINT[sens]}>
                <Input value={form.sensValues} onChange={set("sensValues")} />
              </Field>
            )}
          </DialogContent>
          <DialogActions>
            <Button appearance="primary" disabled={!valid} onClick={run}>
              Run nodal analysis
            </Button>
            <Button onClick={onClose}>Cancel</Button>
          </DialogActions>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
}
