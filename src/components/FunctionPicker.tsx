import { useEffect, useMemo, useState } from "react";
import {
  Badge,
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
  Menu,
  MenuButton,
  MenuDivider,
  MenuGroup,
  MenuGroupHeader,
  MenuItem,
  MenuList,
  MenuPopover,
  MenuTrigger,
  Option,
} from "@fluentui/react-components";
import { LockClosed20Regular, TextEffects20Regular } from "@fluentui/react-icons";
import { CATEGORY_ORDER, FUNCTION_CATALOG, buildFormula, missingArgs, type FnSpec } from "../functions/catalog";
import type { FunctionTarget } from "../services/excelService";
import { a1, a1Range } from "../utils/a1";

/* ------------------------------------------------------------ toolbar drop list */

export function FunctionsMenu({ disabled, onPick, onFreeze }: { disabled?: boolean; onPick: (spec: FnSpec) => void; onFreeze: () => void }) {
  return (
    <Menu>
      <MenuTrigger disableButtonEnhancement>
        <MenuButton size="small" appearance="subtle" icon={<TextEffects20Regular />} disabled={disabled} title="Zee functions" aria-label="Functions">
          <span className="zee-quick-label">Functions</span>
        </MenuButton>
      </MenuTrigger>
      <MenuPopover className="zee-fn-popover">
        <MenuList>
          <MenuGroupHeader>Click → applies to the selected cells, or inserts in the active cell</MenuGroupHeader>
          {CATEGORY_ORDER.map((cat) => (
            <MenuGroup key={cat}>
              <MenuDivider />
              <MenuGroupHeader>{cat}</MenuGroupHeader>
              {FUNCTION_CATALOG.filter((f) => f.category === cat).map((f) => (
                <MenuItem key={f.id} onClick={() => onPick(f)}>
                  <span className="zee-fn-item">
                    <span className="zee-fn-name">
                      ZEE.{f.id}
                      {f.ai && (
                        <Badge size="tiny" appearance="tint" color="brand" className="zee-fn-ai">
                          AI
                        </Badge>
                      )}
                    </span>
                    <span className="zee-fn-brief">{f.brief}</span>
                  </span>
                </MenuItem>
              ))}
            </MenuGroup>
          ))}
          <MenuDivider />
          <MenuItem icon={<LockClosed20Regular />} onClick={onFreeze}>
            <span className="zee-fn-item">
              <span className="zee-fn-name">Freeze results</span>
              <span className="zee-fn-brief">Turn selected formulas into fixed values – no more AI calls</span>
            </span>
          </MenuItem>
        </MenuList>
      </MenuPopover>
    </Menu>
  );
}

/* ------------------------------------------------------------ arguments dialog */

export interface FunctionDialogState {
  spec: FnSpec;
  /** null outside Excel. */
  target: FunctionTarget | null;
  header: boolean;
}

export type ApplyWhere = "selection" | "cell";

export function FunctionDialog({
  state,
  onClose,
  onApply,
}: {
  state: FunctionDialogState | null;
  onClose: () => void;
  onApply: (spec: FnSpec, values: Record<string, string>, header: boolean, where: ApplyWhere) => void;
}) {
  const [values, setValues] = useState<Record<string, string>>({});
  const [header, setHeader] = useState(false);

  useEffect(() => {
    if (!state) return;
    const init: Record<string, string> = {};
    state.spec.args.forEach((a) => {
      if (a.kind === "choice") init[a.name] = a.default ?? a.choices?.[0] ?? "";
    });
    setValues(init);
    setHeader(state.header);
  }, [state]);

  const spec = state?.spec;
  const target = state?.target ?? null;
  const hasData = !!target && !target.isEmpty;
  const canSelection = !!spec && hasData && spec.mode !== "single";

  const missingSel = useMemo(() => (spec && canSelection ? missingArgs(spec, values, true, target!.columnCount) : []), [spec, values, canSelection, target]);
  const missingCell = useMemo(() => (spec ? missingArgs(spec, values, false) : []), [spec, values]);

  const preview = useMemo(() => {
    if (!spec) return "";
    if (canSelection && target) {
      const row = target.rowIndex + (header ? 1 : 0);
      const rows = Math.max(1, target.rowCount - (header ? 1 : 0));
      if (spec.mode === "range") return buildFormula(spec, values, [], a1Range(row, target.columnIndex, rows, target.columnCount, true));
      const cells = Array.from({ length: target.columnCount }, (_, k) => a1(row, target.columnIndex + k));
      return buildFormula(spec, values, cells);
    }
    return buildFormula(spec, values);
  }, [spec, values, canSelection, target, header]);

  if (!spec) return null;
  const set = (name: string, v: string) => setValues((p) => ({ ...p, [name]: v }));
  let inputNo = 0;

  return (
    <Dialog open={!!state} onOpenChange={(_, d) => !d.open && onClose()}>
      <DialogSurface>
        <DialogBody>
          <DialogTitle>
            ZEE.{spec.id} {spec.ai && <Badge appearance="tint" size="small">AI</Badge>}
          </DialogTitle>
          <DialogContent className="zee-form">
            <p className="zee-muted">{spec.brief}.</p>
            <code className="zee-fn-syntax">{spec.syntax}</code>
            {spec.args.map((a) => {
              if (a.kind === "choice")
                return (
                  <Field key={a.name} label={a.label}>
                    <Dropdown value={values[a.name] ?? ""} selectedOptions={[values[a.name] ?? ""]} onOptionSelect={(_, d) => set(a.name, d.optionValue ?? "")}>
                      {(a.choices ?? []).map((c) => (
                        <Option key={c} value={c}>
                          {c}
                        </Option>
                      ))}
                    </Dropdown>
                  </Field>
                );
              const isCellArg = a.kind === "input" || a.kind === "range";
              let placeholder = a.placeholder ?? (a.default ? `default: ${a.default}` : "");
              if (isCellArg && canSelection) {
                if (a.kind === "range") placeholder = "the selected cells";
                else if (!(values[a.name] ?? "").trim()) {
                  const k = inputNo++;
                  placeholder = k < target!.columnCount ? `column ${k + 1} of the selection (each row)` : a.optional ? "leave empty" : "type a cell, e.g. $B$1, or a value";
                }
              }
              return (
                <Field key={a.name} label={a.label} hint={isCellArg ? "Leave empty to use the selection, or type a cell (B2, $B$1) or a value" : undefined}>
                  <Input value={values[a.name] ?? ""} placeholder={placeholder} onChange={(_, d) => set(a.name, d.value)} />
                </Field>
              );
            })}
            {canSelection && <Checkbox checked={header} onChange={(_, d) => setHeader(!!d.checked)} label="First selected row is a header" />}
            <Field label="Formula">
              <code className="zee-fn-syntax">{preview}</code>
            </Field>
            {spec.ai && canSelection && spec.mode === "perRow" && (
              <p className="zee-muted">Rows are sent to Gemini in batches of 25 and answers are cached, so recalculating won't call Gemini again. Use Functions → Freeze results when you're done.</p>
            )}
          </DialogContent>
          <DialogActions>
            <Button appearance="secondary" onClick={onClose}>
              Cancel
            </Button>
            <Button appearance={canSelection ? "secondary" : "primary"} disabled={missingCell.length > 0} onClick={() => onApply(spec, values, header, "cell")} title={missingCell.length ? `Fill in: ${missingCell.map((m) => m.label).join(", ")}` : `Insert into ${target?.activeAddress ?? "the active cell"}`}>
              Insert in active cell
            </Button>
            {canSelection && (
              <Button appearance="primary" disabled={missingSel.length > 0} onClick={() => onApply(spec, values, header, "selection")} title={missingSel.length ? `Fill in: ${missingSel.map((m) => m.label).join(", ")}` : undefined}>
                Apply to selection
              </Button>
            )}
          </DialogActions>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
}
