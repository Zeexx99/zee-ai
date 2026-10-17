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
  Menu,
  MenuButton,
  MenuDivider,
  MenuGroupHeader,
  MenuItem,
  MenuList,
  MenuPopover,
  MenuTrigger,
  Option,
  Textarea,
} from "@fluentui/react-components";
import { Add16Regular, Delete16Regular, Play16Regular, Record20Regular, RecordStop16Regular, Share16Regular, TaskListSquareLtr20Regular } from "@fluentui/react-icons";
import { describeStep, describeTarget, recipeFromText, usesDataSheet, type Recipe, type RecipeStep, type RecordState } from "../services/recipes";

export interface RecipesMenuProps {
  recipes: Recipe[];
  recording: RecordState | null;
  disabled?: boolean;
  canSaveLastPrompt: boolean;
  onStart: () => void;
  onStop: () => void;
  onCancel: () => void;
  onRun: (r: Recipe) => void;
  onDelete: (r: Recipe) => void;
  onShare: (r: Recipe) => void;
  onImport: () => void;
  onSaveLastPrompt: () => void;
}

export function RecipesMenu(p: RecipesMenuProps) {
  const rec = p.recording;
  return (
    <Menu>
      <MenuTrigger disableButtonEnhancement>
        <MenuButton
          size="small"
          appearance="subtle"
          icon={rec ? <Record20Regular className="zee-recording" /> : <TaskListSquareLtr20Regular />}
          disabled={p.disabled && !rec}
          title={rec ? `Recording – ${rec.steps.length} step(s)` : "Recipes – record & replay"}
          aria-label="Recipes"
        >
          <span className="zee-quick-label">{rec ? `Rec ${rec.steps.length}` : "Recipes"}</span>
        </MenuButton>
      </MenuTrigger>
      <MenuPopover>
        <MenuList>
          {rec ? (
            <>
              <MenuItem icon={<RecordStop16Regular />} onClick={p.onStop} disabled={!rec.steps.length}>
                Stop & save ({rec.steps.length} step{rec.steps.length === 1 ? "" : "s"})
              </MenuItem>
              <MenuItem onClick={p.onCancel}>Cancel recording</MenuItem>
            </>
          ) : (
            <>
              <MenuItem icon={<Record20Regular className="zee-recording" />} onClick={p.onStart}>
                Record a new recipe
              </MenuItem>
              <MenuItem icon={<Add16Regular />} onClick={p.onSaveLastPrompt} disabled={!p.canSaveLastPrompt}>
                Save last chat / agent request as recipe
              </MenuItem>
              <MenuItem onClick={p.onImport}>Add a recipe someone shared…</MenuItem>
            </>
          )}
          {p.recipes.length > 0 && <MenuDivider />}
          {p.recipes.length > 0 && <MenuGroupHeader>Saved recipes</MenuGroupHeader>}
          {p.recipes.map((r) => (
            <Menu key={r.id}>
              <MenuTrigger disableButtonEnhancement>
                <MenuItem secondaryContent={`${r.steps.filter((s) => s.enabled !== false).length} steps`} disabled={!!rec}>
                  {r.name}
                </MenuItem>
              </MenuTrigger>
              <MenuPopover>
                <MenuList>
                  <MenuItem icon={<Play16Regular />} onClick={() => p.onRun(r)} disabled={p.disabled}>
                    Run…
                  </MenuItem>
                  <MenuItem icon={<Share16Regular />} onClick={() => p.onShare(r)}>
                    Share (copy as text)
                  </MenuItem>
                  <MenuItem icon={<Delete16Regular />} onClick={() => p.onDelete(r)}>
                    Delete
                  </MenuItem>
                </MenuList>
              </MenuPopover>
            </Menu>
          ))}
        </MenuList>
      </MenuPopover>
    </Menu>
  );
}

function StepList({ steps, toggle }: { steps: RecipeStep[]; toggle?: (i: number, on: boolean) => void }) {
  const labels = steps.map(describeStep);
  return (
    <ol className="zee-recipe-steps">
      {steps.map((s, i) => (
        <li key={i} className={s.enabled === false ? "zee-step-off" : ""}>
          {toggle ? <Checkbox checked={s.enabled !== false} onChange={(_, d) => toggle(i, !!d.checked)} label={`${labels[i]} ${describeTarget(s.target, labels)}`} /> : `${labels[i]} ${describeTarget(s.target, labels)}`}
        </li>
      ))}
    </ol>
  );
}

/* ------------------------------------------------------------ save */

export function RecipeSaveDialog({ steps, onClose, onSave }: { steps: RecipeStep[] | null; onClose: () => void; onSave: (name: string, steps: RecipeStep[]) => void }) {
  const [name, setName] = useState("");
  const [list, setList] = useState<RecipeStep[]>([]);
  useEffect(() => {
    if (!steps) return;
    setList(steps);
    const d = new Date();
    setName(`My recipe ${d.toLocaleDateString(undefined, { day: "numeric", month: "short" })}`);
  }, [steps]);
  if (!steps) return null;
  const on = list.filter((s) => s.enabled !== false).length;
  return (
    <Dialog open onOpenChange={(_, d) => !d.open && onClose()}>
      <DialogSurface className="zee-dialog">
        <DialogBody>
          <DialogTitle>Save recipe</DialogTitle>
          <DialogContent className="zee-form">
            <Field label="Name">
              <Input value={name} onChange={(_, d) => setName(d.value)} placeholder="e.g. Monthly ESP report" />
            </Field>
            <Field label="Steps" hint="Untick a step to leave it out.">
              <StepList steps={list} toggle={(i, v) => setList((p) => p.map((s, k) => (k === i ? { ...s, enabled: v } : s)))} />
            </Field>
            <p className="zee-muted">Next time: open the new data's sheet, then Recipes → {name || "this recipe"} → Run. Steps run on the same columns; more or fewer rows are fine. Columns are matched by their header names.</p>
          </DialogContent>
          <DialogActions>
            <Button appearance="primary" disabled={!name.trim() || !on} onClick={() => onSave(name.trim(), list)}>
              Save
            </Button>
            <Button onClick={onClose}>Discard</Button>
          </DialogActions>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
}

/* ------------------------------------------------------------ run */

export function RecipeRunDialog({ recipe, sheets, active, onClose, onRun }: { recipe: Recipe | null; sheets: string[]; active: string; onClose: () => void; onRun: (r: Recipe, dataSheet: string) => void }) {
  const [sheet, setSheet] = useState(active);
  useEffect(() => setSheet(active), [active, recipe]);
  if (!recipe) return null;
  const needsSheet = usesDataSheet(recipe);
  const prompts = recipe.steps.filter((s) => s.enabled !== false && s.kind === "prompt").length;
  return (
    <Dialog open onOpenChange={(_, d) => !d.open && onClose()}>
      <DialogSurface className="zee-dialog">
        <DialogBody>
          <DialogTitle>Run “{recipe.name}”</DialogTitle>
          <DialogContent className="zee-form">
            <StepList steps={recipe.steps} />
            {needsSheet && (
              <Field label="Data sheet" hint="The sheet with this month's data – same layout as when the recipe was recorded.">
                <Dropdown value={sheet} selectedOptions={[sheet]} onOptionSelect={(_, d) => d.optionValue && setSheet(d.optionValue)}>
                  {sheets.map((s) => (
                    <Option key={s} value={s}>
                      {s}
                    </Option>
                  ))}
                </Dropdown>
              </Field>
            )}
            <p className="zee-muted">
              Steps run one after another without the usual confirmations; each change can be undone from the history list (↶).
              {prompts ? ` ${prompts} step${prompts === 1 ? " asks" : "s ask"} Gemini – agent changes still need your approval.` : ""}
            </p>
          </DialogContent>
          <DialogActions>
            <Button appearance="primary" icon={<Play16Regular />} disabled={needsSheet && !sheet} onClick={() => onRun(recipe, sheet)}>
              Run
            </Button>
            <Button onClick={onClose}>Cancel</Button>
          </DialogActions>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
}

/* ------------------------------------------------------------ import shared */

export function RecipeImportDialog({ open, onClose, onAdd }: { open: boolean; onClose: () => void; onAdd: (r: Recipe) => void }) {
  const [text, setText] = useState("");
  const [error, setError] = useState<string | null>(null);
  useEffect(() => {
    if (open) {
      setText("");
      setError(null);
    }
  }, [open]);
  if (!open) return null;
  const add = () => {
    try {
      onAdd(recipeFromText(text));
    } catch (e) {
      setError((e as Error).message);
    }
  };
  return (
    <Dialog open onOpenChange={(_, d) => !d.open && onClose()}>
      <DialogSurface className="zee-dialog">
        <DialogBody>
          <DialogTitle>Add a shared recipe</DialogTitle>
          <DialogContent className="zee-form">
            <Field label="Paste the recipe text (starts with ZEE-RECIPE)">
              <Textarea value={text} onChange={(_, d) => setText(d.value)} rows={6} resize="vertical" />
            </Field>
            {error && <p className="zee-error-text">{error}</p>}
          </DialogContent>
          <DialogActions>
            <Button appearance="primary" disabled={!text.trim()} onClick={add}>
              Add
            </Button>
            <Button onClick={onClose}>Cancel</Button>
          </DialogActions>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
}
