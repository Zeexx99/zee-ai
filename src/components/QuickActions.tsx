import { Button, Menu, MenuButton, MenuDivider, MenuGroupHeader, MenuItem, MenuList, MenuPopover, MenuTrigger, Toolbar, Tooltip } from "@fluentui/react-components";
import { useCallback, useEffect, useRef, useState, type ReactNode } from "react";
import { Broom20Regular, Calculator20Regular, ChartMultiple20Regular, DataTrending20Regular, MathFormula20Regular, Search20Regular, TableStackRight20Regular } from "@fluentui/react-icons";
import type { Command, CommandGroup } from "./commands";
import { featureNote, type FeatureId } from "../qol/requirements";
import { AGGREGATE_LABELS, type AggregateFn } from "../services/excelService";
import type { FnSpec } from "../functions/catalog";
import { FunctionsMenu } from "./FunctionPicker";

export interface QuickActionHandlers {
  onAggregate: (fn: AggregateFn) => void;
  onClean: () => void;
  onExtract: () => void;
  onFormatTable: () => void;
  onAutofit: () => void;
  onRemoveEmptyRows: () => void;
  onRemoveEmptyColumns: () => void;
  onExplainFormula: () => void;
  onGenerateFormula: () => void;
  onAnalyze: () => void;
  onChart: () => void;
  onDeclineCurve: () => void;
  onProductionQc: () => void;
  onUnits: () => void;
  onEspCheck: () => void;
  onEspAnalysis: () => void;
  onDraw: () => void;
  onPickFunction: (spec: FnSpec) => void;
  onFreeze: () => void;
  onDailyReport: () => void;
  onEspLogTemplate: () => void;
  onEspLogAnalyze: () => void;
  onFormulaAudit: () => void;
  onSuggestRules: () => void;
  onFindDupes: () => void;
  onFillGaps: () => void;
  onFixImport: () => void;
  onExplainSheet: () => void;
  onSheetGuide: () => void;
  onNamedInputs: () => void;
  onImportFile: () => void;
  onMerge: () => void;
  onResample: () => void;
  onNodal: () => void;
}

export interface QuickActionsExtra {
  disabled?: boolean;
  /** Feature availability in this Excel version (greys out items that would fail). */
  can: (id: FeatureId) => boolean;
  /** Rendered at the end of the toolbar (the Recipes menu). */
  recipes?: ReactNode;
  /** Tools pinned as one-click buttons (from Ctrl+K search). */
  pinned?: Command[];
  /** Opens the Ctrl+K tool search. */
  onSearch?: () => void;
}

const GROUP_ICON: Record<CommandGroup, JSX.Element> = {
  Math: <Calculator20Regular />,
  Clean: <Broom20Regular />,
  Analyze: <MathFormula20Regular />,
  Petroleum: <DataTrending20Regular />,
  Draw: <ChartMultiple20Regular />,
  Data: <TableStackRight20Regular />,
  "Zee AI": <Search20Regular />,
};

/** Shows ‹ › arrows (with a fade) when the toolbar is wider than the pane. */
function useScrollArrows() {
  const ref = useRef<HTMLDivElement>(null);
  const [edges, setEdges] = useState({ left: false, right: false });
  const update = useCallback(() => {
    const el = ref.current;
    if (!el) return;
    const left = el.scrollLeft > 2;
    const right = el.scrollLeft + el.clientWidth < el.scrollWidth - 2;
    setEdges((e) => (e.left === left && e.right === right ? e : { left, right }));
  }, []);
  useEffect(() => {
    const el = ref.current;
    if (!el) return;
    update();
    el.addEventListener("scroll", update, { passive: true });
    const ro = typeof ResizeObserver !== "undefined" ? new ResizeObserver(update) : null;
    ro?.observe(el);
    window.addEventListener("resize", update);
    return () => {
      el.removeEventListener("scroll", update);
      ro?.disconnect();
      window.removeEventListener("resize", update);
    };
  }, [update]);
  const scrollBy = (dir: 1 | -1) => ref.current?.scrollBy({ left: dir * Math.max(120, (ref.current?.clientWidth ?? 200) * 0.7), behavior: "smooth" });
  return { ref, edges, scrollBy, update };
}

const MATH: AggregateFn[] = ["SUM", "AVERAGE", "MIN", "MAX"];
const COUNTS: AggregateFn[] = ["COUNT", "COUNTA", "UNIQUE"];

export function QuickActions(p: QuickActionHandlers & QuickActionsExtra) {
  /** Disabled + "needs newer Excel" note when this Excel lacks the API set. */
  const gate = (id: FeatureId, note?: string) => (p.can(id) ? { secondaryContent: note } : { disabled: true, secondaryContent: featureNote(id).replace("Needs a newer Excel", "newer Excel") });
  const strip = useScrollArrows();
  const pinned = p.pinned ?? [];
  const pinKey = pinned.map((c) => c.id).join(",");
  const { update } = strip;
  useEffect(() => update(), [pinKey, update]);
  return (
    <div className="zee-quick-wrap">
      {strip.edges.left && (
        <button type="button" className="zee-scroll-arrow zee-scroll-left" aria-label="Scroll tools left" onClick={() => strip.scrollBy(-1)}>
          ‹
        </button>
      )}
    <Toolbar className="zee-quick" size="small" aria-label="Quick actions" ref={strip.ref}>
      {p.onSearch && (
        <Tooltip content="Search all tools (Ctrl+K)" relationship="label">
          <Button size="small" appearance="subtle" icon={<Search20Regular />} onClick={p.onSearch} aria-label="Search tools" />
        </Tooltip>
      )}
      {pinned.map((c) => (
        <Tooltip key={c.id} content={c.label} relationship="label">
          <Button size="small" appearance="subtle" icon={GROUP_ICON[c.group]} disabled={p.disabled || (!!c.feature && !p.can(c.feature))} onClick={c.run}>
            <span className="zee-quick-label">{c.short ?? c.label.replace(/…$/, "")}</span>
          </Button>
        </Tooltip>
      ))}
      {pinned.length > 0 && <span className="zee-pin-sep" aria-hidden />}
      <FunctionsMenu disabled={p.disabled || !p.can("functions")} onPick={p.onPickFunction} onFreeze={p.onFreeze} />

      <Menu>
        <MenuTrigger disableButtonEnhancement>
          <MenuButton size="small" appearance="subtle" icon={<Calculator20Regular />} disabled={p.disabled} title="Math" aria-label="Math">
            <span className="zee-quick-label">Math</span>
          </MenuButton>
        </MenuTrigger>
        <MenuPopover>
          <MenuList>
            <MenuGroupHeader>Insert below selection</MenuGroupHeader>
            {MATH.map((fn) => (
              <MenuItem key={fn} onClick={() => p.onAggregate(fn)} secondaryContent={`=${fn === "AVERAGE" ? "AVG" : fn}`}>
                {AGGREGATE_LABELS[fn]}
              </MenuItem>
            ))}
            <MenuDivider />
            {COUNTS.map((fn) => (
              <MenuItem key={fn} onClick={() => p.onAggregate(fn)}>
                {AGGREGATE_LABELS[fn]}
              </MenuItem>
            ))}
          </MenuList>
        </MenuPopover>
      </Menu>

      <Menu>
        <MenuTrigger disableButtonEnhancement>
          <MenuButton size="small" appearance="subtle" icon={<Broom20Regular />} disabled={p.disabled} title="Clean" aria-label="Clean">
            <span className="zee-quick-label">Clean</span>
          </MenuButton>
        </MenuTrigger>
        <MenuPopover>
          <MenuList>
            <MenuItem onClick={p.onClean}>Clean data…</MenuItem>
            <MenuItem onClick={p.onExtract}>Split / Extract…</MenuItem>
            <MenuDivider />
            <MenuGroupHeader>Selected table</MenuGroupHeader>
            <MenuItem onClick={p.onFindDupes}>Find duplicates & outliers…</MenuItem>
            <MenuItem onClick={p.onFillGaps}>Fill gaps in time series…</MenuItem>
            <MenuItem onClick={p.onFixImport}>Fix imported data…</MenuItem>
            <MenuDivider />
            <MenuGroupHeader>Blank, 0 or NULL everywhere</MenuGroupHeader>
            <MenuItem onClick={p.onRemoveEmptyColumns}>Remove empty columns</MenuItem>
            <MenuItem onClick={p.onRemoveEmptyRows}>Remove empty rows</MenuItem>
            <MenuDivider />
            <MenuItem onClick={p.onFormatTable}>Format as table</MenuItem>
            <MenuItem onClick={p.onAutofit}>Auto-fit all columns</MenuItem>
          </MenuList>
        </MenuPopover>
      </Menu>

      <Menu>
        <MenuTrigger disableButtonEnhancement>
          <MenuButton size="small" appearance="subtle" icon={<MathFormula20Regular />} disabled={p.disabled} title="Analyze" aria-label="Analyze">
            <span className="zee-quick-label">Analyze</span>
          </MenuButton>
        </MenuTrigger>
        <MenuPopover>
          <MenuList>
            <MenuItem onClick={p.onExplainFormula}>Explain formula in active cell</MenuItem>
            <MenuItem onClick={p.onGenerateFormula}>Generate formula…</MenuItem>
            <MenuDivider />
            <MenuItem onClick={p.onAnalyze}>Summarize selection</MenuItem>
            <MenuItem onClick={p.onChart} {...gate("charts")}>
              Chart from selection…
            </MenuItem>
            <MenuDivider />
            <MenuGroupHeader>Checks</MenuGroupHeader>
            <MenuItem onClick={p.onFormulaAudit} secondaryContent="this sheet">
              Formula audit
            </MenuItem>
            <MenuItem onClick={p.onSuggestRules} {...gate("validation", "AI")}>
              Suggest data validation…
            </MenuItem>
            <MenuDivider />
            <MenuGroupHeader>This sheet</MenuGroupHeader>
            <MenuItem onClick={p.onExplainSheet} secondaryContent="AI">
              Explain this sheet
            </MenuItem>
            <MenuItem onClick={p.onSheetGuide}>Add a sheet guide tab</MenuItem>
            <MenuItem onClick={p.onNamedInputs}>Move hard-coded numbers to inputs…</MenuItem>
          </MenuList>
        </MenuPopover>
      </Menu>

      <Menu>
        <MenuTrigger disableButtonEnhancement>
          <MenuButton size="small" appearance="subtle" icon={<DataTrending20Regular />} disabled={p.disabled} title="Petroleum" aria-label="Petroleum">
            <span className="zee-quick-label">Petroleum</span>
          </MenuButton>
        </MenuTrigger>
        <MenuPopover>
          <MenuList>
            <MenuGroupHeader>Select your data first</MenuGroupHeader>
            <MenuItem onClick={p.onDailyReport} secondaryContent="→ Word">
              Daily production report…
            </MenuItem>
            <MenuItem onClick={p.onDeclineCurve} {...gate("charts", "Arps")}>
              Decline curve analysis…
            </MenuItem>
            <MenuItem onClick={p.onProductionQc}>Production data QC</MenuItem>
            <MenuItem onClick={p.onUnits}>Unit conversion…</MenuItem>
            <MenuDivider />
            <MenuItem onClick={p.onEspAnalysis} {...gate("charts", "calc")}>
              ESP analysis…
            </MenuItem>
            <MenuItem onClick={p.onEspCheck} secondaryContent="AI">
              ESP performance check
            </MenuItem>
            <MenuDivider />
            <MenuGroupHeader>ESP failure log</MenuGroupHeader>
            <MenuItem onClick={p.onEspLogTemplate}>Create a failure log sheet</MenuItem>
            <MenuItem onClick={p.onEspLogAnalyze} {...gate("charts", "MTBF")}>
              Analyze failure log…
            </MenuItem>
            <MenuDivider />
            <MenuGroupHeader>Well performance</MenuGroupHeader>
            <MenuItem onClick={p.onNodal} {...gate("charts", "IPR × VLP")}>
              Nodal analysis…
            </MenuItem>
          </MenuList>
        </MenuPopover>
      </Menu>

      <Menu>
        <MenuTrigger disableButtonEnhancement>
          <MenuButton size="small" appearance="subtle" icon={<ChartMultiple20Regular />} disabled={p.disabled} title="Draw" aria-label="Draw">
            <span className="zee-quick-label">Draw</span>
          </MenuButton>
        </MenuTrigger>
        <MenuPopover>
          <MenuList>
            <MenuGroupHeader>Charts from your columns</MenuGroupHeader>
            <MenuItem onClick={p.onDraw} {...gate("charts")}>
              Plot columns (X vs several Y)…
            </MenuItem>
          </MenuList>
        </MenuPopover>
      </Menu>

      <Menu>
        <MenuTrigger disableButtonEnhancement>
          <MenuButton size="small" appearance="subtle" icon={<TableStackRight20Regular />} disabled={p.disabled} title="Data" aria-label="Data">
            <span className="zee-quick-label">Data</span>
          </MenuButton>
        </MenuTrigger>
        <MenuPopover>
          <MenuList>
            <MenuItem onClick={p.onImportFile} secondaryContent="CSV · LAS · SCADA">
              Import file…
            </MenuItem>
            <MenuItem onClick={p.onMerge} secondaryContent="by well & date">
              Merge well data…
            </MenuItem>
            <MenuDivider />
            <MenuGroupHeader>Selected time series</MenuGroupHeader>
            <MenuItem onClick={p.onResample} secondaryContent="day → month…">
              Resample…
            </MenuItem>
            <MenuItem onClick={p.onFillGaps}>Fill gaps…</MenuItem>
            <MenuItem onClick={p.onFixImport}>Fix imported data…</MenuItem>
          </MenuList>
        </MenuPopover>
      </Menu>

      {p.recipes}
    </Toolbar>
      {strip.edges.right && (
        <button type="button" className="zee-scroll-arrow zee-scroll-right" aria-label="Scroll tools right" onClick={() => strip.scrollBy(1)}>
          ›
        </button>
      )}
    </div>
  );
}
