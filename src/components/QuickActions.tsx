import { Menu, MenuButton, MenuDivider, MenuGroupHeader, MenuItem, MenuList, MenuPopover, MenuTrigger, Toolbar } from "@fluentui/react-components";
import { Broom20Regular, Calculator20Regular, ChartMultiple20Regular, DataTrending20Regular, MathFormula20Regular } from "@fluentui/react-icons";
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
}

const MATH: AggregateFn[] = ["SUM", "AVERAGE", "MIN", "MAX"];
const COUNTS: AggregateFn[] = ["COUNT", "COUNTA", "UNIQUE"];

export function QuickActions(p: QuickActionHandlers & { disabled?: boolean }) {
  return (
    <Toolbar className="zee-quick" size="small" aria-label="Quick actions">
      <FunctionsMenu disabled={p.disabled} onPick={p.onPickFunction} onFreeze={p.onFreeze} />

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
            <MenuItem onClick={p.onChart}>Chart from selection…</MenuItem>
            <MenuDivider />
            <MenuGroupHeader>Checks</MenuGroupHeader>
            <MenuItem onClick={p.onFormulaAudit} secondaryContent="this sheet">
              Formula audit
            </MenuItem>
            <MenuItem onClick={p.onSuggestRules} secondaryContent="AI">
              Suggest data validation…
            </MenuItem>
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
            <MenuItem onClick={p.onDeclineCurve} secondaryContent="Arps">
              Decline curve analysis…
            </MenuItem>
            <MenuItem onClick={p.onProductionQc}>Production data QC</MenuItem>
            <MenuItem onClick={p.onUnits}>Unit conversion…</MenuItem>
            <MenuDivider />
            <MenuItem onClick={p.onEspAnalysis} secondaryContent="calc">
              ESP analysis…
            </MenuItem>
            <MenuItem onClick={p.onEspCheck} secondaryContent="AI">
              ESP performance check
            </MenuItem>
            <MenuDivider />
            <MenuGroupHeader>ESP failure log</MenuGroupHeader>
            <MenuItem onClick={p.onEspLogTemplate}>Create a failure log sheet</MenuItem>
            <MenuItem onClick={p.onEspLogAnalyze} secondaryContent="MTBF">
              Analyze failure log…
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
            <MenuItem onClick={p.onDraw}>Plot columns (X vs several Y)…</MenuItem>
          </MenuList>
        </MenuPopover>
      </Menu>
    </Toolbar>
  );
}
