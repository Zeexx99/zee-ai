import { Menu, MenuButton, MenuDivider, MenuGroupHeader, MenuItem, MenuList, MenuPopover, MenuTrigger, Toolbar } from "@fluentui/react-components";
import { Broom20Regular, Calculator20Regular, ChartMultiple20Regular, DataTrending20Regular, MathFormula20Regular } from "@fluentui/react-icons";
import { AGGREGATE_LABELS, type AggregateFn } from "../services/excelService";

export interface QuickActionHandlers {
  onAggregate: (fn: AggregateFn) => void;
  onClean: () => void;
  onExtract: () => void;
  onFormatTable: () => void;
  onAutofit: () => void;
  onExplainFormula: () => void;
  onGenerateFormula: () => void;
  onAnalyze: () => void;
  onChart: () => void;
  onDeclineCurve: () => void;
  onProductionQc: () => void;
  onUnits: () => void;
  onEspCheck: () => void;
  onDraw: () => void;
}

const MATH: AggregateFn[] = ["SUM", "AVERAGE", "MIN", "MAX"];
const COUNTS: AggregateFn[] = ["COUNT", "COUNTA", "UNIQUE"];

export function QuickActions(p: QuickActionHandlers & { disabled?: boolean }) {
  return (
    <Toolbar className="zee-quick" size="small" aria-label="Quick actions">
      <Menu>
        <MenuTrigger disableButtonEnhancement>
          <MenuButton size="small" appearance="subtle" icon={<Calculator20Regular />} disabled={p.disabled}>
            Math
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
          <MenuButton size="small" appearance="subtle" icon={<Broom20Regular />} disabled={p.disabled}>
            Clean
          </MenuButton>
        </MenuTrigger>
        <MenuPopover>
          <MenuList>
            <MenuItem onClick={p.onClean}>Clean data…</MenuItem>
            <MenuItem onClick={p.onExtract}>Split / Extract…</MenuItem>
            <MenuDivider />
            <MenuItem onClick={p.onFormatTable}>Format as table</MenuItem>
            <MenuItem onClick={p.onAutofit}>Auto-fit columns</MenuItem>
          </MenuList>
        </MenuPopover>
      </Menu>

      <Menu>
        <MenuTrigger disableButtonEnhancement>
          <MenuButton size="small" appearance="subtle" icon={<MathFormula20Regular />} disabled={p.disabled}>
            Analyze
          </MenuButton>
        </MenuTrigger>
        <MenuPopover>
          <MenuList>
            <MenuItem onClick={p.onExplainFormula}>Explain formula in active cell</MenuItem>
            <MenuItem onClick={p.onGenerateFormula}>Generate formula…</MenuItem>
            <MenuDivider />
            <MenuItem onClick={p.onAnalyze}>Summarize selection</MenuItem>
            <MenuItem onClick={p.onChart}>Chart from selection…</MenuItem>
          </MenuList>
        </MenuPopover>
      </Menu>

      <Menu>
        <MenuTrigger disableButtonEnhancement>
          <MenuButton size="small" appearance="subtle" icon={<DataTrending20Regular />} disabled={p.disabled}>
            Petroleum
          </MenuButton>
        </MenuTrigger>
        <MenuPopover>
          <MenuList>
            <MenuGroupHeader>Select your data first</MenuGroupHeader>
            <MenuItem onClick={p.onDeclineCurve} secondaryContent="Arps">
              Decline curve analysis…
            </MenuItem>
            <MenuItem onClick={p.onProductionQc}>Production data QC</MenuItem>
            <MenuItem onClick={p.onUnits}>Unit conversion…</MenuItem>
            <MenuDivider />
            <MenuItem onClick={p.onEspCheck} secondaryContent="AI">
              ESP performance check
            </MenuItem>
          </MenuList>
        </MenuPopover>
      </Menu>

      <Menu>
        <MenuTrigger disableButtonEnhancement>
          <MenuButton size="small" appearance="subtle" icon={<ChartMultiple20Regular />} disabled={p.disabled}>
            Draw
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
