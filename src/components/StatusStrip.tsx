import { useState, type ReactNode } from "react";
import { MessageBar, MessageBarActions, MessageBarBody } from "@fluentui/react-components";

export interface StatusItem {
  key: string;
  intent: "info" | "warning" | "error" | "success";
  body: ReactNode;
  actions?: ReactNode;
}

const RANK: Record<StatusItem["intent"], number> = { error: 0, warning: 1, info: 2, success: 3 };

/** Most important first (errors, then warnings…); ties keep their order. Pure – unit-tested. */
export function orderStatus<T extends Pick<StatusItem, "intent">>(items: T[]): T[] {
  return items.map((it, i) => ({ it, i })).sort((a, b) => RANK[a.it.intent] - RANK[b.it.intent] || a.i - b.i).map((x) => x.it);
}

/** Passive notices (offline, preview, setup, old Excel) in one strip: the most important shows, the rest fold under "+N more". */
export function StatusStrip({ items }: { items: StatusItem[] }) {
  const [expanded, setExpanded] = useState(false);
  if (!items.length) return null;
  const sorted = orderStatus(items);
  const shown = expanded ? sorted : sorted.slice(0, 1);
  const hidden = sorted.length - shown.length;
  return (
    <div className="zee-status" role="status">
      {shown.map((s) => (
        <MessageBar key={s.key} intent={s.intent} className="zee-banner">
          <MessageBarBody>{s.body}</MessageBarBody>
          {s.actions && <MessageBarActions>{s.actions}</MessageBarActions>}
        </MessageBar>
      ))}
      {sorted.length > 1 && (
        <button type="button" className="zee-status-more" onClick={() => setExpanded((v) => !v)} aria-expanded={expanded}>
          {expanded ? "Show less" : `+${hidden} more notice${hidden === 1 ? "" : "s"}`}
        </button>
      )}
    </div>
  );
}
