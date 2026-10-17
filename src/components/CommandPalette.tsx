import { useEffect, useMemo, useRef, useState, type KeyboardEvent } from "react";
import { Button, Dialog, DialogSurface, Input, Tooltip } from "@fluentui/react-components";
import { Search20Regular, Star16Filled, Star16Regular } from "@fluentui/react-icons";
import { searchCommands, type Command } from "./commands";
import type { FeatureId } from "../qol/requirements";

interface Props {
  open: boolean;
  commands: Command[];
  pinned: string[];
  maxPins: number;
  can: (id: FeatureId) => boolean;
  onTogglePin: (id: string) => void;
  onClose: () => void;
}

/** Ctrl+K: type to find any tool, Enter runs it, ☆ pins it to the toolbar. */
export function CommandPalette({ open, commands, pinned, maxPins, can, onTogglePin, onClose }: Props) {
  const [q, setQ] = useState("");
  const [active, setActive] = useState(0);
  const listRef = useRef<HTMLDivElement>(null);
  const results = useMemo(() => searchCommands(commands, q), [commands, q]);

  useEffect(() => {
    if (open) {
      setQ("");
      setActive(0);
    }
  }, [open]);
  useEffect(() => setActive(0), [q]);
  useEffect(() => {
    listRef.current?.querySelector<HTMLElement>(`[data-idx="${active}"]`)?.scrollIntoView({ block: "nearest" });
  }, [active]);

  const enabled = (c: Command) => !c.feature || can(c.feature);
  const run = (c: Command | undefined) => {
    if (!c || !enabled(c)) return;
    onClose();
    // Let the dialog close first so the tool's own dialog can take focus.
    window.setTimeout(c.run, 0);
  };

  const onKey = (e: KeyboardEvent<HTMLInputElement>) => {
    if (e.key === "ArrowDown") {
      e.preventDefault();
      setActive((a) => Math.min(results.length - 1, a + 1));
    } else if (e.key === "ArrowUp") {
      e.preventDefault();
      setActive((a) => Math.max(0, a - 1));
    } else if (e.key === "Enter") {
      e.preventDefault();
      run(results[active]);
    }
  };

  let lastGroup = "";
  return (
    <Dialog open={open} onOpenChange={(_, d) => !d.open && onClose()}>
      <DialogSurface className="zee-cmd-surface" aria-label="Search tools">
        <Input
          autoFocus
          value={q}
          onChange={(_, d) => setQ(d.value)}
          onKeyDown={onKey}
          contentBefore={<Search20Regular />}
          placeholder="Search tools – e.g. nodal, resample, ESP"
          style={{ width: "100%" }}
          aria-controls="zee-cmd-list"
          aria-activedescendant={results[active] ? `zee-cmd-${results[active].id}` : undefined}
        />
        <div className="zee-cmd-list" id="zee-cmd-list" role="listbox" ref={listRef}>
          {results.length === 0 && <p className="zee-muted" style={{ padding: 8 }}>No tool matches “{q}”. Try a shorter word, or ask Zee AI in the chat.</p>}
          {results.map((c, i) => {
            const header = !q && c.group !== lastGroup ? c.group : null;
            lastGroup = c.group;
            const isPinned = pinned.includes(c.id);
            const ok = enabled(c);
            return (
              <div key={c.id}>
                {header && <div className="zee-cmd-group">{header}</div>}
                <div
                  id={`zee-cmd-${c.id}`}
                  data-idx={i}
                  role="option"
                  aria-selected={i === active}
                  aria-disabled={!ok}
                  className="zee-cmd-item"
                  onMouseMove={() => setActive(i)}
                  onClick={() => run(c)}
                >
                  <span className="zee-cmd-label">{c.label}</span>
                  {!ok ? <span className="zee-cmd-hint">newer Excel</span> : q ? <span className="zee-cmd-hint">{c.group}</span> : c.hint && <span className="zee-cmd-hint">{c.hint}</span>}
                  <Tooltip content={isPinned ? "Unpin from toolbar" : `Pin to toolbar (max ${maxPins})`} relationship="label">
                    <Button
                      size="small"
                      appearance="transparent"
                      className={`zee-cmd-star${isPinned ? " zee-cmd-star-on" : ""}`}
                      icon={isPinned ? <Star16Filled /> : <Star16Regular />}
                      onClick={(e) => {
                        e.stopPropagation();
                        onTogglePin(c.id);
                      }}
                    />
                  </Tooltip>
                </div>
              </div>
            );
          })}
        </div>
        <p className="zee-muted" style={{ marginTop: 8, fontSize: 11 }}>
          <span className="zee-kbd">↑</span> <span className="zee-kbd">↓</span> to move · <span className="zee-kbd">Enter</span> to run · ☆ pins a tool to the toolbar · <span className="zee-kbd">Esc</span> to close
        </p>
      </DialogSurface>
    </Dialog>
  );
}
