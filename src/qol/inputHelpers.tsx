/** Chat-box helpers shared by both add-ins: ↑/↓ prompt recall, draft persistence, favourite chips, usage meter. */
import { useEffect, useRef, useState, type KeyboardEvent } from "react";
import { Button, Tooltip } from "@fluentui/react-components";
import { Dismiss12Regular, Star16Filled } from "@fluentui/react-icons";
import { loadDraft, loadHistory, pushHistory, recallStep, saveDraft } from "./promptHistory";
import { estimateCost, formatCost, formatTokens, type DayUsage, type Prices } from "./usageLog";

/**
 * Chat text with draft persistence and ↑/↓ recall of earlier prompts.
 * ↑ works when the box is empty or the caret is on the first line; ↓ when it is on the last line.
 */
export function usePromptBox() {
  const [text, setTextState] = useState<string>(() => loadDraft());
  const recall = useRef({ index: -1, draft: "" });
  const history = useRef<string[]>(loadHistory());

  // Save the unsent text (debounced) so closing the pane doesn't lose it.
  useEffect(() => {
    const t = window.setTimeout(() => saveDraft(text), 300);
    return () => window.clearTimeout(t);
  }, [text]);

  const setText = (v: string) => {
    recall.current.index = -1;
    setTextState(v);
  };

  /** Call from the textarea's onKeyDown first; returns true when the key was used for recall. */
  const onRecallKey = (e: KeyboardEvent<HTMLTextAreaElement>): boolean => {
    if (e.shiftKey || e.altKey || e.ctrlKey || e.metaKey || e.nativeEvent.isComposing) return false;
    if (e.key !== "ArrowUp" && e.key !== "ArrowDown") return false;
    const el = e.currentTarget;
    const before = el.value.slice(0, el.selectionStart ?? 0);
    const after = el.value.slice(el.selectionEnd ?? 0);
    const r = recall.current;
    if (e.key === "ArrowUp") {
      if (before.includes("\n") || !history.current.length) return false;
      if (r.index === -1 && el.value.trim() && el.selectionStart !== 0) return false;
      if (r.index === -1) r.draft = el.value;
      const s = recallStep(history.current, r.index, "older", r.draft);
      if (s.index === r.index) return false;
      e.preventDefault();
      r.index = s.index;
      setTextState(s.text);
      return true;
    }
    if (r.index === -1 || after.includes("\n")) return false;
    e.preventDefault();
    const s = recallStep(history.current, r.index, "newer", r.draft);
    r.index = s.index;
    setTextState(s.text);
    return true;
  };

  /** Call after sending: remembers the prompt and clears the box + draft. */
  const onSent = (sent: string) => {
    history.current = pushHistory(sent);
    recall.current = { index: -1, draft: "" };
    setTextState("");
    saveDraft("");
  };

  return { text, setText, onRecallKey, onSent };
}

/** Favourite (saved) prompts as one-click chips above the chat box. */
export function PromptChips({ prompts, onPick, max = 6 }: { prompts: string[]; onPick: (p: string) => void; max?: number }) {
  if (!prompts.length) return null;
  return (
    <div className="zee-prompt-chips" role="list" aria-label="Favourite prompts">
      {prompts.slice(0, max).map((p, i) => (
        <button key={i} type="button" role="listitem" className="zee-prompt-chip" title={p} onClick={() => onPick(p)}>
          <Star16Filled className="zee-prompt-chip-star" />
          <span>{p.length > 38 ? `${p.slice(0, 36)}…` : p}</span>
        </button>
      ))}
    </div>
  );
}

/** "Sheet1!A1:F200 ×" / "42 words selected ×" – what the next message will include. */
export function ContextChip({ label, ignored, onToggle, title }: { label: string; ignored: boolean; onToggle: (ignore: boolean) => void; title?: string }) {
  return (
    <span className={`zee-chip zee-chip-inline zee-context-chip${ignored ? " zee-chip-ignored" : ""}`} title={title}>
      <span className="zee-context-chip-label">{ignored ? `${label} – not sent` : label}</span>
      <Tooltip content={ignored ? "Include it again" : "Leave it out of the next message"} relationship="label">
        <Button
          size="small"
          appearance="transparent"
          className="zee-chip-x"
          icon={ignored ? <span aria-hidden>↺</span> : <Dismiss12Regular />}
          aria-label={ignored ? "Include again" : "Leave out"}
          onClick={() => onToggle(!ignored)}
        />
      </Tooltip>
    </span>
  );
}

/** "Today 45.2k tokens · ~$0.03" under the chat box. */
export function UsageMeter({ today, prices }: { today: DayUsage; prices: Prices }) {
  const total = today.promptTokens + today.outputTokens;
  if (!total) return null;
  const cost = estimateCost(today, prices);
  return (
    <span className="zee-muted zee-usage-meter" title={`Today: ${today.requests} requests · ${today.promptTokens.toLocaleString()} tokens in · ${today.outputTokens.toLocaleString()} out (details in Settings → Usage)`}>
      Today {formatTokens(total)} tokens{cost !== null ? ` · ~${formatCost(cost)}` : ""}
    </span>
  );
}
