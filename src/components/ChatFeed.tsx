import { useEffect, useMemo, useRef, useState } from "react";
import { Button } from "@fluentui/react-components";
import { ArrowDown16Regular, DataTrending16Regular, Gauge16Regular, History16Regular, Search16Regular, TableSearchRegular, Wrench16Regular } from "@fluentui/react-icons";
import type { ChatMessage } from "../hooks/useChat";
import { MessageBubble } from "./MessageBubble";
import { loadHistory } from "../qol/promptHistory";

/** Starter prompts for daily petroleum work (shown on an empty chat). */
export const WELCOME_CARDS: Array<{ title: string; text: string; prompt: string; icon: JSX.Element }> = [
  {
    title: "Daily production",
    text: "Summarize the selected rates and flag drops",
    prompt: "Summarize the selected daily production data: totals per well, the biggest changes versus the previous day, and any wells with a drop of more than 10%.",
    icon: <DataTrending16Regular />,
  },
  {
    title: "ESP readings",
    text: "Check frequency, current and intake pressure",
    prompt: "Check the selected ESP readings (frequency, motor current, intake pressure, temperature). Point out abnormal values or trends that suggest gas lock, underload or wear.",
    icon: <Gauge16Regular />,
  },
  {
    title: "Well test QC",
    text: "Find gaps, outliers and impossible values",
    prompt: "Run a quality check on the selected well test data: missing dates, duplicates, outliers, negative or impossible values (e.g. water cut above 100%), and list them by row.",
    icon: <TableSearchRegular fontSize={16} />,
  },
  {
    title: "Build a formula",
    text: "Monthly totals per well from daily data",
    prompt: "Write a formula that totals the daily oil rate per well per month from the selected table, and tell me where to put it.",
    icon: <Wrench16Regular />,
  },
];

interface Props {
  messages: ChatMessage[];
  onSuggestion: (t: string) => void;
  /** Put text into the chat box without sending (recent prompts). */
  onPrefill?: (t: string) => void;
  /** Opens Ctrl+K tool search. */
  onSearchTools?: () => void;
}

export function ChatFeed({ messages, onSuggestion, onPrefill, onSearchTools }: Props) {
  const ref = useRef<HTMLDivElement>(null);
  const stick = useRef(true);
  const [showJump, setShowJump] = useState(false);
  // Read once per empty screen: the last prompts sent from this PC.
  const recent = useMemo(() => (messages.length === 0 ? loadHistory().slice(0, 2) : []), [messages.length === 0]); // eslint-disable-line react-hooks/exhaustive-deps

  useEffect(() => {
    const el = ref.current;
    if (el && stick.current) el.scrollTop = el.scrollHeight;
  }, [messages]);

  const jump = () => {
    const el = ref.current;
    if (!el) return;
    stick.current = true;
    setShowJump(false);
    el.scrollTo({ top: el.scrollHeight, behavior: "smooth" });
  };

  return (
    <div className="zee-feed-wrap">
      <div
        className="zee-feed"
        ref={ref}
        role="log"
        aria-live="polite"
        onScroll={(e) => {
          const el = e.currentTarget;
          const atBottom = el.scrollHeight - el.scrollTop - el.clientHeight < 60;
          stick.current = atBottom;
          setShowJump((v) => (v === !atBottom ? v : !atBottom));
        }}
      >
        {messages.length === 0 ? (
          <div className="zee-empty">
            <img src="assets/icon-64.png" alt="" width={44} height={44} />
            <h2>What can I do for your sheet?</h2>
            <p>Select your data, then pick a starter or just ask.</p>
            <div className="zee-welcome-cards">
              {WELCOME_CARDS.map((c) => (
                <button key={c.title} type="button" className="zee-welcome-card" onClick={() => onSuggestion(c.prompt)} title={c.prompt}>
                  <strong>
                    {c.icon} {c.title}
                  </strong>
                  <span>{c.text}</span>
                </button>
              ))}
            </div>
            {recent.length > 0 && (
              <>
                <div className="zee-welcome-h">Recent prompts</div>
                <div className="zee-welcome-cards">
                  {recent.map((r, i) => (
                    <button
                      key={i}
                      type="button"
                      className="zee-welcome-card zee-recent-card"
                      title={`${r}\n\nClick to put it in the chat box`}
                      onClick={() => (onPrefill ? onPrefill(r) : onSuggestion(r))}
                    >
                      <strong>
                        <History16Regular /> Use again
                      </strong>
                      <span>{r}</span>
                    </button>
                  ))}
                </div>
              </>
            )}
            {onSearchTools && (
              <p className="zee-muted zee-welcome-tip">
                <Button size="small" appearance="transparent" icon={<Search16Regular />} onClick={onSearchTools}>
                  Find a tool
                </Button>
                <span className="zee-kbd">Ctrl</span> + <span className="zee-kbd">K</span>
              </p>
            )}
          </div>
        ) : (
          messages.map((m) => <MessageBubble key={m.id} message={m} />)
        )}
      </div>
      {showJump && messages.length > 0 && (
        <Button className="zee-jump" size="small" shape="circular" appearance="primary" icon={<ArrowDown16Regular />} onClick={jump} aria-label="Jump to latest message">
          Latest
        </Button>
      )}
    </div>
  );
}
