import { useEffect, useRef } from "react";
import { Button } from "@fluentui/react-components";
import type { ChatMessage } from "../hooks/useChat";
import { MessageBubble } from "./MessageBubble";

const SUGGESTIONS = [
  "Summarize the selected data and point out anything unusual",
  "Write a formula to look up a value by ID from another column",
  "Create a sample sales table with 10 rows I can insert",
  "Which chart would best show this data, and why?",
];

export function ChatFeed({ messages, onSuggestion }: { messages: ChatMessage[]; onSuggestion: (t: string) => void }) {
  const ref = useRef<HTMLDivElement>(null);
  const stick = useRef(true);

  useEffect(() => {
    const el = ref.current;
    if (el && stick.current) el.scrollTop = el.scrollHeight;
  }, [messages]);

  return (
    <div
      className="zee-feed"
      ref={ref}
      onScroll={(e) => {
        const el = e.currentTarget;
        stick.current = el.scrollHeight - el.scrollTop - el.clientHeight < 60;
      }}
    >
      {messages.length === 0 ? (
        <div className="zee-empty">
          <img src="assets/icon-64.png" alt="" width={48} height={48} />
          <h2>What can I do for your sheet?</h2>
          <p>Select some cells, then ask. Formulas and tables I write come with an Insert button.</p>
          <div className="zee-suggestions">
            {SUGGESTIONS.map((s) => (
              <Button key={s} appearance="outline" size="small" onClick={() => onSuggestion(s)}>
                {s}
              </Button>
            ))}
          </div>
        </div>
      ) : (
        messages.map((m) => <MessageBubble key={m.id} message={m} />)
      )}
    </div>
  );
}
