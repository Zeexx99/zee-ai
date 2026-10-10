import { useState, type KeyboardEvent } from "react";
import { Button, Switch, Textarea, Tooltip } from "@fluentui/react-components";
import { Send20Filled, Stop20Filled } from "@fluentui/react-icons";

interface Props {
  busy: boolean;
  agentMode: boolean;
  onToggleAgent: (v: boolean) => void;
  includeSelection: boolean;
  selectedAddress: string | null;
  onToggleSelection: (v: boolean) => void;
  onSend: (text: string) => void;
  onStop: () => void;
}

export function ChatInput({ busy, agentMode, onToggleAgent, includeSelection, selectedAddress, onToggleSelection, onSend, onStop }: Props) {
  const [text, setText] = useState("");

  const submit = () => {
    const t = text.trim();
    if (!t || busy) return;
    onSend(t);
    setText("");
  };

  const onKeyDown = (e: KeyboardEvent<HTMLTextAreaElement>) => {
    if (e.key === "Enter" && !e.shiftKey && !e.nativeEvent.isComposing) {
      e.preventDefault();
      submit();
    }
  };

  return (
    <div className="zee-input">
      <div className="zee-input-context">
        <Tooltip content="Agent mode: Zee AI can read and change the workbook (you approve each change)" relationship="description">
          <Switch checked={agentMode} onChange={(_, d) => onToggleAgent(d.checked)} label="Agent" />
        </Tooltip>
        <Switch
          checked={includeSelection}
          onChange={(_, d) => onToggleSelection(d.checked)}
          label={
            <span>
              Include selection{" "}
              {includeSelection && selectedAddress && <span className="zee-chip zee-chip-inline">{selectedAddress}</span>}
            </span>
          }
        />
      </div>
      <div className="zee-input-row">
        <Textarea
          className="zee-textarea"
          value={text}
          onChange={(_, d) => setText(d.value)}
          onKeyDown={onKeyDown}
          placeholder={agentMode ? "Tell Zee AI what to do in the workbook… (Enter to send)" : "Ask Zee AI… (Enter to send, Shift+Enter for a new line)"}
          resize="vertical"
          rows={2}
        />
        {busy ? (
          <Tooltip content="Stop generating" relationship="label">
            <Button appearance="secondary" icon={<Stop20Filled />} onClick={onStop} />
          </Tooltip>
        ) : (
          <Tooltip content="Send" relationship="label">
            <Button appearance="primary" icon={<Send20Filled />} onClick={submit} disabled={!text.trim()} />
          </Tooltip>
        )}
      </div>
    </div>
  );
}
