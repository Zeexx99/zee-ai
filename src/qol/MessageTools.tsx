/** Per-message buttons shared by both add-ins: copy / edit & resend, regenerate, and the friendly error bubble. */
import { useState } from "react";
import { Button, Tooltip } from "@fluentui/react-components";
import { ArrowClockwise16Regular, Bug16Regular, Checkmark16Regular, Copy16Regular, Edit16Regular, ErrorCircle16Regular } from "@fluentui/react-icons";
import { copyText, useInsertActions } from "../components/insertActions";
import type { ChatMessage } from "../hooks/useChat";
import { errorHelp } from "./diagnostics";

function SmallCopy({ text, label }: { text: string; label: string }) {
  const [done, setDone] = useState(false);
  return (
    <Tooltip content={done ? "Copied" : label} relationship="label">
      <Button
        size="small"
        appearance="subtle"
        icon={done ? <Checkmark16Regular /> : <Copy16Regular />}
        onClick={async () => {
          if (await copyText(text)) {
            setDone(true);
            window.setTimeout(() => setDone(false), 1500);
          }
        }}
      />
    </Tooltip>
  );
}

/** Copy + "Edit & resend" under a user message (shown on hover / focus). */
export function UserMessageTools({ text }: { text: string }) {
  const actions = useInsertActions();
  return (
    <div className="zee-msg-tools">
      <SmallCopy text={text} label="Copy prompt" />
      <Tooltip content="Edit & resend – puts this prompt back in the chat box" relationship="label">
        <Button size="small" appearance="subtle" icon={<Edit16Regular />} onClick={() => actions.editPrompt(text)} />
      </Tooltip>
    </div>
  );
}

/** "Regenerate" for a finished answer. */
export function RegenerateButton({ message }: { message: ChatMessage }) {
  const actions = useInsertActions();
  return (
    <Tooltip content="Regenerate – ask again with the same prompt" relationship="label">
      <Button size="small" appearance="subtle" icon={<ArrowClockwise16Regular />} disabled={actions.busy} onClick={() => actions.retry(message.id)} />
    </Tooltip>
  );
}

/** Error bubble with what to do next, Retry and Copy diagnostics. */
export function ErrorBubble({ message }: { message: ChatMessage }) {
  const actions = useInsertActions();
  const help = errorHelp(message.text, message.errorKind);
  return (
    <div className="zee-msg zee-msg-assistant">
      <div className="zee-bubble zee-bubble-error">
        <div className="zee-error-line">
          <ErrorCircle16Regular /> <span>{message.text}</span>
        </div>
        {help.hint && <div className="zee-error-hint">{help.hint}</div>}
        <div className="zee-row zee-error-actions">
          <Button size="small" appearance={help.retryable ? "primary" : "secondary"} icon={<ArrowClockwise16Regular />} disabled={actions.busy} onClick={() => actions.retry(message.id)}>
            Retry
          </Button>
          <Tooltip content="Copies version, Office and error details (no keys, no document text) – paste them to Claude or IT" relationship="description">
            <Button size="small" appearance="subtle" icon={<Bug16Regular />} onClick={() => actions.copyDiagnostics(message)}>
              Copy diagnostics
            </Button>
          </Tooltip>
        </div>
      </div>
    </div>
  );
}
