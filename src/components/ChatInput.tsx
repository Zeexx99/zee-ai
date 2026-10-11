import { useState, type KeyboardEvent } from "react";
import { Button, Menu, MenuDivider, MenuGroupHeader, MenuItem, MenuList, MenuPopover, MenuTrigger, Switch, Textarea, Tooltip } from "@fluentui/react-components";
import { Bookmark20Regular, BookmarkAdd20Regular, Delete16Regular, Send20Filled, Stop20Filled } from "@fluentui/react-icons";
import { loadPrompts, savePrompts } from "../services/promptStore";

const short = (p: string) => (p.length > 70 ? `${p.slice(0, 68)}…` : p);

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
  const [prompts, setPrompts] = useState<string[]>(() => loadPrompts());
  const updatePrompts = (list: string[]) => {
    setPrompts(list);
    savePrompts(list);
  };

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
        <Menu positioning="above-end">
          <MenuTrigger disableButtonEnhancement>
            <Tooltip content="Saved prompts" relationship="label">
              <Button appearance="subtle" icon={<Bookmark20Regular />} />
            </Tooltip>
          </MenuTrigger>
          <MenuPopover>
            <MenuList>
              <MenuGroupHeader>Saved prompts – click to use</MenuGroupHeader>
              {prompts.length === 0 && <MenuItem disabled>No saved prompts yet</MenuItem>}
              {prompts.map((p, i) => (
                <MenuItem key={i} title={p} onClick={() => setText(p)}>
                  {short(p)}
                </MenuItem>
              ))}
              <MenuDivider />
              <MenuItem
                icon={<BookmarkAdd20Regular />}
                disabled={!text.trim() || prompts.includes(text.trim())}
                onClick={() => updatePrompts([text.trim(), ...prompts])}
              >
                Save current text as a prompt
              </MenuItem>
              {prompts.length > 0 && (
                <Menu>
                  <MenuTrigger disableButtonEnhancement>
                    <MenuItem icon={<Delete16Regular />}>Remove a prompt</MenuItem>
                  </MenuTrigger>
                  <MenuPopover>
                    <MenuList>
                      {prompts.map((p, i) => (
                        <MenuItem key={i} title={p} onClick={() => updatePrompts(prompts.filter((_, j) => j !== i))}>
                          {short(p)}
                        </MenuItem>
                      ))}
                    </MenuList>
                  </MenuPopover>
                </Menu>
              )}
            </MenuList>
          </MenuPopover>
        </Menu>
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
