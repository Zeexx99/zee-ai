import { useRef, useState, type ClipboardEvent, type DragEvent, type KeyboardEvent } from "react";
import { Button, Menu, MenuDivider, MenuGroupHeader, MenuItem, MenuList, MenuPopover, MenuTrigger, Switch, Textarea, Tooltip } from "@fluentui/react-components";
import { Attach20Regular, Bookmark20Regular, BookmarkAdd20Regular, Delete16Regular, Dismiss12Regular, DocumentPdf16Regular, Image16Regular, Send20Filled, Stop20Filled } from "@fluentui/react-icons";
import { loadPrompts, savePrompts } from "../services/promptStore";
import type { Attachment } from "../services/geminiService";
import { ACCEPT_ATTR, mimeOf, pickFiles, readAsBase64 } from "../utils/files";

const short = (p: string) => (p.length > 70 ? `${p.slice(0, 68)}…` : p);

interface Props {
  busy: boolean;
  agentMode: boolean;
  onToggleAgent: (v: boolean) => void;
  includeSelection: boolean;
  selectedAddress: string | null;
  onToggleSelection: (v: boolean) => void;
  onSend: (text: string, attachments?: Attachment[]) => void;
  onStop: () => void;
}

export function ChatInput({ busy, agentMode, onToggleAgent, includeSelection, selectedAddress, onToggleSelection, onSend, onStop }: Props) {
  const [text, setText] = useState("");
  const [files, setFiles] = useState<Attachment[]>([]);
  const [fileError, setFileError] = useState<string | null>(null);
  const [dragging, setDragging] = useState(false);
  const fileInput = useRef<HTMLInputElement>(null);
  const [prompts, setPrompts] = useState<string[]>(() => loadPrompts());
  const updatePrompts = (list: string[]) => {
    setPrompts(list);
    savePrompts(list);
  };

  const addFiles = async (incoming: File[]) => {
    if (!incoming.length) return;
    const { accepted, error } = pickFiles(incoming, files);
    setFileError(error);
    try {
      const read = await Promise.all(
        accepted.map(async (f, i): Promise<Attachment> => {
          const mimeType = mimeOf(f);
          const name = f.name && f.name !== "image.png" ? f.name : `pasted-image-${files.length + i + 1}.${mimeType.split("/")[1] ?? "png"}`;
          return { name, mimeType, size: f.size, data: await readAsBase64(f) };
        })
      );
      setFiles((cur) => [...cur, ...read]);
    } catch (e) {
      setFileError((e as Error).message);
    }
  };

  const onPaste = (e: ClipboardEvent<HTMLTextAreaElement>) => {
    const pasted = Array.from(e.clipboardData?.files ?? []);
    if (pasted.length) {
      e.preventDefault();
      void addFiles(pasted);
    }
  };

  const onDrop = (e: DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    setDragging(false);
    void addFiles(Array.from(e.dataTransfer?.files ?? []));
  };

  const submit = () => {
    const t = text.trim();
    if ((!t && !files.length) || busy) return;
    onSend(t, files.length ? files : undefined);
    setText("");
    setFiles([]);
    setFileError(null);
  };

  const onKeyDown = (e: KeyboardEvent<HTMLTextAreaElement>) => {
    if (e.key === "Enter" && !e.shiftKey && !e.nativeEvent.isComposing) {
      e.preventDefault();
      submit();
    }
  };

  return (
    <div
      className={`zee-input${dragging ? " zee-input-drop" : ""}`}
      onDragOver={(e) => {
        if (Array.from(e.dataTransfer?.types ?? []).includes("Files")) {
          e.preventDefault();
          setDragging(true);
        }
      }}
      onDragLeave={() => setDragging(false)}
      onDrop={onDrop}
    >
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
      {(files.length > 0 || fileError) && (
        <div className="zee-attachments">
          {files.map((f, i) => (
            <span key={`${f.name}-${i}`} className="zee-chip zee-attachment" title={`${f.name} · ${Math.max(1, Math.round(f.size / 1000))} KB`}>
              {f.mimeType === "application/pdf" ? <DocumentPdf16Regular /> : <Image16Regular />}
              <span className="zee-attachment-name">{f.name}</span>
              <Button size="small" appearance="transparent" icon={<Dismiss12Regular />} aria-label={`Remove ${f.name}`} onClick={() => setFiles((cur) => cur.filter((_, j) => j !== i))} />
            </span>
          ))}
          {fileError && <span className="zee-attachment-error">{fileError}</span>}
        </div>
      )}
      <div className="zee-input-row">
        <input
          ref={fileInput}
          type="file"
          accept={ACCEPT_ATTR}
          multiple
          hidden
          onChange={(e) => {
            void addFiles(Array.from(e.target.files ?? []));
            e.target.value = "";
          }}
        />
        <Tooltip content="Attach images or PDFs (or paste / drop them here)" relationship="label">
          <Button appearance="subtle" icon={<Attach20Regular />} onClick={() => fileInput.current?.click()} disabled={busy} />
        </Tooltip>
        <Textarea
          className="zee-textarea"
          value={text}
          onChange={(_, d) => setText(d.value)}
          onKeyDown={onKeyDown}
          onPaste={onPaste}
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
            <Button appearance="primary" icon={<Send20Filled />} onClick={submit} disabled={!text.trim() && !files.length} />
          </Tooltip>
        )}
      </div>
    </div>
  );
}
