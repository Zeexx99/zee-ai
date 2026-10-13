import { memo, useMemo, useState, type ReactNode } from "react";
import ReactMarkdown, { type Components, type Options } from "react-markdown";
import remarkGfm from "remark-gfm";
import rehypeHighlight from "rehype-highlight";
import { common } from "lowlight";
import excel from "highlight.js/lib/languages/excel";
import { Button, Menu, MenuItem, MenuList, MenuPopover, MenuTrigger, Spinner, SplitButton, Tooltip, type MenuButtonProps } from "@fluentui/react-components";
import { Add16Regular, ArrowRight16Regular, Attach16Regular, Checkmark16Regular, Copy16Regular, Dismiss16Regular, DocumentPdf16Regular, ErrorCircle16Regular, Eye16Regular, Globe16Regular, Image16Regular, Prohibited16Regular, Table16Regular, Warning16Regular, Wrench16Regular } from "@fluentui/react-icons";
import { ErrorBubble, RegenerateButton, UserMessageTools } from "../qol/MessageTools";
import type { ActionLog, ChatMessage, PendingApproval } from "../hooks/useChat";
import { formulasFromCode, hastTableToMatrix, hastText, parseDelimited, type HastLike } from "../utils/markdown";
import { copyText, openExternal, useInsertActions } from "./insertActions";
import { modelLabel } from "../services/settingsService";

const rehypePlugins: Options["rehypePlugins"] = [
  [rehypeHighlight, { languages: { ...common, excel }, aliases: { excel: ["formula", "xlsx", "xls"] }, detect: false }],
];
const remarkPlugins: Options["remarkPlugins"] = [remarkGfm];

function CopyButton({ text }: { text: string }) {
  const [done, setDone] = useState(false);
  return (
    <Tooltip content={done ? "Copied" : "Copy"} relationship="label">
      <Button
        size="small"
        appearance="subtle"
        icon={done ? <Checkmark16Regular /> : <Copy16Regular />}
        onClick={async () => {
          if (await copyText(text)) {
            setDone(true);
            setTimeout(() => setDone(false), 1500);
          }
        }}
      />
    </Tooltip>
  );
}

function CodeBlock({ node, children, streaming }: { node?: HastLike; children: ReactNode; streaming: boolean }) {
  const actions = useInsertActions();
  const code = node?.children?.find((c) => c.tagName === "code");
  const classes = (code?.properties?.className as string[] | undefined) ?? [];
  const lang = (classes.find((c) => c.startsWith("language-")) ?? "").replace("language-", "").toLowerCase();
  const text = hastText(code).replace(/\n$/, "");
  const formulas = ["excel", "formula", "xlsx", "xls", ""].includes(lang) ? formulasFromCode(text) : [];
  const isTable = lang === "csv" || lang === "tsv";

  return (
    <div className="zee-code">
      <div className="zee-code-bar">
        <span className="zee-code-lang">{lang || "text"}</span>
        <span className="zee-spacer" />
        {formulas.length > 0 && (
          <Button size="small" appearance="primary" icon={<Add16Regular />} disabled={streaming} onClick={() => actions.insertFormulas(formulas)}>
            {formulas.length > 1 ? `Insert ${formulas.length} formulas` : "Insert at cursor"}
          </Button>
        )}
        {isTable && (
          <Button size="small" appearance="primary" icon={<Add16Regular />} disabled={streaming} onClick={() => actions.insertTable(parseDelimited(text))}>
            Insert table
          </Button>
        )}
        <CopyButton text={text} />
      </div>
      <pre>{children}</pre>
    </div>
  );
}

function TableBlock({ node, children, streaming }: { node?: HastLike; children: ReactNode; streaming: boolean }) {
  const actions = useInsertActions();
  const matrix = useMemo(() => (node ? hastTableToMatrix(node) : []), [node]);
  const tsv = useMemo(() => matrix.map((r) => r.join("\t")).join("\n"), [matrix]);
  return (
    <div className="zee-table">
      <div className="zee-table-scroll">
        <table>{children}</table>
      </div>
      <div className="zee-code-bar">
        <span className="zee-code-lang">
          table · {Math.max(0, matrix.length - 1)} rows × {matrix[0]?.length ?? 0} cols
        </span>
        <span className="zee-spacer" />
        <Menu positioning="below-end">
          <MenuTrigger disableButtonEnhancement>
            {(triggerProps: MenuButtonProps) => (
              <SplitButton
                size="small"
                appearance="primary"
                icon={<Add16Regular />}
                disabled={streaming || matrix.length === 0}
                menuButton={triggerProps}
                primaryActionButton={{ onClick: () => actions.insertTable(matrix), title: "Insert at the selected cell" }}
              >
                Insert table
              </SplitButton>
            )}
          </MenuTrigger>
          <MenuPopover>
            <MenuList>
              <MenuItem icon={<Add16Regular />} onClick={() => actions.insertTable(matrix)}>
                Insert at the selected cell
              </MenuItem>
              <MenuItem icon={<Table16Regular />} disabled={matrix.length < 2} onClick={() => actions.insertTable(matrix, true)}>
                Insert as Excel table (filters, banded rows)
              </MenuItem>
            </MenuList>
          </MenuPopover>
        </Menu>
        <CopyButton text={tsv} />
      </div>
    </div>
  );
}

const ACTION_ICON: Record<ActionLog["status"], JSX.Element> = {
  running: <Spinner size="extra-tiny" />,
  done: <Checkmark16Regular />,
  read: <Eye16Regular />,
  rejected: <Prohibited16Regular />,
  error: <ErrorCircle16Regular />,
};

function ActionList({ actions }: { actions: ActionLog[] }) {
  if (!actions.length) return null;
  return (
    <ul className="zee-actions">
      {actions.map((a, i) => (
        <li key={i} className={`zee-action zee-action-${a.status}`} title={a.detail}>
          {ACTION_ICON[a.status]} <span>{a.description}</span>
          {a.detail && <span className="zee-action-detail"> – {a.detail}</span>}
        </li>
      ))}
    </ul>
  );
}

function ApprovalCard({ pending }: { pending: PendingApproval }) {
  const actions = useInsertActions();
  return (
    <div className="zee-approval">
      <div className="zee-approval-title">
        <Wrench16Regular /> Zee AI wants to change your workbook:
      </div>
      <ol>
        {pending.calls.map((c, i) => (
          <li key={i}>
            {c.description}
            {c.overwrites && (
              <div className="zee-approval-warn">
                <Warning16Regular /> Replaces existing data in {c.overwrites.replace(/^.*!/, "")}
              </div>
            )}
            {c.preview && (
              <table className="zee-approval-preview">
                <tbody>
                  {c.preview.map((r, ri) => (
                    <tr key={ri}>
                      {r.map((v, ci) => (
                        <td key={ci}>{v.length > 24 ? `${v.slice(0, 22)}…` : v}</td>
                      ))}
                    </tr>
                  ))}
                </tbody>
              </table>
            )}
          </li>
        ))}
      </ol>
      <div className="zee-row">
        <Button appearance="primary" size="small" icon={<Checkmark16Regular />} onClick={() => actions.respondToApproval(true)}>
          Apply
        </Button>
        <Button size="small" icon={<Dismiss16Regular />} onClick={() => actions.respondToApproval(false)}>
          Reject
        </Button>
        <span className="zee-muted">You can undo with ↶</span>
      </div>
    </div>
  );
}

function SourceList({ sources }: { sources: NonNullable<ChatMessage["sources"]> }) {
  return (
    <div className="zee-sources">
      <span className="zee-sources-title">
        <Globe16Regular /> Sources
      </span>
      <ol>
        {sources.slice(0, 10).map((src) => (
          <li key={src.uri}>
            <a
              href={src.uri}
              onClick={(e) => {
                e.preventDefault();
                openExternal(src.uri);
              }}
            >
              {src.title}
            </a>
          </li>
        ))}
      </ol>
    </div>
  );
}

function ContinueButton() {
  const actions = useInsertActions();
  return (
    <div className="zee-row zee-continue">
      <Button size="small" appearance="primary" icon={<ArrowRight16Regular />} onClick={() => actions.continueAgent()}>
        Continue
      </Button>
    </div>
  );
}

function formatSize(bytes: number): string {
  return bytes >= 1_000_000 ? `${(bytes / 1_000_000).toFixed(1)} MB` : `${Math.max(1, Math.round(bytes / 1000))} KB`;
}

function formatTokens(n: number): string {
  return n >= 1000 ? `${(n / 1000).toFixed(1)}k` : String(n);
}

function MessageBubbleImpl({ message }: { message: ChatMessage }) {
  const streaming = message.status === "streaming";
  const components = useMemo<Components>(
    () => ({
      pre: ({ node, children }) => (
        <CodeBlock node={node as unknown as HastLike} streaming={streaming}>
          {children}
        </CodeBlock>
      ),
      table: ({ node, children }) => (
        <TableBlock node={node as unknown as HastLike} streaming={streaming}>
          {children}
        </TableBlock>
      ),
      a: ({ href, children }) => (
        <a
          href={href}
          target="_blank"
          rel="noreferrer"
          onClick={(e) => {
            // Task panes can't open new tabs reliably (desktop Excel ignores target=_blank or navigates the pane away).
            if (href && /^https?:/i.test(href)) {
              e.preventDefault();
              openExternal(href);
            }
          }}
        >
          {children}
        </a>
      ),
    }),
    [streaming]
  );

  if (message.role === "user") {
    return (
      <div className="zee-msg zee-msg-user">
        <div className="zee-bubble">{message.text}</div>
        {message.attachments?.map((a) => (
          <div key={a.name} className="zee-chip" title={`${a.mimeType} · ${formatSize(a.size)}`}>
            {a.mimeType === "application/pdf" ? <DocumentPdf16Regular /> : a.mimeType.startsWith("image/") ? <Image16Regular /> : <Attach16Regular />} {a.name}
          </div>
        ))}
        {message.contextLabel && <div className="zee-chip">📎 {message.contextLabel}</div>}
        <UserMessageTools text={message.text} />
      </div>
    );
  }

  if (message.status === "error") return <ErrorBubble message={message} />;

  const working = streaming && !message.text && !message.pending;
  return (
    <div className="zee-msg zee-msg-assistant">
      {message.mode === "agent" && message.actions && <ActionList actions={message.actions} />}
      {message.pending && <ApprovalCard pending={message.pending} />}
      <div className="zee-bubble zee-markdown">
        {message.text ? (
          <ReactMarkdown remarkPlugins={remarkPlugins} rehypePlugins={rehypePlugins} components={components}>
            {message.text}
          </ReactMarkdown>
        ) : working ? (
          <Spinner size="tiny" label={message.mode === "agent" ? "Working on your workbook…" : "Thinking…"} labelPosition="after" />
        ) : message.pending ? (
          <span className="zee-muted">Waiting for your approval…</span>
        ) : null}
        {streaming && message.text && <span className="zee-caret" />}
      </div>
      {!streaming && message.sources && message.sources.length > 0 && <SourceList sources={message.sources} />}
      {!streaming && message.canContinue && <ContinueButton />}
      {!streaming && (
        <div className="zee-meta">
          {message.model && <span>{modelLabel(message.model)}</span>}
          {message.mode === "agent" && <span className="zee-chip zee-chip-inline">agent</span>}
          {message.usage && (
            <span>
              {formatTokens(message.usage.promptTokens)} in · {formatTokens(message.usage.outputTokens)} out
            </span>
          )}
          <span className="zee-spacer" />
          <RegenerateButton message={message} />
          <CopyButton text={message.text} />
        </div>
      )}
    </div>
  );
}

export const MessageBubble = memo(MessageBubbleImpl);
