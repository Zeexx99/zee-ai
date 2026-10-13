/** Settings sections: usage & cost, backup (export / import settings) and updates. */
import { useMemo, useState } from "react";
import { Button, Checkbox, Field, Input, MessageBar, MessageBarBody, Spinner, Textarea } from "@fluentui/react-components";
import { ArrowClockwise16Regular, ArrowDownload16Regular, ArrowUpload16Regular } from "@fluentui/react-icons";
import { copyText } from "../components/insertActions";
import { APP } from "./app";
import { buildSettingsExport, downloadFile, parseSettingsExport } from "./settingsTransfer";
import { estimateCost, formatCost, formatTokens, usageForDay, usageForMonth, type Prices } from "./usageLog";
import { VERSION, buildLabel, fetchLiveBuild, isNewerBuild, reloadLatest } from "./version";

type Status = { intent: "success" | "error" | "info" | "warning"; text: string } | null;

interface UsageLike {
  requests: number;
  promptTokens: number;
  outputTokens: number;
}

export function UsageSection({ session, prices, onPrices }: { session: UsageLike; prices: Prices; onPrices: (p: Prices) => void }) {
  // Re-read on every render of the drawer (cheap) so the numbers are current when it opens.
  const today = usageForDay();
  const month = usageForMonth();
  const rows: Array<[string, UsageLike]> = [
    ["This session", session],
    ["Today", today],
    ["This month", month],
  ];
  const num = (v: string) => {
    const n = Number(v);
    return Number.isFinite(n) && n >= 0 ? n : 0;
  };
  return (
    <section>
      <h3>Usage</h3>
      <table className="zee-usage-table">
        <thead>
          <tr>
            <th />
            <th>Requests</th>
            <th>Tokens in</th>
            <th>Tokens out</th>
            <th>Est. cost</th>
          </tr>
        </thead>
        <tbody>
          {rows.map(([label, u]) => {
            const cost = estimateCost(u, prices);
            return (
              <tr key={label}>
                <th>{label}</th>
                <td>{u.requests.toLocaleString()}</td>
                <td>{formatTokens(u.promptTokens)}</td>
                <td>{formatTokens(u.outputTokens)}</td>
                <td>{cost === null ? "–" : formatCost(cost)}</td>
              </tr>
            );
          })}
        </tbody>
      </table>
      <div className="zee-row">
        <Field label="Price per 1M input tokens (USD)" className="zee-grow">
          <Input type="number" min={0} step={0.05} value={prices.inPerM ? String(prices.inPerM) : ""} placeholder="e.g. 0.30" onChange={(_, d) => onPrices({ ...prices, inPerM: num(d.value) })} />
        </Field>
        <Field label="Price per 1M output tokens (USD)" className="zee-grow">
          <Input type="number" min={0} step={0.05} value={prices.outPerM ? String(prices.outPerM) : ""} placeholder="e.g. 2.50" onChange={(_, d) => onPrices({ ...prices, outPerM: num(d.value) })} />
        </Field>
      </div>
      <p className="zee-muted">
        Optional: copy your model’s prices from Google’s Gemini pricing page to see an estimated cost here and under the chat box. Counted on this PC only; your real bill is in Google AI
        Studio / Cloud console.
      </p>
    </section>
  );
}

export function BackupSection<T extends object>({
  settings,
  defaults,
  onImport,
  loadPrompts,
  savePrompts,
}: {
  settings: T;
  defaults: T;
  onImport: (patch: Partial<T>) => void;
  loadPrompts: () => string[];
  savePrompts: (p: string[]) => void;
}) {
  const [includeSecrets, setIncludeSecrets] = useState(false);
  const [status, setStatus] = useState<Status>(null);
  const [paste, setPaste] = useState("");
  const [showPaste, setShowPaste] = useState(false);
  const [checking, setChecking] = useState(false);
  const [update, setUpdate] = useState<Status>(null);
  const built = useMemo(() => buildLabel(), []);

  const doImport = (text: string, source: string) => {
    try {
      const r = parseSettingsExport(text, defaults);
      const keys = Object.keys(r.patch);
      if (!keys.length) throw new Error("The file contains no settings this add-in uses.");
      onImport(r.patch);
      if (r.prompts?.length) {
        const mine = loadPrompts();
        savePrompts([...mine, ...r.prompts.filter((p) => !mine.includes(p))]);
      }
      const cross = r.from !== APP.id && r.from !== "unknown" ? ` (made in the ${r.from === "word" ? "Word" : "Excel"} add-in – shared settings were applied)` : "";
      setStatus({
        intent: "success",
        text: `Imported ${keys.length} setting${keys.length === 1 ? "" : "s"}${r.prompts?.length ? ` and ${r.prompts.length} saved prompts` : ""} from ${source}${cross}.${"apiKey" in r.patch ? "" : " The API key wasn't in the file – enter it in Connection."}`,
      });
      setPaste("");
      setShowPaste(false);
    } catch (e) {
      setStatus({ intent: "error", text: (e as Error).message });
    }
  };

  return (
    <>
      <section>
        <h3>Backup &amp; move to another PC</h3>
        <p className="zee-muted">Export your settings, quick actions, protected terms and saved prompts to a file, then import it on the other PC (or in the other Zee AI add-in).</p>
        <Checkbox checked={includeSecrets} onChange={(_, d) => setIncludeSecrets(!!d.checked)} label="Include my API key and proxy token (keep the file private)" />
        <div className="zee-row">
          <Button
            size="small"
            icon={<ArrowDownload16Regular />}
            onClick={async () => {
              const json = buildSettingsExport(settings, { includeSecrets, prompts: loadPrompts() });
              const name = `zee-ai-${APP.id}-settings.json`;
              const saved = downloadFile(name, json);
              const copied = await copyText(json);
              setStatus({
                intent: saved || copied ? "success" : "error",
                text: saved || copied ? `${saved ? `Saved ${name} (check Downloads). ` : ""}${copied ? "Also copied to the clipboard – you can paste it on the other PC." : ""}` : "Couldn't save or copy the file here.",
              });
            }}
          >
            Export settings
          </Button>
          <label className="zee-file-pick">
            <span className="zee-muted">
              <ArrowUpload16Regular /> Import a file:
            </span>
            <input
              type="file"
              accept=".json,application/json"
              onChange={async (e) => {
                const f = e.target.files?.[0];
                e.target.value = "";
                if (f) doImport(await f.text(), `“${f.name}”`);
              }}
            />
          </label>
          <Button size="small" appearance="subtle" onClick={() => setShowPaste((v) => !v)}>
            {showPaste ? "Cancel paste" : "Paste instead…"}
          </Button>
        </div>
        {showPaste && (
          <>
            <Textarea rows={4} resize="vertical" value={paste} placeholder="Paste the exported settings text here" onChange={(_, d) => setPaste(d.value)} />
            <div className="zee-row">
              <Button size="small" appearance="primary" disabled={!paste.trim()} onClick={() => doImport(paste, "the pasted text")}>
                Import
              </Button>
            </div>
          </>
        )}
        {status && (
          <MessageBar intent={status.intent}>
            <MessageBarBody>{status.text}</MessageBarBody>
          </MessageBar>
        )}
      </section>

      <section>
        <h3>Updates</h3>
        <p className="zee-muted">
          {APP.name} {VERSION}
          {built ? ` · built ${built}` : " · development build"}
        </p>
        <div className="zee-row">
          <Button
            size="small"
            disabled={checking}
            icon={checking ? <Spinner size="extra-tiny" /> : undefined}
            onClick={async () => {
              setChecking(true);
              const live = await fetchLiveBuild();
              setChecking(false);
              if (!live) setUpdate({ intent: "warning", text: "Couldn't check (offline, or this is the local development server)." });
              else if (isNewerBuild(live)) setUpdate({ intent: "info", text: `A newer build (${live.version}, ${buildLabel(live.build)}) is online – click “Reload latest version”.` });
              else setUpdate({ intent: "success", text: "You're on the latest version." });
            }}
          >
            Check for updates
          </Button>
          <Button size="small" appearance="primary" icon={<ArrowClockwise16Regular />} onClick={() => void reloadLatest()}>
            Reload latest version
          </Button>
        </div>
        {update && (
          <MessageBar intent={update.intent}>
            <MessageBarBody>{update.text}</MessageBarBody>
          </MessageBar>
        )}
        <p className="zee-muted">“Reload latest version” clears {APP.host}’s cached copy of this pane and loads the newest published build – no need to close {APP.host}.</p>
      </section>
    </>
  );
}
