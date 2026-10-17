import { useEffect, useMemo, useState } from "react";
import {
  Button,
  Combobox,
  Divider,
  DrawerBody,
  DrawerHeader,
  DrawerHeaderTitle,
  Dropdown,
  Field,
  Input,
  Link,
  MessageBar,
  MessageBarBody,
  Option,
  OptionGroup,
  OverlayDrawer,
  Radio,
  RadioGroup,
  Spinner,
  Switch,
} from "@fluentui/react-components";
import { ArrowSync16Regular, Dismiss24Regular, Eye20Regular, EyeOff20Regular, Person16Regular } from "@fluentui/react-icons";
import { ACCENTS, DEFAULT_MODELS, DEFAULT_SETTINGS, modelHint, modelLabel, type AuthMode, type Density, type ThemeMode, type ZeeSettings } from "../services/settingsService";
import { loadPrompts, savePrompts } from "../services/promptStore";
import { BackupSection, UsageSection } from "../qol/BackupSection";
import { getGoogleSession, signInWithGoogle, signOutGoogle, type GoogleSession } from "../services/authService";
import { listModels, testConnection } from "../services/geminiService";
import { clearFunctionCache } from "../functions/zeeFunctions";
import type { UsageTotals } from "../App";
import { openExternal } from "./insertActions";
import { ExcelSupportSection } from "../qol/ExcelSupport";

interface Props {
  open: boolean;
  settings: ZeeSettings;
  usage: UsageTotals;
  onChange: (patch: Partial<ZeeSettings>) => void;
  /** Settings from an export file (Backup & move). */
  onImport: (patch: Partial<ZeeSettings>) => void;
  onClose: () => void;
}

type Status = { intent: "success" | "error" | "info"; text: string } | null;

const TEMPS: Array<{ value: string; label: string }> = [
  { value: "auto", label: "Auto (model default – recommended)" },
  { value: "0", label: "0 – precise, repeatable" },
  { value: "0.4", label: "0.4 – balanced" },
  { value: "1", label: "1.0 – creative" },
];

function minutesLeft(s: GoogleSession | null): number {
  return s ? Math.max(0, Math.round((s.expiresAt - Date.now()) / 60000)) : 0;
}

export function SettingsDrawer({ open, settings, usage, onChange, onImport, onClose }: Props) {
  const [showKey, setShowKey] = useState(false);
  const [status, setStatus] = useState<Status>(null);
  const [testing, setTesting] = useState(false);
  const [refreshing, setRefreshing] = useState(false);
  const [signingIn, setSigningIn] = useState(false);
  const [session, setSession] = useState<GoogleSession | null>(() => getGoogleSession());
  const [modelText, setModelText] = useState(settings.model);

  useEffect(() => setModelText(settings.model), [settings.model]);
  useEffect(() => {
    if (open) {
      setSession(getGoogleSession());
      setStatus(null);
    }
  }, [open]);

  const extraModels = useMemo(
    () => settings.fetchedModels.filter((id) => !DEFAULT_MODELS.some((m) => m.id === id)),
    [settings.fetchedModels]
  );
  const signedIn = !!session && minutesLeft(session) > 0;

  const runTest = async () => {
    setTesting(true);
    setStatus(null);
    try {
      setStatus({ intent: "success", text: await testConnection(settings) });
    } catch (e) {
      setStatus({ intent: "error", text: (e as Error).message });
    } finally {
      setTesting(false);
    }
  };

  const refreshModels = async () => {
    setRefreshing(true);
    try {
      const ids = await listModels(settings.apiKey);
      onChange({ fetchedModels: ids });
      setStatus({ intent: "info", text: `Loaded ${ids.length} models from your account.` });
    } catch (e) {
      setStatus({ intent: "error", text: (e as Error).message });
    } finally {
      setRefreshing(false);
    }
  };

  const signIn = async () => {
    setSigningIn(true);
    setStatus(null);
    try {
      const s = await signInWithGoogle(settings.googleClientId);
      setSession(s);
      setStatus({ intent: "success", text: `Signed in${s.email ? ` as ${s.email}` : ""}.` });
    } catch (e) {
      setStatus({ intent: "error", text: (e as Error).message });
    } finally {
      setSigningIn(false);
    }
  };

  const commitModel = (id: string) => {
    const v = id.trim();
    if (v && v !== settings.model) onChange({ model: v });
  };

  return (
    <OverlayDrawer open={open} position="end" size="full" onOpenChange={(_, d) => !d.open && onClose()}>
      <DrawerHeader>
        <DrawerHeaderTitle action={<Button appearance="subtle" aria-label="Close" icon={<Dismiss24Regular />} onClick={onClose} />}>
          Settings
        </DrawerHeaderTitle>
      </DrawerHeader>
      <DrawerBody className="zee-settings">
        <section>
          <h3>Connection</h3>
          <RadioGroup value={settings.authMode} onChange={(_, d) => onChange({ authMode: d.value as AuthMode })}>
            <Radio value="apiKey" label="Gemini API key (bring your own key)" />
            <Radio value="google" label="Google account (Vertex AI, your Cloud project quota)" />
            <Radio value="proxy" label="Team proxy (key kept on your server)" />
          </RadioGroup>

          {settings.authMode === "proxy" ? (
            <div className="zee-form">
              <Field label="Proxy URL" hint="Your deployed Zee AI proxy (see proxy/README.md), e.g. https://zee-proxy.yourteam.workers.dev">
                <Input value={settings.proxyUrl} onChange={(_, d) => onChange({ proxyUrl: d.value })} placeholder="https://…" />
              </Field>
              <Field label="Team access token" hint="The shared secret your admin configured on the proxy.">
                <Input type="password" value={settings.proxyToken} onChange={(_, d) => onChange({ proxyToken: d.value })} autoComplete="off" />
              </Field>
            </div>
          ) : settings.authMode === "apiKey" ? (
            <Field
              label="Gemini API key"
              hint={
                <>
                  Get a free key at{" "}
                  <Link inline onClick={() => openExternal("https://aistudio.google.com/apikey")}>
                    aistudio.google.com/apikey
                  </Link>
                  . Stored only on this computer, never in the workbook.
                </>
              }
            >
              <Input
                type={showKey ? "text" : "password"}
                value={settings.apiKey}
                onChange={(_, d) => onChange({ apiKey: d.value })}
                placeholder="AIza…"
                autoComplete="off"
                contentAfter={
                  <Button
                    size="small"
                    appearance="transparent"
                    aria-label={showKey ? "Hide key" : "Show key"}
                    icon={showKey ? <EyeOff20Regular /> : <Eye20Regular />}
                    onClick={() => setShowKey(!showKey)}
                  />
                }
              />
            </Field>
          ) : (
            <div className="zee-form">
              <Field label="OAuth Client ID (Web application)" hint="Google Cloud console → APIs & Services → Credentials">
                <Input value={settings.googleClientId} onChange={(_, d) => onChange({ googleClientId: d.value })} placeholder="1234-abc.apps.googleusercontent.com" />
              </Field>
              <Field label="Google Cloud project ID" hint="Vertex AI API must be enabled; usage is billed to this project.">
                <Input value={settings.gcpProjectId} onChange={(_, d) => onChange({ gcpProjectId: d.value })} placeholder="my-project-123" />
              </Field>
              <Field label="Vertex AI location" hint='"global" works for most Gemini models; or e.g. us-central1, europe-west4.'>
                <Input value={settings.gcpLocation} onChange={(_, d) => onChange({ gcpLocation: d.value })} placeholder="global" />
              </Field>
              <div className="zee-row">
                {signedIn ? (
                  <>
                    <span className="zee-chip">
                      <Person16Regular /> {session?.email ?? "Signed in"} · {minutesLeft(session)} min left
                    </span>
                    <Button
                      size="small"
                      onClick={() => {
                        signOutGoogle();
                        setSession(null);
                      }}
                    >
                      Sign out
                    </Button>
                  </>
                ) : (
                  <Button appearance="primary" onClick={signIn} disabled={signingIn || !settings.googleClientId.trim()} icon={signingIn ? <Spinner size="tiny" /> : undefined}>
                    Sign in with Google
                  </Button>
                )}
              </div>
            </div>
          )}

          <div className="zee-row">
            <Button onClick={runTest} disabled={testing} icon={testing ? <Spinner size="tiny" /> : undefined}>
              Test connection
            </Button>
          </div>
          {status && (
            <MessageBar intent={status.intent}>
              <MessageBarBody>{status.text}</MessageBarBody>
            </MessageBar>
          )}
        </section>

        <Divider />

        <section>
          <h3>Model</h3>
          <Field label="Gemini model" hint={modelHint(settings.model) ?? "Custom model ID"}>
            <div className="zee-row">
              <Combobox
                className="zee-grow"
                freeform
                value={modelText}
                selectedOptions={[settings.model]}
                onChange={(e) => setModelText(e.target.value)}
                onBlur={() => commitModel(modelText)}
                onOptionSelect={(_, d) => {
                  if (d.optionValue) {
                    setModelText(d.optionValue);
                    commitModel(d.optionValue);
                  }
                }}
              >
                <OptionGroup label="Recommended">
                  {DEFAULT_MODELS.map((m) => (
                    <Option key={m.id} value={m.id} text={m.id}>
                      {m.label}
                    </Option>
                  ))}
                </OptionGroup>
                {extraModels.length > 0 && (
                  <OptionGroup label="From your account">
                    {extraModels.map((id) => (
                      <Option key={id} value={id}>
                        {id}
                      </Option>
                    ))}
                  </OptionGroup>
                )}
              </Combobox>
              {settings.authMode === "apiKey" && (
                <Button
                  aria-label="Refresh model list"
                  title="Load the models available to your key"
                  icon={refreshing ? <Spinner size="tiny" /> : <ArrowSync16Regular />}
                  disabled={refreshing || !settings.apiKey.trim()}
                  onClick={refreshModels}
                />
              )}
            </div>
          </Field>
          <Field label="Creativity (temperature)">
            <Dropdown
              value={TEMPS.find((t) => t.value === (settings.temperature === null ? "auto" : String(settings.temperature)))?.label ?? String(settings.temperature)}
              selectedOptions={[settings.temperature === null ? "auto" : String(settings.temperature)]}
              onOptionSelect={(_, d) => onChange({ temperature: d.optionValue === "auto" ? null : Number(d.optionValue) })}
            >
              {TEMPS.map((t) => (
                <Option key={t.value} value={t.value}>
                  {t.label}
                </Option>
              ))}
            </Dropdown>
          </Field>
        </section>

        <Divider />

        <section>
          <h3>Appearance</h3>
          <Field label="Theme">
            <RadioGroup layout="horizontal" value={settings.theme} onChange={(_, d) => onChange({ theme: d.value as ThemeMode })} aria-label="Theme">
              <Radio value="dark" label="Dark (default)" />
              <Radio value="light" label="Light" />
              <Radio value="auto" label="Match Office" />
            </RadioGroup>
          </Field>
          <p className="zee-muted">"Match Office" follows your Office theme (File → Account → Office Theme). Quick switch: ⋯ in the header → Light / Dark mode.</p>
          <Field label="Accent colour">
            <div className="zee-accents" role="group" aria-label="Accent colour">
              {ACCENTS.map((a) => (
                <button key={a.id} type="button" className="zee-accent-swatch" aria-pressed={settings.accent === a.id} onClick={() => onChange({ accent: a.id })}>
                  <span className="zee-accent-dot" style={{ background: a.hex }} />
                  {a.label}
                </button>
              ))}
            </div>
          </Field>
          <Field label="Density">
            <RadioGroup layout="horizontal" value={settings.density} onChange={(_, d) => onChange({ density: d.value as Density })} aria-label="Density">
              <Radio value="comfortable" label="Comfortable" />
              <Radio value="compact" label="Compact (small screens)" />
            </RadioGroup>
          </Field>
          <p className="zee-muted">Windows high-contrast mode is followed automatically when it's on.</p>
        </section>

        <Divider />

        <section>
          <h3>Agent mode</h3>
          <Switch checked={settings.agentMode} onChange={(_, d) => onChange({ agentMode: d.checked })} label="Let Zee AI act on the workbook (read, write, format, charts)" />
          <Switch
            checked={settings.autoApprove}
            disabled={!settings.agentMode}
            onChange={(_, d) => onChange({ autoApprove: d.checked })}
            label="Apply changes without asking (not recommended)"
          />
          <Switch
            checked={settings.agentDryRun}
            disabled={!settings.agentMode || settings.autoApprove}
            onChange={(_, d) => onChange({ agentDryRun: d.checked })}
            label="Dry run: show cells before → after in the approval card"
          />
          <p className="zee-muted">Every change is listed for approval first and can be undone step by step with ↶.</p>
        </section>

        <Divider />

        <section>
          <h3>Gemini tools</h3>
          <Switch
            checked={settings.webSearch}
            onChange={(_, d) => onChange({ webSearch: d.checked })}
            label="Search the web (Google Search) – answers list their sources"
          />
          <Switch
            checked={settings.codeExecution}
            onChange={(_, d) => onChange({ codeExecution: d.checked })}
            label="Run Python for calculations in chat (code execution)"
          />
          <p className="zee-muted">
            Web search may add cost on your Google account (grounded requests are billed separately above the free allowance). In agent mode, web search is offered as a
            tool; code execution applies to chat answers.
          </p>
        </section>

        <Divider />

        <section>
          <h3>Sheet context</h3>
          <Switch checked={settings.includeSelection} onChange={(_, d) => onChange({ includeSelection: d.checked })} label="Include the selected range in prompts by default" />
          <Switch
            checked={settings.includeWorkbookOverview}
            onChange={(_, d) => onChange({ includeWorkbookOverview: d.checked })}
            label="Include a map of all sheets (names, headers, tables)"
          />
          <Field label="Max cells sent per prompt" hint="Larger selections are truncated to this many cells to protect speed and quota.">
            <Input
              type="number"
              min={50}
              max={20000}
              step={50}
              value={String(settings.maxContextCells)}
              onChange={(_, d) => {
                const n = Number(d.value);
                if (Number.isFinite(n)) onChange({ maxContextCells: Math.min(20000, Math.max(50, Math.round(n))) });
              }}
            />
          </Field>
          <Switch
            checked={settings.bigDataMode}
            onChange={(_, d) => onChange({ bigDataMode: d.checked })}
            label="Big-data mode: summarize very large selections (column statistics + a sample)"
          />
          <Field label="Big-data mode from (rows)" hint="Selections with at least this many rows are summarized over every row instead of cut off. Shown as “big-data mode” on the message.">
            <Input
              type="number"
              min={1000}
              step={10000}
              disabled={!settings.bigDataMode}
              value={String(settings.bigDataRows)}
              onChange={(_, d) => {
                const n = Number(d.value);
                if (Number.isFinite(n)) onChange({ bigDataRows: Math.min(1_000_000, Math.max(1000, Math.round(n))) });
              }}
            />
          </Field>
        </section>

        <Divider />

        <section>
          <h3>Cost guard</h3>
          <Field label="Ask before sending prompts larger than (tokens)" hint="0 = never ask. ~4 characters ≈ 1 token.">
            <Input
              type="number"
              min={0}
              step={1000}
              value={String(settings.costGuardTokens)}
              onChange={(_, d) => {
                const n = Number(d.value);
                if (Number.isFinite(n) && n >= 0) onChange({ costGuardTokens: Math.round(n) });
              }}
            />
          </Field>
        </section>

        <Divider />

        <section>
          <h3>=ZEE.ASK() worksheet function</h3>
          <p className="zee-muted">
            Type <code>=ZEE.ASK("Classify this failure as electrical or mechanical", B2)</code> in any cell and fill down. Answers are cached so recalculation doesn't cost
            extra.
          </p>
          <p className="zee-muted">
            Petroleum functions (instant, no AI): <code>ZEE.VOGEL</code>, <code>ZEE.VOGEL_QMAX</code>, <code>ZEE.PI</code>, <code>ZEE.RS</code>, <code>ZEE.PB</code>,{" "}
            <code>ZEE.BO</code>, <code>ZEE.ARPS</code>, <code>ZEE.ARPS_CUM</code>, <code>ZEE.TDH</code>, <code>ZEE.HHP</code>, <code>ZEE.AFFINITY</code>,{" "}
            <code>ZEE.SG_MIX</code>, <code>ZEE.CONVERT</code>. Gas wells: <code>ZEE.Z</code>, <code>ZEE.BG</code>, <code>ZEE.GAS_VISC</code>,{" "}
            <code>ZEE.DELIV</code>, <code>ZEE.DELIV_Q</code>. Type <code>=ZEE.</code> in a cell to see their arguments.
          </p>
          <div className="zee-row">
            <Button
              size="small"
              onClick={() => {
                clearFunctionCache();
                setStatus({ intent: "info", text: "=ZEE.ASK() cache cleared. Recalculate (Ctrl+Alt+F9) to refresh answers." });
              }}
            >
              Clear answer cache
            </Button>
          </div>
        </section>

        <Divider />

        <UsageSection session={usage} prices={{ inPerM: settings.priceInPerM, outPerM: settings.priceOutPerM }} onPrices={(p) => onChange({ priceInPerM: p.inPerM, priceOutPerM: p.outPerM })} />
        <p className="zee-muted">
          Current model: {modelLabel(settings.model)}. Quotas and billing are managed in Google AI Studio (API key) or the Google Cloud console (Vertex AI).
        </p>

        <Divider />

        <BackupSection settings={settings} defaults={DEFAULT_SETTINGS} onImport={onImport} loadPrompts={loadPrompts} savePrompts={savePrompts} />

        <Divider />

        <ExcelSupportSection />
      </DrawerBody>
    </OverlayDrawer>
  );
}
