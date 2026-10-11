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
import { DEFAULT_MODELS, modelHint, modelLabel, type AuthMode, type ThemeMode, type ZeeSettings } from "../services/settingsService";
import { getGoogleSession, signInWithGoogle, signOutGoogle, type GoogleSession } from "../services/authService";
import { listModels, testConnection } from "../services/geminiService";
import { clearFunctionCache } from "../functions/zeeFunctions";
import type { UsageTotals } from "../App";
import { openExternal } from "./insertActions";

interface Props {
  open: boolean;
  settings: ZeeSettings;
  usage: UsageTotals;
  onChange: (patch: Partial<ZeeSettings>) => void;
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

export function SettingsDrawer({ open, settings, usage, onChange, onClose }: Props) {
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
          <RadioGroup layout="horizontal" value={settings.theme} onChange={(_, d) => onChange({ theme: d.value as ThemeMode })} aria-label="Theme">
            <Radio value="auto" label="Match Office" />
            <Radio value="light" label="Light" />
            <Radio value="dark" label="Dark" />
          </RadioGroup>
          <p className="zee-muted">"Match Office" follows your Office theme (File → Account → Office Theme). Tip: the moon/sun button in the header switches quickly.</p>
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
          <p className="zee-muted">Every change is listed for approval first and can be undone step by step with ↶.</p>
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

        <section>
          <h3>Usage this session</h3>
          <div className="zee-usage">
            <div>
              <strong>{usage.requests}</strong>
              <span>requests</span>
            </div>
            <div>
              <strong>{usage.promptTokens.toLocaleString()}</strong>
              <span>input tokens</span>
            </div>
            <div>
              <strong>{usage.outputTokens.toLocaleString()}</strong>
              <span>output tokens</span>
            </div>
          </div>
          <p className="zee-muted">
            Current model: {modelLabel(settings.model)}. Quotas and billing are managed in Google AI Studio (API key) or the Google Cloud console (Vertex AI).
          </p>
        </section>
      </DrawerBody>
    </OverlayDrawer>
  );
}
