import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import {
  Badge,
  Button,
  FluentProvider,
  Menu,
  MenuItem,
  MenuList,
  MenuPopover,
  MenuTrigger,
  MessageBar,
  MessageBarActions,
  MessageBarBody,
  MenuDivider,
  MenuGroupHeader,
  ProgressBar,
  Spinner,
  Toast,
  ToastBody,
  ToastTitle,
  Toaster,
  Tooltip,
  useId,
  useToastController,
  type ToastIntent,
} from "@fluentui/react-components";
import {
  Add20Regular,
  ArrowDownload20Regular,
  ArrowUndo16Regular,
  ArrowUndo20Regular,
  Copy20Regular,
  History20Regular,
  Settings20Regular,
  Table20Regular,
  WeatherMoon20Regular,
  WeatherSunny20Regular,
} from "@fluentui/react-icons";

import { darkTheme, lightTheme, useAutoDark, useBodyThemeClass } from "./theme";
import { useChat, type ChatMessage } from "./hooks/useChat";
import { ChatFeed } from "./components/ChatFeed";
import { ChatInput } from "./components/ChatInput";
import { QuickActions } from "./components/QuickActions";
// Imported eagerly (not lazy-loaded): a separate chunk breaks when Excel holds a cached page from an older deploy.
import { SettingsDrawer } from "./components/SettingsDrawer";
import {
  ChartDialog,
  CleanDialog,
  ConfirmDialog,
  DeclineCurveDialog,
  DrawDialog,
  EspDialog,
  ExtractDialog,
  GenerateFormulaDialog,
  UnitDialog,
  type ConfirmState,
  type DcaChoice,
  type DrawChoice,
  type EspDialogChoice,
  type ExtractChoice,
} from "./components/Dialogs";
import { InsertActionsContext, copyText, type InsertActions } from "./components/insertActions";
import { UpdateBanners, useUpdates } from "./qol/UpdateBanners";
import { DailyReportDialog, EspLogDialog, RulesDialog, type DailyReportTarget } from "./components/ToolDialogs";
import {
  ESP_FAILURE_PROMPT,
  applyColumnRules,
  createEspFailureLog,
  datesIn,
  profileSelectionForRules,
  readHeaderedSelection,
  runDailyReport,
  runEspFailureAnalysis,
  runFormulaAudit,
  type HeaderedSelection,
  type RulesTarget,
} from "./services/toolActions";
import { RULES_SCHEMA, RULES_SYSTEM_PROMPT, profilesForPrompt, sanitizeRules, type ColumnRule } from "./utils/columnRules";
import { AUDIT_LABELS } from "./utils/formulaAudit";
import type { DailyMapping } from "./utils/dailyReport";
import type { EspLogMapping } from "./utils/espFailures";
import { recordUsage, usageForDay } from "./qol/usageLog";
import { useOfflineQueue } from "./qol/offline";
import { buildDiagnostics } from "./qol/diagnostics";
import { timeAgo } from "./utils/undoHistory";
import { describeLines } from "./utils/clean";
import { FunctionDialog, type ApplyWhere, type FunctionDialogState } from "./components/FunctionPicker";
import { buildFormula, canApplyDirectly, guessHeader, inputCount, type FnSpec } from "./functions/catalog";

import { isConfigured, loadSettings, modelLabel, saveSettings, type ZeeSettings } from "./services/settingsService";
import type { Attachment } from "./services/geminiService";
import {
  AGGREGATE_LABELS,
  OverwriteRequiredError,
  cleanSelection,
  autofitAllColumns,
  applyEmptyLines,
  getChangeHistory,
  revertChange,
  scanEmptyLines,
  undoBackTo,
  type HistoryItem,
  createChart,
  getWorkbookOverview,
  insertConvertedColumns,
  formatSelectionAsTable,
  freezeSelection,
  getFunctionTarget,
  writeFormulaColumn,
  type FunctionTarget,
  addSheet,
  formatRangeAt,
  writeBlock,
  getPlotSource,
  plotColumns,
  type PlotSource,
  formatSelectionForPrompt,
  getActiveCellInfo,
  getSelectedAddress,
  getSelectionContext,
  getSheetOverview,
  insertAggregate,
  insertFormulas,
  insertResultColumn,
  isExcelAvailable,
  onUndoAvailabilityChange,
  readColumnSelection,
  undoLastChange,
  writeMatrixAtActiveCell,
  type AggregateFn,
  type ChartKind,
} from "./services/excelService";
import { ESP_PROMPT, getSelectionHeaders, runDeclineByWell, runDeclineCurve, runEspAnalysis, runProductionQc } from "./services/petroActions";
import { UNIT_CONVERSIONS } from "./utils/petro";
import { estimateTokens, generateJson, type Usage } from "./services/geminiService";
import { explainFormulaPrompt, generateFormulaPrompt, transformPrompt, TRANSFORM_SYSTEM_PROMPT } from "./services/prompts";
import { a1, a1Range, columnLetter, stripSheet } from "./utils/a1";
import { extractValue, EXTRACT_LABELS } from "./utils/extract";
import type { CleanOptions } from "./utils/clean";
import type { Cell } from "./utils/markdown";

export interface UsageTotals {
  requests: number;
  promptTokens: number;
  outputTokens: number;
}

const AI_BATCH = 50;

export default function App() {
  /* ------------------------------------------------------------ settings & theme */
  const [settings, setSettings] = useState<ZeeSettings>(() => loadSettings());
  const settingsRef = useRef(settings);
  settingsRef.current = settings;
  const getSettings = useCallback(() => settingsRef.current, []);
  const updateSettings = useCallback((patch: Partial<ZeeSettings>) => {
    setSettings((prev) => {
      const next = { ...prev, ...patch };
      saveSettings(next);
      return next;
    });
  }, []);
  const autoDark = useAutoDark();
  const dark = settings.theme === "auto" ? autoDark : settings.theme === "dark";
  useBodyThemeClass(dark);
  const inExcel = useMemo(() => isExcelAvailable(), []);

  /* ------------------------------------------------------------ toasts */
  const toasterId = useId("toaster");
  const { dispatchToast } = useToastController(toasterId);
  const toast = useCallback(
    (intent: ToastIntent, title: string, body?: string) =>
      dispatchToast(
        <Toast>
          <ToastTitle>{title}</ToastTitle>
          {body && <ToastBody>{body}</ToastBody>}
        </Toast>,
        { intent, timeout: intent === "error" ? 8000 : 4000 }
      ),
    [dispatchToast]
  );
  const toastError = useCallback((e: unknown) => toast("error", "Something went wrong", (e as Error)?.message ?? String(e)), [toast]);

  /* ------------------------------------------------------------ usage */
  const [usage, setUsage] = useState<UsageTotals>({ requests: 0, promptTokens: 0, outputTokens: 0 });
  const [today, setToday] = useState(() => usageForDay());
  const addUsage = useCallback((u: Usage) => {
    setUsage((p) => ({ requests: p.requests + 1, promptTokens: p.promptTokens + u.promptTokens, outputTokens: p.outputTokens + u.outputTokens }));
    setToday(recordUsage(u));
  }, []);
  const prices = useMemo(() => ({ inPerM: settings.priceInPerM, outPerM: settings.priceOutPerM }), [settings.priceInPerM, settings.priceOutPerM]);
  const updates = useUpdates();

  /* ------------------------------------------------------------ selection tracking */
  const [selectedAddress, setSelectedAddress] = useState<string | null>(null);
  /** × on the selection chip: leave the selection out of the next message (resets when the selection changes). */
  const [ignoreSelection, setIgnoreSelection] = useState(false);
  useEffect(() => {
    if (!inExcel) return;
    let timer = 0;
    const refresh = () => {
      window.clearTimeout(timer);
      timer = window.setTimeout(() => {
        getSelectedAddress()
          .then((a) => {
            setSelectedAddress(a.replace(/^'(.*)'!/, "$1!"));
            setIgnoreSelection(false);
          })
          .catch(() => setSelectedAddress("multiple areas"));
      }, 150);
    };
    refresh();
    Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, refresh);
    return () => {
      window.clearTimeout(timer);
      Office.context.document.removeHandlerAsync(Office.EventType.DocumentSelectionChanged, { handler: refresh });
    };
  }, [inExcel]);

  /* ------------------------------------------------------------ undo */
  const [undoLabel, setUndoLabel] = useState<string | null>(null);
  const [undoDepth, setUndoDepth] = useState(0);
  useEffect(
    () =>
      onUndoAvailabilityChange((label, depth) => {
        setUndoLabel(label);
        setUndoDepth(depth);
      }),
    []
  );
  const undo = async () => {
    try {
      const label = await undoLastChange();
      toast("success", "Undone", `Reverted: ${label}`);
    } catch (e) {
      toastError(e);
    }
  };

  /* ------------------------------------------------------------ change history */
  const [history, setHistory] = useState<HistoryItem[]>([]);
  const revertOne = async (h: HistoryItem) => {
    try {
      const label = await revertChange(h.index);
      toast("success", "Reverted", label);
    } catch (e) {
      toastError(e);
    }
  };
  const undoTo = async (h: HistoryItem) => {
    const count = undoDepth - h.index;
    const ok = await askConfirm({
      title: `Undo ${count} change${count === 1 ? "" : "s"}?`,
      message: `This undoes everything from the newest change back to “${h.label}” (${timeAgo(h.at)}).`,
      confirmLabel: "Undo",
    });
    if (!ok) return;
    try {
      const n = await undoBackTo(h.index);
      toast("success", `Undid ${n} change${n === 1 ? "" : "s"}`, `Back to before “${h.label}”.`);
    } catch (e) {
      toastError(e);
    }
  };

  /* ------------------------------------------------------------ confirm (promise based) */
  const [confirmState, setConfirmState] = useState<ConfirmState | null>(null);
  const confirmResolver = useRef<((ok: boolean) => void) | null>(null);
  const askConfirm = useCallback(
    (state: ConfirmState) =>
      new Promise<boolean>((resolve) => {
        confirmResolver.current?.(false);
        confirmResolver.current = resolve;
        setConfirmState(state);
      }),
    []
  );
  const onConfirmResult = (ok: boolean) => {
    confirmResolver.current?.(ok);
    confirmResolver.current = null;
    setConfirmState(null);
  };

  /** Cost guard: ask before sending a large prompt. */
  const confirmLarge = useCallback(
    (tokens: number) =>
      askConfirm({
        title: "Large request",
        message: `This message is about ${tokens.toLocaleString()} tokens (selection + sheet map + chat history). Send it anyway? You can lower "Max cells sent" or start a new chat to reduce cost.`,
        confirmLabel: "Send",
      }),
    [askConfirm]
  );

  const { messages, busy, send, retry, stop, clear, respondToApproval, continueAgent } = useChat({ getSettings, onUsage: addUsage, confirmLarge });

  // Esc stops a running answer from anywhere in the pane.
  const busyRef = useRef(false);
  busyRef.current = busy;
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if (e.key === "Escape" && busyRef.current) stop();
    };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, [stop]);

  /** Edit & resend: put a prompt back into the chat box. */
  const [inject, setInject] = useState<{ text: string; n: number } | null>(null);
  const editPrompt = useCallback((text: string) => setInject({ text, n: Date.now() }), []);

  /** Run a write; if it would overwrite data, ask first and retry with overwrite=true. */
  const withOverwrite = useCallback(
    async <T,>(fn: (overwrite: boolean) => Promise<T>): Promise<T | undefined> => {
      try {
        return await fn(false);
      } catch (e) {
        if (!(e instanceof OverwriteRequiredError)) throw e;
        const ok = await askConfirm({
          title: "Replace existing data?",
          message: `${stripSheet(e.address)} already contains data. Replace it? (You can undo with ↶ afterwards.)`,
          confirmLabel: "Replace",
        });
        return ok ? fn(true) : undefined;
      }
    },
    [askConfirm]
  );

  /* ------------------------------------------------------------ chat */
  const [settingsOpen, setSettingsOpen] = useState(false);

  const ensureConfigured = useCallback(() => {
    if (isConfigured(settingsRef.current)) return true;
    setSettingsOpen(true);
    toast("warning", "Connect Gemini first", "Add an API key or sign in with Google.");
    return false;
  }, [toast]);

  /** Returns [promptSuffix, contextLabel] for the current selection. */
  const selectionContext = useCallback(
    async (force = false): Promise<[string, string | undefined]> => {
      if (!inExcel || (!force && !settingsRef.current.includeSelection)) return ["", undefined];
      try {
        const ctx = await getSelectionContext(settingsRef.current.maxContextCells);
        const label = `${ctx.sheetName}!${stripSheet(ctx.isEmpty ? ctx.selectionAddress : ctx.address)}${ctx.truncated ? " (truncated)" : ""}`;
        return [`\n\n${formatSelectionForPrompt(ctx)}`, label];
      } catch (e) {
        toast("warning", "Selection not included", (e as Error).message);
        return ["", undefined];
      }
    },
    [inExcel, toast]
  );

  /** Compact map of every sheet (names, headers, tables) when enabled. */
  const workbookContext = useCallback(async (): Promise<string> => {
    if (!inExcel || !settingsRef.current.includeWorkbookOverview) return "";
    try {
      return `\n\n${await getWorkbookOverview()}`;
    } catch {
      return "";
    }
  }, [inExcel]);

  const sendChat = useCallback(
    async (text: string, attachments?: Attachment[]) => {
      if (!ensureConfigured()) return;
      const skip = ignoreSelection;
      setIgnoreSelection(false);
      const [[ctx, label], overview] = await Promise.all([skip ? Promise.resolve<[string, string | undefined]>(["", undefined]) : selectionContext(), workbookContext()]);
      const display = text || (attachments?.length ? "(see attached)" : "");
      const ask = text || "Extract the content of the attached file(s). Put tables in markdown tables so they can be inserted.";
      await send({ display, prompt: ask + ctx + overview, contextLabel: label, attachments });
    },
    [ensureConfigured, ignoreSelection, selectionContext, send, workbookContext]
  );

  /** Regenerate / Retry. Attached files aren't kept in the history, so those prompts go back to the chat box instead. */
  const retryAnswer = useCallback(
    (id: string) => {
      const i = messages.findIndex((m) => m.id === id);
      const user = i > 0 ? messages[i - 1] : undefined;
      if (user?.attachments?.length) {
        editPrompt(user.text === "(see attached)" ? "" : user.text);
        toast("info", "Attach the files again", "Files aren't kept after sending – attach them again and press Enter.");
        return;
      }
      void retry(id);
    },
    [editPrompt, messages, retry, toast]
  );

  const offline = useOfflineQueue(messages, busy, retryAnswer);

  const copyDiagnostics = useCallback(
    async (m: ChatMessage) => {
      const s = settingsRef.current;
      const report = buildDiagnostics({
        error: m.text,
        kind: m.errorKind,
        model: m.model ?? s.model,
        authMode: s.authMode,
        extra: {
          Mode: m.mode ?? (s.agentMode ? "agent" : "chat"),
          "Web search / code execution": `${s.webSearch ? "on" : "off"} / ${s.codeExecution ? "on" : "off"}`,
          "ExcelApi 1.12 / 1.17": inExcel
            ? ["1.12", "1.17"].map((v) => (Office.context.requirements.isSetSupported("ExcelApi", v) ? "yes" : "no")).join(" / ")
            : "not in Excel",
        },
      });
      if (await copyText(report)) toast("success", "Diagnostics copied", "Paste them to Claude or IT. No keys or sheet data are included.");
      else toast("error", "Couldn't copy", report.slice(0, 300));
    },
    [inExcel, toast]
  );

  /* ------------------------------------------------------------ insert actions for chat output */
  const insertActions = useMemo<InsertActions>(
    () => ({
      insertFormulas: async (formulas) => {
        try {
          const addr = await withOverwrite((ow) => insertFormulas(formulas, { overwrite: ow }));
          if (addr) toast("success", "Formula inserted", stripSheet(addr));
        } catch (e) {
          toastError(e);
        }
      },
      insertTable: async (matrix: Cell[][], asTable?: boolean) => {
        try {
          const addr = await withOverwrite((ow) => writeMatrixAtActiveCell(matrix, { overwrite: ow, header: true, asTable, label: asTable ? "insert Excel table" : "insert table" }));
          if (addr) toast("success", asTable ? "Excel table inserted" : "Table inserted", `${stripSheet(addr)}${asTable ? " · filters and banded rows on (undo with ↶)" : ""}`);
        } catch (e) {
          toastError(e);
        }
      },
      respondToApproval,
      continueAgent,
      retry: retryAnswer,
      editPrompt,
      copyDiagnostics,
      busy,
    }),
    [busy, continueAgent, copyDiagnostics, editPrompt, respondToApproval, retryAnswer, toast, toastError, withOverwrite]
  );

  /* ------------------------------------------------------------ export chat */
  const chatAsText = useCallback(
    () =>
      messages
        .filter((m) => m.status !== "streaming")
        .map((m) => {
          const who = m.role === "user" ? "You" : `Zee AI${m.mode === "agent" ? " (agent)" : ""}`;
          const ctx = m.contextLabel ? ` [${m.contextLabel}]` : "";
          const acts = m.actions?.length ? `\n\nActions:\n${m.actions.map((a) => `- ${a.status}: ${a.description}`).join("\n")}` : "";
          return `### ${who}${ctx}\n\n${m.text}${acts}`;
        })
        .join("\n\n---\n\n"),
    [messages]
  );
  const exportCopy = async () => {
    if (await copyText(chatAsText())) toast("success", "Chat copied", "Paste it into Word, Teams, an email or a text file.");
    else toast("error", "Couldn't copy", "The clipboard is blocked here – use Export to a new sheet instead.");
  };
  const exportSheet = async () => {
    try {
      const sheet = await addSheet("Chat export");
      const rows: Cell[][] = [
        ["#", "From", "Context", "Message"],
        ...messages
          .filter((m) => m.status !== "streaming")
          .map((m, i): Cell[] => [i + 1, m.role === "user" ? "You" : "Zee AI", m.contextLabel ?? "", m.text.slice(0, 32000)]),
      ];
      await writeBlock(sheet, "A1", rows, { header: true });
      await formatRangeAt(`D2:D${rows.length}`, { wrap: true }, sheet);
      toast("success", "Chat exported", `${rows.length - 1} messages → sheet "${sheet}"`);
    } catch (e) {
      toastError(e);
    }
  };

  /* ------------------------------------------------------------ quick actions */
  const [cleanOpen, setCleanOpen] = useState(false);
  const [extractOpen, setExtractOpen] = useState(false);
  const [genOpen, setGenOpen] = useState(false);
  const [chartOpen, setChartOpen] = useState(false);
  const [unitOpen, setUnitOpen] = useState(false);
  const [dcaHeaders, setDcaHeaders] = useState<string[] | null>(null);
  const [plotSource, setPlotSource] = useState<PlotSource | null>(null);
  const [job, setJob] = useState<{ text: string; abort?: AbortController; progress?: number } | null>(null);
  /** Progress callback for long scans (updates the job banner, at most ~10× per second). */
  const progressFor = (abort: AbortController) => {
    let last = 0;
    return (fraction: number, text: string) => {
      const now = Date.now();
      if (now - last < 100 && fraction < 1) return;
      last = now;
      setJob({ text, abort, progress: Math.max(0, Math.min(1, fraction)) });
    };
  };

  const onAggregate = async (fn: AggregateFn) => {
    try {
      const res = await withOverwrite((ow) => insertAggregate(fn, ow));
      if (!res) return;
      const preview = res
        .slice(0, 3)
        .map((r) => `${r.address} = ${typeof r.value === "number" ? r.value.toLocaleString(undefined, { maximumFractionDigits: 4 }) : String(r.value)}`)
        .join(", ");
      toast("success", `${AGGREGATE_LABELS[fn]} inserted`, preview + (res.length > 3 ? ` … (+${res.length - 3})` : ""));
    } catch (e) {
      toastError(e);
    }
  };

  const onClean = async (o: CleanOptions) => {
    setCleanOpen(false);
    const abort = new AbortController();
    const onProgress = progressFor(abort);
    setJob({ text: "Checking what would change…", abort, progress: 0 });
    try {
      // Preview first: count the changes and show a few examples before anything is written.
      const preview = await cleanSelection(o, { dryRun: true, onProgress, signal: abort.signal });
      if (!preview.cellsChanged) {
        toast("info", "Already clean", `${stripSheet(preview.address)} · ${preview.cellsScanned.toLocaleString()} cells checked – nothing to change.`);
        return;
      }
      setJob(null);
      const examples = (preview.samples ?? []).map((x) => `${x.address}: “${x.before}” → “${x.after}”`).join("\n");
      const ok = await askConfirm({
        title: `Clean ${preview.cellsChanged.toLocaleString()} cell${preview.cellsChanged === 1 ? "" : "s"}?`,
        message: `${stripSheet(preview.address)} · ${preview.cellsScanned.toLocaleString()} cells checked. For example:\n${examples}${preview.cellsScanned > 50_000 ? "\n\nThis range is too large for undo." : "\n\nYou can undo with ↶."}`,
        confirmLabel: "Clean",
      });
      if (!ok) return;
      // Writing can't be cancelled half-way (it would leave a partly cleaned range without undo).
      setJob({ text: "Cleaning…", progress: 0 });
      const r = await cleanSelection(o, { onProgress: (f, t) => setJob({ text: t, progress: f }) });
      toast(
        r.cellsChanged ? "success" : "info",
        r.cellsChanged ? `Cleaned ${r.cellsChanged.toLocaleString()} cells` : "Already clean",
        `${stripSheet(r.address)} · ${r.cellsScanned.toLocaleString()} cells scanned${r.cellsChanged && !r.undoAvailable ? " · undo not available for very large ranges" : ""}`
      );
    } catch (e) {
      if (abort.signal.aborted) toast("info", "Cancelled", "Nothing was changed.");
      else toastError(e);
    } finally {
      setJob(null);
    }
  };

  const onRemoveEmpty = (kind: "rows" | "columns") => async () => {
    const abort = new AbortController();
    const onProgress = progressFor(abort);
    setJob({ text: `Looking for empty ${kind}…`, abort, progress: 0 });
    try {
      // Scan first, show what would be deleted, delete only after the user confirms.
      const plan = await scanEmptyLines(kind, { onProgress, signal: abort.signal });
      setJob(null);
      if (!plan.hits.length) {
        toast("info", `No empty ${kind} found`, `${stripSheet(plan.address)} · ${plan.scanned.toLocaleString()} ${kind} checked.`);
        return;
      }
      const b = plan.block;
      const which = describeLines(plan.hits, kind === "rows" ? (i) => String(b.rowIndex + i + 1) : (i) => columnLetter(b.columnIndex + i));
      const n = plan.hits.length;
      const ok = await askConfirm({
        title: `Remove ${n.toLocaleString()} empty ${n === 1 ? kind.slice(0, -1) : kind}?`,
        message: `On “${plan.sheetName}” (${stripSheet(plan.address)}), ${kind === "rows" ? "rows" : "columns"} ${which} contain only blanks, zeros or N/A-type values and will be deleted.${
          plan.undo ? " You can undo with ↶." : " This sheet is too large for undo – save a copy first if unsure."
        }`,
        confirmLabel: `Remove ${n.toLocaleString()}`,
      });
      if (!ok) return;
      setJob({ text: `Removing ${n.toLocaleString()} empty ${kind}…`, progress: 0 });
      const r = await applyEmptyLines(plan, { onProgress: (f, t) => setJob({ text: t, progress: f }) });
      toast(
        r.removed ? "success" : "info",
        r.removed ? `Removed ${r.removed.toLocaleString()} empty ${kind}` : `No empty ${kind} found`,
        `${stripSheet(r.address)} · ${r.scanned.toLocaleString()} ${kind} checked${r.removed && !r.undoAvailable ? " · undo not available for very large ranges" : ""}`
      );
    } catch (e) {
      if (abort.signal.aborted) toast("info", "Cancelled", "Nothing was changed.");
      else toastError(e);
    } finally {
      setJob(null);
    }
  };

  const onExtract = async (choice: ExtractChoice) => {
    setExtractOpen(false);
    try {
      const src = await readColumnSelection(choice.header, choice.kind === "ai" ? 2000 : 50000);
      let results: string[];
      let label: string;
      if (choice.kind === "ai") {
        if (!ensureConfigured()) return;
        label = choice.instruction.length > 30 ? `${choice.instruction.slice(0, 30)}…` : choice.instruction;
        if (src.values.length > AI_BATCH) {
          const est = estimateTokens(src.values.join("\n")) * 2 + Math.ceil(src.values.length / AI_BATCH) * 300;
          const ok = await askConfirm({
            title: "Run Gemini on every row?",
            message: `${src.values.length.toLocaleString()} rows → ${Math.ceil(src.values.length / AI_BATCH)} requests, roughly ${est.toLocaleString()} tokens. Continue?`,
            confirmLabel: "Run",
          });
          if (!ok) return;
        }
        const abort = new AbortController();
        results = [];
        for (let i = 0; i < src.values.length; i += AI_BATCH) {
          if (abort.signal.aborted) throw new Error("Cancelled.");
          setJob({ text: `Gemini is processing rows ${i + 1}–${Math.min(i + AI_BATCH, src.values.length)} of ${src.values.length}…`, abort });
          const batch = src.values.slice(i, i + AI_BATCH);
          const { data, usage: u } = await generateJson<string[]>({
            settings: settingsRef.current,
            systemInstruction: TRANSFORM_SYSTEM_PROMPT,
            prompt: transformPrompt(choice.instruction, batch),
            schema: { type: "ARRAY", items: { type: "STRING" } },
            signal: abort.signal,
          });
          if (u) addUsage(u);
          const arr = Array.isArray(data) ? data.map((v) => (v === null || v === undefined ? "" : String(v))) : [];
          while (arr.length < batch.length) arr.push("");
          results.push(...arr.slice(0, batch.length));
        }
      } else {
        const kind = choice.kind;
        label = EXTRACT_LABELS[kind];
        results = src.values.map((v) => extractValue(v, kind));
      }
      setJob({ text: "Writing results…" });
      const where = await insertResultColumn(src, results, label);
      const hits = results.filter(Boolean).length;
      toast("success", `${label}: ${hits} of ${results.length} rows`, `New column at ${stripSheet(where)}`);
    } catch (e) {
      if ((e as Error).message !== "Stopped.") toastError(e);
    } finally {
      setJob(null);
    }
  };

  const onExplainFormula = async () => {
    if (!ensureConfigured()) return;
    try {
      const cell = await getActiveCellInfo();
      if (!cell.hasFormula) {
        toast("info", "No formula here", `${stripSheet(cell.address)} contains a value, not a formula. Click a cell with a formula.`);
        return;
      }
      const overview = await getSheetOverview().catch(() => "");
      await send({ display: `Explain the formula in ${stripSheet(cell.address)}`, prompt: explainFormulaPrompt(cell.address, cell.formula, overview), contextLabel: cell.address, forceChat: true });
    } catch (e) {
      toastError(e);
    }
  };

  const onGenerateFormula = async (desc: string) => {
    setGenOpen(false);
    if (!ensureConfigured()) return;
    try {
      const [cell, overview, [sel, label]] = await Promise.all([
        inExcel ? getActiveCellInfo() : Promise.resolve(null),
        inExcel ? getSheetOverview().catch(() => "") : Promise.resolve(""),
        selectionContext(),
      ]);
      const address = cell ? stripSheet(cell.address) : "A1";
      await send({ display: `Generate formula: ${desc}`, prompt: generateFormulaPrompt(desc, address, overview, sel), contextLabel: label ?? address, forceChat: true });
    } catch (e) {
      toastError(e);
    }
  };

  const onAnalyze = async () => {
    if (!ensureConfigured()) return;
    const [ctx, label] = await selectionContext(true);
    if (!ctx) return;
    const ask = "Summarize the selected data: what it contains, key statistics, notable patterns or outliers, and 2–3 useful formulas or charts I could add.";
    await send({ display: "Summarize selection", prompt: ask + ctx, contextLabel: label, forceChat: true });
  };

  /* ------------------------------------------------------------ charts & petroleum tools */
  const onChart = async (kind: ChartKind, title: string) => {
    setChartOpen(false);
    try {
      const addr = await getSelectedAddress();
      const r = await createChart(addr, kind, title || undefined);
      toast("success", "Chart created", r);
    } catch (e) {
      toastError(e);
    }
  };

  const openDeclineCurve = async () => {
    try {
      const { headers } = await getSelectionHeaders();
      if (headers.length < 2) {
        toast("info", "Select time and rate", "Select at least two columns: date/time and rate (with or without a header row).");
        return;
      }
      setDcaHeaders(headers);
    } catch (e) {
      toastError(e);
    }
  };

  const onDeclineCurve = async (c: DcaChoice) => {
    setDcaHeaders(null);
    setJob({ text: c.wellColumn !== undefined ? "Fitting decline curves for every well…" : "Fitting decline curves…" });
    try {
      if (c.wellColumn !== undefined) {
        const w = await runDeclineByWell(c);
        toast(w.failed ? "warning" : "success", `Fitted ${w.fitted} of ${w.wells} wells`, `${w.failed ? `${w.failed} wells had fewer than 3 valid points · ` : ""}summary → sheet "${w.sheet}"`);
        return;
      }
      const r = await runDeclineCurve(c);
      const b = r.best;
      toast(
        "success",
        `Best fit: ${b.model} (R² ${b.r2.toFixed(3)})`,
        `qi ${b.qi.toPrecision(4)}, Di ${b.di.toPrecision(3)}, b ${b.b.toFixed(2)} · ${r.pointsUsed} points${r.dropped ? `, ${r.dropped} skipped` : ""} → sheet "${r.sheet}"`
      );
    } catch (e) {
      toastError(e);
    } finally {
      setJob(null);
    }
  };

  const [espSource, setEspSource] = useState<PlotSource | null>(null);
  const openEspAnalysis = async () => {
    try {
      setEspSource(await getPlotSource());
    } catch (e) {
      toastError(e);
    }
  };
  const onEspAnalysis = async (c: EspDialogChoice) => {
    const src = espSource;
    setEspSource(null);
    if (!src) return;
    setJob({ text: "Analysing ESP data…" });
    try {
      const r = await runEspAnalysis({ source: src, columns: c.columns, options: c.options, curveAddress: c.curveAddress });
      toast(r.status.startsWith("OK") ? "success" : "warning", r.status, `${r.rows} rows analysed, ${r.flagged} flagged → sheet "${r.sheet}"`);
    } catch (e) {
      toastError(e);
    } finally {
      setJob(null);
    }
  };

  const onProductionQc = async () => {
    setJob({ text: "Checking production data…" });
    try {
      const r = await runProductionQc();
      toast(r.issues ? "warning" : "success", r.issues ? `${r.issues} issues found` : "No issues found", `${r.columns} columns checked → sheet "${r.sheet}"`);
    } catch (e) {
      toastError(e);
    } finally {
      setJob(null);
    }
  };

  const onUnits = async (id: string, header: boolean) => {
    setUnitOpen(false);
    const conv = UNIT_CONVERSIONS.find((u) => u.id === id);
    if (!conv) return;
    try {
      const r = await insertConvertedColumns(conv, header);
      toast("success", `Converted ${r.converted.toLocaleString()} values`, `${conv.label} → ${stripSheet(r.address)}`);
    } catch (e) {
      toastError(e);
    }
  };

  const onEspCheck = async () => {
    if (!ensureConfigured()) return;
    const [ctx, label] = await selectionContext(true);
    if (!ctx) return;
    await send({ display: "ESP performance check", prompt: ESP_PROMPT + ctx, contextLabel: label, forceChat: true });
  };

  const openDraw = async () => {
    try {
      setPlotSource(await getPlotSource());
    } catch (e) {
      toastError(e);
    }
  };

  const onDraw = async (c: DrawChoice) => {
    const src = plotSource;
    setPlotSource(null);
    if (!src) return;
    setJob({ text: "Drawing chart…" });
    try {
      const r = await plotColumns({
        source: src,
        xIndex: c.xIndex,
        yIndexes: c.yIndexes,
        chartType: c.chartType,
        title: c.title,
        secondaryYIndexes: c.secondaryYIndexes,
        logLeft: c.logLeft,
        logRight: c.logRight,
      });
      toast(
        r.warnings.length ? "warning" : "success",
        "Chart created",
        `${r.title} on ${src.sheetName}${r.helperSheet ? ` · linked data copied to sheet "${r.helperSheet}" (this chart type needs adjacent columns)` : ""}${r.warnings.length ? ` · Note: ${r.warnings.join("; ")}` : ""}`
      );
    } catch (e) {
      toastError(e);
    } finally {
      setJob(null);
    }
  };


  /* ------------------------------------------------------------ daily report, ESP failure log, audit, validation */
  const toolSel = useRef<HeaderedSelection | null>(null);
  const [dailyTarget, setDailyTarget] = useState<DailyReportTarget | null>(null);
  const openDailyReport = async () => {
    try {
      const sel = await readHeaderedSelection("production data (Date, Well, Oil… columns)");
      toolSel.current = sel;
      setDailyTarget({ headers: sel.headers, datesFor: (col) => datesIn(sel, col) });
    } catch (e) {
      toastError(e);
    }
  };
  const onDailyReport = async (m: DailyMapping, reportDate?: number) => {
    setDailyTarget(null);
    const sel = toolSel.current;
    if (!sel) return;
    setJob({ text: "Building the daily report…" });
    try {
      const r = await runDailyReport(sel, m, reportDate);
      const copied = await copyText(r.packageText);
      toast(
        "success",
        `Daily report ${r.pkg.dateLabel} – ${r.summary.wells.length} wells`,
        `Sheet "${r.sheet}"${r.hasChart ? " with chart" : ""}. Now in Word: Zee AI → Templates → Daily report from Excel${copied ? " (also copied to the clipboard, in case Word asks you to paste it)" : ""}.`
      );
    } catch (e) {
      toastError(e);
    } finally {
      setJob(null);
    }
  };

  const onEspLogTemplate = async () => {
    try {
      const sheet = await createEspFailureLog();
      toast("success", "ESP failure log created", `Sheet "${sheet}": one row per ESP run. Run life fills in by itself; Status, Component and Cause have drop-down lists.`);
    } catch (e) {
      toastError(e);
    }
  };
  const [espLogHeaders, setEspLogHeaders] = useState<string[] | null>(null);
  const openEspLog = async () => {
    try {
      const sel = await readHeaderedSelection("ESP failure log");
      toolSel.current = sel;
      setEspLogHeaders(sel.headers);
    } catch (e) {
      toastError(e);
    }
  };
  const onEspLog = async (m: EspLogMapping) => {
    setEspLogHeaders(null);
    const sel = toolSel.current;
    if (!sel) return;
    setJob({ text: "Analysing the ESP failure log…" });
    try {
      const r = await runEspFailureAnalysis(sel, m);
      const rep = r.report;
      toast("success", `MTBF ${rep.mtbf ?? "–"} days · ${rep.failures} failures in ${rep.installs} runs`, `Report on sheet "${r.sheet}".${rep.skipped ? ` ${rep.skipped} rows skipped (no well or install date).` : ""}`);
      setJob(null);
      if (rep.failures && isConfigured(settingsRef.current)) {
        await send({ display: "ESP failure analysis – summary & recommendations", prompt: `${ESP_FAILURE_PROMPT}\n\n${r.digest}`, contextLabel: `${sel.block.address} → ${r.sheet}`, forceChat: true });
      }
    } catch (e) {
      toastError(e);
    } finally {
      setJob(null);
    }
  };

  const onFormulaAudit = async () => {
    const abort = new AbortController();
    setJob({ text: "Auditing formulas…", abort, progress: 0 });
    try {
      const r = await runFormulaAudit({ onProgress: progressFor(abort), signal: abort.signal });
      setJob(null);
      const parts = Object.entries(r.counts).map(([k, n]) => `${n} ${AUDIT_LABELS[k as keyof typeof AUDIT_LABELS].toLowerCase()}`);
      toast(r.issues.length ? "warning" : "success", r.issues.length ? `${r.issues.length} findings in ${r.formulaCells.toLocaleString()} formulas` : `All ${r.formulaCells.toLocaleString()} formulas look fine`, `${parts.join(" · ") || "No problems"} → sheet "${r.sheet}"`);
      const serious = r.issues.filter((i) => i.severity !== "low");
      if (serious.length && isConfigured(settingsRef.current)) {
        const ok = await askConfirm({
          title: "Ask Zee AI how to fix them?",
          message: `${serious.length} error / inconsistency finding${serious.length === 1 ? "" : "s"} on "${r.source}". Zee AI can explain the likely cause and the corrected formula for each (up to 40).`,
          confirmLabel: "Explain fixes",
        });
        if (ok) {
          const list = serious
            .slice(0, 40)
            .map((i) => `- ${i.address}: ${AUDIT_LABELS[i.kind]} – ${i.detail} Formula: ${i.formula || "(value)"}`)
            .join("\n");
          await send({
            display: `Formula audit of "${r.source}" – how to fix`,
            prompt: `These problems were found by a formula audit of sheet "${r.source}". For each, explain the likely cause in one line and give the corrected formula in an \`\`\`excel block when possible. Group similar ones.\n\n${list}`,
            contextLabel: r.source,
            forceChat: true,
          });
        }
      }
    } catch (e) {
      if (abort.signal.aborted) toast("info", "Cancelled", "Nothing was changed.");
      else toastError(e);
    } finally {
      setJob(null);
    }
  };

  const rulesTarget = useRef<RulesTarget | null>(null);
  const [rules, setRules] = useState<ColumnRule[] | null>(null);
  const onSuggestRules = async () => {
    if (!ensureConfigured()) return;
    const abort = new AbortController();
    setJob({ text: "Profiling the columns…", abort });
    try {
      const t = await profileSelectionForRules();
      rulesTarget.current = t;
      setJob({ text: `Gemini is suggesting rules for ${t.profiles.length} columns…`, abort });
      const { data, usage: u } = await generateJson<{ rules: ColumnRule[] }>({
        settings: settingsRef.current,
        systemInstruction: RULES_SYSTEM_PROMPT,
        prompt: `Columns of ${t.sel.block.address} (${t.sel.rows.length} data rows):\n${profilesForPrompt(t.profiles)}`,
        schema: RULES_SCHEMA,
        signal: abort.signal,
      });
      if (u) addUsage(u);
      setRules(sanitizeRules(data?.rules, t.profiles));
    } catch (e) {
      if (abort.signal.aborted) toast("info", "Cancelled", "Nothing was changed.");
      else toastError(e);
    } finally {
      setJob(null);
    }
  };
  const onApplyRules = async (chosen: ColumnRule[], opts: { highlight: boolean; validate: boolean }) => {
    setRules(null);
    const t = rulesTarget.current;
    if (!t) return;
    setJob({ text: "Checking the columns…" });
    try {
      const res = await applyColumnRules(t, chosen, opts);
      const bad = res.filter((r) => r.violations);
      const total = bad.reduce((n, r) => n + r.violations, 0);
      toast(
        total ? "warning" : "success",
        total ? `${total.toLocaleString()} value${total === 1 ? "" : "s"} break the rules` : "All values pass",
        bad.length
          ? bad.map((r) => `${r.column}: ${r.violations} (${r.examples.join(", ")}${r.violations > r.examples.length ? "…" : ""})`).join(" · ") + (opts.highlight ? " – highlighted in red." : "")
          : `${res.length} column${res.length === 1 ? "" : "s"} checked${opts.validate ? "; new entries are validated" : ""}.`
      );
    } catch (e) {
      toastError(e);
    } finally {
      setJob(null);
    }
  };

  /* ------------------------------------------------------------ =ZEE function picker */
  const [fnDialog, setFnDialog] = useState<FunctionDialogState | null>(null);

  const applyFunction = async (spec: FnSpec, values: Record<string, string>, header: boolean, where: ApplyWhere, target: FunctionTarget | null) => {
    setFnDialog(null);
    try {
      if (spec.ai && !ensureConfigured()) return;
      if (where === "cell" || !target || target.isEmpty || spec.mode === "single") {
        const f = buildFormula(spec, values);
        const addr = await withOverwrite((ow) => insertFormulas([f], { overwrite: ow, label: `insert ZEE.${spec.id}` }));
        if (addr) toast("success", `ZEE.${spec.id} inserted`, `${stripSheet(addr)}  ${f}`);
        return;
      }

      const first = target.rowIndex + (header ? 1 : 0);
      const rows = target.rowCount - (header ? 1 : 0);
      if (rows < 1) throw new Error("The selection only has a header row – select the data rows too.");
      if (rows > 20000) throw new Error("Select at most 20,000 rows at a time.");
      const outCol = target.columnIndex + target.columnCount;
      const headerCell = header ? [`ZEE.${spec.id}`] : [];

      let cells: string[];
      let reserveRows = 0;
      if (spec.mode === "range") {
        if (spec.id === "FLAG" && rows > 1000) throw new Error("ZEE.FLAG checks up to 1,000 rows at a time – select fewer rows.");
        const range = a1Range(first, target.columnIndex, rows, target.columnCount, true);
        cells = [...headerCell, buildFormula(spec, values, [], range)];
        if (spec.id === "FLAG") reserveRows = rows + headerCell.length; // spills one result per row
      } else {
        const n = inputCount(spec);
        const formulas = Array.from({ length: rows }, (_, r) => {
          const refs = Array.from({ length: Math.min(n, target.columnCount) }, (_, k) => a1(first + r, target.columnIndex + k));
          return buildFormula(spec, values, refs);
        });
        if (spec.ai && rows > 50) {
          const ok = await askConfirm({
            title: `Run ZEE.${spec.id} on ${rows.toLocaleString()} rows?`,
            message: `Rows are sent to Gemini in batches of 25 (about ${Math.ceil(rows / 25).toLocaleString()} requests); repeated values and recalculations come from the cache. Continue?`,
            confirmLabel: "Run",
          });
          if (!ok) return;
        }
        cells = [...headerCell, ...formulas];
      }

      const addr = await withOverwrite((ow) =>
        writeFormulaColumn(target.sheetName, target.rowIndex, outCol, cells, {
          overwrite: ow,
          header,
          reserveRows,
          label: `ZEE.${spec.id} column`,
        })
      );
      if (addr) toast("success", `ZEE.${spec.id} applied`, `${stripSheet(addr)} · ${spec.mode === "range" ? "1 formula" : `${rows.toLocaleString()} rows`}`);
    } catch (e) {
      toastError(e);
    }
  };

  const onPickFunction = async (spec: FnSpec) => {
    let target: FunctionTarget | null = null;
    if (inExcel) {
      try {
        target = await getFunctionTarget();
      } catch (e) {
        toastError(e);
        return;
      }
    }
    const header = !!target && !target.isEmpty && guessHeader(target.sample);
    // Nothing to type and data is selected → apply straight away.
    if (target && !target.isEmpty && canApplyDirectly(spec, target.columnCount)) {
      await applyFunction(spec, {}, header, "selection", target);
      return;
    }
    setFnDialog({ spec, target, header });
  };

  const onFreeze = async () => {
    try {
      const r = await freezeSelection();
      toast("success", `Froze ${r.frozen.toLocaleString()} formulas`, `${stripSheet(r.address)} now holds fixed values (undo with ↶).`);
    } catch (e) {
      toastError(e);
    }
  };

  const simple = (fn: () => Promise<string>, title: string) => async () => {
    try {
      const r = await fn();
      toast("success", title, stripSheet(r));
    } catch (e) {
      toastError(e);
    }
  };

  /* ------------------------------------------------------------ render */
  return (
    // Note: layout classes live on an inner div, NOT on FluentProvider. Fluent copies the provider's
    // className onto its portal containers (toasts, dialogs); a full-height white class there
    // covered the whole pane with an empty white layer.
    <FluentProvider theme={dark ? darkTheme : lightTheme} style={{ height: "100%" }}>
      <div className={`zee-app ${dark ? "zee-dark" : "zee-light"}`}>
      <InsertActionsContext.Provider value={insertActions}>
        <header className="zee-header">
          <img src="assets/icon-32.png" alt="" width={22} height={22} />
          <span className="zee-title">Zee AI</span>
          <Tooltip content="Change model in Settings" relationship="description">
            <Badge
              appearance="tint"
              size="small"
              className="zee-model-badge"
              role="button"
              tabIndex={0}
              onClick={() => setSettingsOpen(true)}
              onKeyDown={(e) => (e.key === "Enter" || e.key === " ") && setSettingsOpen(true)}
            >
              {modelLabel(settings.model)}
            </Badge>
          </Tooltip>
          <span className="zee-spacer" />
          <Tooltip content={undoLabel ? `Undo: ${undoLabel}${undoDepth > 1 ? ` (${undoDepth} steps available)` : ""}` : "Nothing to undo"} relationship="label">
            <Button appearance="subtle" size="small" icon={<ArrowUndo20Regular />} disabled={!undoLabel} onClick={undo}>
              {undoDepth > 1 ? undoDepth : undefined}
            </Button>
          </Tooltip>
          <Menu onOpenChange={(_, d) => d.open && setHistory(getChangeHistory())}>
            <MenuTrigger disableButtonEnhancement>
              <Tooltip content="Change history – revert one of Zee AI's recent changes" relationship="label">
                <Button appearance="subtle" size="small" icon={<History20Regular />} disabled={!undoDepth} />
              </Tooltip>
            </MenuTrigger>
            <MenuPopover className="zee-history-popover">
              <MenuList>
                <MenuGroupHeader>Zee AI's recent changes (newest first)</MenuGroupHeader>
                {history.map((h) => (
                  <Menu key={`${h.index}-${h.at}`}>
                    <MenuTrigger disableButtonEnhancement>
                      <MenuItem secondaryContent={timeAgo(h.at)}>{h.label}</MenuItem>
                    </MenuTrigger>
                    <MenuPopover>
                      <MenuList>
                        <MenuItem icon={<ArrowUndo16Regular />} disabled={!h.alone} onClick={() => revertOne(h)}>
                          {h.alone ? "Revert just this change" : "Revert just this (newer changes on the same sheet – use the option below)"}
                        </MenuItem>
                        <MenuItem onClick={() => undoTo(h)}>
                          Undo back to here ({undoDepth - h.index} change{undoDepth - h.index === 1 ? "" : "s"})
                        </MenuItem>
                      </MenuList>
                    </MenuPopover>
                  </Menu>
                ))}
                <MenuDivider />
                <MenuItem disabled>Only changes made in this session ({history.length} of max 25)</MenuItem>
              </MenuList>
            </MenuPopover>
          </Menu>
          <Tooltip content={dark ? "Switch to light mode" : "Switch to dark mode"} relationship="label">
            <Button
              appearance="subtle"
              size="small"
              icon={dark ? <WeatherSunny20Regular /> : <WeatherMoon20Regular />}
              onClick={() => updateSettings({ theme: dark ? "light" : "dark" })}
            />
          </Tooltip>
          <Menu>
            <MenuTrigger disableButtonEnhancement>
              <Tooltip content="Export chat" relationship="label">
                <Button appearance="subtle" size="small" icon={<ArrowDownload20Regular />} disabled={!messages.length} />
              </Tooltip>
            </MenuTrigger>
            <MenuPopover>
              <MenuList>
                <MenuItem icon={<Copy20Regular />} onClick={exportCopy}>
                  Copy chat as text
                </MenuItem>
                <MenuItem icon={<Table20Regular />} onClick={exportSheet} disabled={!inExcel}>
                  Export chat to a new sheet
                </MenuItem>
              </MenuList>
            </MenuPopover>
          </Menu>
          <Tooltip content="New chat" relationship="label">
            <Button appearance="subtle" size="small" icon={<Add20Regular />} onClick={clear} disabled={!messages.length} />
          </Tooltip>
          <Tooltip content="Settings" relationship="label">
            <Button appearance="subtle" size="small" icon={<Settings20Regular />} onClick={() => setSettingsOpen(true)} />
          </Tooltip>
        </header>

        <UpdateBanners newerVersion={updates.newerVersion} notes={updates.notes} onDismissNotes={updates.dismissNotes} />
        {!offline.online && (
          <MessageBar intent="warning" className="zee-banner">
            <MessageBarBody>
              You're offline. {offline.queued ? `${offline.queued} request${offline.queued === 1 ? "" : "s"} will be sent` : "Requests that fail are sent"} again automatically when the connection is back.
            </MessageBarBody>
          </MessageBar>
        )}

        {!inExcel && (
          <MessageBar intent="warning" className="zee-banner">
            <MessageBarBody>Preview mode – open Zee AI from Excel to read and write cells. Chat still works.</MessageBarBody>
          </MessageBar>
        )}
        {inExcel && !isConfigured(settings) && (
          <MessageBar intent="info" className="zee-banner">
            <MessageBarBody>Connect a Gemini API key or Google account to start.</MessageBarBody>
            <MessageBarActions>
              <Button size="small" onClick={() => setSettingsOpen(true)}>
                Set up
              </Button>
            </MessageBarActions>
          </MessageBar>
        )}

        <QuickActions
          disabled={busy || !!job}
          onAggregate={onAggregate}
          onClean={() => setCleanOpen(true)}
          onExtract={() => setExtractOpen(true)}
          onFormatTable={simple(formatSelectionAsTable, "Formatted as table")}
          onAutofit={simple(autofitAllColumns, "All columns auto-fitted")}
          onRemoveEmptyRows={onRemoveEmpty("rows")}
          onRemoveEmptyColumns={onRemoveEmpty("columns")}
          onExplainFormula={onExplainFormula}
          onGenerateFormula={() => setGenOpen(true)}
          onAnalyze={onAnalyze}
          onChart={() => setChartOpen(true)}
          onDeclineCurve={openDeclineCurve}
          onProductionQc={onProductionQc}
          onUnits={() => setUnitOpen(true)}
          onEspCheck={onEspCheck}
          onEspAnalysis={openEspAnalysis}
          onDraw={openDraw}
          onPickFunction={onPickFunction}
          onFreeze={onFreeze}
          onDailyReport={openDailyReport}
          onEspLogTemplate={onEspLogTemplate}
          onEspLogAnalyze={openEspLog}
          onFormulaAudit={onFormulaAudit}
          onSuggestRules={onSuggestRules}
        />

        {job && (
          <MessageBar intent="info" className="zee-banner" layout={job.progress !== undefined ? "multiline" : "auto"}>
            <MessageBarBody>
              <Spinner size="extra-tiny" /> {job.text}
              {job.progress !== undefined && <ProgressBar className="zee-job-progress" value={job.progress} max={1} thickness="large" />}
            </MessageBarBody>
            {job.abort && (
              <MessageBarActions>
                <Button size="small" onClick={() => job.abort?.abort()}>
                  Cancel
                </Button>
              </MessageBarActions>
            )}
          </MessageBar>
        )}

        <ChatFeed messages={messages} onSuggestion={sendChat} />

        <ChatInput
          busy={busy}
          agentMode={settings.agentMode}
          onToggleAgent={(v) => updateSettings({ agentMode: v })}
          includeSelection={settings.includeSelection}
          selectedAddress={selectedAddress}
          onToggleSelection={(v) => updateSettings({ includeSelection: v })}
          ignoreSelection={ignoreSelection}
          onIgnoreSelection={setIgnoreSelection}
          today={today}
          prices={prices}
          inject={inject}
          onSend={sendChat}
          onStop={stop}
        />

        {settingsOpen && (
          <SettingsDrawer
            open={settingsOpen}
            settings={settings}
            usage={usage}
            onChange={updateSettings}
            onImport={(patch) => {
              updateSettings({ ...patch, theme: ["auto", "light", "dark"].includes(String(patch.theme ?? settings.theme)) ? patch.theme ?? settings.theme : "auto" });
              toast("success", "Settings imported", "Check Connection – the API key is only included when it was exported with it.");
            }}
            onClose={() => setSettingsOpen(false)}
          />
        )}
        <CleanDialog open={cleanOpen} onClose={() => setCleanOpen(false)} onRun={onClean} />
        <ExtractDialog open={extractOpen} onClose={() => setExtractOpen(false)} onRun={onExtract} />
        <GenerateFormulaDialog open={genOpen} onClose={() => setGenOpen(false)} onRun={onGenerateFormula} />
        <ChartDialog open={chartOpen} onClose={() => setChartOpen(false)} onRun={onChart} />
        <UnitDialog open={unitOpen} onClose={() => setUnitOpen(false)} onRun={onUnits} />
        <DeclineCurveDialog open={!!dcaHeaders} headers={dcaHeaders ?? []} onClose={() => setDcaHeaders(null)} onRun={onDeclineCurve} />
        <EspDialog
          source={espSource}
          onClose={() => setEspSource(null)}
          onRun={onEspAnalysis}
          onPickSelection={() => getSelectedAddress().catch(() => null)}
        />
        <DrawDialog source={plotSource} onClose={() => setPlotSource(null)} onRun={onDraw} />
        <FunctionDialog state={fnDialog} onClose={() => setFnDialog(null)} onApply={(spec, values, header, where) => applyFunction(spec, values, header, where, fnDialog?.target ?? null)} />
        <DailyReportDialog target={dailyTarget} onClose={() => setDailyTarget(null)} onRun={onDailyReport} />
        <EspLogDialog headers={espLogHeaders} onClose={() => setEspLogHeaders(null)} onRun={onEspLog} />
        <RulesDialog rules={rules} onClose={() => setRules(null)} onApply={onApplyRules} />
        <ConfirmDialog state={confirmState} onResult={onConfirmResult} />
        <Toaster toasterId={toasterId} position="top-end" />
      </InsertActionsContext.Provider>
      </div>
    </FluentProvider>
  );
}
