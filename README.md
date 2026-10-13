# Zee AI – Gemini copilot and agent for Excel

Zee AI is an Excel task-pane add-in (Windows, Mac, Excel on the web) built with React, TypeScript, Fluent UI v9 and Office.js, powered by Google Gemini.

## Features

| | |
|---|---|
| **Chat** | Streaming answers with Markdown and code highlighting. It sees your selection (with real cell addresses) and, optionally, a map of every sheet (names, headers, tables, named ranges). |
| **Agent mode** | Toggle **Agent** above the input box. Gemini can read and find cells, write values/formulas, format, insert/delete rows and columns, clear, copy, sort, filter, find & replace, freeze panes, add conditional formatting, data validation (drop-downs), named ranges, PivotTables, charts, tables and sheets. Petroleum tools: decline-curve fit, production QC report, unit-conversion columns and ESP analysis. **Every change is listed for your approval** (with a preview of the values) and can be undone. You see each step live while it works; if it reaches its step limit, click **Continue**. |
| **Attachments** | Click 📎, paste a screenshot (Ctrl+V) or drop files on the input: images (PNG, JPG, WebP, HEIC) and PDFs, up to 5 files / 14 MB per message. Ask e.g. "digitise this pump curve into a table" or "extract the well-test results". Files are sent with that message only and are not saved in the chat history. |
| **Web search & Python** | Settings → **Gemini tools**. *Search the web* grounds answers in Google Search and lists the sources (in agent mode it becomes a `web_search` tool). *Run Python* lets Gemini calculate with code; the code and its output appear in the answer. Both are off by default. |
| **Multi-step undo** | ↶ reverts Zee AI's changes one step at a time (up to 25). It covers cells, formatting (ranges up to 20,000 cells), inserted/deleted rows and columns, sorts, find & replace, conditional formats, validation, names, freeze panes, filters, new sheets, charts and tables. |
| **Functions drop list** | Click **Functions** in the toolbar. Every `=ZEE.*` function is listed with a one-line description. If cells with data are selected, clicking a function fills a new column of formulas to the right (one per row). Otherwise it inserts the formula into the active cell. When a function needs input, such as categories or units, a small form asks for it and previews the formula. **Freeze results** turns formulas into fixed values so they stop calling Gemini. |
| **AI worksheet functions** | `ZEE.ASK`, `ZEE.EXTRACT`, `ZEE.CLASSIFY`, `ZEE.CLEAN`, `ZEE.TRANSLATE`, `ZEE.SUMMARIZE`, `ZEE.FLAG` (spills one flag per row), `ZEE.FORMULA`, and `ZEE.EXPLAIN`. They avoid spamming Gemini: answers are cached, filled-down cells are batched 25 rows per request with at most 4 requests at a time, and empty cells are never sent. |
| **Field functions (no AI, free)** | `ZEE.UNITS(value, "from", "to")` converts psi, bar, kPa, bbl/d, m3/d, Mscf/d, °F/°C, API/SG, psi/ft, hp/kW and more. The others are `ZEE.CONVERT`, `VOGEL`, `VOGEL_QMAX`, `PI`, `RS`, `PB`, `BO`, `ARPS`, `ARPS_CUM`, `TDH`, `HHP`, `AFFINITY` and `SG_MIX`. |
| **Insert buttons** | Formulas in chat get **Insert at cursor**, and tables get **Insert table**. You're asked before anything is overwritten. |
| **Math** | AutoSum-style Sum, Average, Min, Max, Count, Count non-empty, Count unique. |
| **Clean** | Trim, remove hidden characters, convert numbers stored as text, fix case, fix dates. Also Split / Extract (names, emails, phones, or a custom AI instruction per row), Format as table. **Remove empty columns** and **Remove empty rows** delete columns/rows that are entirely blank, 0 or NULL/N/A/NaN (scans the whole sheet automatically – no selection needed; a header cell is ignored for columns; undoable). **Auto-fit all columns** fits every column on the sheet, and Zee AI auto-fits all columns after every write. |
| **Analyze** | Explain the formula in the active cell, generate a formula, summarize the selection, **chart from selection**. |
| **Petroleum** | **Decline curve analysis**: Arps exponential, harmonic and hyperbolic fit, a forecast, time-to-limit and EUR, and a chart on a new sheet. **Production data QC**: missing, negative, outlier and stuck-gauge values, and date gaps or duplicates, reported on a new sheet. **Unit conversion**: bbl↔m³, Mscf/d→m³/d, psi↔bar/kPa, °F↔°C, ft↔m, °API↔SG, hp↔kW. **ESP performance check**: an AI diagnosis covering TDH, affinity laws, operating range, gas interference and more. |
| **Per-workbook history** | The chat is saved for each workbook and comes back when you reopen it. When browser storage fills up, the oldest messages are dropped first. |
| **Cost guard** | Asks before sending a very large prompt, and before running AI on many rows. It also shows token usage per message and per session. |
| **Auth** | Gemini API key, Google sign-in (Vertex AI quota), or a **team proxy** that keeps the key on a server (`proxy/`), with per-person tokens, an origin lock, a daily request cap and a size limit. |
| **Previews & progress** | Clean and Remove empty rows / columns show what will change ("Remove 412 empty rows: 5, 9, 12–20 and 3 more?") before anything happens. Long scans show a progress bar with **Cancel**. |
| **Change history** | The clock icon next to ↶ lists Zee AI's recent changes with times. Revert just one (when no newer change touched that sheet) or undo back to any point. |
| **Excel tables from chat** | Tables in answers: **Insert table** at the selected cell, or **Insert as Excel table** (filters + banded rows, undoable). |
| **Quality of life** | Selection chip (e.g. `Sheet1!A1:F200`) with × to leave it out of the next message, ↑/↓ recalls earlier prompts, favourite prompts as chips, the unsent draft survives closing the pane, copy / edit & resend / regenerate on every message, Esc stops, friendly errors with Retry + Copy diagnostics, today's token usage (and cost if you enter prices), "What's new" after updates, a banner when a newer version is online, and Settings → Backup & updates (export / import settings, Reload latest version). |
| **update-all.bat** | Builds, tests and publishes the Excel **and** Word add-ins in one go, waits until GitHub Pages serves both new versions, then clears Office's add-in cache. |
| **Daily production report** | Petroleum → Daily production report: select production data (Date, Well, Oil, optional Water / Gas). A new sheet gets per-well rates, change vs each well's previous reading, water cut, GOR, totals, highlights (drops ≥20%, water-cut jumps, zero rates, missing wells) and an oil-by-well chart. The result is handed to Zee AI for Word (Templates → Daily report from Excel), which writes the report text and can insert the chart. |
| **ESP failure log** | Petroleum → ESP failure log → *Create a failure log sheet* (table with drop-downs for status, component and cause; run life fills itself), or *Analyze failure log…*: run life, MTBF (= run days / failures), failures by cause / component / model / well, repeat-failure wells, monthly trend and charts on a new sheet, followed by an AI summary with recommendations. Also available to the agent (`esp_failure_analysis`). |
| **Formula audit** | Analyze → Formula audit checks every formula on the sheet: error values, `#REF!`, formulas that differ from the rest of their column, typed values in formula columns and hard-coded numbers. Report sheet with clickable cell links; optionally Zee AI explains the fixes. Agent tool `formula_audit`. |
| **AI data validation** | Analyze → Suggest data validation: Zee profiles each column and proposes engineering limits (water cut 0–100, rates ≥ 0, Hz 30–70…) or allowed lists. Review / edit them, then highlight the values that break them (live conditional format) and optionally add Excel data validation. Every rule is undoable. |
| **Offline queue** | Requests that fail because the connection dropped are sent again automatically when it comes back. |
| **Duplicates & outliers** | Clean → Find duplicates & outliers…: on the selected table, finds duplicate rows (case/space-insensitive), spellings of the same value that differ (`Well-12`, `W 12`, `well12`, Arabic digits) and outliers (robust z-score). Choose to highlight them (yellow / blue / orange), unify spellings to the most common one, and/or delete duplicates (first kept). Each is an undo step. |
| **Fill gaps in time series** | Clean → Fill gaps in time series…: pick the date column and optionally the well column. Zee detects daily / weekly / monthly spacing, adds missing dates per well and fills numbers by straight line or previous value (or adds dates only). Gaps longer than a limit (default 31 steps) are left alone. The result goes to a new sheet with added rows in blue and filled cells in yellow, plus a "Zee: filled" note column. |
| **Fix imported data** | Clean → Fix imported data…: per column, converts numbers stored as text (thousands separators, decimal comma, (negatives), Arabic digits), moves a unit stuck to numbers into the header (`Oil` → `Oil (bbl)`), turns mixed date texts into real dates (day-first unless the data proves otherwise) and tidies text. Shows the counts and examples first; formulas are never touched; undoable. |
| **Explain this sheet / sheet guide** | Analyze → Explain this sheet: Zee maps every formula (input vs calculated columns, formula patterns, what depends on what, links to other sheets and names) and the AI explains the purpose, inputs, calculation flow, outputs and risks. *Add a sheet guide tab* writes the same map to a "Guide" sheet – no AI needed. |
| **Hard-coded numbers → inputs** | Analyze → Move hard-coded numbers to inputs…: lists the numbers typed inside formulas (5.615, 0.433, 141.5… get oilfield names like `FT3_PER_BBL`). Tick and rename them; Zee creates an **Inputs** sheet with named cells and rewrites the formulas (`=B2*5.615` → `=B2*FT3_PER_BBL`). One undo step. Also offered at the end of a formula audit. |
| **Agent dry run** | The approval card shows the real cells that will change – current value struck through, new value in green – and the content that a clear / delete would remove. Settings → Agent mode → *Dry run* (on by default). |
| **Big-data mode** | When the selection has at least 100,000 rows (Settings → Sheet context, adjustable or off), Zee sends statistics over **every** row (counts, blanks, min/max/mean, date span, top values) plus ~50 sample rows instead of a truncated block. The message shows "big-data mode". The agent can do the same on any range with `profile_range`. |

---

## Option A – Hosted on GitHub Pages (recommended, no local server)

1. Install **Git for Windows** if you don't have it: <https://git-scm.com/download/win>.
2. Double-click **`publish-github.bat`**. It will:
   - ask for your GitHub username and open GitHub so you can create an empty public repo named `zee-ai`
   - generate `manifest.prod.xml` for `https://<you>.github.io/zee-ai`
   - push the code (a GitHub sign-in window may appear)
   - open the repo's **Settings → Pages**: set **Source = GitHub Actions**
3. Wait for the **Actions** tab to show a green deploy (about 2 minutes).
4. Double-click **`use-online.bat`**. It checks that the site is up and registers the hosted manifest with Excel.
5. Close every Excel window, reopen Excel, and click **Home → Zee AI**.

To update later, change the code and run `publish-github.bat` again. Every push redeploys automatically.

> The repo must be public on a free GitHub plan. The code contains no keys: API keys are typed into Settings and stored only on your PC.

## Option B – Local development (`https://localhost:3100`)

```bat
setup.bat        :: first time: npm install + localhost certificate + build + open Excel
serve-zee.bat    :: every time: start the local server (keep its window open). It also renews the certificate.
start-zee.bat    :: rebuild after code changes, restart the server and open Excel
diagnose.bat     :: writes diag2-log.txt if something doesn't load
```

---

## Connect Gemini

Open **⚙ Settings** in the pane.

- **API key**: paste a key from <https://aistudio.google.com/apikey> and click **Test connection**.
- **Google account (Vertex AI)**:
  1. Enable the Vertex AI API.
  2. Create a *Web application* OAuth client with JavaScript origin = your add-in origin (`https://localhost:3100` or `https://<you>.github.io`) and redirect URI = `<add-in URL>/auth/callback.html`.
  3. Give your account the *Vertex AI User* role, then sign in.
- **Team proxy**: deploy `proxy/worker.js` (see `proxy/README.md`), then paste its URL and the team token.

Models: Gemini 3.8 Flash (default), 3.1 Pro preview, 3.5 Flash-Lite, and 2.5 Pro/Flash (legacy). Use ⟳ to load all models your key can use.

## Project structure

```
manifest.xml                 Dev manifest (localhost:3100), shared runtime + =ZEE.* custom functions
scripts/make-manifest.mjs    Generates manifest.prod.xml for a hosted URL
.github/workflows/deploy.yml Builds and deploys to GitHub Pages on push
public/                      icons, OAuth dialog pages, functions.json (custom-function metadata)
proxy/                       Optional Cloudflare Worker team proxy
src/
  App.tsx                    UI orchestration, quick actions, dialogs, confirmations
  hooks/useChat.ts           Chat + agent wiring (approvals, streaming, Continue), cost guard, history
  services/
    agentLoop.ts             The agent loop (model step → tools → approval → repeat), retries; unit-tested
    agentService.ts          Agent tools (declarations, descriptions, execution)
    workbookOps.ts           Agent workbook operations: rows/columns, sort, filter, find/replace, CF, validation, names, PivotTables
    excelService.ts          Office.js helpers: context, writes, undo stack (incl. formatting), charts, tables, sheets
    geminiService.ts         REST client: streaming, JSON mode, tool calling, retries, proxy/Vertex/key
    petroActions.ts          Decline-curve and QC report sheets, ESP prompt
    chatStore.ts             Per-workbook chat history
    settingsService.ts / authService.ts / prompts.ts
  functions/zeeFunctions.ts  =ZEE.ASK() implementation (cache + throttling)
  utils/                     Pure, unit-tested helpers (A1, clean, extract, markdown, SSE, petro, agent arguments, attachments)
```

## Troubleshooting

| Symptom | Fix |
|---|---|
| "We can't load the add-in" (local) | Run `serve-zee.bat` and keep it open, then close and reopen Excel. |
| "We can't load the add-in" (hosted) | Check that the Actions deploy succeeded and `https://<you>.github.io/zee-ai/taskpane.html` opens in a browser. |
| `=ZEE.ASK()` shows `#NAME?` | The add-in hasn't loaded in this workbook yet: open the Zee AI pane once, then press Ctrl+Alt+F9. |
| `=ZEE.ASK()` shows `#N/A` | Hover over the cell to see the message (usually "connect Gemini in Settings" or a quota error). |
| Agent made an unwanted change | Click ↶ (repeatedly, for several steps). |
| Port 3000/3100 in use | Zee AI uses 3100. `diagnose.bat` shows which process holds a port. |
| Old version keeps loading | Close Excel and delete `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`. |
