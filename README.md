# Zee AI – Gemini copilot and agent for Excel

Zee AI is an Excel task-pane add-in (Windows, Mac, Excel on the web) built with React, TypeScript, Fluent UI v9 and Office.js, powered by Google Gemini.

## Features

| | |
|---|---|
| **Chat** | Streaming answers with Markdown and code highlighting. It sees your selection (with real cell addresses) and, optionally, a map of every sheet (names, headers, tables, named ranges). |
| **Agent mode** | Toggle **Agent** above the input box. Gemini can read ranges, write values/formulas, format, create charts, tables and sheets, and fit decline curves. **Every change is listed for your approval** (with a preview of the values) and can be undone. |
| **Multi-step undo** | ↶ reverts Zee AI's changes one step at a time (up to 25). It covers cells, inserted columns, new sheets, charts and tables. Formatting is not covered. |
| **`=ZEE.ASK()`** | Worksheet function, for example `=ZEE.ASK("Classify this failure as electrical, mechanical or reservoir", B2)`. Fill it down a column. Answers are cached and calls are throttled. |
| **Insert buttons** | Formulas in chat get **Insert at cursor**, and tables get **Insert table**. You're asked before anything is overwritten. |
| **Math** | AutoSum-style Sum, Average, Min, Max, Count, Count non-empty, Count unique. |
| **Clean** | Trim, remove hidden characters, convert numbers stored as text, fix case, fix dates. Also Split / Extract (names, emails, phones, or a custom AI instruction per row), Format as table, Auto-fit. |
| **Analyze** | Explain the formula in the active cell, generate a formula, summarize the selection, **chart from selection**. |
| **Petroleum** | **Decline curve analysis**: Arps exponential, harmonic and hyperbolic fit, a forecast, time-to-limit and EUR, and a chart on a new sheet. **Production data QC**: missing, negative, outlier and stuck-gauge values, and date gaps or duplicates, reported on a new sheet. **Unit conversion**: bbl↔m³, Mscf/d→m³/d, psi↔bar/kPa, °F↔°C, ft↔m, °API↔SG, hp↔kW. **ESP performance check**: an AI diagnosis covering TDH, affinity laws, operating range, gas interference and more. |
| **Per-workbook history** | The chat is saved for each workbook and comes back when you reopen it. |
| **Cost guard** | Asks before sending a very large prompt, and before running AI on many rows. It also shows token usage per message and per session. |
| **Auth** | Gemini API key, Google sign-in (Vertex AI quota), or a **team proxy** that keeps the key on a server (`proxy/`). |

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
  hooks/useChat.ts           Chat + agent loop (tool calls, approvals), cost guard, history
  services/
    agentService.ts          Agent tools (declarations, descriptions, execution)
    excelService.ts          Office.js helpers: context, writes, undo stack, charts, tables, sheets
    geminiService.ts         REST client: streaming, JSON mode, tool calling, retries, proxy/Vertex/key
    petroActions.ts          Decline-curve and QC report sheets, ESP prompt
    chatStore.ts             Per-workbook chat history
    settingsService.ts / authService.ts / prompts.ts
  functions/zeeFunctions.ts  =ZEE.ASK() implementation (cache + throttling)
  utils/                     Pure, unit-tested helpers (A1, clean, extract, markdown, SSE, petro)
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
