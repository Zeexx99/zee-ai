import React, { Component, type ReactNode } from "react";
import { createRoot } from "react-dom/client";
import App from "./App";
import "./styles.css";
import { registerCustomFunctions } from "./functions/zeeFunctions";

declare global {
  interface Window {
    __zeeLog?: (type: string, msg: string) => void;
    __zeeShowError?: (msg: string) => void;
  }
}
const log = (type: string, msg: string) => window.__zeeLog?.(type, msg);

/** Shows render-time crashes instead of a blank pane. */
class ErrorBoundary extends Component<{ children: ReactNode }, { error: Error | null }> {
  state = { error: null as Error | null };
  static getDerivedStateFromError(error: Error) {
    return { error };
  }
  componentDidCatch(error: Error, info: { componentStack?: string | null }) {
    log("react-error", `${error.stack ?? error.message}\n${info.componentStack ?? ""}`);
  }
  render() {
    if (this.state.error) {
      return (
        <pre style={{ whiteSpace: "pre-wrap", color: "#b42318", padding: 12, font: "12px Consolas, monospace" }}>
          Zee AI crashed while rendering:{"\n"}
          {this.state.error.stack ?? this.state.error.message}
        </pre>
      );
    }
    return this.props.children;
  }
}

let rendered = false;
function render(reason: string) {
  if (rendered) return;
  rendered = true;
  const el = document.getElementById("root");
  if (!el) return;
  log("render", `rendering (${reason}); host=${typeof Office !== "undefined" ? String(Office.context?.host) : "no-office"}`);
  createRoot(el).render(
    <React.StrictMode>
      <ErrorBoundary>
        <App />
      </ErrorBoundary>
    </React.StrictMode>
  );
}

log("main", "main.tsx executing");
// Office.onReady resolves both inside Office and in a plain browser (host = null),
// so the UI can also be previewed at https://localhost:3100/taskpane.html.
if (typeof Office !== "undefined") {
  Office.onReady(() => {
    registerCustomFunctions(); // =ZEE.ASK() – shared runtime, so the pane and the functions share settings
    render("Office.onReady");
  });
  // Safety net: never leave the pane blank if Office.onReady stalls.
  setTimeout(() => {
    if (!rendered) log("warn", "Office.onReady did not resolve within 8s – rendering anyway");
    render("timeout");
  }, 8000);
} else {
  log("warn", "office.js not available");
  render("no-office");
}
