/** "New version online – Reload" and "What's new" banners (shown at the top of the pane). */
import { useCallback, useEffect, useState } from "react";
import { Button, MessageBar, MessageBarActions, MessageBarBody, MessageBarTitle } from "@fluentui/react-components";
import { BUILD, VERSION, fetchLiveBuild, isNewerBuild, markVersionSeen, notesToShow, readLastSeenVersion, reloadLatest } from "./version";
import { APP, type ReleaseNote } from "./app";
import { readText } from "./store";

const hasStoredSettings = () => readText(APP.settingsKey) !== "";

const CHECK_EVERY_MS = 30 * 60_000;
const FOCUS_THROTTLE_MS = 5 * 60_000;

/** Polls build.json (on load, every 30 min and when the pane regains focus) and tracks release notes. */
export function useUpdates() {
  const [newer, setNewer] = useState<string | null>(null);
  const [notes, setNotes] = useState<ReleaseNote[]>(() => {
    let last = readLastSeenVersion();
    if (!last) {
      // No record yet: a brand-new install has nothing to announce; an existing user is updating
      // from a version that predates "What's new", so show the current notes.
      if (!hasStoredSettings()) {
        markVersionSeen();
        return [];
      }
      last = "0";
    }
    return notesToShow(last, VERSION);
  });

  useEffect(() => {
    if (!BUILD) return; // dev server
    let alive = true;
    let lastCheck = 0;
    const check = async () => {
      if (Date.now() - lastCheck < FOCUS_THROTTLE_MS) return;
      lastCheck = Date.now();
      const live = await fetchLiveBuild();
      if (alive && live && isNewerBuild(live)) setNewer(live.version);
    };
    void check();
    const timer = window.setInterval(() => {
      lastCheck = 0;
      void check();
    }, CHECK_EVERY_MS);
    const onFocus = () => void check();
    window.addEventListener("focus", onFocus);
    return () => {
      alive = false;
      window.clearInterval(timer);
      window.removeEventListener("focus", onFocus);
    };
  }, []);

  const dismissNotes = useCallback(() => {
    markVersionSeen();
    setNotes([]);
  }, []);

  // Seeing a version without notes (e.g. a patch release) still counts as seen.
  useEffect(() => {
    if (!notes.length && readLastSeenVersion() !== VERSION) markVersionSeen();
  }, [notes.length]);

  return { newerVersion: newer, notes, dismissNotes };
}

export function UpdateBanners({ newerVersion, notes, onDismissNotes }: { newerVersion: string | null; notes: ReleaseNote[]; onDismissNotes: () => void }) {
  const [reloading, setReloading] = useState(false);
  const [expanded, setExpanded] = useState(false);
  const items = notes.flatMap((n) => n.items);
  const shown = expanded ? items : items.slice(0, 3);
  return (
    <>
      {newerVersion && (
        <MessageBar intent="info" className="zee-banner">
          <MessageBarBody>
            A newer version{newerVersion !== VERSION ? ` (${newerVersion})` : ""} is online.
          </MessageBarBody>
          <MessageBarActions>
            <Button
              size="small"
              appearance="primary"
              disabled={reloading}
              onClick={() => {
                setReloading(true);
                void reloadLatest();
              }}
            >
              {reloading ? "Reloading…" : "Reload now"}
            </Button>
          </MessageBarActions>
        </MessageBar>
      )}
      {items.length > 0 && (
        <MessageBar intent="success" className="zee-banner zee-whatsnew" layout="multiline">
          <MessageBarBody>
            <MessageBarTitle>Updated to {VERSION} – what’s new</MessageBarTitle>
            <ul className="zee-whatsnew-list">
              {shown.map((t, i) => (
                <li key={i}>{t}</li>
              ))}
            </ul>
            {items.length > 3 && (
              <Button size="small" appearance="transparent" onClick={() => setExpanded((v) => !v)}>
                {expanded ? "Show less" : `Show all ${items.length}`}
              </Button>
            )}
          </MessageBarBody>
          <MessageBarActions>
            <Button size="small" onClick={onDismissNotes}>
              Got it
            </Button>
          </MessageBarActions>
        </MessageBar>
      )}
    </>
  );
}
