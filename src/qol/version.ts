/** Version info, "What's new" after an update, "new version online" check and "Reload latest version". */
import { APP, RELEASE_NOTES, type ReleaseNote } from "./app";
import { readText, writeText } from "./store";

export const VERSION: string = typeof __ZEE_VERSION__ !== "undefined" ? __ZEE_VERSION__ : "dev";
export const BUILD: string = typeof __ZEE_BUILD__ !== "undefined" ? __ZEE_BUILD__ : "";

const SEEN_KEY = `${APP.storagePrefix}.lastSeenVersion`;

/** -1 / 0 / 1 for dotted version numbers ("1.10.0" > "1.9.2"). Non-numeric parts count as 0. */
export function compareVersions(a: string, b: string): number {
  const pa = a.split(".").map((x) => parseInt(x, 10) || 0);
  const pb = b.split(".").map((x) => parseInt(x, 10) || 0);
  for (let i = 0; i < Math.max(pa.length, pb.length); i++) {
    const d = (pa[i] ?? 0) - (pb[i] ?? 0);
    if (d) return d > 0 ? 1 : -1;
  }
  return 0;
}

/**
 * Release notes to show after an update: every note newer than `lastSeen` up to `current`.
 * Nothing on a first install (lastSeen empty) or when the version didn't change.
 */
export function notesToShow(lastSeen: string, current: string, notes: ReleaseNote[] = RELEASE_NOTES): ReleaseNote[] {
  if (!lastSeen || compareVersions(current, lastSeen) <= 0) return [];
  return notes.filter((n) => compareVersions(n.version, lastSeen) > 0 && compareVersions(n.version, current) <= 0);
}

export function readLastSeenVersion(): string {
  return readText(SEEN_KEY);
}

export function markVersionSeen(v: string = VERSION): void {
  writeText(SEEN_KEY, v);
}

export interface LiveBuild {
  version: string;
  build: string;
}

/** True when the build online is newer than the one running (both are ISO timestamps). */
export function isNewerBuild(live: LiveBuild | null, runningBuild: string = BUILD): boolean {
  if (!live || !runningBuild) return false;
  const a = Date.parse(live.build);
  const b = Date.parse(runningBuild);
  return Number.isFinite(a) && Number.isFinite(b) && a - b > 60_000;
}

/** Reads build.json next to the page (written by vite.config.ts at build time). null when unavailable (dev server, offline). */
export async function fetchLiveBuild(signal?: AbortSignal): Promise<LiveBuild | null> {
  try {
    const url = new URL("build.json", window.location.href);
    url.searchParams.set("t", Date.now().toString(36));
    const res = await fetch(url.toString(), { cache: "no-store", signal });
    if (!res.ok) return null;
    const j = (await res.json()) as Partial<LiveBuild>;
    return typeof j.build === "string" && typeof j.version === "string" ? { version: j.version, build: j.build } : null;
  } catch {
    return null;
  }
}

/**
 * Loads the newest published version: clears any Cache Storage entries, then reloads the page with
 * a cache-busting parameter (the code files have hashed names, so a fresh page pulls fresh code).
 */
export async function reloadLatest(): Promise<void> {
  try {
    if (typeof caches !== "undefined") {
      const keys = await caches.keys();
      await Promise.all(keys.map((k) => caches.delete(k)));
    }
  } catch {
    /* not available in this webview */
  }
  try {
    const regs = (await navigator.serviceWorker?.getRegistrations?.()) ?? [];
    await Promise.all(regs.map((r) => r.unregister()));
  } catch {
    /* none */
  }
  const url = new URL(window.location.href);
  url.searchParams.set("zeev", Date.now().toString(36));
  window.location.replace(url.toString());
}

/** "13 Oct 2026, 14:51" for the running build (empty in dev). */
export function buildLabel(build: string = BUILD): string {
  const t = Date.parse(build);
  return Number.isFinite(t) ? new Date(t).toLocaleString() : "";
}
