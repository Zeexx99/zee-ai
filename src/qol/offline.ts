/**
 * Offline queue: answers that failed with a network error during this session are sent again
 * automatically when the connection comes back (the browser's "online" event). Each failed request is
 * retried once per reconnection, so a connection that is "up" but blocked doesn't cause a retry loop.
 */
import { useEffect, useRef, useState } from "react";

interface MsgLike {
  id: string;
  status: string;
  errorKind?: string;
}

/** Failed-with-network-error answers among `live` ids (pure – unit tested). */
export function networkFailures(messages: MsgLike[], live: Set<string>): string[] {
  return messages.filter((m) => m.status === "error" && m.errorKind === "network" && live.has(m.id)).map((m) => m.id);
}

export function useOfflineQueue(messages: MsgLike[], busy: boolean, retry: (id: string) => void) {
  const [online, setOnline] = useState(() => (typeof navigator === "undefined" ? true : navigator.onLine !== false));
  /** Answers that were generated in this session (not loaded from history). */
  const live = useRef(new Set<string>());
  /** Snapshot of the queue taken when the connection came back. */
  const toRetry = useRef<string[]>([]);
  const [draining, setDraining] = useState(0);

  useEffect(() => {
    messages.forEach((m) => m.status === "streaming" && live.current.add(m.id));
  }, [messages]);

  const queued = networkFailures(messages, live.current);
  const queuedRef = useRef(queued);
  queuedRef.current = queued;

  useEffect(() => {
    const on = () => {
      setOnline(true);
      toRetry.current = [...queuedRef.current];
      if (toRetry.current.length) setDraining((n) => n + 1);
    };
    const off = () => setOnline(false);
    window.addEventListener("online", on);
    window.addEventListener("offline", off);
    return () => {
      window.removeEventListener("online", on);
      window.removeEventListener("offline", off);
    };
  }, []);

  useEffect(() => {
    if (!draining || busy || !online) return;
    const next = toRetry.current.find((id) => queuedRef.current.includes(id));
    if (!next) {
      toRetry.current = [];
      return;
    }
    // Give the network a moment to settle after reconnecting.
    const t = window.setTimeout(() => {
      toRetry.current = toRetry.current.filter((id) => id !== next);
      live.current.delete(next);
      retry(next);
    }, 1500);
    return () => window.clearTimeout(t);
  }, [draining, busy, online, retry, queued.join("|")]);

  return { online, queued: queued.length };
}
