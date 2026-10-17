import { useMemo } from "react";
import { liveFeatureReport } from "./requirements";

/** Settings → "This Excel": which Zee AI features this Excel version can run. */
export function ExcelSupportSection() {
  const report = useMemo(() => liveFeatureReport(), []);
  return (
    <section>
      <h3>This Excel</h3>
      {!report ? (
        <p className="zee-muted">Open Zee AI inside Excel to see which features your Excel version supports.</p>
      ) : (
        <>
          <p className="zee-muted">
            Office API level: <b>ExcelApi {report.highest ?? "unknown"}</b>. Features your Excel can't run are greyed out in the menus (and hidden from the agent) instead of failing.
          </p>
          <ul className="zee-feature-list">
            {report.features.map((f) => (
              <li key={f.id} title={`${f.set} ${f.version}`}>
                <span aria-hidden>{f.ok ? "✅" : "⛔"}</span>
                <span>
                  {f.label}
                  {!f.ok && <span className="zee-muted"> – needs {f.set} {f.version} (Microsoft 365 / Excel on the web)</span>}
                </span>
              </li>
            ))}
          </ul>
        </>
      )}
    </section>
  );
}
