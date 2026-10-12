import React, { useEffect, useState } from 'react';

// Mirrors ReadinessReport in src/services/readiness.ts.
interface ReadinessCheck {
  id: string;
  setting: string;
  value: string | null;
  enforced: boolean;
  title: string;
  open: string;
}

interface ReadinessReport {
  ready: boolean;
  checks: ReadinessCheck[];
}

const RUNBOOK_SECTION = 'Credential envelope row binding';

/**
 * Post-deploy operator checks, shown at the top of the Admin view and never
 * collapsed. The runbook tells the operator to close the legacy credential
 * envelope window after a deploy; this is the one place that reports whether
 * that has happened.
 */
export default function ReadinessChecks() {
  const [report, setReport] = useState<ReadinessReport | null>(null);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    let cancelled = false;
    fetch('/api/manage/readiness', { credentials: 'same-origin' })
      .then(async (r) => {
        if (!r.ok) {
          const data = (await r.json().catch(() => ({}))) as { error?: string };
          throw new Error(data.error ?? `HTTP ${r.status}`);
        }
        return (await r.json()) as ReadinessReport;
      })
      .then((data) => {
        if (!cancelled) setReport(data);
      })
      .catch((err: unknown) => {
        if (!cancelled) setError(err instanceof Error ? err.message : String(err));
      });
    return () => { cancelled = true; };
  }, []);

  const open = report !== null && !report.ready;

  return (
    <div className={`card readiness-card${open ? ' readiness-card--open' : ''}`}>
      <h2>Readiness</h2>
      {error && <div className="error-msg">Readiness could not be read: {error}</div>}
      {!report && !error && <div className="empty-state">Checking…</div>}
      {report && (
        <>
          <p className="section-hint">
            {report.ready
              ? 'Every post-deploy check is enforced. The legacy credential envelope window is closed.'
              : `The legacy credential envelope window is still open. Finish the procedure under "${RUNBOOK_SECTION}" in docs/operations-runbook.md, then set the variables it names.`}
          </p>
          <ul className="readiness-list">
            {report.checks.map((check) => (
              <li
                key={check.id}
                className={`readiness-check ${check.enforced ? 'readiness-check--ok' : 'readiness-check--open'}`}
              >
                <span className="readiness-status">{check.enforced ? 'Enforced' : 'Open'}</span>
                <div className="readiness-info">
                  <span className="readiness-title">{check.title}</span>
                  <span className="readiness-desc">
                    <code>{check.setting}</code>
                    {check.value === null ? ' is unset. ' : <> is set to <code>{check.value}</code>. </>}
                    {check.enforced ? 'Legacy envelopes are refused.' : check.open}
                  </span>
                </div>
              </li>
            ))}
          </ul>
        </>
      )}
    </div>
  );
}
