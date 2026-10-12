import { isLegacyEnvelopeReadAllowed, isLegacyIdentityBindingAllowed } from './credentialCrypto.js';

/**
 * Operator readiness checks: configuration the runbook tells an operator to
 * finish after a deploy, reported so that the admin console notices when it
 * has not happened. Each check is a documented operator action with no other
 * surface watching it, which is how a short migration window becomes
 * permanent.
 *
 * `enforced` comes from the same helpers the storage layer consults, so the
 * report says what the server does, not what the variable looks like. `value`
 * is the raw setting so a near miss (`yes`, `1`, `True `) is visible next to
 * the verdict.
 */
export interface ReadinessCheck {
  /** Stable identifier for the check. */
  id: 'envelope-row-binding' | 'session-identity-binding';
  /** Environment variable that closes the check. */
  setting: string;
  /** Raw configured value, or null when the variable is unset. */
  value: string | null;
  /** Whether the server is enforcing the stricter behaviour. */
  enforced: boolean;
  /** Short name for the check. */
  title: string;
  /** What the server does while the check is open. */
  open: string;
}

export interface ReadinessReport {
  /** True only when every check is enforced. */
  ready: boolean;
  checks: ReadinessCheck[];
}

function rawSetting(value: string | undefined): string | null {
  return value === undefined ? null : value;
}

/** Collect the readiness checks against the current process environment. */
export function collectReadinessChecks(): ReadinessReport {
  const checks: ReadinessCheck[] = [
    {
      id: 'envelope-row-binding',
      setting: 'MCP_ENVELOPE_REQUIRE_AAD',
      value: rawSetting(process.env.MCP_ENVELOPE_REQUIRE_AAD),
      enforced: !isLegacyEnvelopeReadAllowed(),
      title: 'Credential envelope row binding',
      open: 'Credential envelopes written before row binding are still readable, so a storage writer can move one between rows (threat model row 2.2).',
    },
    {
      id: 'session-identity-binding',
      setting: 'MCP_SESSION_REQUIRE_IDENTITY_BINDING',
      value: rawSetting(process.env.MCP_SESSION_REQUIRE_IDENTITY_BINDING),
      enforced: !isLegacyIdentityBindingAllowed(),
      title: 'Session identity binding',
      open: 'Session envelopes bound only to their row are still readable, so a storage writer can edit the identity columns of a row that has not been rebound (threat model row 2.3).',
    },
  ];
  return { ready: checks.every((c) => c.enforced), checks };
}
