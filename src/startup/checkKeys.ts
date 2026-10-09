import { cryptoKeyProblems } from '../services/credentialCrypto.js';

/**
 * Container startup gate for the application keys.
 *
 * credentialCrypto reads MCP_SESSION_HMAC_KEY and MCP_DATA_ENCRYPTION_KEY on
 * first use, so an instance missing either one used to boot, pass /health, and
 * fail on the first session write. The image's CMD runs this before the
 * Functions host (see Dockerfile): a bad key exits non-zero, the replica never
 * becomes ready, and the new revision fails to provision while the previous
 * one keeps serving.
 *
 * Lives outside src/functions so the Functions host never loads it as an entry
 * point, and so tests can call it without spawning a process.
 */
export function runStartupKeyCheck(
  env: NodeJS.ProcessEnv,
  log: (line: string) => void,
): number {
  const problems = cryptoKeyProblems(env);
  if (problems.length === 0) {
    log('[startup] MCP_SESSION_HMAC_KEY and MCP_DATA_ENCRYPTION_KEY validated');
    return 0;
  }
  for (const problem of problems) log(`[startup] FATAL: ${problem}`);
  log('[startup] Refusing to start. Generate a key with `openssl rand -hex 32`; see "Application keys" in docs/operations-runbook.md.');
  return 1;
}

if (require.main === module) {
  process.exitCode = runStartupKeyCheck(process.env, (line) => console.error(line));
}
