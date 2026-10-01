/**
 * Scope/documentation drift guard.
 *
 * `GRAPH_SCOPES` in src/services/graphClient.ts is the exact set of scopes MSAL
 * names on every auth-code exchange and silent refresh. Per the README's own
 * rule, *everything in that array must be consented before first sign-in* — a
 * scope requested but not granted fails authentication for the whole tenant.
 *
 * A new tenant learns which permissions to consent to from the README permission
 * table, not from the source. So if a scope lives in `GRAPH_SCOPES` but not in
 * the table, the documented consent set is incomplete and a by-the-book operator
 * onboards a broken instance. That is exactly what happened with
 * `Directory.Read.All`: it sat in `GRAPH_SCOPES` (backing Global Admin detection
 * via `/me/transitiveMemberOf`) while appearing in none of the tenant docs.
 *
 * This test is the liveness check on that write path. It reads the real source
 * and the real README and fails CI if they disagree, so the two lists cannot
 * silently drift again. It intentionally reads files from disk rather than
 * importing constants, because the README is the artifact that can rot.
 */
import { readFileSync } from 'fs';
import { join } from 'path';

const repoRoot = join(__dirname, '..', '..');
const graphClientSrc = readFileSync(join(repoRoot, 'src', 'services', 'graphClient.ts'), 'utf8');
const readme = readFileSync(join(repoRoot, 'README.md'), 'utf8');

/** The scope strings inside the `export const GRAPH_SCOPES = [ ... ]` literal. */
function parseGraphScopes(src: string): string[] {
  const m = src.match(/export const GRAPH_SCOPES\s*=\s*\[([\s\S]*?)\]/);
  if (!m) {
    throw new Error(
      'Could not locate the `export const GRAPH_SCOPES = [ ... ]` array in graphClient.ts. ' +
        'If it was renamed or restructured, update scopeDocDrift.test.ts to match.'
    );
  }
  return [...m[1].matchAll(/['"]([^'"]+)['"]/g)].map((x) => x[1]);
}

/**
 * The scope names in the first column of the README "Delegated Graph
 * permissions" table. The table is a superset of GRAPH_SCOPES — it also lists
 * admin-consent-only, tool-specific scopes that never appear in the request.
 */
function parseDocumentedTableScopes(md: string): Set<string> {
  const lines = md.split('\n');
  const headerIdx = lines.findIndex((l) => /^\|\s*Permission\s*\|\s*Backs\s*\|/.test(l));
  if (headerIdx === -1) {
    throw new Error(
      'Could not find the "| Permission | Backs |" table header in README.md. ' +
        'If the permissions table moved or was reformatted, update scopeDocDrift.test.ts.'
    );
  }
  const scopes = new Set<string>();
  // Skip the header row and the |---|---| separator; read data rows until the
  // table ends (first line that is not a table row).
  for (let j = headerIdx + 2; j < lines.length; j++) {
    const line = lines[j];
    if (!line.trimStart().startsWith('|')) break;
    const firstCell = line.split('|')[1] ?? '';
    for (const tok of firstCell.matchAll(/`([^`]+)`/g)) scopes.add(tok[1]);
  }
  return scopes;
}

/**
 * The scopes named in the prose "At the time of writing `GRAPH_SCOPES`
 * contains ..." sentence — a second hand-maintained copy of the array.
 */
function parseProseScopes(md: string): Set<string> {
  const anchor = 'At the time of writing `GRAPH_SCOPES` contains';
  const start = md.indexOf(anchor);
  if (start === -1) {
    throw new Error(
      `Could not find the "${anchor} ..." sentence in README.md. ` +
        'If it was reworded, update scopeDocDrift.test.ts (or the sentence).'
    );
  }
  // Scope names contain dots but never ". " (dot-space), so the first ". "
  // reliably marks the end of the enumeration sentence.
  const rest = md.slice(start);
  const end = rest.indexOf('. ');
  const sentence = end === -1 ? rest : rest.slice(0, end);
  const scopes = new Set<string>();
  for (const tok of sentence.matchAll(/`([^`]+)`/g)) {
    if (tok[1] !== 'GRAPH_SCOPES') scopes.add(tok[1]);
  }
  return scopes;
}

describe('GRAPH_SCOPES ↔ documentation drift', () => {
  const graphScopes = parseGraphScopes(graphClientSrc);

  it('parses a non-empty GRAPH_SCOPES array from the source', () => {
    expect(graphScopes.length).toBeGreaterThan(0);
    expect(graphScopes).toContain('Directory.Read.All');
  });

  it('documents every GRAPH_SCOPES entry in the README permission table', () => {
    const documented = parseDocumentedTableScopes(readme);
    const missing = graphScopes.filter((s) => !documented.has(s));
    expect(missing).toEqual([]);
    if (missing.length) {
      throw new Error(
        `These scopes are requested by MSAL (GRAPH_SCOPES) but missing from the ` +
          `README permission table, so a new tenant would not know to consent to them: ${missing.join(', ')}`
      );
    }
  });

  it('keeps the README prose enumeration in sync with GRAPH_SCOPES', () => {
    const prose = parseProseScopes(readme);
    // Every array entry must be named in the sentence, and the sentence must
    // name nothing extra — both directions, since either can drift.
    const missingFromProse = graphScopes.filter((s) => !prose.has(s));
    const extraInProse = [...prose].filter((s) => !graphScopes.includes(s));
    expect({ missingFromProse, extraInProse }).toEqual({ missingFromProse: [], extraInProse: [] });
  });
});
