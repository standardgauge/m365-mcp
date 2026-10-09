import { TableClient, TableEntity, odata } from '@azure/data-tables';
import { getTableClient as getStorageTableClient } from './storageClient.js';

const GLOBAL_TABLE = 'GlobalDenyList';
const USER_TABLE = 'UserDenyList';

export type DenyListType = 'sharepoint' | 'mail' | 'calendar' | 'onedrive' | 'onenote' | 'contacts' | 'teams';

export interface DenyListEntry {
  partitionKey: string;
  rowKey: string;
  path: string;
  description: string;
  addedBy: string;
  addedByName?: string;
  addedAt: string;
}

function getTableClient(tableName: string): TableClient {
  return getStorageTableClient(tableName);
}

/** Creates the table if it doesn't already exist (409 = already exists, safe to ignore). */
async function ensureTable(tableName: string): Promise<void> {
  const client = getTableClient(tableName);
  try {
    await client.createTable();
  } catch (err: unknown) {
    const status = (err as { statusCode?: number }).statusCode;
    if (status !== 409) throw err;
  }
}

// ── Encoding helpers ──────────────────────────────────────────────────────────

/** Azure Table Storage row keys cannot contain certain characters; base64-encode paths. */
function encodeRowKey(path: string): string {
  return Buffer.from(path).toString('base64').replace(/=/g, '').replace(/\+/g, '-').replace(/\//g, '_');
}

// ── Global deny list (admin-managed) ─────────────────────────────────────────

export async function addGlobalDenyEntry(
  tenantId: string,
  type: DenyListType,
  path: string,
  addedBy: string,
  description = '',
  addedByName?: string
): Promise<void> {
  await ensureTable(GLOBAL_TABLE);
  const client = getTableClient(GLOBAL_TABLE);
  const pk = `${tenantId}:${type}`;
  const entity: TableEntity<DenyListEntry> = {
    partitionKey: pk,
    rowKey: encodeRowKey(path),
    path,
    description,
    addedBy,
    addedByName: addedByName ?? '',
    addedAt: new Date().toISOString(),
  };
  await client.upsertEntity(entity, 'Replace');
}

export async function removeGlobalDenyEntry(
  tenantId: string,
  type: DenyListType,
  path: string
): Promise<void> {
  await ensureTable(GLOBAL_TABLE);
  const client = getTableClient(GLOBAL_TABLE);
  const pk = `${tenantId}:${type}`;
  await client.deleteEntity(pk, encodeRowKey(path));
}

export async function listGlobalDenyEntries(
  tenantId: string,
  type: DenyListType
): Promise<DenyListEntry[]> {
  await ensureTable(GLOBAL_TABLE);
  const client = getTableClient(GLOBAL_TABLE);
  const pk = `${tenantId}:${type}`;
  const results: DenyListEntry[] = [];
  const entities = client.listEntities<DenyListEntry>({
    queryOptions: { filter: odata`PartitionKey eq ${pk}` },
  });
  for await (const entity of entities) {
    results.push(entity as DenyListEntry);
  }
  return results;
}

// ── Per-user deny list ────────────────────────────────────────────────────────

/** Partition key format: `{userId}:{type}` */
function userPk(userId: string, type: DenyListType): string {
  return `${userId}:${type}`;
}

export async function addUserDenyEntry(
  userId: string,
  type: DenyListType,
  path: string,
  addedByName?: string
): Promise<void> {
  await ensureTable(USER_TABLE);
  const client = getTableClient(USER_TABLE);
  const entity: TableEntity<DenyListEntry> = {
    partitionKey: userPk(userId, type),
    rowKey: encodeRowKey(path),
    path,
    description: '',
    addedBy: userId,
    addedByName: addedByName ?? '',
    addedAt: new Date().toISOString(),
  };
  await client.upsertEntity(entity, 'Replace');
}

export async function removeUserDenyEntry(
  userId: string,
  type: DenyListType,
  path: string
): Promise<void> {
  await ensureTable(USER_TABLE);
  const client = getTableClient(USER_TABLE);
  await client.deleteEntity(userPk(userId, type), encodeRowKey(path));
}

export async function listUserDenyEntries(
  userId: string,
  type: DenyListType
): Promise<DenyListEntry[]> {
  await ensureTable(USER_TABLE);
  const client = getTableClient(USER_TABLE);
  const pk = userPk(userId, type);
  const results: DenyListEntry[] = [];
  const entities = client.listEntities<DenyListEntry>({
    queryOptions: { filter: odata`PartitionKey eq ${pk}` },
  });
  for await (const entity of entities) {
    results.push(entity as DenyListEntry);
  }
  return results;
}

export async function clearUserDenyList(
  userId: string,
  type: DenyListType
): Promise<void> {
  const entries = await listUserDenyEntries(userId, type);
  if (entries.length === 0) return;
  const client = getTableClient(USER_TABLE);
  await Promise.all(entries.map((e) => client.deleteEntity(e.partitionKey, e.rowKey)));
}

// ── Config-driven default deny lists ──────────────────────────────────────────

/**
 * Per-type env var holding a comma-separated list of default denied paths that
 * are enforced on top of the admin-managed table entries.
 *
 * These exist so a deployment can guarantee sensitive containers are
 * unreachable *before* an admin has seeded the table, and so that clearing the
 * table can never un-deny them. They are enforced identically to global/user
 * entries — same case-insensitive prefix matching — but live in deployment
 * config rather than storage, which makes them tamper-resistant.
 *
 * Unset → no defaults for that type (the canonical multi-tenant behaviour: a
 * fresh install denies nothing until an admin configures it). The Example deploy
 * sets `DEFAULT_MAIL_DENY_FOLDERS="Finance,HR,Legal,IR,Management"` on the
 * Container App (see infra/main.bicep).
 */
const DEFAULT_DENY_ENV_VARS: Partial<Record<DenyListType, string>> = {
  mail: 'DEFAULT_MAIL_DENY_FOLDERS',
  sharepoint: 'DEFAULT_SHAREPOINT_DENY_PATHS',
};

/**
 * Returns the deployment-configured default denied paths for `type`, parsed
 * from the matching env var. Empty when the var is unset or blank. Read fresh
 * on each call so config changes (and tests) take effect without a restart.
 */
export function getDefaultDenyPaths(type: DenyListType): string[] {
  // `type` is a fixed union and the keys/values below are hard-coded literals,
  // so neither lookup is attacker-controlled.
  // eslint-disable-next-line security/detect-object-injection
  const envVar = DEFAULT_DENY_ENV_VARS[type];
  if (!envVar) return [];
  // eslint-disable-next-line security/detect-object-injection
  const raw = process.env[envVar];
  if (!raw) return [];
  return raw
    .split(',')
    .map((p) => p.trim())
    .filter((p) => p.length > 0);
}

// ── Combined path-check helpers ───────────────────────────────────────────────

/**
 * Whose per-user (tier 2) deny lists apply to an operation. Usually the
 * signed-in caller alone. When the caller reaches another user's mailbox or
 * calendar through delegation, it is the caller *and* the owner of that data
 * (see `resolveDenySubject` in mailboxOwner.ts): an owner who hides a folder
 * from AI hides it from every agent, not only their own.
 */
export type DenySubject = string | readonly string[];

/** Default + global + every subject's per-user entries, as a flat path list. */
async function loadDeniedPaths(
  tenantId: string,
  subject: DenySubject,
  type: DenyListType
): Promise<string[]> {
  const userIds = typeof subject === 'string' ? [subject] : [...new Set(subject)];
  const [globalEntries, ...userEntryLists] = await Promise.all([
    listGlobalDenyEntries(tenantId, type),
    ...userIds.map((id) => listUserDenyEntries(id, type)),
  ]);
  return [
    ...getDefaultDenyPaths(type),
    ...[...globalEntries, ...userEntryLists.flat()].map((e) => e.path),
  ];
}

/**
 * Strips the Graph API drive prefix so paths from two different sources can be
 * compared on equal footing.
 *
 * Graph parentReference.path values look like:
 *   /drives/<driveId>/root:/Documents/SubFolder
 * Drive-relative user inputs (args.path for write tools) look like:
 *   Documents/SubFolder  or  /Documents/SubFolder
 *
 * Both reduce to /Documents/SubFolder so deny-list entries recorded in Graph
 * format correctly block write operations that supply a drive-relative path.
 * Paths that already lack the Graph prefix (mail folder names, plain SP paths)
 * are returned unchanged after a leading-slash is added if missing.
 *
 * Unicode NFC normalization is applied so that a deny entry and a runtime path
 * that are canonically equivalent but encoded differently (e.g. a precomposed
 * "é" vs "e" + combining accent) still match. Without it a folder could appear
 * blocked in the admin UI yet never match at runtime (F14).
 */
export function canonicalizePath(path: string): string {
  const normalized = path.normalize('NFC');
  const rootMarker = '/root:';
  const idx = normalized.indexOf(rootMarker);
  if (idx !== -1) {
    const rest = normalized.slice(idx + rootMarker.length);
    return rest.startsWith('/') ? rest : `/${rest}`;
  }
  return normalized.startsWith('/') ? normalized : `/${normalized}`;
}

/**
 * Returns true if `path` matches or is under any denied path.
 * Both sides are canonicalized before comparison to handle the mismatch
 * between Graph parentReference.path format and drive-relative write args.
 * Comparison is case-insensitive to prevent bypass via case manipulation
 * (SharePoint, OneDrive, and Exchange paths are all case-insensitive).
 */
function matchesDenyList(path: string, deniedPaths: string[]): boolean {
  const normalizedPath = canonicalizePath(path).toLowerCase();
  return deniedPaths.some((denied) => {
    const normalizedDenied = canonicalizePath(denied).toLowerCase();
    if (normalizedPath === normalizedDenied) return true;
    const prefix = normalizedDenied.endsWith('/') ? normalizedDenied : `${normalizedDenied}/`;
    return normalizedPath.startsWith(prefix);
  });
}

/**
 * Checks both the global and per-user deny list for `path`.
 * Returns true if the path is blocked for this user — or, when `userId` names
 * several subjects (delegated access), for any of them.
 */
export async function isPathDenied(
  tenantId: string,
  userId: DenySubject,
  type: DenyListType,
  path: string
): Promise<boolean> {
  try {
    const allDenied = await loadDeniedPaths(tenantId, userId, type);
    return matchesDenyList(path, allDenied);
  } catch (err: unknown) {
    const message = err instanceof Error ? err.message : String(err);
    console.error('[denyList] Storage unavailable — failing closed (denying access):', message);
    return true; // Fail closed: treat as denied when storage is down
  }
}

/**
 * True if `path` sits inside a folder that a deny entry names — either by
 * prefix / exact match (the ordinary {@link matchesDenyList} rule) or by a
 * denied folder appearing as an *ancestor segment* anywhere in the path.
 *
 * Search hits carry a full site-relative path (e.g.
 * `/Shared Documents/Finance/deal.xlsx`), so a deny entry recorded as a bare
 * folder name (`Finance`) — or as a path that omits the document-library
 * segment (`/Finance`) — would never match by prefix alone. Matching each
 * denied entry's segment run against the path's ancestor segments closes that
 * gap, so a denied folder is honored regardless of the entry's recorded shape.
 * Segment comparison is exact per component, so `IR` blocks a folder literally
 * named `IR` but not `Investor IR Notes` — no accidental over-blocking.
 */
function matchesSearchDenyList(path: string, deniedPaths: string[]): boolean {
  if (matchesDenyList(path, deniedPaths)) return true;
  const segs = canonicalizePath(path).toLowerCase().split('/').filter(Boolean);
  return deniedPaths.some((denied) => {
    const dsegs = canonicalizePath(denied).toLowerCase().split('/').filter(Boolean);
    if (dsegs.length === 0) return false;
    for (let i = 0; i + dsegs.length <= segs.length; i++) {
      let run = true;
      for (let j = 0; j < dsegs.length; j++) {
        // Numeric loop indices into local string arrays — not attacker-keyed.
        // eslint-disable-next-line security/detect-object-injection
        if (segs[i + j] !== dsegs[j]) { run = false; break; }
      }
      // Only an ANCESTOR match denies here (there is a segment after the run);
      // an exact leaf match is already handled by matchesDenyList above.
      if (run && i + dsegs.length < segs.length) return true;
    }
    return false;
  });
}

/**
 * Deny-list filter for Microsoft Search hits (`filterDeniedPaths`'s sibling for
 * the search surface). Unlike folder listings, search returns items from deep
 * inside the tree, so a hit under a denied folder must be dropped even when the
 * deny entry names only the folder (bare name or partial path). Uses
 * {@link matchesSearchDenyList} for that ancestor-segment matching while drawing
 * on the same default + global + per-user deny sources as
 * {@link filterDeniedPaths}. Callers must set each item's `path` to a resolved
 * site-relative path; items with no matchable path are the caller's to drop
 * (fail closed) before this runs. Fails closed (drops all) when storage is down.
 */
export async function filterDeniedSearchHits<
  T extends { path?: string; webUrl?: string; id?: string; name?: string }
>(tenantId: string, userId: string, items: T[]): Promise<T[]> {
  try {
    const [globalEntries, userEntries] = await Promise.all([
      listGlobalDenyEntries(tenantId, 'sharepoint'),
      listUserDenyEntries(userId, 'sharepoint'),
    ]);
    const allDenied = [
      ...getDefaultDenyPaths('sharepoint'),
      ...[...globalEntries, ...userEntries].map((e) => e.path),
    ];
    return items.filter((item) => {
      const primary = item.path ?? item.webUrl ?? item.id ?? '';
      if (primary && matchesSearchDenyList(primary, allDenied)) return false;
      if (item.name && matchesDenyList(item.name, allDenied)) return false;
      return true;
    });
  } catch (err: unknown) {
    const message = err instanceof Error ? err.message : String(err);
    console.error('[denyList] Storage unavailable — failing closed (denying all search hits):', message);
    return []; // Fail closed: deny all rather than allow all when storage is down
  }
}

/**
 * Filters `items` to exclude anything whose `path` or `webUrl` is blocked by
 * the two-tier deny list.
 */
export async function filterDeniedPaths<
  T extends { path?: string; webUrl?: string; id?: string; name?: string }
>(tenantId: string, userId: DenySubject, type: DenyListType, items: T[]): Promise<T[]> {
  try {
    const allDenied = await loadDeniedPaths(tenantId, userId, type);

    return items.filter((item) => {
      // Exclude when the primary identifier (path/webUrl/id) OR the display
      // name is denied. Name matching lets an admin block a shared/sensitive
      // calendar or folder by its human-readable name rather
      // than its opaque Graph ID.
      const primary = item.path ?? item.webUrl ?? item.id ?? '';
      if (primary && matchesDenyList(primary, allDenied)) return false;
      if (item.name && matchesDenyList(item.name, allDenied)) return false;
      return true;
    });
  } catch (err: unknown) {
    const message = err instanceof Error ? err.message : String(err);
    console.error('[denyList] Storage unavailable — failing closed (denying all access):', message);
    return []; // Fail closed: deny all rather than allow all when storage is down
  }
}
