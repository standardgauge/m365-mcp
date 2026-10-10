import { TableClient, TableEntity } from '@azure/data-tables';
import { getTableClient as getStorageTableClient } from './storageClient.js';

const TABLE = 'serviceSettings';
const ROW_KEY = 'enabledServices';

const DEFAULT_SERVICES = ['mail', 'sharepoint'];

// Full list of known service keys (for reference/validation)
export const ALL_SERVICE_KEYS = ['mail', 'sharepoint', 'calendar', 'onedrive', 'onenote', 'contacts', 'teams'];

function getTableClient(): TableClient {
  return getStorageTableClient(TABLE);
}

async function ensureTable(): Promise<void> {
  const client = getTableClient();
  try {
    await client.createTable();
  } catch (err: unknown) {
    const status = (err as { statusCode?: number }).statusCode;
    if (status !== 409) throw err;
  }
}

export async function getEnabledServices(tenantId: string): Promise<string[]> {
  await ensureTable();
  const client = getTableClient();
  try {
    const entity = await client.getEntity<{ services: string }>(tenantId, ROW_KEY);
    return JSON.parse(entity.services) as string[];
  } catch (err: unknown) {
    const status = (err as { statusCode?: number }).statusCode;
    if (status === 404) return [...DEFAULT_SERVICES];
    throw err;
  }
}

export async function setEnabledServices(tenantId: string, services: string[]): Promise<void> {
  await ensureTable();
  const client = getTableClient();
  const entity: TableEntity = {
    partitionKey: tenantId,
    rowKey: ROW_KEY,
    services: JSON.stringify(services),
  };
  await client.upsertEntity(entity, 'Replace');
}

// ── Read-only services ─────────────────────────────────────────
// A service listed here stays enabled for reads but has all write/mutating
// operations blocked (e.g. calendar in read-only mode: list/get events work,
// create/update/delete are refused). Opt-in per tenant; default is empty so
// existing tenants are unaffected. Enforced in policyEnforcement (HTTP routes)
// and mcpEndpoint (MCP tools/call) so neither surface is a bypass.
const READ_ONLY_ROW_KEY = 'readOnlyServices';

export async function getReadOnlyServices(tenantId: string): Promise<string[]> {
  await ensureTable();
  const client = getTableClient();
  try {
    const entity = await client.getEntity<{ services: string }>(tenantId, READ_ONLY_ROW_KEY);
    return JSON.parse(entity.services) as string[];
  } catch (err: unknown) {
    const status = (err as { statusCode?: number }).statusCode;
    if (status === 404) return [];
    throw err;
  }
}

export async function setReadOnlyServices(tenantId: string, services: string[]): Promise<void> {
  await ensureTable();
  const client = getTableClient();
  const entity: TableEntity = {
    partitionKey: tenantId,
    rowKey: READ_ONLY_ROW_KEY,
    services: JSON.stringify(services),
  };
  await client.upsertEntity(entity, 'Replace');
}

const ALLOWED_SITES_ROW_KEY = 'allowedSites';

export async function getAllowedSites(tenantId: string): Promise<Array<{ id: string; name: string }>> {
  await ensureTable();
  const client = getTableClient();
  try {
    const entity = await client.getEntity<{ sites: string }>(tenantId, ALLOWED_SITES_ROW_KEY);
    return JSON.parse(entity.sites) as Array<{ id: string; name: string }>;
  } catch (err: unknown) {
    const status = (err as { statusCode?: number }).statusCode;
    if (status === 404) return [];
    throw err;
  }
}

export async function setAllowedSites(tenantId: string, sites: Array<{ id: string; name: string }>): Promise<void> {
  await ensureTable();
  const client = getTableClient();
  const entity: TableEntity = {
    partitionKey: tenantId,
    rowKey: ALLOWED_SITES_ROW_KEY,
    sites: JSON.stringify(sites),
  };
  await client.upsertEntity(entity, 'Replace');
}
