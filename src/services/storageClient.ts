import { TableClient, TableServiceClient } from '@azure/data-tables';
import type { TokenCredential } from '@azure/core-auth';
import { CachedTokenCredential, ContainerAppManagedIdentityCredential } from './managedIdentity.js';

/**
 * One place that decides how the server reaches Table Storage.
 *
 *   AZURE_STORAGE_TABLE_ENDPOINT      https://<account>.table.core.windows.net
 *   AZURE_STORAGE_IDENTITY_CLIENT_ID  client id of the user-assigned identity
 *                                     that holds a data role on the account;
 *                                     unset uses the system-assigned identity
 *   AZURE_STORAGE_CONNECTION_STRING   account-key connection string
 *
 * The endpoint wins when set: requests carry an Entra token for the app's
 * managed identity, so no storage key exists anywhere in the deployment and the
 * account can run with shared-key access disabled. This is what infra/ deploys.
 * The connection string is the fallback for local development (Azurite) and
 * for an instance whose infra predates the endpoint. With neither set, the
 * Azurite emulator is assumed.
 */

export const AZURITE_CONNECTION_STRING = 'UseDevelopmentStorage=true';

let credential: TokenCredential | null = null;

function storageCredential(): TokenCredential {
  if (!credential) {
    const clientId = process.env.AZURE_STORAGE_IDENTITY_CLIENT_ID?.trim() || undefined;
    credential = new CachedTokenCredential(new ContainerAppManagedIdentityCredential(clientId));
  }
  return credential;
}

function tableEndpoint(): string | undefined {
  return process.env.AZURE_STORAGE_TABLE_ENDPOINT?.trim().replace(/\/+$/, '') || undefined;
}

function connectionString(): string {
  return process.env.AZURE_STORAGE_CONNECTION_STRING ?? AZURITE_CONNECTION_STRING;
}

function insecureAllowed(conn: string): boolean {
  // Azurite uses HTTP; the SDK rejects plain-HTTP endpoints unless this flag is set.
  return conn === AZURITE_CONNECTION_STRING || conn.includes('DefaultEndpointsProtocol=http;');
}

/** True when an Azure storage account is configured, by endpoint or connection string. */
export function isStorageConfigured(): boolean {
  return Boolean(tableEndpoint() || process.env.AZURE_STORAGE_CONNECTION_STRING);
}

export function getTableClient(tableName: string): TableClient {
  const endpoint = tableEndpoint();
  if (endpoint) return new TableClient(endpoint, tableName, storageCredential());
  const conn = connectionString();
  return TableClient.fromConnectionString(conn, tableName, { allowInsecureConnection: insecureAllowed(conn) });
}

export function getTableServiceClient(): TableServiceClient {
  const endpoint = tableEndpoint();
  if (endpoint) return new TableServiceClient(endpoint, storageCredential());
  const conn = connectionString();
  return TableServiceClient.fromConnectionString(conn, { allowInsecureConnection: insecureAllowed(conn) });
}

/** Test seam: forget the cached credential so env changes take effect. */
export function __resetStorageClientForTests(): void {
  credential = null;
}
