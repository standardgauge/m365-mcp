/**
 * Secrets in the shipped templates stay out of the Container App.
 *
 * Threat model section 6: with the storage key, the data key, the HMAC key and
 * the client secret all held as Container App secrets, anyone who could read
 * the app's secrets held every user's refresh token. The templates now put the
 * three app secrets in Key Vault, bind them as references, and reach storage
 * with the app's identity on an account that refuses shared-key auth. This
 * holds them there: a template that reintroduces a value-backed secret for one
 * of these, a storage key, or a connection string goes red.
 */
import { readFileSync, readdirSync } from 'fs';
import { join } from 'path';

const root = join(__dirname, '..', '..');
const read = (p: string): string => readFileSync(join(root, p), 'utf8');

const KEY_VAULT_SECRETS = ['azure-client-secret', 'mcp-session-hmac-key', 'mcp-data-encryption-key'];

// Templates that deploy the Container App, discovered by content as in
// containerPortInvariant.test.ts.
const appTemplates = readdirSync(join(root, 'infra'))
  .filter((f) => f.endsWith('.bicep'))
  .map((f) => join('infra', f))
  .filter((p) => /Microsoft\.App\/containerApps@/.test(read(p)));

function secretBlock(bicep: string, name: string): string {
  const at = bicep.indexOf(`name: '${name}'`);
  expect(at).toBeGreaterThan(-1);
  return bicep.slice(at, bicep.indexOf('}', at));
}

describe('infra secrets invariant', () => {
  it('finds the Container App templates', () => {
    expect(appTemplates.length).toBeGreaterThan(0);
  });

  describe.each(appTemplates)('%s', (path) => {
    const bicep = read(path);

    it.each(KEY_VAULT_SECRETS)('binds %s as a Key Vault reference through the runtime identity', (name) => {
      const block = secretBlock(bicep, name);
      expect(block).toMatch(/keyVaultUrl:/);
      expect(block).toMatch(/identity: runtimeIdentity\.id/);
      expect(block).not.toMatch(/\bvalue:/);
    });

    it('carries no storage key or connection string', () => {
      expect(bicep).not.toMatch(/AccountKey=/);
      expect(bicep).not.toMatch(/storageAccount\.listKeys\(/);
      expect(bicep).not.toMatch(/AZURE_STORAGE_CONNECTION_STRING/);
    });

    it('points the app at the table endpoint with its runtime identity', () => {
      expect(bicep).toMatch(/name: 'AZURE_STORAGE_TABLE_ENDPOINT',\s+value: storageAccount\.properties\.primaryEndpoints\.table/);
      expect(bicep).toMatch(/name: 'AZURE_STORAGE_IDENTITY_CLIENT_ID',\s+value: runtimeIdentity\.properties\.clientId/);
    });
  });

  it('the full-stack template disables shared-key access on the account it creates', () => {
    const main = read('infra/main.bicep');
    expect(main).toMatch(/allowSharedKeyAccess: false/);
  });

  it('the vault is RBAC-mode with purge protection', () => {
    const kv = read('infra/key-vault.bicep');
    expect(kv).toMatch(/enableRbacAuthorization: true/);
    expect(kv).toMatch(/enablePurgeProtection: true/);
  });
});
