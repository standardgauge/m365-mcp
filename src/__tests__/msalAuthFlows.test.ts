/**
 * The MSAL-backed auth flows, run through the real @azure/msal-node.
 *
 * Every other test that touches graphClient mocks MSAL out entirely, so none of
 * them would notice a library upgrade that changed request shapes, cache
 * serialization, or how a cached refresh token is found and used. Here MSAL is
 * the genuine package; only its network client is swapped for FakeEntra, an
 * in-process stand-in for the Entra endpoints (fixtures/fakeEntra.ts), and the
 * Table Storage row the cache plugin reads and writes is an in-memory string.
 *
 * Covered:
 *   - authorization code: the auth URL (with PKCE challenge and nonce), the
 *     code redemption (verifier presented, ID-token nonce checked), and the
 *     cache write that follows it
 *   - token cache persistence: what the plugin saves is what the next access
 *     loads, and a cache serialized by msal-node 2.16.3 (the version deployed
 *     before the move to 7) still yields its account and refresh token, so
 *     signed-in users are not logged out by the upgrade, and is re-keyed so the
 *     rotated refresh token is the one used next time (msalCacheKeys.ts)
 *   - refresh: an expired access token is renewed with the cached refresh token
 *     and the rotated refresh token is persisted; a rejected refresh surfaces
 *     as an error the /api/auth/refresh handler turns into a 401
 *   - device code: GET /api/auth/device returns the user code, polls through
 *     authorization_pending, and stores a session once the grant succeeds
 */

import { createHash, randomBytes } from 'crypto';
import { readFileSync } from 'fs';
import { join } from 'path';

process.env.MCP_SESSION_HMAC_KEY = randomBytes(32).toString('hex');
process.env.MCP_DATA_ENCRYPTION_KEY = randomBytes(32).toString('hex');

import { jest } from '@jest/globals';
import type { HttpRequest, InvocationContext } from '@azure/functions';
import {
  FakeEntra,
  FAKE_CLIENT_ID,
  FAKE_HOME_ACCOUNT_ID,
  FAKE_TENANT_ID,
  FAKE_UPN,
  FAKE_USER_OID,
  idToken,
} from './fixtures/fakeEntra.js';

process.env.AZURE_CLIENT_ID = FAKE_CLIENT_ID;
process.env.AZURE_CLIENT_SECRET = 'fake-client-secret';
process.env.AZURE_TENANT_ID = FAKE_TENANT_ID;
process.env.OAUTH_REDIRECT_URI = 'https://mcp.example.com/api/auth/callback';

// ── Real MSAL, fake network ───────────────────────────────────────────────────

// Swapped per test; the network client below forwards to whatever is current.
const mockEntraRef: { current: FakeEntra } = { current: new FakeEntra() };

jest.mock('@azure/msal-node', () => {
  const actual = jest.requireActual<typeof import('@azure/msal-node')>('@azure/msal-node');
  const networkClient = {
    sendGetRequestAsync: (url: string, options?: unknown) =>
      mockEntraRef.current.sendGetRequestAsync(url, options as never),
    sendPostRequestAsync: (url: string, options?: unknown) =>
      mockEntraRef.current.sendPostRequestAsync(url, options as never),
  };
  const withFakeNetwork = (config: import('@azure/msal-node').Configuration) => ({
    ...config,
    system: { ...config.system, networkClient: networkClient as never },
  });
  class ConfidentialClientApplication extends actual.ConfidentialClientApplication {
    constructor(config: import('@azure/msal-node').Configuration) {
      super(withFakeNetwork(config));
    }
  }
  class PublicClientApplication extends actual.PublicClientApplication {
    constructor(config: import('@azure/msal-node').Configuration) {
      super(withFakeNetwork(config));
    }
  }
  return { ...actual, ConfidentialClientApplication, PublicClientApplication };
});

// ── Table Storage row for the MSAL cache ──────────────────────────────────────

const mockCacheRow: { value: string | null; saves: number } = { value: null, saves: 0 };

jest.mock('../services/tableStorage.js', () => ({
  loadMsalCache: jest.fn(async () => mockCacheRow.value),
  saveMsalCache: jest.fn(async (data: string) => {
    mockCacheRow.value = data;
    mockCacheRow.saves += 1;
  }),
}));

// ── Collaborators of the device code handler ──────────────────────────────────

const mockStoreSession = jest.fn<(session: Record<string, unknown>) => Promise<void>>();
jest.mock('../services/tokenCache.js', () => ({
  storeSession: (session: Record<string, unknown>) => mockStoreSession(session),
}));

jest.mock('@microsoft/microsoft-graph-client', () => ({
  Client: {
    initWithMiddleware: jest.fn(() => ({
      api: jest.fn(() => ({
        select: jest.fn(() => ({
          get: jest.fn(async () => ({
            id: 'graph-user-id',
            displayName: 'Adele Vance',
            mail: 'adele.vance@fabrikam.com',
            userPrincipalName: 'adele.vance@fabrikam.com',
          })),
        })),
      })),
    })),
  },
}));

jest.mock('isomorphic-fetch', () => ({}));
jest.mock('@azure/functions', () => ({ app: { http: jest.fn() } }));

// ── Modules under test ────────────────────────────────────────────────────────

import { app } from '@azure/functions';
import { InteractionRequiredAuthError } from '@azure/msal-node';
import {
  GRAPH_SCOPES,
  acquireTokenByCode,
  acquireTokenSilent,
  getAuthCodeUrl,
} from '../services/graphClient.js';
import { currentCredentialKey, migrateLegacyCredentialKeys } from '../services/msalCacheKeys.js';
import '../functions/auth/deviceLogin.js';

type Handler = (req: HttpRequest, ctx: InvocationContext) => Promise<{ status?: number; jsonBody?: unknown }>;
const registrations = (app.http as unknown as jest.Mock).mock.calls as Array<[string, { handler: Handler }]>;
const deviceLoginReg = registrations.find(([name]) => name === 'deviceLogin');
if (!deviceLoginReg) throw new Error('deviceLogin handler was not registered');
const deviceLogin = deviceLoginReg[1].handler;

// Serialized by msal-node 2.16.3 from an auth code sign-in against FakeEntra.
// The id token is a placeholder in the file (secret scanning blocks anything
// JWT-shaped) and filled in here with the one FakeEntra issues.
const LEGACY_CACHE = readFileSync(
  join(__dirname, 'fixtures', 'msal-node-2.16.3-cache.json'),
  'utf8'
).replace('"__FAKE_ID_TOKEN__"', JSON.stringify(idToken()));

// A PKCE pair and nonce as /api/auth/login makes them.
const CODE_VERIFIER = randomBytes(32).toString('base64url');
const CODE_CHALLENGE = createHash('sha256').update(CODE_VERIFIER).digest('base64url');
const NONCE = 'test-nonce-0123456789abcdef';
const BINDING = { codeVerifier: CODE_VERIFIER, nonce: NONCE };

/** Arm FakeEntra as if the authorize request carried CODE_CHALLENGE and NONCE. */
function armAuthorize(): void {
  mockEntraRef.current.codeChallenge = CODE_CHALLENGE;
  mockEntraRef.current.authorizeNonce = NONCE;
}

function persistedRefreshTokens(): string[] {
  const cache = JSON.parse(mockCacheRow.value ?? '{}') as {
    RefreshToken?: Record<string, { secret: string }>;
  };
  return Object.values(cache.RefreshToken ?? {}).map((rt) => rt.secret);
}

beforeEach(() => {
  mockEntraRef.current = new FakeEntra();
  mockCacheRow.value = null;
  mockCacheRow.saves = 0;
  mockStoreSession.mockReset();
  mockStoreSession.mockResolvedValue(undefined);
  jest.spyOn(console, 'warn').mockImplementation(() => undefined);
  jest.spyOn(console, 'log').mockImplementation(() => undefined);
});

afterEach(() => {
  jest.restoreAllMocks();
});

describe('authorization code flow', () => {
  it('builds an auth URL for the tenant, client, scopes, state, PKCE challenge and nonce', async () => {
    const url = new URL(
      await getAuthCodeUrl('opaque-state', { codeChallenge: CODE_CHALLENGE, nonce: NONCE })
    );

    expect(url.origin + url.pathname).toBe(
      `https://login.microsoftonline.com/${FAKE_TENANT_ID}/oauth2/v2.0/authorize`
    );
    expect(url.searchParams.get('client_id')).toBe(FAKE_CLIENT_ID);
    expect(url.searchParams.get('state')).toBe('opaque-state');
    expect(url.searchParams.get('redirect_uri')).toBe(process.env.OAUTH_REDIRECT_URI);
    const scopes = url.searchParams.get('scope')!.split(' ');
    for (const scope of GRAPH_SCOPES) expect(scopes).toContain(scope);
    expect(scopes).toContain('offline_access');
    expect(url.searchParams.get('code_challenge')).toBe(CODE_CHALLENGE);
    expect(url.searchParams.get('code_challenge_method')).toBe('S256');
    expect(url.searchParams.get('nonce')).toBe(NONCE);
  });

  it('redeems the code and persists the refresh token through the cache plugin', async () => {
    armAuthorize();
    const result = await acquireTokenByCode('fake-auth-code', BINDING);

    expect(result.accessToken).toBe('fake-access-token-1');
    expect(result.homeAccountId).toBe(FAKE_HOME_ACCOUNT_ID);
    expect(result.expiresOn).toBeInstanceOf(Date);

    const [req] = mockEntraRef.current.tokenRequests;
    expect(req.form.get('grant_type')).toBe('authorization_code');
    expect(req.form.get('code')).toBe('fake-auth-code');
    expect(req.form.get('client_secret')).toBe('fake-client-secret');
    expect(req.form.get('code_verifier')).toBe(CODE_VERIFIER);

    expect(mockCacheRow.saves).toBeGreaterThan(0);
    expect(persistedRefreshTokens()).toEqual(['fake-refresh-token-1']);
  });

  it('fails when the verifier does not match the challenge, and caches nothing', async () => {
    armAuthorize();

    await expect(
      acquireTokenByCode('fake-auth-code', {
        codeVerifier: randomBytes(32).toString('base64url'),
        nonce: NONCE,
      })
    ).rejects.toThrow(/invalid_grant/);
    expect(persistedRefreshTokens()).toEqual([]);
  });

  it('rejects an ID token whose nonce differs from the one sent at login', async () => {
    armAuthorize();
    mockEntraRef.current.authorizeNonce = 'nonce-from-another-flow';

    await expect(acquireTokenByCode('fake-auth-code', BINDING)).rejects.toThrow(/nonce/i);
    expect(persistedRefreshTokens()).toEqual([]);
  });

  it('rejects an ID token that carries no nonce', async () => {
    armAuthorize();
    mockEntraRef.current.authorizeNonce = null;

    await expect(acquireTokenByCode('fake-auth-code', BINDING)).rejects.toThrow(/nonce/i);
    expect(persistedRefreshTokens()).toEqual([]);
  });
});

describe('silent refresh from the persisted cache', () => {
  it('renews an expired access token with the cached refresh token and persists the rotation', async () => {
    armAuthorize();
    await acquireTokenByCode('fake-auth-code', BINDING);
    const tokensBefore = mockEntraRef.current.tokenRequests.length;

    // FakeEntra issues 60s access tokens, inside MSAL's refresh buffer, so the
    // silent call has to go to the token endpoint.
    const result = await acquireTokenSilent(FAKE_HOME_ACCOUNT_ID);

    expect(result.accessToken).toBe('fake-access-token-2');
    const refresh = mockEntraRef.current.tokenRequests.slice(tokensBefore);
    expect(refresh).toHaveLength(1);
    expect(refresh[0].form.get('grant_type')).toBe('refresh_token');
    expect(refresh[0].form.get('refresh_token')).toBe('fake-refresh-token-1');
    expect(persistedRefreshTokens()).toEqual(['fake-refresh-token-2']);
  });

  it('serves a still-valid access token from the cache without a network call', async () => {
    mockEntraRef.current.accessTokenLifetimeSeconds = 3600;
    armAuthorize();
    await acquireTokenByCode('fake-auth-code', BINDING);
    const tokensBefore = mockEntraRef.current.tokenRequests.length;

    const result = await acquireTokenSilent(FAKE_HOME_ACCOUNT_ID);

    expect(result.accessToken).toBe('fake-access-token-1');
    expect(mockEntraRef.current.tokenRequests.length).toBe(tokensBefore);
  });

  it('reads a cache written by msal-node 2.16.3 and refreshes from its refresh token', async () => {
    mockCacheRow.value = LEGACY_CACHE;
    // The fixture holds token pair 1; number the fresh pairs so they are distinct.
    mockEntraRef.current.issued = 100;

    const first = await acquireTokenSilent(FAKE_HOME_ACCOUNT_ID);

    expect(first.accessToken).toBe('fake-access-token-101');
    const [req] = mockEntraRef.current.tokenRequests;
    expect(req.form.get('grant_type')).toBe('refresh_token');
    expect(req.form.get('refresh_token')).toBe('fake-refresh-token-1');
    expect(req.form.get('client_id')).toBe(FAKE_CLIENT_ID);

    // One account, and the rotated refresh token replaced the old one rather
    // than landing beside it under msal-node 7's key format.
    const cache = JSON.parse(mockCacheRow.value!) as {
      Account: Record<string, { username: string }>;
      IdToken: Record<string, unknown>;
    };
    expect(Object.values(cache.Account).map((a) => a.username)).toEqual([FAKE_UPN]);
    expect(Object.keys(cache.IdToken)).toHaveLength(1);
    expect(persistedRefreshTokens()).toEqual(['fake-refresh-token-101']);

    // So the next refresh presents the rotated token, not the one from before
    // the upgrade.
    const second = await acquireTokenSilent(FAKE_HOME_ACCOUNT_ID);
    expect(second.accessToken).toBe('fake-access-token-102');
    expect(mockEntraRef.current.tokenRequests[1].form.get('refresh_token')).toBe('fake-refresh-token-101');
    expect(persistedRefreshTokens()).toEqual(['fake-refresh-token-102']);
  });

  it('leaves a cache written by the installed msal-node unchanged when re-keying', async () => {
    // If this fails after an MSAL upgrade, the credential key format moved
    // again: update currentCredentialKey in msalCacheKeys.ts to match.
    armAuthorize();
    await acquireTokenByCode('fake-auth-code', BINDING);
    await acquireTokenSilent(FAKE_HOME_ACCOUNT_ID);
    const written = mockCacheRow.value!;

    expect(migrateLegacyCredentialKeys(written)).toBe(written);
  });

  it('re-keys every credential in the 2.16.3 fixture to the current format', () => {
    const migrated = JSON.parse(migrateLegacyCredentialKeys(LEGACY_CACHE)) as Record<
      string,
      Record<string, Record<string, string>>
    >;
    for (const section of ['IdToken', 'AccessToken', 'RefreshToken']) {
      const entries = Object.entries(migrated[section]);
      expect(entries).toHaveLength(1);
      const [key, credential] = entries[0];
      expect(key).toBe(currentCredentialKey(credential));
    }
  });

  it('passes through input it cannot interpret', () => {
    expect(migrateLegacyCredentialKeys('not json')).toBe('not json');
    expect(migrateLegacyCredentialKeys('null')).toBe('null');
    const odd = JSON.stringify({ RefreshToken: { k: null }, AccessToken: 'x' });
    expect(migrateLegacyCredentialKeys(odd)).toBe(odd);
  });

  it('rejects when Entra refuses the refresh token', async () => {
    // Any rejection here is what /api/auth/refresh turns into a 401 with a
    // loginUrl; the error class is MSAL's business.
    mockCacheRow.value = LEGACY_CACHE;
    mockEntraRef.current.refreshError = 'invalid_grant';

    await expect(acquireTokenSilent(FAKE_HOME_ACCOUNT_ID)).rejects.toThrow('invalid_grant');
  });

  it('classifies interaction_required from Entra as InteractionRequiredAuthError', async () => {
    mockCacheRow.value = LEGACY_CACHE;
    mockEntraRef.current.refreshError = 'interaction_required';

    await expect(acquireTokenSilent(FAKE_HOME_ACCOUNT_ID)).rejects.toBeInstanceOf(
      InteractionRequiredAuthError
    );
  });

  it('asks for re-authentication when the account is not in the cache', async () => {
    mockCacheRow.value = LEGACY_CACHE;

    await expect(acquireTokenSilent(`not-cached.${FAKE_TENANT_ID}`)).rejects.toThrow(
      'Re-authentication required'
    );
    expect(mockEntraRef.current.tokenRequests).toHaveLength(0);
  });
});

describe('device code flow (GET /api/auth/device)', () => {
  function ctx(): InvocationContext {
    return { log: jest.fn(), error: jest.fn(), warn: jest.fn() } as unknown as InvocationContext;
  }

  it('returns the user code, polls through authorization_pending, and stores the session', async () => {
    const stored = new Promise<Record<string, unknown>>((resolve) => {
      mockStoreSession.mockImplementation(async (session) => resolve(session));
    });

    const res = await deviceLogin({} as HttpRequest, ctx());

    expect(res.status).toBe(200);
    expect(res.jsonBody).toEqual({
      status: 'pending',
      userCode: 'FAKE-CODE',
      verificationUri: 'https://microsoft.com/devicelogin',
      message: expect.stringContaining('FAKE-CODE'),
    });

    const session = await stored;
    expect(session).toMatchObject({
      homeAccountId: FAKE_HOME_ACCOUNT_ID,
      tenantId: FAKE_TENANT_ID,
      accessToken: 'fake-access-token-1',
      email: 'adele.vance@fabrikam.com',
    });
    expect(typeof session.sessionToken).toBe('string');

    const grants = mockEntraRef.current.tokenRequests.map((r) => r.form.get('grant_type'));
    expect(grants).toEqual(['device_code', 'device_code']);
    expect(mockEntraRef.current.tokenRequests[1].form.get('device_code')).toBe('fake-device-code');
  });

  it('logs and stores nothing when the user is from another tenant', async () => {
    const context = ctx();
    const rejected = new Promise<void>((resolve) => {
      (context.error as jest.Mock).mockImplementation(() => resolve());
    });

    const res = await deviceLogin({} as HttpRequest, context);
    expect(res.status).toBe(200);

    // The handler reads AZURE_TENANT_ID only once the grant completes, after
    // the response has gone back, so changing it here is what that check sees.
    process.env.AZURE_TENANT_ID = 'aaaaaaaa-bbbb-4ccc-8ddd-eeeeeeeeeeee';
    try {
      await rejected;
    } finally {
      process.env.AZURE_TENANT_ID = FAKE_TENANT_ID;
    }
    expect((context.error as jest.Mock).mock.calls[0][0]).toMatch(/tenant mismatch/);
    expect(mockStoreSession).not.toHaveBeenCalled();
  });
});

// Identity fields FakeEntra puts in the id token, asserted once so a change to
// the fixture shows up here rather than as a confusing mismatch above.
it('FakeEntra identity matches the committed 2.16.3 cache fixture', () => {
  const cache = JSON.parse(LEGACY_CACHE) as { Account: Record<string, { local_account_id: string }> };
  expect(Object.values(cache.Account).map((a) => a.local_account_id)).toEqual([FAKE_USER_OID]);
});
