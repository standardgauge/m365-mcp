import {
  ConfidentialClientApplication,
  PublicClientApplication,
  Configuration,
  AuthorizationCodeRequest,
  AuthorizationUrlRequest,
  LogLevel,
  ICachePlugin,
  TokenCacheContext,
} from '@azure/msal-node';
import { Client, ClientOptions } from '@microsoft/microsoft-graph-client';
import 'isomorphic-fetch';
import {
  loadMsalCachePartition,
  saveMsalCachePartition,
  MsalCacheConflictError,
} from './tableStorage.js';
import { migrateLegacyCredentialKeys } from './msalCacheKeys.js';

export interface MsalConfig {
  clientId: string;
  clientSecret: string;
  tenantId: string;
  redirectUri: string;
}

function getMsalConfig(): MsalConfig {
  const clientId = process.env.AZURE_CLIENT_ID;
  const clientSecret = process.env.AZURE_CLIENT_SECRET;
  const tenantId = process.env.AZURE_TENANT_ID;
  const redirectUri = process.env.OAUTH_REDIRECT_URI;

  if (!clientId || !clientSecret || !tenantId || !redirectUri) {
    throw new Error(
      'Missing required env vars: AZURE_CLIENT_ID, AZURE_CLIENT_SECRET, AZURE_TENANT_ID, OAUTH_REDIRECT_URI'
    );
  }

  return { clientId, clientSecret, tenantId, redirectUri };
}

// Minimal dynamic-consent trigger set requested by MSAL (auth-code + silent refresh).
//
// This is deliberately NOT the full list of delegated permissions the tools use. On the
// v2.0 endpoint the access token's `scp` claim carries ALL Graph delegated permissions that
// have been consented for the resource — not just the subset named here (see
// https://learn.microsoft.com/entra/identity-platform/scopes-oidc: "the returned token
// contains all scopes granted for that resource"). The calendar (Calendars.ReadWrite),
// OneNote (Notes.*), and Teams read (Channel.ReadBasic.All / ChannelMessage.Read.All /
// Team.ReadBasic.All) tools already rely on this — none of those scopes is listed here, yet
// the tools work because they're granted via app-registration admin consent.
//
// The scheduling + Teams tools follow the same pattern: their scopes
// (Calendars.Read.Shared, MailboxSettings.Read, Place.Read.All, ChatMessage.Send,
// ChannelMessage.Send, Team.ReadBasic.All, Channel.ReadBasic.All) are granted via admin
// consent — documented in README.md and docs/entra-setup.md — and flow into the token
// automatically on the next silent refresh once consent lands, with no re-login.
// (create_event's isOnlineMeeting flag needs no new scope: it only sets fields on the
// POST /me/events calendar API — already covered by Calendars.ReadWrite — and never calls
// the standalone /onlineMeetings API, so OnlineMeetings.ReadWrite is not required.)
//
// DO NOT add new tool permissions to this list. Adding a scope that a live tenant has not yet
// consented to makes acquireTokenSilent below request it and throw interaction_required,
// breaking EVERY tool call for already-signed-in users across all deploy forks (example,
// example, Example) mid-rollout. New permissions are granted at the app-registration level, not here.
export const GRAPH_SCOPES = [
  'Sites.ReadWrite.All',
  'Files.ReadWrite.All',
  'Mail.ReadWrite',
  'Mail.ReadBasic',
  'User.Read',
  'Contacts.ReadWrite',
  'Directory.Read.All',
];

// ── Token cache: one MSAL app and one Table row per account ──
//
// MSAL keeps its cache in memory on the app object and calls the plugin around
// every access. A single app shared by every user would hold every user's
// refresh token in one place, and two requests for different users would
// interleave their load and save on the same in-memory cache. So each account
// gets its own ConfidentialClientApplication, whose plugin reads and writes
// only that account's row (tableStorage.ts, "MSAL cache persistence"), and
// token acquisition for one account runs one call at a time on this replica.
//
// Writes are conditional on the ETag the plugin last read. If another replica
// refreshed the same account in between, our write loses and theirs stands:
// both caches hold a refresh token issued moments apart for the same account,
// and the next access loads the winner's.

function msalLogger(): NonNullable<Configuration['system']>['loggerOptions'] {
  return {
    loggerCallback(loglevel: LogLevel, message: string, containsPii: boolean) {
      if (!containsPii) {
        console.log('[MSAL]', message);
      }
    },
    piiLoggingEnabled: false,
    // Warning (not Verbose) in production: Verbose floods logs with
    // per-request MSAL internals that add noise without diagnostic value
    // here, and keeps the log surface minimal (F11).
    logLevel: LogLevel.Warning,
  };
}

function confidentialConfig(cfg: MsalConfig, cachePlugin?: ICachePlugin): Configuration {
  return {
    auth: {
      clientId: cfg.clientId,
      authority: `https://login.microsoftonline.com/${cfg.tenantId}`,
      clientSecret: cfg.clientSecret,
    },
    ...(cachePlugin ? { cache: { cachePlugin } } : {}),
    system: { loggerOptions: msalLogger() },
  };
}

interface AccountCachePlugin extends ICachePlugin {
  /** Whether the last load found no usable row for the account. */
  readonly rowMissing: boolean;
}

/**
 * Cache plugin bound to one account's row. It only ever updates a row it
 * found: rows are created by sign-in (acquireTokenByCode), so a row deleted to
 * sign the account out is not recreated from what this replica still holds in
 * memory.
 */
function accountCachePlugin(homeAccountId: string): AccountCachePlugin {
  let etag: string | undefined;
  let rowMissing = false;
  return {
    get rowMissing() {
      return rowMissing;
    },
    async beforeCacheAccess(cacheContext: TokenCacheContext): Promise<void> {
      try {
        const row = await loadMsalCachePartition(homeAccountId);
        etag = row.etag;
        rowMissing = row.data === null;
        if (row.data) {
          // A cache written by msal-node 2.x keeps its old credential keys until
          // re-keyed here; see msalCacheKeys.ts for what goes wrong otherwise.
          cacheContext.tokenCache.deserialize(migrateLegacyCredentialKeys(row.data));
        }
      } catch (err) {
        // A storage error is not a sign-out: carry on with what is in memory,
        // as before. The write below is still conditional on the last ETag.
        console.error('[MSAL] Failed to load cache from Table Storage:', err);
      }
    },
    async afterCacheAccess(cacheContext: TokenCacheContext): Promise<void> {
      if (!cacheContext.cacheHasChanged || rowMissing || etag === undefined) return;
      try {
        etag = await saveMsalCachePartition(homeAccountId, cacheContext.tokenCache.serialize(), etag);
      } catch (err) {
        if (err instanceof MsalCacheConflictError) {
          console.warn('[MSAL] Cache row changed under this write; keeping the newer row');
        } else {
          console.error('[MSAL] Failed to save cache to Table Storage:', err);
        }
      }
    },
  };
}

// Per-account apps, least recently used first. Bounded so a replica that has
// served many users does not keep an app for each of them forever; an evicted
// account just gets a new app (and one metadata lookup) on its next request.
const MAX_ACCOUNT_APPS = 1000;
interface AccountApp {
  app: ConfidentialClientApplication;
  cache: AccountCachePlugin;
}
const accountApps = new Map<string, AccountApp>();
const accountLocks = new Map<string, Promise<unknown>>();

function accountAppKey(homeAccountId: string): string {
  const cfg = getMsalConfig();
  return `${cfg.clientId}:${cfg.tenantId}:${homeAccountId}`;
}

function getAccountMsalApp(homeAccountId: string): AccountApp {
  const key = accountAppKey(homeAccountId);
  let entry = accountApps.get(key);
  if (entry) {
    accountApps.delete(key);
  } else {
    const cache = accountCachePlugin(homeAccountId);
    entry = { app: new ConfidentialClientApplication(confidentialConfig(getMsalConfig(), cache)), cache };
    if (accountApps.size >= MAX_ACCOUNT_APPS) {
      accountApps.delete(accountApps.keys().next().value!);
    }
  }
  accountApps.set(key, entry);
  return entry;
}

/** Runs `fn` after any earlier call for the same account on this replica. */
function withAccountLock<T>(homeAccountId: string, fn: () => Promise<T>): Promise<T> {
  const previous = accountLocks.get(homeAccountId) ?? Promise.resolve();
  const run = previous.then(fn, fn);
  const settled = run.catch(() => undefined);
  accountLocks.set(homeAccountId, settled);
  void settled.then(() => {
    if (accountLocks.get(homeAccountId) === settled) accountLocks.delete(homeAccountId);
  });
  return run;
}

// The app used to build sign-in URLs. It has no cache plugin and never holds a
// token: code redemption runs on a fresh app per sign-in (acquireTokenByCode).
const msalApps = new Map<string, ConfidentialClientApplication>();
const publicMsalApps = new Map<string, PublicClientApplication>();

export function getMsalApp(config?: MsalConfig): ConfidentialClientApplication {
  const cfg = config ?? getMsalConfig();
  const key = `${cfg.clientId}:${cfg.tenantId}`;

  if (!msalApps.has(key)) {
    msalApps.set(key, new ConfidentialClientApplication(confidentialConfig(cfg)));
  }

  return msalApps.get(key)!;
}

export function getPublicMsalApp(config?: MsalConfig): PublicClientApplication {
  const cfg = config ?? getMsalConfig();
  const key = `${cfg.clientId}:${cfg.tenantId}`;

  if (!publicMsalApps.has(key)) {
    const msalConfig: Configuration = {
      auth: {
        clientId: cfg.clientId,
        authority: `https://login.microsoftonline.com/${cfg.tenantId}`,
      },
      system: { loggerOptions: msalLogger() },
    };
    publicMsalApps.set(key, new PublicClientApplication(msalConfig));
  }

  return publicMsalApps.get(key)!;
}

/**
 * What binds one authorization-code redemption to the browser that started it,
 * beyond `state`. `login` generates both halves and keeps the secret ones in
 * HttpOnly cookies; `callback` reads them back.
 *
 * - PKCE (RFC 7636, S256): the authorize URL carries SHA-256(verifier); the
 *   token request must present the verifier, so a code intercepted or injected
 *   from elsewhere cannot be redeemed against this browser's flow.
 * - `nonce`: sent on the authorize URL, echoed by Entra in the ID token, and
 *   compared by MSAL in acquireTokenByCode; a mismatch throws.
 */
export interface AuthCodeUrlBinding {
  codeChallenge: string;
  nonce: string;
}

export interface AuthCodeRedemptionBinding {
  codeVerifier: string;
  nonce: string;
}

export async function getAuthCodeUrl(state: string, binding: AuthCodeUrlBinding): Promise<string> {
  const msalApp = getMsalApp();
  const params: AuthorizationUrlRequest = {
    scopes: GRAPH_SCOPES,
    redirectUri: getMsalConfig().redirectUri,
    state,
    codeChallenge: binding.codeChallenge,
    codeChallengeMethod: 'S256',
    nonce: binding.nonce,
  };
  return msalApp.getAuthCodeUrl(params);
}

export async function acquireTokenByCode(
  code: string,
  binding: AuthCodeRedemptionBinding
): Promise<{
  accessToken: string;
  homeAccountId: string;
  expiresOn: Date | null;
}> {
  const cfg = getMsalConfig();
  // A fresh app, so its cache holds this sign-in and nothing else; that cache
  // becomes the account's row below.
  const msalApp = new ConfidentialClientApplication(confidentialConfig(cfg));

  const request: AuthorizationCodeRequest = {
    code,
    scopes: GRAPH_SCOPES,
    redirectUri: cfg.redirectUri,
    codeVerifier: binding.codeVerifier,
    // With nonce set, MSAL requires the ID token's nonce claim to match it.
    nonce: binding.nonce,
  };

  const response = await msalApp.acquireTokenByCode(request);

  if (!response?.accessToken || !response.account) {
    throw new Error('Failed to acquire token by authorization code');
  }

  const homeAccountId = response.account.homeAccountId;
  await withAccountLock(homeAccountId, () =>
    persistSignIn(homeAccountId, msalApp.getTokenCache().serialize()),
  );

  return {
    accessToken: response.accessToken,
    homeAccountId,
    expiresOn: response.expiresOn,
  };
}

const SIGN_IN_SAVE_ATTEMPTS = 3;

/**
 * Replaces the account's row with the cache from a fresh sign-in. Still
 * conditional: it re-reads the ETag and retries if another replica writes the
 * row in between, so the sign-in's tokens are what is left in the row.
 */
async function persistSignIn(homeAccountId: string, cacheData: string): Promise<void> {
  for (let attempt = 1; ; attempt++) {
    const { etag } = await loadMsalCachePartition(homeAccountId);
    try {
      await saveMsalCachePartition(homeAccountId, cacheData, etag);
      return;
    } catch (err) {
      if (!(err instanceof MsalCacheConflictError) || attempt >= SIGN_IN_SAVE_ATTEMPTS) throw err;
    }
  }
}

export async function acquireTokenSilent(
  homeAccountId: string
): Promise<{ accessToken: string; expiresOn: Date | null }> {
  return withAccountLock(homeAccountId, async () => {
    const { app: msalApp, cache } = getAccountMsalApp(homeAccountId);
    const accounts = await msalApp.getTokenCache().getAllAccounts();
    // MSAL merges a loaded cache into what it already holds in memory, so an
    // app that served this account before still lists it after its row is
    // gone. The row is what counts: no row, no account, and the app (with
    // its stale tokens) is dropped.
    const account = cache.rowMissing
      ? undefined
      : accounts.find((a) => a.homeAccountId === homeAccountId);

    if (!account) {
      if (cache.rowMissing) accountApps.delete(accountAppKey(homeAccountId));
      throw new Error(
        `No cached account found for homeAccountId ${homeAccountId}. Re-authentication required.`
      );
    }

    const response = await msalApp.acquireTokenSilent({
      scopes: GRAPH_SCOPES,
      account,
    });

    if (!response?.accessToken) {
      throw new Error('Token refresh failed. Re-authentication required.');
    }

    return {
      accessToken: response.accessToken,
      expiresOn: response.expiresOn,
    };
  });
}

export function createGraphClient(accessToken: string): Client {
  // Graph SDK v3 uses initWithMiddleware + the AuthenticationProvider interface
  // (promise-based getAccessToken), not the legacy callback-style Client.init.
  const clientOptions: ClientOptions = {
    authProvider: {
      getAccessToken: () => Promise.resolve(accessToken),
    },
  };
  return Client.initWithMiddleware(clientOptions);
}
