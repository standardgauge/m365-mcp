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
import { saveMsalCache, loadMsalCache } from './tableStorage.js';

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

// MSAL cache plugin — persists token cache to Azure Table Storage
const cachePlugin: ICachePlugin = {
  async beforeCacheAccess(cacheContext: TokenCacheContext): Promise<void> {
    try {
      const cached = await loadMsalCache();
      if (cached) {
        cacheContext.tokenCache.deserialize(cached);
      } else {
        console.warn('[MSAL] No cache data returned from storage — MSAL operating with empty cache');
      }
    } catch (err) {
      console.error('[MSAL] Failed to load cache from Table Storage:', err);
    }
  },
  async afterCacheAccess(cacheContext: TokenCacheContext): Promise<void> {
    if (cacheContext.cacheHasChanged) {
      try {
        await saveMsalCache(cacheContext.tokenCache.serialize());
      } catch (err) {
        console.error('[MSAL] Failed to save cache to Table Storage:', err);
      }
    }
  },
};

// Singleton MSAL app per client+tenant combination
const msalApps = new Map<string, ConfidentialClientApplication>();
const publicMsalApps = new Map<string, PublicClientApplication>();

export function getMsalApp(config?: MsalConfig): ConfidentialClientApplication {
  const cfg = config ?? getMsalConfig();
  const key = `${cfg.clientId}:${cfg.tenantId}`;

  if (!msalApps.has(key)) {
    const msalConfig: Configuration = {
      auth: {
        clientId: cfg.clientId,
        authority: `https://login.microsoftonline.com/${cfg.tenantId}`,
        clientSecret: cfg.clientSecret,
      },
      cache: {
        cachePlugin,
      },
      system: {
        loggerOptions: {
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
        },
      },
    };
    msalApps.set(key, new ConfidentialClientApplication(msalConfig));
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
      system: {
        loggerOptions: {
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
        },
      },
    };
    publicMsalApps.set(key, new PublicClientApplication(msalConfig));
  }

  return publicMsalApps.get(key)!;
}

export async function getAuthCodeUrl(state: string): Promise<string> {
  const msalApp = getMsalApp();
  const params: AuthorizationUrlRequest = {
    scopes: GRAPH_SCOPES,
    redirectUri: getMsalConfig().redirectUri,
    state,
  };
  return msalApp.getAuthCodeUrl(params);
}

export async function acquireTokenByCode(code: string): Promise<{
  accessToken: string;
  homeAccountId: string;
  expiresOn: Date | null;
}> {
  const msalApp = getMsalApp();
  const cfg = getMsalConfig();

  const request: AuthorizationCodeRequest = {
    code,
    scopes: GRAPH_SCOPES,
    redirectUri: cfg.redirectUri,
  };

  const response = await msalApp.acquireTokenByCode(request);

  if (!response?.accessToken || !response.account) {
    throw new Error('Failed to acquire token by authorization code');
  }

  return {
    accessToken: response.accessToken,
    homeAccountId: response.account.homeAccountId,
    expiresOn: response.expiresOn,
  };
}

export async function acquireTokenSilent(
  homeAccountId: string
): Promise<{ accessToken: string; expiresOn: Date | null }> {
  const msalApp = getMsalApp();
  const tokenCache = msalApp.getTokenCache();
  const accounts = await tokenCache.getAllAccounts();
  const account = accounts.find((a) => a.homeAccountId === homeAccountId);

  if (!account) {
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
