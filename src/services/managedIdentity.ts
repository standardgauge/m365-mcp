import type { AccessToken, TokenCredential } from '@azure/core-auth';

/**
 * Token source for one of the Container App's managed identities, read from
 * the endpoint Container Apps injects (IDENTITY_ENDPOINT and IDENTITY_HEADER).
 * Kept to this instead of @azure/identity, which would pull a second, older
 * @azure/msal-node into the image.
 *
 * With no client id the endpoint answers for the system-assigned identity; a
 * client id selects a user-assigned identity attached to the app.
 */
export class ContainerAppManagedIdentityCredential implements TokenCredential {
  constructor(private readonly clientId?: string) {}

  async getToken(scopes: string | string[]): Promise<AccessToken> {
    const endpoint = process.env.IDENTITY_ENDPOINT;
    const header = process.env.IDENTITY_HEADER;
    if (!endpoint || !header) {
      throw new Error('IDENTITY_ENDPOINT / IDENTITY_HEADER not set: no managed identity available');
    }
    const scope = Array.isArray(scopes) ? scopes[0] : scopes;
    const url = new URL(endpoint);
    url.searchParams.set('resource', scope.replace(/\/\.default$/, ''));
    url.searchParams.set('api-version', '2019-08-01');
    if (this.clientId) url.searchParams.set('client_id', this.clientId);
    const res = await fetch(url, { headers: { 'X-IDENTITY-HEADER': header } });
    if (!res.ok) {
      throw new Error(`managed identity token request failed: HTTP ${res.status}`);
    }
    const body = (await res.json()) as { access_token?: string; expires_on?: string | number };
    if (!body.access_token || body.expires_on === undefined) {
      throw new Error('managed identity token response missing access_token or expires_on');
    }
    return { token: body.access_token, expiresOnTimestamp: Number(body.expires_on) * 1000 };
  }
}

/** Refresh a cached token this long before it expires. */
export const TOKEN_REFRESH_MARGIN_MS = 5 * 60 * 1000;

/**
 * Holds the last token per scope and shares one in-flight request between
 * concurrent callers. A client pipeline caches its own token, but several
 * storage modules build a TableClient per call, so without this every table
 * operation would make its own round trip to the identity endpoint.
 */
export class CachedTokenCredential implements TokenCredential {
  private readonly cache = new Map<string, Promise<AccessToken>>();

  constructor(private readonly inner: TokenCredential) {}

  async getToken(scopes: string | string[]): Promise<AccessToken> {
    const key = Array.isArray(scopes) ? scopes.join(' ') : scopes;
    const cached = this.cache.get(key);
    if (cached) {
      try {
        const token = await cached;
        if (token.expiresOnTimestamp - TOKEN_REFRESH_MARGIN_MS > Date.now()) return token;
      } catch {
        // A failed request is not cached; fall through and ask again.
      }
      if (this.cache.get(key) !== cached) return this.getToken(scopes);
    }
    const pending = this.inner.getToken(scopes).then((token) => {
      if (!token) throw new Error('managed identity returned no token');
      return token;
    });
    this.cache.set(key, pending);
    pending.catch(() => {
      if (this.cache.get(key) === pending) this.cache.delete(key);
    });
    return pending;
  }
}
