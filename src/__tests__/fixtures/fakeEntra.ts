/**
 * An in-process stand-in for the Entra v2.0 endpoints MSAL talks to.
 *
 * It implements MSAL's INetworkModule, so it is handed to the real
 * @azure/msal-node through `system.networkClient` and the library runs its own
 * request building, response parsing and cache writes unmodified. Nothing here
 * imports from MSAL: the shapes are the wire format, which is what a library
 * upgrade has to keep honouring.
 *
 * Served: instance discovery, OpenID configuration, the device code endpoint,
 * and the token endpoint for the authorization_code, device_code and
 * refresh_token grants. Any other URL throws, so a request the test did not
 * anticipate fails the test rather than being answered with something vague.
 */

export const FAKE_TENANT_ID = '72f988bf-0000-4000-8000-00000000c0de';
export const FAKE_CLIENT_ID = '11111111-2222-4333-8444-555555555555';
export const FAKE_USER_OID = '99999999-8888-4777-8666-555555555555';
export const FAKE_HOME_ACCOUNT_ID = `${FAKE_USER_OID}.${FAKE_TENANT_ID}`;
export const FAKE_UPN = 'adele.vance@fabrikam.com';

const HOST = 'https://login.microsoftonline.com';

interface NetworkResponse<T> {
  headers: Record<string, string>;
  body: T;
  status: number;
}

interface NetworkRequestOptions {
  headers?: Record<string, string>;
  body?: string;
}

export interface TokenRequestRecord {
  url: string;
  form: URLSearchParams;
}

function b64url(value: string): string {
  return Buffer.from(value, 'utf8').toString('base64url');
}

/** An unsigned id token for the fake user. Built at run time, never committed:
 *  a JWT-shaped string in the tree trips secret scanning. */
export function idToken(): string {
  const now = Math.floor(Date.now() / 1000);
  const header = b64url(JSON.stringify({ alg: 'none', typ: 'JWT' }));
  const claims = b64url(
    JSON.stringify({
      aud: FAKE_CLIENT_ID,
      iss: `${HOST}/${FAKE_TENANT_ID}/v2.0`,
      iat: now,
      nbf: now,
      exp: now + 3600,
      name: 'Adele Vance',
      oid: FAKE_USER_OID,
      preferred_username: FAKE_UPN,
      sub: 'fake-subject',
      tid: FAKE_TENANT_ID,
      ver: '2.0',
    })
  );
  return `${header}.${claims}.fake-signature`;
}

function json<T>(body: T, status = 200): NetworkResponse<T> {
  return { headers: { 'content-type': 'application/json' }, body, status };
}

export class FakeEntra {
  /** Every POST to the token endpoint, in order. */
  readonly tokenRequests: TokenRequestRecord[] = [];
  /** Device code polls answered `authorization_pending` before succeeding. */
  devicePendingPolls = 1;
  /** Lifetime of issued access tokens. Under MSAL's 300s refresh buffer means
   *  the next silent call has to use the refresh token. */
  accessTokenLifetimeSeconds = 60;
  /** When set, the refresh_token grant answers with this OAuth error. */
  refreshError: string | null = null;

  /** Tokens issued so far; the next pair is numbered one higher. */
  issued = 0;

  async sendGetRequestAsync<T>(url: string, _options?: NetworkRequestOptions): Promise<NetworkResponse<T>> {
    const u = new URL(url);
    if (u.pathname === '/common/discovery/instance') {
      return json({
        tenant_discovery_endpoint: `${HOST}/${FAKE_TENANT_ID}/v2.0/.well-known/openid-configuration`,
        'api-version': '1.1',
        metadata: [
          {
            preferred_network: 'login.microsoftonline.com',
            preferred_cache: 'login.windows.net',
            aliases: ['login.microsoftonline.com', 'login.windows.net', 'login.microsoft.com', 'sts.windows.net'],
          },
        ],
      }) as NetworkResponse<T>;
    }
    if (u.pathname === `/${FAKE_TENANT_ID}/v2.0/.well-known/openid-configuration`) {
      return json({
        token_endpoint: `${HOST}/${FAKE_TENANT_ID}/oauth2/v2.0/token`,
        authorization_endpoint: `${HOST}/${FAKE_TENANT_ID}/oauth2/v2.0/authorize`,
        device_authorization_endpoint: `${HOST}/${FAKE_TENANT_ID}/oauth2/v2.0/devicecode`,
        end_session_endpoint: `${HOST}/${FAKE_TENANT_ID}/oauth2/v2.0/logout`,
        issuer: `${HOST}/${FAKE_TENANT_ID}/v2.0`,
        jwks_uri: `${HOST}/${FAKE_TENANT_ID}/discovery/v2.0/keys`,
      }) as NetworkResponse<T>;
    }
    throw new Error(`FakeEntra: unexpected GET ${url}`);
  }

  async sendPostRequestAsync<T>(url: string, options?: NetworkRequestOptions): Promise<NetworkResponse<T>> {
    const u = new URL(url);
    const form = new URLSearchParams(options?.body ?? '');

    if (u.pathname === `/${FAKE_TENANT_ID}/oauth2/v2.0/devicecode`) {
      return json({
        user_code: 'FAKE-CODE',
        device_code: 'fake-device-code',
        verification_uri: 'https://microsoft.com/devicelogin',
        expires_in: 900,
        // Zero so the pending poll does not slow the suite down.
        interval: 0,
        message: 'To sign in, enter FAKE-CODE at https://microsoft.com/devicelogin',
      }) as NetworkResponse<T>;
    }

    if (u.pathname === `/${FAKE_TENANT_ID}/oauth2/v2.0/token`) {
      this.tokenRequests.push({ url, form });
      const grant = form.get('grant_type');

      if (grant === 'device_code' && this.devicePendingPolls > 0) {
        this.devicePendingPolls -= 1;
        return json({ error: 'authorization_pending', error_description: 'pending' }, 400) as NetworkResponse<T>;
      }
      if (grant === 'refresh_token' && this.refreshError) {
        return json(
          { error: this.refreshError, error_description: `${this.refreshError}: fake`, error_codes: [70000] },
          400
        ) as NetworkResponse<T>;
      }
      if (!['authorization_code', 'device_code', 'refresh_token'].includes(grant ?? '')) {
        throw new Error(`FakeEntra: unexpected grant_type ${grant}`);
      }

      this.issued += 1;
      return json({
        token_type: 'Bearer',
        scope: (form.get('scope') ?? '').split(' ').filter((s) => !['openid', 'profile', 'offline_access'].includes(s)).join(' '),
        expires_in: this.accessTokenLifetimeSeconds,
        ext_expires_in: this.accessTokenLifetimeSeconds,
        access_token: `fake-access-token-${this.issued}`,
        refresh_token: `fake-refresh-token-${this.issued}`,
        id_token: idToken(),
        client_info: b64url(JSON.stringify({ uid: FAKE_USER_OID, utid: FAKE_TENANT_ID })),
      }) as NetworkResponse<T>;
    }

    throw new Error(`FakeEntra: unexpected POST ${url}`);
  }
}
