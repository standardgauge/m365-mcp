import { app, HttpRequest, HttpResponseInit, InvocationContext } from '@azure/functions';
import { getAuthCodeUrl } from '../../services/graphClient.js';
import { createHash, randomBytes } from 'crypto';
import { withSecurity } from '../../services/securityHeaders.js';
import { withRateLimit } from '../../services/rateLimit.js';

// install_nonce is a PKCE-style code_challenge: SHA256(verifier) in hex.
// The verifier (16 random bytes = 32 hex chars) stays on the CLI; only the
// 64-char SHA256 digest travels in the browser URL. install-poll requires the
// raw verifier — SHA256(verifier) must match the stored challenge. An attacker
// who sees the challenge in the URL cannot poll without inverting SHA256.
const NONCE_FORMAT = /^[a-f0-9]{64}$/;
const DEVICE_LABEL_FORMAT = /^[a-zA-Z0-9._\- ]{1,64}$/;

/**
 * GET /api/auth/login[?install_nonce=XXX]
 *
 * Initiates the MSAL authorization-code flow. Generates a random `state`, a
 * PKCE code verifier and an OpenID Connect `nonce`, stores each in a
 * short-lived HttpOnly cookie, then redirects the browser to the Microsoft
 * identity platform login page with `state`, the S256 code challenge and the
 * nonce. The callback checks `state` and hands the verifier and nonce to MSAL
 * for the code redemption (see AuthCodeUrlBinding in graphClient.ts).
 *
 * If `install_nonce` is provided (SHA256(verifier) from install-mcp.sh),
 * stores the PKCE code_challenge in a short-lived HttpOnly cookie so the OAuth
 * callback can attach the resulting session to it. The install-poll endpoint
 * requires the raw verifier — callers who only know the challenge cannot
 * retrieve the session (PKCE-style binding; see).
 *
 * No table-storage write happens here. The polling client treats "not found"
 * as "still waiting", so there's no race between the script's first poll and
 * the browser navigating to this URL.
 */
async function login(
  request: HttpRequest,
  context: InvocationContext
): Promise<HttpResponseInit> {
  try {
    const state = randomBytes(16).toString('hex');
    // RFC 7636: 32 random bytes give a 43-character base64url verifier.
    const codeVerifier = randomBytes(32).toString('base64url');
    const codeChallenge = createHash('sha256').update(codeVerifier).digest('base64url');
    const nonce = randomBytes(16).toString('hex');
    const authUrl = await getAuthCodeUrl(state, { codeChallenge, nonce });

    const cookies: string[] = [
      `oauth_state=${state}; HttpOnly; Secure; SameSite=Lax; Path=/; Max-Age=600`,
      `oauth_pkce=${codeVerifier}; HttpOnly; Secure; SameSite=Lax; Path=/; Max-Age=600`,
      `oauth_nonce=${nonce}; HttpOnly; Secure; SameSite=Lax; Path=/; Max-Age=600`,
    ];

    const installNonce = request.query.get('install_nonce');
    if (installNonce) {
      if (!NONCE_FORMAT.test(installNonce)) {
        return { status: 400, jsonBody: { error: 'install_nonce must be 64 hex characters (SHA256 of verifier)' } };
      }
      cookies.push(
        `install_nonce=${installNonce}; HttpOnly; Secure; SameSite=Lax; Path=/; Max-Age=600`
      );
    }

    const deviceLabel = request.query.get('device_label');
    if (deviceLabel) {
      if (!DEVICE_LABEL_FORMAT.test(deviceLabel)) {
        return { status: 400, jsonBody: { error: 'device_label must be 1-64 characters (alphanumeric, spaces, hyphens, dots, underscores)' } };
      }
      cookies.push(
        `device_label=${encodeURIComponent(deviceLabel)}; HttpOnly; Secure; SameSite=Lax; Path=/; Max-Age=600`
      );
    }

    return {
      status: 302,
      headers: {
        Location: authUrl,
      },
      cookies: cookies.map((c) => parseSetCookie(c)),
    };
  } catch (err: unknown) {
    // The exception text stays in the log; the response says nothing about why.
    const message = err instanceof Error ? err.message : 'Unknown error';
    context.error('login error:', message);
    return { status: 500, jsonBody: { error: 'Failed to initiate login flow' } };
  }
}

/**
 * Parse a Set-Cookie header string into the structured Cookie object that
 * Azure Functions v4 expects. Only handles the attributes we set.
 */
function parseSetCookie(setCookie: string): {
  name: string;
  value: string;
  httpOnly?: boolean;
  secure?: boolean;
  sameSite?: 'Lax' | 'Strict' | 'None';
  path?: string;
  maxAge?: number;
} {
  const [nameValue, ...attrs] = setCookie.split(';').map((s) => s.trim());
  const eq = nameValue.indexOf('=');
  const name = nameValue.slice(0, eq);
  const value = nameValue.slice(eq + 1);

  const cookie: ReturnType<typeof parseSetCookie> = { name, value };
  for (const attr of attrs) {
    const [k, v] = attr.split('=').map((s) => s.trim());
    const key = k.toLowerCase();
    if (key === 'httponly') cookie.httpOnly = true;
    else if (key === 'secure') cookie.secure = true;
    else if (key === 'samesite') cookie.sameSite = v as 'Lax' | 'Strict' | 'None';
    else if (key === 'path') cookie.path = v;
    else if (key === 'max-age') cookie.maxAge = parseInt(v, 10);
  }
  return cookie;
}

app.http('login', {
  methods: ['GET'],
  authLevel: 'anonymous',
  route: 'api/auth/login',
  handler: withSecurity(withRateLimit('login', login)),
});
