import { app, HttpRequest, HttpResponseInit, InvocationContext } from '@azure/functions';
import { acquireTokenByCode, createGraphClient } from '../../services/graphClient.js';
import { storeSession, SESSION_TTL_MS } from '../../services/tokenCache.js';
import type { UserSession } from '../../services/tokenCache.js';
import { extractTenantId } from '../../services/tenantUtils.js';
import { attachSessionToInstallNonce } from '../../services/tableStorage.js';
import { randomBytes } from 'crypto';
import { withSecurity } from '../../services/securityHeaders.js';
import { resolveFrontendUrl } from '../../services/frontendUrl.js';

const NONCE_TTL_MS = 5 * 60 * 1000;

/**
 * GET /api/auth/callback
 *
 * Handles the redirect from the Microsoft identity platform after the user
 * authenticates. Validates the OAuth state parameter (CSRF protection),
 * exchanges the authorization code for tokens, fetches the user profile
 * via Graph, generates a session token, stores the session, then redirects
 * to the admin UI.
 */
async function callback(
  request: HttpRequest,
  context: InvocationContext
): Promise<HttpResponseInit> {
  try {
    const error = request.query.get('error');
    const errorDescription = request.query.get('error_description');

    if (error) {
      context.error('OAuth error from identity platform:', error, errorDescription);
      return {
        status: 400,
        jsonBody: { error, description: errorDescription ?? 'No description provided' },
      };
    }

    // ── Validate OAuth state parameter (CSRF protection) ──
    const stateParam = request.query.get('state');
    const cookieHeader = request.headers.get('cookie') ?? '';
    const stateMatch = cookieHeader.match(/(?:^|;\s*)oauth_state=([^;]+)/);
    const stateCookie = stateMatch ? stateMatch[1] : null;

    if (!stateParam || !stateCookie || stateParam !== stateCookie) {
      context.error('OAuth state mismatch — possible CSRF attack');
      return {
        status: 403,
        jsonBody: { error: 'State parameter mismatch. Please try logging in again.' },
      };
    }

    // Optional: install-nonce cookie set by /api/auth/login when called from
    // install-mcp.sh. If present, the install script is polling install-poll
    // and we need to attach the new session to the nonce slot it reserved.
    const nonceMatch = cookieHeader.match(/(?:^|;\s*)install_nonce=([^;]+)/);
    const installNonce = nonceMatch ? nonceMatch[1] : null;

    const deviceLabelMatch = cookieHeader.match(/(?:^|;\s*)device_label=([^;]+)/);
    const deviceLabel = deviceLabelMatch ? decodeURIComponent(deviceLabelMatch[1]) : undefined;

    const code = request.query.get('code');
    if (!code) {
      return { status: 400, jsonBody: { error: 'Missing authorization code in callback' } };
    }

    // Exchange the code for tokens
    const { accessToken, homeAccountId, expiresOn } = await acquireTokenByCode(code);
    const tenantId = extractTenantId(homeAccountId);

    // Enforce single-tenant: reject users from foreign tenants
    const expectedTenantId = process.env.AZURE_TENANT_ID;
    if (expectedTenantId && tenantId !== expectedTenantId) {
      context.error(`Tenant mismatch: user tenant ${tenantId} does not match expected ${expectedTenantId}`);
      return {
        status: 403,
        jsonBody: { error: 'Access denied. Your account belongs to a different tenant.' },
      };
    }

    // Fetch the authenticated user's profile
    const graph = createGraphClient(accessToken);
    const me = await graph
      .api('/me')
      .select('id,displayName,mail,userPrincipalName')
      .get();

    // Generate a crypto-random session token — this is the authentication credential
    const sessionToken = randomBytes(32).toString('hex');

    const session: UserSession = {
      userId: me.id,
      homeAccountId,
      displayName: me.displayName ?? me.userPrincipalName,
      email: me.mail ?? me.userPrincipalName,
      tenantId,
      accessToken,
      expiresAt: expiresOn ? expiresOn.getTime() : Date.now() + 3_600_000,
      sessionToken,
      sessionCreatedAt: Date.now(),
      sessionAbsoluteCreatedAt: Date.now(),
      deviceLabel,
    };

    await storeSession(session);

    // If this OAuth flow was started by install-mcp.sh, attach the new
    // session to its install-nonce slot so the polling install script can
    // pick it up. We don't fail the OAuth flow if attach fails — the user
    // can still use the admin UI; only the install script flow is degraded.
    if (installNonce) {
      const attached = await attachSessionToInstallNonce(installNonce, {
        sessionToken,
        userId: session.userId,
        email: session.email,
        displayName: session.displayName,
        deviceLabel: session.deviceLabel,
        expiresAt: Date.now() + NONCE_TTL_MS,
      });
      if (!attached) {
        context.warn(
          `install_nonce attach failed — nonce expired or never reserved (userId=${session.userId})`
        );
      }
    }

    // The admin SPA is served from the site root by the serveAdmin catch-all.
    // resolveFrontendUrl refuses a value whose path the Functions host reserves
    // (notably '/admin'), since such a redirect 404s before reaching any
    // function. See docs/operations-runbook.md → "Reserved paths".
    const frontendUrl = resolveFrontendUrl(process.env.FRONTEND_URL, (m) =>
      context.error(m)
    );
    return {
      status: 302,
      headers: {
        // No userId query param: the SPA resolves identity from /api/auth/me
        // against the session cookie set below, and putting the GUID in the URL
        // only leaks it into browser history and the onward Referer.
        Location: frontendUrl,
      },
      // Azure Functions v4 cookie API — each cookie is a separate Set-Cookie header
      // (comma-joining Set-Cookie is invalid per HTTP spec)
      cookies: [
        {
          name: 'mcp_session',
          value: sessionToken,
          httpOnly: true,
          secure: true,
          sameSite: 'Lax' as const,
          path: '/',
          maxAge: SESSION_TTL_MS / 1000,
        },
        {
          name: 'user_id',
          value: session.userId,
          secure: true,
          sameSite: 'Lax' as const,
          path: '/',
          maxAge: SESSION_TTL_MS / 1000,
        },
        {
          name: 'user_name',
          value: encodeURIComponent(session.displayName),
          secure: true,
          sameSite: 'Lax' as const,
          path: '/',
          maxAge: SESSION_TTL_MS / 1000,
        },
        {
          name: 'oauth_state',
          value: '',
          httpOnly: true,
          secure: true,
          sameSite: 'Lax' as const,
          path: '/',
          maxAge: 0,
        },
        {
          name: 'install_nonce',
          value: '',
          httpOnly: true,
          secure: true,
          sameSite: 'Lax' as const,
          path: '/',
          maxAge: 0,
        },
        {
          name: 'device_label',
          value: '',
          httpOnly: true,
          secure: true,
          sameSite: 'Lax' as const,
          path: '/',
          maxAge: 0,
        },
      ],
    };
  } catch (err: unknown) {
    const message = err instanceof Error ? err.message : 'Unknown error';
    context.error('callback error:', message);
    return { status: 500, jsonBody: { error: 'Authentication callback failed' } };
  }
}

app.http('callback', {
  methods: ['GET'],
  authLevel: 'anonymous',
  route: 'api/auth/callback',
  handler: withSecurity(callback),
});
