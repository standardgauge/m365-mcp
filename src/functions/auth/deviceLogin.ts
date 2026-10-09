import { app, HttpRequest, HttpResponseInit, InvocationContext } from '@azure/functions';
import { getPublicMsalApp, GRAPH_SCOPES, createGraphClient } from '../../services/graphClient.js';
import { storeSession } from '../../services/tokenCache.js';
import type { UserSession } from '../../services/tokenCache.js';
import { extractTenantId } from '../../services/tenantUtils.js';
import { randomBytes } from 'crypto';
import { withSecurity } from '../../services/securityHeaders.js';
import { withRateLimit } from '../../services/rateLimit.js';
import { auditActor, auditTenantId, logAccess } from '../../services/auditLog.js';

/** A device-code sign-in that did not produce a session. */
function auditDeviceLoginFailure(reason: string, resource?: string): void {
  logAccess({
    tenantId: auditTenantId(),
    userId: '',
    userEmail: '',
    operation: 'auth.device_login',
    resource,
    result: 'denied',
    reason,
    source: 'http',
  });
}

/**
 * GET /api/auth/device
 *
 * Initiates a Device Code Flow login. Returns the userCode and verificationUri
 * immediately so the caller can display them. MSAL polls in the background; once
 * the user completes authentication the session is stored and the polling promise
 * resolves with the token response.
 *
 * Use this when the browser-redirect flow is blocked by tenant CA/SSO policies.
 *
 * The outcome is audited as `auth.device_login` when the flow finishes, which
 * can be up to fifteen minutes after this request returned.
 */
async function deviceLogin(
  _request: HttpRequest,
  context: InvocationContext
): Promise<HttpResponseInit> {
  const msalApp = getPublicMsalApp();

  // We need to surface the device code info to the caller before MSAL starts
  // polling, so we capture it via the callback and resolve a promise with it.
  let resolveDeviceCode!: (info: { userCode: string; verificationUri: string; message: string }) => void;
  const deviceCodeInfo = new Promise<{ userCode: string; verificationUri: string; message: string }>(
    (resolve) => { resolveDeviceCode = resolve; }
  );

  // Start the device code flow — this call does NOT resolve until the user
  // authenticates (or it times out). We run it in the background.
  const tokenPromise = msalApp.acquireTokenByDeviceCode({
    scopes: GRAPH_SCOPES,
    deviceCodeCallback: (response) => {
      resolveDeviceCode({
        userCode: response.userCode,
        verificationUri: response.verificationUri,
        message: response.message,
      });
    },
  });

  // Wait for the callback to fire (happens within the first few ms, before any polling)
  const codeInfo = await deviceCodeInfo;

  // Store the session once the user completes auth — fire and forget from the
  // perspective of this HTTP response, which returns immediately with the code.
  tokenPromise
    .then(async (response) => {
      if (!response?.accessToken || !response.account) {
        context.error('deviceLogin: token response missing accessToken or account');
        auditDeviceLoginFailure('token response missing accessToken or account');
        return;
      }

      const graph = createGraphClient(response.accessToken);
      const me = await graph
        .api('/me')
        .select('id,displayName,mail,userPrincipalName')
        .get();

      const tenantId = extractTenantId(response.account.homeAccountId);

      // Enforce single-tenant: reject users from foreign tenants
      const expectedTenantId = process.env.AZURE_TENANT_ID;
      if (expectedTenantId && tenantId !== expectedTenantId) {
        context.error(`deviceLogin: tenant mismatch — user tenant ${tenantId} does not match expected ${expectedTenantId}`);
        auditDeviceLoginFailure('foreign tenant', `tenant:${tenantId}`);
        return;
      }

      const sessionToken = randomBytes(32).toString('hex');
      const session: UserSession = {
        userId: me.id,
        homeAccountId: response.account.homeAccountId,
        displayName: me.displayName ?? me.userPrincipalName,
        email: me.mail ?? me.userPrincipalName,
        tenantId,
        accessToken: response.accessToken,
        expiresAt: response.expiresOn ? response.expiresOn.getTime() : Date.now() + 3_600_000,
        sessionToken,
        sessionCreatedAt: Date.now(),
        sessionAbsoluteCreatedAt: Date.now(),
      };

      await storeSession(session);
      // The address is the one that started the flow, not the device the user
      // signed in on: the server never sees the latter.
      logAccess({ ...auditActor(session), operation: 'auth.device_login', result: 'allowed', source: 'http' });
      context.log(`deviceLogin: session stored for ${session.email} (${session.userId})`);
    })
    .catch((err: unknown) => {
      const message = err instanceof Error ? err.message : String(err);
      context.error('deviceLogin: token acquisition failed:', message);
      auditDeviceLoginFailure('token acquisition failed');
    });

  return {
    status: 200,
    jsonBody: {
      status: 'pending',
      userCode: codeInfo.userCode,
      verificationUri: codeInfo.verificationUri,
      message: codeInfo.message,
    },
  };
}

app.http('deviceLogin', {
  methods: ['GET'],
  authLevel: 'anonymous',
  route: 'api/auth/device',
  handler: withSecurity(withRateLimit('device', deviceLogin)),
});
