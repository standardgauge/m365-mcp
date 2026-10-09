import { app, HttpRequest, HttpResponseInit, InvocationContext } from '@azure/functions';
import { createHash } from 'crypto';
import { consumeInstallNonce } from '../../services/tableStorage.js';
import { withSecurity } from '../../services/securityHeaders.js';
import { withRateLimit } from '../../services/rateLimit.js';
import { auditTenantId, logAccess } from '../../services/auditLog.js';

// The verifier is 16 random bytes in hex (32 chars). The CLI keeps the verifier
// and sends only SHA256(verifier) — the code_challenge — in the login URL.
// An attacker who intercepts the challenge cannot poll: they'd need to invert SHA256.
const VERIFIER_FORMAT = /^[a-f0-9]{32}$/;

/**
 * GET /api/auth/install-poll?nonce_verifier=XXX
 *
 * Companion to install-mcp.sh's hands-free auth flow. PKCE-style binding:
 * the CLI generates a random verifier, sends SHA256(verifier) as install_nonce
 * in the login URL, and presents the raw verifier here. The server re-derives
 * the challenge and looks up the session — only the entity that generated the
 * verifier can retrieve the attached session.
 *
 * Responses:
 *   - 200 + {sessionToken, userId, email, displayName} once the OAuth callback
 *     has attached the session. The record is consumed (one-time-use) on read.
 *   - 202 if the OAuth flow has not completed yet.
 *   - 400 if the nonce_verifier parameter is missing or malformed.
 *   - 410 if the nonce is unknown, expired, or already consumed.
 *
 * The script hands the returned sessionToken to the stdio shim, which keeps it
 * in the OS credential store; the MCP client config names only the store.
 * No further use of the userId is needed for auth — the token is the credential.
 */
async function installPoll(
  request: HttpRequest,
  _context: InvocationContext
): Promise<HttpResponseInit> {
  const verifier = request.query.get('nonce_verifier');
  if (!verifier) {
    return { status: 400, jsonBody: { error: 'Missing nonce_verifier query parameter' } };
  }
  if (!VERIFIER_FORMAT.test(verifier)) {
    return { status: 400, jsonBody: { error: 'nonce_verifier must be 32 hex characters' } };
  }

  // Derive the code_challenge from the verifier. The login endpoint embedded
  // SHA256(verifier) in the cookie; the callback stored the session under
  // sha256(SHA256(verifier)). Only the entity that generated verifier can
  // reproduce this chain — a chosen-challenge attacker cannot.
  const challenge = createHash('sha256').update(verifier).digest('hex');

  const result = await consumeInstallNonce(challenge);
  if (result === null) {
    // A verifier whose record has expired. Unknown verifiers read as pending
    // (see consumeInstallNonce), so they are not recorded here.
    logAccess({
      tenantId: auditTenantId(),
      userId: '',
      userEmail: '',
      operation: 'auth.install_handoff',
      resource: 'poll',
      result: 'denied',
      reason: 'install nonce expired',
      source: 'http',
    });
    return {
      status: 410,
      jsonBody: { error: 'Nonce unknown, expired, or already consumed' },
    };
  }
  if (result === 'pending') {
    return {
      status: 202,
      jsonBody: { status: 'pending' },
    };
  }

  logAccess({
    tenantId: auditTenantId(result.tenantId),
    userId: result.userId,
    userEmail: result.email,
    deviceLabel: result.deviceLabel,
    operation: 'auth.install_handoff',
    resource: 'poll',
    result: 'allowed',
    reason: 'session token handed to the installer',
    source: 'http',
  });

  return {
    status: 200,
    jsonBody: {
      sessionToken: result.sessionToken,
      userId: result.userId,
      email: result.email,
      displayName: result.displayName,
    },
  };
}

app.http('install-poll', {
  methods: ['GET'],
  authLevel: 'anonymous',
  route: 'api/auth/install-poll',
  handler: withSecurity(withRateLimit('install-poll', installPoll)),
});
