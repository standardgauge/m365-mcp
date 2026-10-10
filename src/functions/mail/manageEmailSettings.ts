/**
 * Per-user email output mode management endpoint.
 *
 * GET  /api/mail/settings
 *   Returns the calling user's email output mode.
 *   Admins may pass ?userId=<target> to query another user.
 *   The response carries the effective mode plus `enforced` / `enforcedBy`
 *   so a client can tell a locked mode from a chosen one.
 *
 * POST /api/mail/settings
 *   Body: { emailOutputMode: 'draft' | 'send', userId?: string }
 *   Sets the email output mode.
 *   - Non-admins can only update their own setting (self-service).
 *   - Global Admins can set the mode for any user.
 *   - Refused with 403 for every caller while an administrator enforces
 *     draft mode for the tenant or for the target user. The
 *     refusal is written to the audit log as a denied access. Enforcement
 *     is lifted through /api/manage/email-output-policy, not here, so a
 *     client that can reach this route (the REST bridge exposes it as
 *     set_email_output_mode) can never talk its way out of the lock.
 *
 * The default mode for any user without an explicit setting is 'draft'.
 */

import { app, HttpRequest, HttpResponseInit, InvocationContext } from '@azure/functions';
import { getUserEmailSettings, setUserEmailSettings } from '../../services/userEmailSettings.js';
import type { EmailOutputMode } from '../../services/userEmailSettings.js';
import { getTenantId } from '../../services/tokenCache.js';
import { authenticateRequest, auditAdminRefusal, checkGlobalAdmin } from '../../services/authMiddleware.js';
import { withSecurity } from '../../services/securityHeaders.js';
import { auditSnapshot, logAccess } from '../../services/auditLog.js';
import { SessionStoreUnavailableError, SESSION_STORE_RETRY_AFTER_S } from '../../services/sessionStoreError.js';

async function manageEmailSettingsHandler(
  request: HttpRequest,
  context: InvocationContext,
): Promise<HttpResponseInit> {
  try {
    const auth = await authenticateRequest(request);
    if (!auth) {
      return { status: 401, jsonBody: { error: 'Authentication required' } };
    }

    const tenantId = await getTenantId(auth.userId);
    const isAdmin = await checkGlobalAdmin(auth.userId);

    if (request.method === 'GET') {
      const targetUserId = request.query.get('userId') ?? auth.userId;
      if (!isAdmin && targetUserId !== auth.userId) {
        return { status: 403, jsonBody: { error: 'You can only view your own email settings' } };
      }
      const settings = await getUserEmailSettings(tenantId, targetUserId);
      return {
        status: 200,
        jsonBody: {
          userId: targetUserId,
          emailOutputMode: settings.emailOutputMode,
          preferredEmailOutputMode: settings.preferredEmailOutputMode,
          enforced: settings.enforced === true,
          enforcedBy: settings.enforcedBy ?? null,
        },
      };
    }

    // POST — update email output mode
    const body = (await request.json()) as { emailOutputMode?: string; userId?: string };
    const targetUserId = body.userId ?? auth.userId;

    if (!isAdmin && targetUserId !== auth.userId) {
      auditAdminRefusal(auth, 'set_email_output_mode', targetUserId);
      return { status: 403, jsonBody: { error: 'You can only modify your own email settings' } };
    }

    if (!body.emailOutputMode || !['draft', 'send'].includes(body.emailOutputMode)) {
      return { status: 400, jsonBody: { error: 'emailOutputMode must be "draft" or "send"' } };
    }

    // Enforced draft mode. Checked after validation so a malformed body never
    // costs a storage read, and before the write so nothing lands while a policy applies.
    const current = await getUserEmailSettings(tenantId, targetUserId);
    if (current.enforced) {
      const scope = current.enforcedBy === 'tenant' ? 'for this tenant' : 'for this user';
      const error =
        `Email output mode is enforced to draft by an administrator ${scope} and cannot be changed here. ` +
        'Lift the policy under Email Output Policy in the admin UI first.';
      logAccess({
        tenantId,
        userId: auth.userId,
        userEmail: auth.session.email,
        deviceLabel: auth.session.deviceLabel,
        operation: 'set_email_output_mode',
        resource: targetUserId,
        result: 'denied',
        reason: `draft mode enforced by ${current.enforcedBy} policy; requested '${body.emailOutputMode}'`,
        source: 'http',
      });
      return { status: 403, jsonBody: { error, enforced: true, enforcedBy: current.enforcedBy } };
    }

    await setUserEmailSettings(tenantId, targetUserId, {
      emailOutputMode: body.emailOutputMode as EmailOutputMode,
    });
    logAccess({
      tenantId,
      userId: auth.userId,
      userEmail: auth.session.email,
      deviceLabel: auth.session.deviceLabel,
      operation: 'set_email_output_mode',
      resource: targetUserId,
      result: 'allowed',
      source: 'http',
      before: auditSnapshot({ emailOutputMode: current.preferredEmailOutputMode }),
      after: auditSnapshot({ emailOutputMode: body.emailOutputMode }),
    });

    return {
      status: 200,
      jsonBody: { ok: true, userId: targetUserId, emailOutputMode: body.emailOutputMode },
    };
  } catch (err: unknown) {
    const message = err instanceof Error ? err.message : String(err);
    context.error('manageEmailSettings error:', message);
    if (err instanceof SessionStoreUnavailableError) {
      return { status: 503, headers: { 'Retry-After': String(SESSION_STORE_RETRY_AFTER_S) }, jsonBody: { error: 'Session store unavailable, retry shortly' } };
    }
    return { status: 500, jsonBody: { error: message } };
  }
}

app.http('manageEmailSettings', {
  methods: ['GET', 'POST'],
  authLevel: 'anonymous',
  route: 'api/mail/settings',
  handler: withSecurity(manageEmailSettingsHandler),
});
