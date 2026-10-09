import { app, HttpRequest, HttpResponseInit, InvocationContext } from '@azure/functions';
import {
  getEmailOutputModeEnforcement,
  getEmailOutputModePolicy,
  setEmailOutputModePolicy,
} from '../../services/userEmailSettings.js';
import { getTenantId } from '../../services/tokenCache.js';
import { authenticateConsoleRequest, checkGlobalAdmin } from '../../services/authMiddleware.js';
import { withSecurity } from '../../services/securityHeaders.js';
import { logAccess } from '../../services/auditLog.js';

/**
 * Admin endpoint for the enforced draft-mode policy.
 *
 * The self-service email output mode can be flipped by the
 * user, or by an agent acting as the user through `set_email_output_mode`.
 * That makes draft mode a guard against unintended sends only: an injected
 * prompt can switch to 'send' and then `send_mail`. This endpoint is how a
 * Global Admin pins the mode so that neither the MCP tool nor
 * POST /api/mail/settings can move it.
 *
 * GET  /api/manage/email-output-policy[?userId=<target>]
 *   Returns the tenant-wide policy. With ?userId, also returns that user's
 *   policy and the resolved enforcement (tenant wins over user).
 *   Requires Global Administrator role.
 *
 * POST /api/manage/email-output-policy
 *   Body: { scope: 'tenant', enforceDraft: boolean }
 *      or { scope: 'user', userId: string, enforceDraft: boolean }
 *   Sets or clears the policy. Requires Global Administrator role.
 *
 * Admin-only for both reads and writes, like /api/manage/mail-config: a user
 * cannot see or lift a policy that applies to them. Every change is written
 * to the audit log table and echoed to the console audit stream.
 */
async function manageEmailOutputPolicy(
  request: HttpRequest,
  context: InvocationContext,
): Promise<HttpResponseInit> {
  try {
    const auth = await authenticateConsoleRequest(request);
    if (!auth) {
      return { status: 401, jsonBody: { error: 'Authentication required' } };
    }

    const isAdmin = await checkGlobalAdmin(auth.userId);
    if (!isAdmin) {
      return { status: 403, jsonBody: { error: 'Global Administrator role required' } };
    }

    const tenantId = await getTenantId(auth.userId);

    if (request.method === 'GET') {
      const targetUserId = request.query.get('userId');
      if (!targetUserId) {
        const tenant = await getEmailOutputModePolicy(tenantId, { scope: 'tenant' });
        return { status: 200, jsonBody: { tenant } };
      }
      const enforcement = await getEmailOutputModeEnforcement(tenantId, targetUserId);
      return {
        status: 200,
        jsonBody: {
          userId: targetUserId,
          tenant: enforcement.tenant,
          user: enforcement.user,
          enforced: enforcement.enforced,
          enforcedBy: enforcement.enforcedBy,
        },
      };
    }

    // POST — set or clear a policy
    const body = (await request.json()) as {
      scope?: string;
      userId?: string;
      enforceDraft?: unknown;
    };

    if (body.scope !== 'tenant' && body.scope !== 'user') {
      return { status: 400, jsonBody: { error: 'scope must be "tenant" or "user"' } };
    }
    if (body.scope === 'user' && (typeof body.userId !== 'string' || body.userId.length === 0)) {
      return { status: 400, jsonBody: { error: 'Missing userId in request body for scope "user"' } };
    }
    if (typeof body.enforceDraft !== 'boolean') {
      return { status: 400, jsonBody: { error: 'Missing or invalid enforceDraft field (must be boolean)' } };
    }

    const target = body.scope === 'tenant'
      ? { scope: 'tenant' as const }
      : { scope: 'user' as const, userId: body.userId as string };
    const policy = await setEmailOutputModePolicy(tenantId, target, body.enforceDraft, auth.userId);

    const resource = target.scope === 'tenant' ? 'tenant' : `user:${target.userId}`;
    logAccess({
      tenantId,
      userId: auth.userId,
      userEmail: auth.session.email,
      deviceLabel: auth.session.deviceLabel,
      operation: 'set_email_output_policy',
      resource,
      result: 'allowed',
      reason: `enforceDraft=${body.enforceDraft}`,
      source: 'http',
    });
    console.log(
      `[audit] email-output-policy-set admin=${auth.userId} target=${resource} enforceDraft=${body.enforceDraft} ts=${policy.updatedAt}`,
    );
    context.log(
      `[admin] email output policy updated: admin=${auth.userId} target=${resource} enforceDraft=${body.enforceDraft}`,
    );

    return {
      status: 200,
      jsonBody: {
        ok: true,
        scope: target.scope,
        ...(target.scope === 'user' ? { userId: target.userId } : {}),
        enforceDraft: policy.enforceDraft,
        updatedAt: policy.updatedAt,
        updatedBy: policy.updatedBy,
      },
    };
  } catch (err: unknown) {
    const message = err instanceof Error ? err.message : String(err);
    context.error('manageEmailOutputPolicy error:', message);
    return { status: 500, jsonBody: { error: message } };
  }
}

app.http('manageEmailOutputPolicy', {
  methods: ['GET', 'POST'],
  authLevel: 'anonymous',
  route: 'api/manage/email-output-policy',
  handler: withSecurity(manageEmailOutputPolicy),
});
