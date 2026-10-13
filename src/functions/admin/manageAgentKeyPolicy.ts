import { app, HttpRequest, HttpResponseInit, InvocationContext } from '@azure/functions';
import { authenticateConsoleRequest, authorizeAdmin } from '../../services/authMiddleware.js';
import { auditActor, auditSnapshot, logAccess } from '../../services/auditLog.js';
import { withSecurity } from '../../services/securityHeaders.js';
import { getTenantIdFromSession } from '../../services/tokenCache.js';
import { ALL_SERVICE_KEYS } from '../../services/serviceSettings.js';
import { AGENT_KEY_MAX_LIFETIME_DAYS, getTenantAgentKeyPolicy, normalizeTenantAgentKeyPolicy, setTenantAgentKeyPolicy } from '../../services/agentKeys.js';
import type { TenantAgentKeyPolicy } from '../../services/agentKeys.js';

/**
 * Tenant-wide policy for agent API keys (services/agentKeys.ts).
 *
 * GET  /api/manage/agent-key-policy   any signed-in console user, so the key
 *                                     form can show what a key may be granted
 * POST /api/manage/agent-key-policy   Global Administrator
 *      { allowKeys?, disabledServices?, readOnlyServices?, maxKeyLifetimeDays? }
 *
 * The policy is read on every key-authenticated request, so a change here
 * applies to existing keys at once; nothing has to be re-minted.
 */
async function manageAgentKeyPolicy(request: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> {
  try {
    const auth = await authenticateConsoleRequest(request);
    if (!auth) return { status: 401, jsonBody: { error: 'Authentication required' } };
    const tenantId = getTenantIdFromSession(auth.session);

    if (request.method === 'GET') {
      return { status: 200, jsonBody: await getTenantAgentKeyPolicy(tenantId) };
    }

    const isAdmin = await authorizeAdmin(auth, 'policy.agent_key_policy.set', 'tenant');
    if (!isAdmin) return { status: 403, jsonBody: { error: 'Global Administrator role required' } };

    const body = (await request.json().catch(() => null)) as Partial<Record<keyof TenantAgentKeyPolicy, unknown>> | null;
    if (!body || typeof body !== 'object') return { status: 400, jsonBody: { error: 'JSON body required' } };
    if (body.allowKeys === undefined && body.disabledServices === undefined && body.readOnlyServices === undefined && body.maxKeyLifetimeDays === undefined) {
      return { status: 400, jsonBody: { error: 'Provide allowKeys, disabledServices, readOnlyServices and/or maxKeyLifetimeDays' } };
    }
    if (body.maxKeyLifetimeDays !== undefined) {
      const d = Number(body.maxKeyLifetimeDays);
      if (!Number.isInteger(d) || d < 1 || d > AGENT_KEY_MAX_LIFETIME_DAYS) {
        return { status: 400, jsonBody: { error: `maxKeyLifetimeDays must be an integer from 1 to ${AGENT_KEY_MAX_LIFETIME_DAYS}` } };
      }
    }
    for (const field of ['disabledServices', 'readOnlyServices'] as const) {
      // eslint-disable-next-line security/detect-object-injection -- field iterates a literal tuple
      const v = body[field];
      if (v === undefined) continue;
      if (!Array.isArray(v) || v.some((s) => typeof s !== 'string')) {
        return { status: 400, jsonBody: { error: `${field} must be an array of strings` } };
      }
      const unknown = (v as string[]).filter((s) => !ALL_SERVICE_KEYS.includes(s));
      if (unknown.length > 0) return { status: 400, jsonBody: { error: `Unknown service key(s) in ${field}: ${unknown.join(', ')}` } };
    }
    if (body.allowKeys !== undefined && typeof body.allowKeys !== 'boolean') {
      return { status: 400, jsonBody: { error: 'allowKeys must be a boolean' } };
    }

    const before = await getTenantAgentKeyPolicy(tenantId);
    const next = normalizeTenantAgentKeyPolicy({
      allowKeys: body.allowKeys === undefined ? before.allowKeys : (body.allowKeys as boolean),
      disabledServices: body.disabledServices === undefined ? before.disabledServices : (body.disabledServices as string[]),
      readOnlyServices: body.readOnlyServices === undefined ? before.readOnlyServices : (body.readOnlyServices as string[]),
      maxKeyLifetimeDays: body.maxKeyLifetimeDays === undefined ? before.maxKeyLifetimeDays : Number(body.maxKeyLifetimeDays),
    });
    await setTenantAgentKeyPolicy(tenantId, next);
    logAccess({
      ...auditActor(auth.session),
      operation: 'policy.agent_key_policy.set',
      resource: 'tenant',
      result: 'allowed',
      source: 'http',
      before: auditSnapshot(before),
      after: auditSnapshot(next),
    });
    return { status: 200, jsonBody: next };
  } catch (err: unknown) {
    const message = err instanceof Error ? err.message : String(err);
    context.error('manageAgentKeyPolicy error:', message);
    return { status: 500, jsonBody: { error: 'Internal server error' } };
  }
}

app.http('manageAgentKeyPolicy', {
  methods: ['GET', 'POST'],
  authLevel: 'anonymous',
  route: 'api/manage/agent-key-policy',
  handler: withSecurity(manageAgentKeyPolicy),
});
