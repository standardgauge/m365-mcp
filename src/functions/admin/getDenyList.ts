import { app, HttpRequest, HttpResponseInit, InvocationContext } from '@azure/functions';
import {
  listGlobalDenyEntries,
  addGlobalDenyEntry,
  removeGlobalDenyEntry,
  listUserDenyEntries,
  addUserDenyEntry,
  removeUserDenyEntry,
  clearUserDenyList,
  agentKeyDenySubject,
  isAgentKeyDenySubject,
  DENY_LIST_TYPES,
  type DenyListType,
} from '../../services/denyList.js';
import { getTenantId, getTenantIdFromSession } from '../../services/tokenCache.js';
import { authenticateConsoleRequest, authorizeAdmin, auditAdminRefusal, checkGlobalAdmin } from '../../services/authMiddleware.js';
import type { ConsoleAuthResult } from '../../services/authMiddleware.js';
import { findAgentKey, getTenantAgentKeyPolicy } from '../../services/agentKeys.js';
import { auditActor, auditSnapshot, logAccess } from '../../services/auditLog.js';

/** The fields of a deny entry worth recording, or null when there is none. */
function entrySnapshot(entry: { path: string; description?: string } | undefined): string {
  return auditSnapshot(entry ? { path: entry.path, description: entry.description ?? '' } : null);
}
import { withSecurity } from '../../services/securityHeaders.js';

async function denyListGlobal(
  request: HttpRequest,
  context: InvocationContext
): Promise<HttpResponseInit> {
  try {
    const type = (request.query.get('type') ?? 'sharepoint') as DenyListType;
    const auth = await authenticateConsoleRequest(request);
    if (!auth) {
      return { status: 401, jsonBody: { error: 'Authentication required' } };
    }
    const addedBy = auth.userId;
    const tenantId = await getTenantId(addedBy);

    if (request.method === 'GET') {
      const entries = await listGlobalDenyEntries(tenantId, type);
      return { status: 200, jsonBody: { entries } };
    }

    // POST and DELETE require Global Admin
    const operation = request.method === 'POST' ? 'policy.deny_list.global.add' : 'policy.deny_list.global.remove';
    const isAdmin = await authorizeAdmin(auth, operation);
    if (!isAdmin) {
      return { status: 403, jsonBody: { error: 'Global Administrator role required' } };
    }

    const body = await request.json() as { type?: DenyListType; path?: string; description?: string; addedByName?: string };
    const entryType = body.type ?? type;
    const path = body.path;
    if (!path) return { status: 400, jsonBody: { error: 'Missing path' } };

    const before = (await listGlobalDenyEntries(tenantId, entryType)).find((e) => e.path === path);
    if (request.method === 'POST') {
      await addGlobalDenyEntry(tenantId, entryType, path, addedBy, body.description ?? '', body.addedByName);
    } else {
      await removeGlobalDenyEntry(tenantId, entryType, path);
    }
    logAccess({
      ...auditActor(auth.session),
      operation,
      resource: `${entryType}:${path}`,
      result: 'allowed',
      source: 'http',
      before: entrySnapshot(before),
      after: entrySnapshot(request.method === 'POST' ? { path, description: body.description } : undefined),
    });
    return { status: 200, jsonBody: { ok: true } };
  } catch (err: unknown) {
    const message = err instanceof Error ? err.message : String(err);
    context.error('denyListGlobal error:', message);
    return { status: request.method === 'GET' ? 200 : 500, jsonBody: request.method === 'GET' ? { entries: [] } : { error: message } };
  }
}

/**
 * One agent key's own deny entries: `keyId` on a deny-list/user
 * request. Stored as tier 2 rows under the key's subject (denyList.ts), so a
 * key can be kept out of a folder, mail folder or calendar its owner still
 * sees.
 *
 * The key's owner may read and change them. A Global Administrator may read
 * any key's and add to it, but not remove, since removing an entry widens the
 * key and an administrator can only narrow another user's key (the same rule
 * as PATCH /api/manage/agent-keys). An owner removing an entry is a widening
 * too, so it is refused when the tenant's key policy has ownersMayWiden off.
 */
async function denyListAgentKey(
  request: HttpRequest,
  auth: ConsoleAuthResult,
  keyId: string,
  type: DenyListType,
  body: { path?: string; addedByName?: string } | null,
): Promise<HttpResponseInit> {
  if (!DENY_LIST_TYPES.includes(type)) return { status: 400, jsonBody: { error: `Unknown deny-list type: ${type}` } };
  const tenantId = getTenantIdFromSession(auth.session);
  const record = await findAgentKey(tenantId, keyId);
  if (!record) return { status: 404, jsonBody: { error: 'No such key' } };
  const isOwner = record.userId === auth.userId;
  const subject = agentKeyDenySubject(record.keyId);

  if (request.method === 'GET') {
    if (!isOwner && !(await authorizeAdmin(auth, 'admin.deny_list.agent_key.read', `key:${keyId}`))) {
      return { status: 403, jsonBody: { error: 'You can only view your own keys' } };
    }
    const entries = await listUserDenyEntries(subject, type);
    return { status: 200, jsonBody: { entries } };
  }

  const adding = request.method === 'POST';
  const operation = adding ? 'policy.deny_list.agent_key.add' : 'policy.deny_list.agent_key.remove';
  const path = body?.path;
  if (!path) return { status: 400, jsonBody: { error: 'Missing path' } };
  const resource = `key:${keyId}/${type}:${path}`;
  if (record.revokedAt) return { status: 409, jsonBody: { error: 'The key is revoked; mint a new one' } };

  if (!isOwner) {
    if (!(await authorizeAdmin(auth, operation, resource))) {
      return { status: 403, jsonBody: { error: 'You can only change your own keys' } };
    }
    if (!adding) {
      auditAdminRefusal(auth, operation, resource);
      return { status: 403, jsonBody: { error: "An administrator can only narrow another user's key" } };
    }
  } else if (!adding && !(await getTenantAgentKeyPolicy(tenantId)).ownersMayWiden) {
    logAccess({
      ...auditActor(auth.session),
      operation,
      resource,
      result: 'denied',
      reason: 'tenant policy allows owners to narrow keys only',
      source: 'http',
    });
    return { status: 403, jsonBody: { error: 'Your organization only allows a key to be narrowed; mint a new key for wider access' } };
  }

  const before = (await listUserDenyEntries(subject, type)).find((e) => e.path === path);
  if (adding) {
    await addUserDenyEntry(subject, type, path, body?.addedByName);
  } else {
    await removeUserDenyEntry(subject, type, path);
  }
  logAccess({
    ...auditActor(auth.session),
    operation,
    resource,
    result: 'allowed',
    source: 'http',
    before: entrySnapshot(before),
    after: entrySnapshot(adding ? { path } : undefined),
  });
  return { status: 200, jsonBody: { ok: true } };
}

async function denyListUser(
  request: HttpRequest,
  context: InvocationContext
): Promise<HttpResponseInit> {
  try {
    const type = (request.query.get('type') ?? 'sharepoint') as DenyListType;
    const auth = await authenticateConsoleRequest(request);
    if (!auth) {
      return { status: 401, jsonBody: { error: 'Authentication required' } };
    }
    const requesterId = auth.userId;

    if (request.method === 'GET') {
      const keyId = request.query.get('keyId');
      if (keyId) return await denyListAgentKey(request, auth, keyId, type, null);
      const targetUserId = request.query.get('targetUserId') ?? requesterId;
      if (isAgentKeyDenySubject(targetUserId)) return { status: 400, jsonBody: { error: 'Use keyId for a key\'s deny list' } };
      if (!targetUserId) return { status: 400, jsonBody: { error: 'Missing targetUserId or x-user-id' } };
      // IDOR protection: non-admins can only read their own deny list
      if (targetUserId !== requesterId) {
        const isAdmin = await authorizeAdmin(auth, 'admin.deny_list.user.read', `user:${targetUserId}`);
        if (!isAdmin) {
          return { status: 403, jsonBody: { error: 'You can only view your own deny list' } };
        }
      }
      const entries = await listUserDenyEntries(targetUserId, type);
      return { status: 200, jsonBody: { entries } };
    }

    const body = await request.json() as { type?: DenyListType; path?: string; targetUserId?: string; keyId?: string; addedByName?: string };
    const entryType = body.type ?? type;
    if (body.keyId) return await denyListAgentKey(request, auth, body.keyId, entryType, body);
    const path = body.path;
    const targetUserId = body.targetUserId ?? requesterId;
    if (!path) return { status: 400, jsonBody: { error: 'Missing path' } };
    if (!targetUserId) return { status: 400, jsonBody: { error: 'Missing targetUserId' } };
    // A key's rows share this table under a synthetic subject; they are
    // reached through keyId, where the narrow-only rules apply, never here.
    if (isAgentKeyDenySubject(targetUserId)) return { status: 400, jsonBody: { error: 'Use keyId for a key\'s deny list' } };

    const operation = request.method === 'POST' ? 'policy.deny_list.user.add' : 'policy.deny_list.user.remove';
    const resource = `user:${targetUserId}/${entryType}:${path}`;

    // IDOR protection: non-admins can only modify their own deny list
    if (targetUserId !== requesterId) {
      const isAdmin = await authorizeAdmin(auth, operation, resource);
      if (!isAdmin) {
        return { status: 403, jsonBody: { error: 'You can only modify your own deny list' } };
      }
    }

    const before = (await listUserDenyEntries(targetUserId, entryType)).find((e) => e.path === path);
    if (request.method === 'POST') {
      await addUserDenyEntry(targetUserId, entryType, path, body.addedByName);
    } else {
      await removeUserDenyEntry(targetUserId, entryType, path);
    }
    logAccess({
      ...auditActor(auth.session),
      operation,
      resource,
      result: 'allowed',
      source: 'http',
      before: entrySnapshot(before),
      after: entrySnapshot(request.method === 'POST' ? { path } : undefined),
    });
    return { status: 200, jsonBody: { ok: true } };
  } catch (err: unknown) {
    const message = err instanceof Error ? err.message : String(err);
    context.error('denyListUser error:', message);
    return { status: request.method === 'GET' ? 200 : 500, jsonBody: request.method === 'GET' ? { entries: [] } : { error: message } };
  }
}

async function denyListUserClear(
  request: HttpRequest,
  context: InvocationContext
): Promise<HttpResponseInit> {
  try {
    const auth = await authenticateConsoleRequest(request);
    if (!auth) {
      return { status: 401, jsonBody: { error: 'Authentication required' } };
    }
    const body = await request.json() as { targetUserId?: string };
    const targetUserId = body.targetUserId ?? auth.userId;
    if (!targetUserId) return { status: 400, jsonBody: { error: 'Missing targetUserId' } };
    if (isAgentKeyDenySubject(targetUserId)) return { status: 400, jsonBody: { error: 'Use keyId for a key\'s deny list' } };

    // IDOR protection: non-admins can only clear their own deny list
    if (targetUserId !== auth.userId) {
      const isAdmin = await authorizeAdmin(auth, 'policy.deny_list.user.clear', `user:${targetUserId}`);
      if (!isAdmin) {
        return { status: 403, jsonBody: { error: 'You can only clear your own deny list' } };
      }
    }

    const [sharepoint, mail] = await Promise.all([
      listUserDenyEntries(targetUserId, 'sharepoint'),
      listUserDenyEntries(targetUserId, 'mail'),
    ]);
    await Promise.all([
      clearUserDenyList(targetUserId, 'sharepoint'),
      clearUserDenyList(targetUserId, 'mail'),
    ]);
    logAccess({
      ...auditActor(auth.session),
      operation: 'policy.deny_list.user.clear',
      resource: `user:${targetUserId}`,
      result: 'allowed',
      source: 'http',
      before: auditSnapshot({ sharepoint: sharepoint.map((e) => e.path), mail: mail.map((e) => e.path) }),
      after: auditSnapshot({ sharepoint: [], mail: [] }),
    });
    return { status: 200, jsonBody: { ok: true } };
  } catch (err: unknown) {
    const message = err instanceof Error ? err.message : String(err);
    context.error('denyListUserClear error:', message);
    return { status: 500, jsonBody: { error: message } };
  }
}

async function isGlobalAdmin(
  request: HttpRequest,
  _context: InvocationContext
): Promise<HttpResponseInit> {
  const auth = await authenticateConsoleRequest(request);
  if (!auth) return { status: 200, jsonBody: { isAdmin: false } };
  const isAdmin = await checkGlobalAdmin(auth.userId);
  return { status: 200, jsonBody: { isAdmin } };
}

app.http('denyListGlobal', {
  methods: ['GET', 'POST', 'DELETE'],
  authLevel: 'anonymous',
  route: 'api/manage/deny-list/global',
  handler: withSecurity(denyListGlobal),
});

app.http('denyListUser', {
  methods: ['GET', 'POST', 'DELETE'],
  authLevel: 'anonymous',
  route: 'api/manage/deny-list/user',
  handler: withSecurity(denyListUser),
});

app.http('denyListUserClear', {
  methods: ['POST'],
  authLevel: 'anonymous',
  route: 'api/manage/deny-list/user/clear',
  handler: withSecurity(denyListUserClear),
});

app.http('isGlobalAdmin', {
  methods: ['GET'],
  authLevel: 'anonymous',
  route: 'api/manage/is-global-admin',
  handler: withSecurity(isGlobalAdmin),
});
