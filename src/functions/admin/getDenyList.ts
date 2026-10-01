import { app, HttpRequest, HttpResponseInit, InvocationContext } from '@azure/functions';
import {
  listGlobalDenyEntries,
  addGlobalDenyEntry,
  removeGlobalDenyEntry,
  listUserDenyEntries,
  addUserDenyEntry,
  removeUserDenyEntry,
  clearUserDenyList,
  type DenyListType,
} from '../../services/denyList.js';
import { getTenantId } from '../../services/tokenCache.js';
import { authenticateRequest, checkGlobalAdmin } from '../../services/authMiddleware.js';
import { withSecurity } from '../../services/securityHeaders.js';

async function denyListGlobal(
  request: HttpRequest,
  context: InvocationContext
): Promise<HttpResponseInit> {
  try {
    const type = (request.query.get('type') ?? 'sharepoint') as DenyListType;
    const auth = await authenticateRequest(request);
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
    const isAdmin = await checkGlobalAdmin(addedBy);
    if (!isAdmin) {
      return { status: 403, jsonBody: { error: 'Global Administrator role required' } };
    }

    const body = await request.json() as { type?: DenyListType; path?: string; description?: string; addedByName?: string };
    const entryType = body.type ?? type;
    const path = body.path;
    if (!path) return { status: 400, jsonBody: { error: 'Missing path' } };

    if (request.method === 'POST') {
      await addGlobalDenyEntry(tenantId, entryType, path, addedBy, body.description ?? '', body.addedByName);
    } else {
      await removeGlobalDenyEntry(tenantId, entryType, path);
    }
    return { status: 200, jsonBody: { ok: true } };
  } catch (err: unknown) {
    const message = err instanceof Error ? err.message : String(err);
    context.error('denyListGlobal error:', message);
    return { status: request.method === 'GET' ? 200 : 500, jsonBody: request.method === 'GET' ? { entries: [] } : { error: message } };
  }
}

async function denyListUser(
  request: HttpRequest,
  context: InvocationContext
): Promise<HttpResponseInit> {
  try {
    const type = (request.query.get('type') ?? 'sharepoint') as DenyListType;
    const auth = await authenticateRequest(request);
    if (!auth) {
      return { status: 401, jsonBody: { error: 'Authentication required' } };
    }
    const requesterId = auth.userId;

    if (request.method === 'GET') {
      const targetUserId = request.query.get('targetUserId') ?? requesterId;
      if (!targetUserId) return { status: 400, jsonBody: { error: 'Missing targetUserId or x-user-id' } };
      // IDOR protection: non-admins can only read their own deny list
      if (targetUserId !== requesterId) {
        const isAdmin = await checkGlobalAdmin(requesterId);
        if (!isAdmin) {
          return { status: 403, jsonBody: { error: 'You can only view your own deny list' } };
        }
      }
      const entries = await listUserDenyEntries(targetUserId, type);
      return { status: 200, jsonBody: { entries } };
    }

    const body = await request.json() as { type?: DenyListType; path?: string; targetUserId?: string; addedByName?: string };
    const entryType = body.type ?? type;
    const path = body.path;
    const targetUserId = body.targetUserId ?? requesterId;
    if (!path) return { status: 400, jsonBody: { error: 'Missing path' } };
    if (!targetUserId) return { status: 400, jsonBody: { error: 'Missing targetUserId' } };

    // IDOR protection: non-admins can only modify their own deny list
    if (targetUserId !== requesterId) {
      const isAdmin = await checkGlobalAdmin(requesterId);
      if (!isAdmin) {
        return { status: 403, jsonBody: { error: 'You can only modify your own deny list' } };
      }
    }

    if (request.method === 'POST') {
      await addUserDenyEntry(targetUserId, entryType, path, body.addedByName);
    } else {
      await removeUserDenyEntry(targetUserId, entryType, path);
    }
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
    const auth = await authenticateRequest(request);
    if (!auth) {
      return { status: 401, jsonBody: { error: 'Authentication required' } };
    }
    const body = await request.json() as { targetUserId?: string };
    const targetUserId = body.targetUserId ?? auth.userId;
    if (!targetUserId) return { status: 400, jsonBody: { error: 'Missing targetUserId' } };

    // IDOR protection: non-admins can only clear their own deny list
    if (targetUserId !== auth.userId) {
      const isAdmin = await checkGlobalAdmin(auth.userId);
      if (!isAdmin) {
        return { status: 403, jsonBody: { error: 'You can only clear your own deny list' } };
      }
    }

    await Promise.all([
      clearUserDenyList(targetUserId, 'sharepoint'),
      clearUserDenyList(targetUserId, 'mail'),
    ]);
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
  const auth = await authenticateRequest(request);
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
