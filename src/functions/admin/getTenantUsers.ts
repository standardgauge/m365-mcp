import { app, HttpRequest, HttpResponseInit, InvocationContext } from '@azure/functions';
import { authenticateRequest, checkGlobalAdmin } from '../../services/authMiddleware.js';
import { getValidAccessTokenForSession, listActiveSessions, SESSION_TTL_MS } from '../../services/tokenCache.js';
import { createGraphClient } from '../../services/graphClient.js';
import { withSecurity } from '../../services/securityHeaders.js';

// Cap pagination to defend against malformed or cyclic @odata.nextLink chains.
// 50 pages × 999 users/page = ~50k users; well above any plausible tenant size.
const MAX_PAGES = 50;

interface TenantUser {
  userId: string;
  email: string;
  displayName: string;
  accountEnabled: boolean;
  userType: 'Member' | 'Guest' | string;
  licensed: boolean;
  installed: boolean;
}

interface GraphUser {
  id: string;
  displayName?: string | null;
  userPrincipalName?: string | null;
  mail?: string | null;
  accountEnabled?: boolean;
  userType?: string | null;
  assignedLicenses?: Array<{ skuId: string }>;
}

// In-memory cache, keyed by tenantId. Tenant user lists rarely change minute-to-minute.
const CACHE_TTL_MS = 5 * 60 * 1000;
const cache = new Map<string, { fetchedAt: number; users: GraphUser[] }>();

async function fetchAllTenantUsers(accessToken: string): Promise<GraphUser[]> {
  const graph = createGraphClient(accessToken);
  const select = 'id,displayName,userPrincipalName,mail,accountEnabled,userType,assignedLicenses';
  const collected: GraphUser[] = [];
  const visitedLinks = new Set<string>();

  let response: { value?: GraphUser[]; '@odata.nextLink'?: string } = await graph
    .api('/users')
    .header('ConsistencyLevel', 'eventual')
    .select(select)
    .top(999)
    .get();

  for (let page = 0; page < MAX_PAGES; page++) {
    if (Array.isArray(response.value)) collected.push(...response.value);
    const nextLink = response['@odata.nextLink'];
    if (!nextLink) return collected;
    if (visitedLinks.has(nextLink)) {
      console.warn('[getTenantUsers] Detected cyclic @odata.nextLink, halting pagination');
      return collected;
    }
    visitedLinks.add(nextLink);
    response = await graph.api(nextLink).header('ConsistencyLevel', 'eventual').get();
  }

  console.warn(`[getTenantUsers] Hit MAX_PAGES (${MAX_PAGES}) ceiling, halting pagination`);
  return collected;
}

async function getTenantUsersFromCacheOrFetch(tenantId: string, accessToken: string): Promise<GraphUser[]> {
  const cached = cache.get(tenantId);
  if (cached && Date.now() - cached.fetchedAt < CACHE_TTL_MS) {
    return cached.users;
  }
  const users = await fetchAllTenantUsers(accessToken);
  cache.set(tenantId, { fetchedAt: Date.now(), users });
  return users;
}

async function getTenantUsers(
  request: HttpRequest,
  context: InvocationContext,
): Promise<HttpResponseInit> {
  try {
    const auth = await authenticateRequest(request);
    if (!auth) {
      return { status: 401, jsonBody: { error: 'Authentication required' } };
    }

    const isAdmin = await checkGlobalAdmin(auth.userId);
    if (!isAdmin) {
      return { status: 403, jsonBody: { error: 'Global Administrator role required' } };
    }

    const includeAll = request.query.get('includeAll') === 'true';
    const tenantId = auth.session.tenantId;
    const accessToken = await getValidAccessTokenForSession(auth.session);

    const [graphUsers, sessions] = await Promise.all([
      getTenantUsersFromCacheOrFetch(tenantId, accessToken),
      listActiveSessions(),
    ]);

    // Only count sessions whose idle TTL hasn't expired — otherwise the badge
    // marks users as "installed" based on stale storage rows from past installs.
    const now = Date.now();
    const installedUserIds = new Set(
      sessions
        .filter((s) => s.sessionCreatedAt > 0 && now - s.sessionCreatedAt < SESSION_TTL_MS)
        .map((s) => s.userId),
    );

    const mapped: TenantUser[] = graphUsers.map((u) => ({
      userId: u.id,
      displayName: u.displayName ?? u.userPrincipalName ?? '',
      email: u.mail ?? u.userPrincipalName ?? '',
      accountEnabled: u.accountEnabled ?? false,
      userType: u.userType ?? 'Member',
      licensed: Array.isArray(u.assignedLicenses) && u.assignedLicenses.length > 0,
      installed: installedUserIds.has(u.id),
    }));

    const filtered = includeAll
      ? mapped
      : mapped.filter((u) => u.accountEnabled && u.userType === 'Member' && u.licensed);

    filtered.sort((a, b) => a.displayName.localeCompare(b.displayName));

    return {
      status: 200,
      jsonBody: {
        users: filtered,
        totalCount: mapped.length,
        filteredCount: filtered.length,
        cached: cache.get(tenantId) !== undefined,
      },
    };
  } catch (err: unknown) {
    const message = err instanceof Error ? err.message : String(err);
    context.error('getTenantUsers error:', message);
    return { status: 500, jsonBody: { error: message } };
  }
}

app.http('getTenantUsers', {
  methods: ['GET'],
  authLevel: 'anonymous',
  route: 'api/manage/tenant-users',
  handler: withSecurity(getTenantUsers),
});
