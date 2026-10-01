import { app, HttpRequest, HttpResponseInit, InvocationContext } from '@azure/functions';
import { getValidAccessTokenForSession, getTenantId } from '../../services/tokenCache.js';
import { createGraphClient } from '../../services/graphClient.js';
import { filterDeniedSearchHits } from '../../services/denyList.js';
import { getAllowedSites } from '../../services/serviceSettings.js';
import {
  SEARCH_DRIVE_ITEM_FIELDS,
  searchFetchSize,
  siteRelativePathFromWebUrl,
  filterHitsToAllowedSites,
} from '../../services/sharepointSearch.js';
import { withPolicyEnforcement } from '../../services/policyEnforcement.js';
import { assertOpaqueId } from '../../services/opaqueId.js';
import type { AuthResult } from '../../services/authMiddleware.js';
import { withSecurity } from '../../services/securityHeaders.js';

interface SearchResult {
  id: string;
  name: string;
  webUrl: string;
  siteId: string | null;
  path: string | undefined;
  size: number | null;
  createdDateTime: string | null;
  lastModifiedDateTime: string | null;
  summary: string | null;
}

async function searchSharepointHandler(
  request: HttpRequest,
  context: InvocationContext,
  auth: AuthResult,
): Promise<HttpResponseInit> {
  const userId = auth.userId;

  const query = request.query.get('q');
  const siteId = request.query.get('siteId');
  const maxResults = Math.min(parseInt(request.query.get('maxResults') ?? '25', 10), 50);

  if (!query) {
    return { status: 400, jsonBody: { error: 'q query param is required' } };
  }
  if (siteId) assertOpaqueId(siteId, 'siteId');

  const accessToken = await getValidAccessTokenForSession(auth.session);
  const graph = createGraphClient(accessToken);

  const tenantId = await getTenantId(userId);
  // The allow-list / deny-list are enforced as post-filters below. The old
  // `contentSources` source constraint is invalid for a driveItem query and made
  // every search error, so scoping is post-filter only; over-fetch a
  // wider window when a scope is active so filtering still fills maxResults.
  const allowedSites = await getAllowedSites(tenantId);
  // An explicit siteId (validated in-allow-list upstream) narrows the post-filter
  // to that single site; otherwise the full allow-list applies (empty = allow all).
  const scopeSites = siteId ? [{ id: siteId }] : allowedSites;
  const size = searchFetchSize(maxResults, scopeSites.length > 0);

  const searchPayload = {
    requests: [
      {
        entityTypes: ['driveItem'],
        query: { queryString: query },
        fields: [...SEARCH_DRIVE_ITEM_FIELDS],
        from: 0,
        size,
      },
    ],
  };

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const response: any = await graph.api('/search/query').post(searchPayload);

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const hits: any[] = response.value?.[0]?.hitsContainers?.[0]?.hits ?? [];

  const items: SearchResult[] = hits.map((hit) => {
    const res = hit.resource ?? {};
    return {
      id: res.id ?? '',
      name: res.name ?? '',
      webUrl: res.webUrl ?? '',
      siteId: res.parentReference?.siteId ?? null,
      // Site-relative path resolved from webUrl so listing-style deny entries
      // match (parentReference.path is drive-relative and lacks the document
      // library segment; an absolute webUrl never matches a path-style entry).
      path: siteRelativePathFromWebUrl(res.webUrl) ?? undefined,
      size: res.size ?? null,
      createdDateTime: res.createdDateTime ?? null,
      lastModifiedDateTime: res.lastModifiedDateTime ?? null,
      summary: hit.summary ?? null,
    };
  });

  // Security boundary: drop hits outside the scope, then drop hits whose path
  // could not be resolved (fail closed — an unmatchable path must not slip past
  // the deny list), then deny-list. Slice to maxResults last (over-fetch).
  const siteFiltered = filterHitsToAllowedSites(items, scopeSites);
  const pathResolved = siteFiltered.filter((it) => typeof it.path === 'string' && it.path.length > 0);
  const allowed = (await filterDeniedSearchHits(tenantId, userId, pathResolved)).slice(0, maxResults);

  return { status: 200, jsonBody: { results: allowed, count: allowed.length } };
}

app.http('searchSharepoint', {
  methods: ['GET'],
  authLevel: 'anonymous',
  route: 'api/sharepoint/search',
  handler: withSecurity(withPolicyEnforcement('sharepoint', searchSharepointHandler, {
    getSiteId: (req) => req.query.get('siteId') ?? undefined,
  })),
});
