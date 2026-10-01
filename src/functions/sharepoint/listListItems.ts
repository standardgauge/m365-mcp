import { app, HttpRequest, HttpResponseInit, InvocationContext } from '@azure/functions';
import { getValidAccessTokenForSession } from '../../services/tokenCache.js';
import { createGraphClient } from '../../services/graphClient.js';
import { withPolicyEnforcement } from '../../services/policyEnforcement.js';
import { assertOpaqueId } from '../../services/opaqueId.js';
import { collectPage } from '../../services/graphPaging.js';
import type { AuthResult } from '../../services/authMiddleware.js';
import { withSecurity } from '../../services/securityHeaders.js';

// Per-page $top used when walking the list-items collection for offset paging.
const LIST_ITEMS_PAGE_SIZE = 200;

async function listListItemsHandler(
  request: HttpRequest,
  context: InvocationContext,
  auth: AuthResult,
): Promise<HttpResponseInit> {
  const siteId = request.query.get('siteId');
  const listId = request.query.get('listId');
  const maxResults = Math.min(Math.max(1, parseInt(request.query.get('maxResults') ?? '50', 10) || 50), 100);
  const offset = Math.max(0, parseInt(request.query.get('offset') ?? '0', 10) || 0);

  if (!siteId || !listId) {
    return { status: 400, jsonBody: { error: 'siteId and listId query params are required' } };
  }
  assertOpaqueId(siteId, 'siteId');
  assertOpaqueId(listId, 'listId');

  const accessToken = await getValidAccessTokenForSession(auth.session);
  const graph = createGraphClient(accessToken);

  // Offset paging: the collection has no $skip, so walk nextLink pages
  // internally to reach an item deep in a multi-thousand-item list.
  const listPath = `/sites/${siteId}/lists/${listId}/items`;
  const fetchFirst = () => graph.api(listPath).expand('fields').top(LIST_ITEMS_PAGE_SIZE).get();
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const fetchNext = (link: string) => graph.api(link).get() as Promise<any>;
  const page = await collectPage(fetchFirst, fetchNext, offset, maxResults);

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const items = (page.items as any[]).map((item: any) => ({
    id: item.id,
    fields: item.fields ?? {},
    createdDateTime: item.createdDateTime,
    lastModifiedDateTime: item.lastModifiedDateTime,
  }));

  return { status: 200, jsonBody: { items, count: items.length, hasMore: page.hasMore, nextOffset: page.nextOffset } };
}

app.http('listListItems', {
  methods: ['GET'],
  authLevel: 'anonymous',
  route: 'api/sharepoint/lists/items',
  handler: withSecurity(withPolicyEnforcement('sharepoint', listListItemsHandler, {
    getSiteId: (req) => req.query.get('siteId') ?? undefined,
  })),
});
