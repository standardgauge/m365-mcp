import { app, HttpRequest, HttpResponseInit, InvocationContext } from '@azure/functions';
import { getValidAccessTokenForSession } from '../../services/tokenCache.js';
import { createGraphClient } from '../../services/graphClient.js';
import { withPolicyEnforcement } from '../../services/policyEnforcement.js';
import { assertOpaqueId } from '../../services/opaqueId.js';
import type { AuthResult } from '../../services/authMiddleware.js';
import { withSecurity } from '../../services/securityHeaders.js';

async function listSharepointListsHandler(
  request: HttpRequest,
  context: InvocationContext,
  auth: AuthResult,
): Promise<HttpResponseInit> {
  const siteId = request.query.get('siteId');
  if (!siteId) {
    return { status: 400, jsonBody: { error: 'siteId query param is required' } };
  }
  assertOpaqueId(siteId, 'siteId');

  const accessToken = await getValidAccessTokenForSession(auth.session);
  const graph = createGraphClient(accessToken);

  const result = await graph
    .api(`/sites/${siteId}/lists`)
    .select('id,displayName,description,lastModifiedDateTime,list')
    .top(100)
    .get();

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const lists = (result.value ?? []).map((l: any) => ({
    id: l.id,
    displayName: l.displayName,
    description: l.description ?? '',
    template: l.list?.template ?? '',
    itemCount: l.list?.contentTypesEnabled ? undefined : undefined,
    lastModifiedDateTime: l.lastModifiedDateTime,
  }));

  return { status: 200, jsonBody: { lists, count: lists.length } };
}

app.http('listSharepointLists', {
  methods: ['GET'],
  authLevel: 'anonymous',
  route: 'api/sharepoint/lists',
  handler: withSecurity(withPolicyEnforcement('sharepoint', listSharepointListsHandler, {
    getSiteId: (req) => req.query.get('siteId') ?? undefined,
  })),
});
