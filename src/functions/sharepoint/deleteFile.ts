import { app, HttpRequest, HttpResponseInit, InvocationContext } from '@azure/functions';
import { getValidAccessTokenForSession } from '../../services/tokenCache.js';
import { createGraphClient } from '../../services/graphClient.js';
import { withPolicyEnforcement, checkAllowedSite, checkDenyList } from '../../services/policyEnforcement.js';
import { assertOpaqueId } from '../../services/opaqueId.js';
import type { AuthResult } from '../../services/authMiddleware.js';
import { withSecurity } from '../../services/securityHeaders.js';

async function deleteSharepointFileHandler(
  request: HttpRequest,
  context: InvocationContext,
  auth: AuthResult,
): Promise<HttpResponseInit> {
  const userId = auth.userId;

  const itemId = request.params['itemId'];
  const siteId = request.query.get('siteId');
  const driveId = request.query.get('driveId');

  if (!itemId || !siteId) {
    return { status: 400, jsonBody: { error: 'itemId path param and siteId query param are required' } };
  }
  assertOpaqueId(itemId, 'itemId');
  assertOpaqueId(siteId, 'siteId');
  if (driveId) assertOpaqueId(driveId, 'driveId');

  // Enforce allowedSites policy
  const siteViolation = await checkAllowedSite(userId, siteId);
  if (siteViolation) {
    return { status: siteViolation.status, jsonBody: { error: siteViolation.error } };
  }

  const accessToken = await getValidAccessTokenForSession(auth.session);
  const graph = createGraphClient(accessToken);

  const dp = driveId
    ? `/sites/${siteId}/drives/${driveId}`
    : `/sites/${siteId}/drive`;

  // Resolve file path for deny-list check
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const meta: any = await graph.api(`${dp}/items/${itemId}`).select('name,parentReference').get();
  const filePath = `${meta.parentReference?.path ?? ''}/${meta.name}`;
  const denyViolation = await checkDenyList(userId, 'sharepoint', filePath);
  if (denyViolation) {
    return { status: denyViolation.status, jsonBody: { error: denyViolation.error } };
  }

  await graph.api(`${dp}/items/${itemId}`).delete();

  return { status: 200, jsonBody: { status: 'deleted', itemId } };
}

app.http('deleteSharepointFile', {
  methods: ['DELETE'],
  authLevel: 'anonymous',
  route: 'api/sharepoint/files/{itemId}/delete',
  handler: withSecurity(withPolicyEnforcement('sharepoint', deleteSharepointFileHandler, {
    getSiteId: (req) => req.query.get('siteId') ?? undefined,
  })),
});
