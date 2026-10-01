import { app, HttpRequest, HttpResponseInit, InvocationContext } from '@azure/functions';
import { getValidAccessTokenForSession } from '../../services/tokenCache.js';
import { createGraphClient } from '../../services/graphClient.js';
import { withPolicyEnforcement, checkAllowedSite, checkDenyList } from '../../services/policyEnforcement.js';
import { assertOpaqueId } from '../../services/opaqueId.js';
import type { AuthResult } from '../../services/authMiddleware.js';
import { withSecurity } from '../../services/securityHeaders.js';

interface WriteRequest {
  siteId: string;
  path: string;
  content: string;
  driveId?: string;
  contentType?: string;
}

async function writeSharepointFileHandler(
  request: HttpRequest,
  context: InvocationContext,
  auth: AuthResult,
): Promise<HttpResponseInit> {
  const userId = auth.userId;

  const body = (await request.json()) as WriteRequest;
  if (!body.siteId || !body.path || body.content === undefined) {
    return { status: 400, jsonBody: { error: 'siteId, path, and content are required' } };
  }
  assertOpaqueId(body.siteId, 'siteId');
  if (body.driveId) assertOpaqueId(body.driveId, 'driveId');

  // Enforce allowedSites policy (body-level siteId not available to wrapper)
  const siteViolation = await checkAllowedSite(userId, body.siteId);
  if (siteViolation) {
    return { status: siteViolation.status, jsonBody: { error: siteViolation.error } };
  }

  // Enforce deny-list policy
  const denyViolation = await checkDenyList(userId, 'sharepoint', body.path);
  if (denyViolation) {
    return { status: denyViolation.status, jsonBody: { error: denyViolation.error } };
  }

  const accessToken = await getValidAccessTokenForSession(auth.session);
  const graph = createGraphClient(accessToken);

  const drivePath = body.driveId
    ? `/sites/${body.siteId}/drives/${body.driveId}`
    : `/sites/${body.siteId}/drive`;

  const apiPath = `${drivePath}/root:/${body.path.replace(/^\//, '')}:/content`;
  const ct = body.contentType ?? 'text/plain';

  const result = await graph
    .api(apiPath)
    .header('Content-Type', ct)
    .put(Buffer.from(body.content, 'utf-8'));

  return {
    status: 200,
    jsonBody: {
      id: result.id,
      name: result.name,
      webUrl: result.webUrl,
      size: result.size,
      status: 'written',
    },
  };
}

app.http('writeSharepointFile', {
  methods: ['PUT'],
  authLevel: 'anonymous',
  route: 'api/sharepoint/files',
  handler: withSecurity(withPolicyEnforcement('sharepoint', writeSharepointFileHandler)),
});
