import { app, HttpRequest, HttpResponseInit, InvocationContext } from '@azure/functions';
import { getValidAccessTokenForSession } from '../../services/tokenCache.js';
import { createGraphClient } from '../../services/graphClient.js';
import { withPolicyEnforcement, checkAllowedSite, checkDenyList } from '../../services/policyEnforcement.js';
import { assertOpaqueId } from '../../services/opaqueId.js';
import type { AuthResult } from '../../services/authMiddleware.js';
import { withSecurity } from '../../services/securityHeaders.js';

interface CreateFolderRequest {
  siteId: string;
  name: string;
  parentId?: string;
  driveId?: string;
}

async function createSharepointFolderHandler(
  request: HttpRequest,
  context: InvocationContext,
  auth: AuthResult,
): Promise<HttpResponseInit> {
  const userId = auth.userId;

  const body = (await request.json()) as CreateFolderRequest;
  if (!body.siteId || !body.name) {
    return { status: 400, jsonBody: { error: 'siteId and name are required' } };
  }
  assertOpaqueId(body.siteId, 'siteId');
  if (body.driveId) assertOpaqueId(body.driveId, 'driveId');
  if (body.parentId) assertOpaqueId(body.parentId, 'parentId');

  // Enforce allowedSites policy (body-level siteId not available to wrapper)
  const siteViolation = await checkAllowedSite(userId, body.siteId);
  if (siteViolation) {
    return { status: siteViolation.status, jsonBody: { error: siteViolation.error } };
  }

  const accessToken = await getValidAccessTokenForSession(auth.session);
  const graph = createGraphClient(accessToken);

  const drivePath = body.driveId
    ? `/sites/${body.siteId}/drives/${body.driveId}`
    : `/sites/${body.siteId}/drive`;

  // Resolve the parent path to check deny list before creating
  if (body.parentId) {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const parentMeta: any = await graph.api(`${drivePath}/items/${body.parentId}`).select('name,parentReference').get();
    const parentPath = `${parentMeta.parentReference?.path ?? ''}/${parentMeta.name}`;
    const denyViolation = await checkDenyList(userId, 'sharepoint', `${parentPath}/${body.name}`);
    if (denyViolation) {
      return { status: denyViolation.status, jsonBody: { error: denyViolation.error } };
    }
  }

  const apiPath = body.parentId
    ? `${drivePath}/items/${body.parentId}/children`
    : `${drivePath}/root/children`;

  const result = await graph.api(apiPath).post({
    name: body.name,
    folder: {},
    '@microsoft.graph.conflictBehavior': 'fail',
  });

  return {
    status: 201,
    jsonBody: {
      id: result.id,
      name: result.name,
      webUrl: result.webUrl,
      status: 'created',
    },
  };
}

app.http('createSharepointFolder', {
  methods: ['POST'],
  authLevel: 'anonymous',
  route: 'api/sharepoint/folders',
  handler: withSecurity(withPolicyEnforcement('sharepoint', createSharepointFolderHandler)),
});
