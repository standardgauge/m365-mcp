import { app, HttpRequest, HttpResponseInit, InvocationContext } from '@azure/functions';
import { getValidAccessTokenForSession } from '../../services/tokenCache.js';
import { createGraphClient } from '../../services/graphClient.js';
import { withPolicyEnforcement, checkDenyList } from '../../services/policyEnforcement.js';
import type { AuthResult } from '../../services/authMiddleware.js';
import { withSecurity } from '../../services/securityHeaders.js';

interface WriteRequest {
  path: string;
  content: string;
  contentType?: string;
}

async function writeOneDriveFileHandler(
  request: HttpRequest,
  context: InvocationContext,
  auth: AuthResult,
): Promise<HttpResponseInit> {
  const userId = auth.userId;

  const body = (await request.json()) as WriteRequest;
  if (!body.path || body.content === undefined) {
    return { status: 400, jsonBody: { error: 'path and content are required' } };
  }

  // Enforce deny-list policy
  const denyViolation = await checkDenyList(userId, 'onedrive', body.path);
  if (denyViolation) {
    return { status: denyViolation.status, jsonBody: { error: denyViolation.error } };
  }

  const accessToken = await getValidAccessTokenForSession(auth.session);
  const graph = createGraphClient(accessToken);

  // PUT to /me/drive/root:/{path}:/content creates or overwrites
  const apiPath = `/me/drive/root:/${body.path.replace(/^\//, '')}:/content`;
  const contentType = body.contentType ?? 'text/plain';

  const result = await graph
    .api(apiPath)
    .header('Content-Type', contentType)
    .put(Buffer.from(body.content, 'utf-8'));

  return {
    status: 200,
    jsonBody: {
      id: result.id,
      name: result.name,
      webUrl: result.webUrl,
      size: result.size,
      lastModifiedDateTime: result.lastModifiedDateTime,
      status: 'written',
    },
  };
}

app.http('writeOneDriveFile', {
  methods: ['PUT'],
  authLevel: 'anonymous',
  route: 'api/onedrive/files',
  handler: withSecurity(withPolicyEnforcement('onedrive', writeOneDriveFileHandler)),
});
