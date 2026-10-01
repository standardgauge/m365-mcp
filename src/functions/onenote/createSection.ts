import { app, HttpRequest, HttpResponseInit, InvocationContext } from '@azure/functions';
import { getValidAccessTokenForSession } from '../../services/tokenCache.js';
import { createGraphClient } from '../../services/graphClient.js';
import { withPolicyEnforcement, checkDenyList } from '../../services/policyEnforcement.js';
import { assertOpaqueId } from '../../services/opaqueId.js';
import type { AuthResult } from '../../services/authMiddleware.js';
import { withSecurity } from '../../services/securityHeaders.js';

async function createSectionHandler(
  request: HttpRequest,
  context: InvocationContext,
  auth: AuthResult,
): Promise<HttpResponseInit> {
  const userId = auth.userId;

  const body = (await request.json()) as { notebookId: string; displayName: string };
  if (!body.notebookId || !body.displayName) {
    return { status: 400, jsonBody: { error: 'notebookId and displayName are required' } };
  }
  assertOpaqueId(body.notebookId, 'notebookId');

  // Deny-list check on notebookId (body-level, not available to wrapper)
  const denyViolation = await checkDenyList(userId, 'onenote', body.notebookId);
  if (denyViolation) {
    return { status: denyViolation.status, jsonBody: { error: denyViolation.error } };
  }

  const accessToken = await getValidAccessTokenForSession(auth.session);
  const graph = createGraphClient(accessToken);

  const result = await graph
    .api(`/me/onenote/notebooks/${body.notebookId}/sections`)
    .post({ displayName: body.displayName });

  return {
    status: 201,
    jsonBody: {
      id: result.id,
      displayName: result.displayName,
      self: result.self,
      status: 'created',
    },
  };
}

app.http('createSection', {
  methods: ['POST'],
  authLevel: 'anonymous',
  route: 'api/onenote/sections',
  handler: withSecurity(withPolicyEnforcement('onenote', createSectionHandler)),
});
