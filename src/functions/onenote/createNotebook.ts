import { app, HttpRequest, HttpResponseInit, InvocationContext } from '@azure/functions';
import { getValidAccessTokenForSession } from '../../services/tokenCache.js';
import { createGraphClient } from '../../services/graphClient.js';
import { withPolicyEnforcement } from '../../services/policyEnforcement.js';
import type { AuthResult } from '../../services/authMiddleware.js';
import { withSecurity } from '../../services/securityHeaders.js';

async function createNotebookHandler(
  request: HttpRequest,
  context: InvocationContext,
  auth: AuthResult,
): Promise<HttpResponseInit> {
  const body = (await request.json()) as { displayName: string };
  if (!body.displayName) {
    return { status: 400, jsonBody: { error: 'displayName is required' } };
  }

  const accessToken = await getValidAccessTokenForSession(auth.session);
  const graph = createGraphClient(accessToken);

  const result = await graph.api('/me/onenote/notebooks').post({
    displayName: body.displayName,
  });

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

app.http('createNotebook', {
  methods: ['POST'],
  authLevel: 'anonymous',
  route: 'api/onenote/notebooks',
  handler: withSecurity(withPolicyEnforcement('onenote', createNotebookHandler)),
});
