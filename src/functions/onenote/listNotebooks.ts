import { app, HttpRequest, HttpResponseInit, InvocationContext } from '@azure/functions';
import { getValidAccessTokenForSession, getTenantId } from '../../services/tokenCache.js';
import { createGraphClient } from '../../services/graphClient.js';
import { filterDeniedPaths } from '../../services/denyList.js';
import { withPolicyEnforcement } from '../../services/policyEnforcement.js';
import type { AuthResult } from '../../services/authMiddleware.js';
import { withSecurity } from '../../services/securityHeaders.js';

interface NotebookItem {
  id: string;
  name: string;
  createdDateTime: string;
}

async function listNotebooksHandler(
  request: HttpRequest,
  context: InvocationContext,
  auth: AuthResult,
): Promise<HttpResponseInit> {
  const userId = auth.userId;
  const accessToken = await getValidAccessTokenForSession(auth.session);

  const graph = createGraphClient(accessToken);

  const result = await graph
    .api('/me/onenote/notebooks')
    .select('id,displayName,createdDateTime')
    .top(100)
    .get();

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const notebooks: NotebookItem[] = (result.value ?? []).map((n: any) => ({
    id: n.id,
    name: n.displayName,
    createdDateTime: n.createdDateTime,
  }));

  const allowed = await filterDeniedPaths(await getTenantId(userId), userId, 'onenote', notebooks);

  return { status: 200, jsonBody: { notebooks: allowed, count: allowed.length } };
}

app.http('listNotebooks', {
  methods: ['GET'],
  authLevel: 'anonymous',
  route: 'api/onenote/notebooks',
  handler: withSecurity(withPolicyEnforcement('onenote', listNotebooksHandler)),
});
