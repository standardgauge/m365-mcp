import { app, HttpRequest, HttpResponseInit, InvocationContext } from '@azure/functions';
import { getValidAccessTokenForSession, getTenantId } from '../../services/tokenCache.js';
import { createGraphClient } from '../../services/graphClient.js';
import { isPathDenied } from '../../services/denyList.js';
import { withPolicyEnforcement } from '../../services/policyEnforcement.js';
import { resolveSectionNotebook } from '../../services/containerResolver.js';
import { assertOpaqueId } from '../../services/opaqueId.js';
import type { AuthResult } from '../../services/authMiddleware.js';
import { withSecurity } from '../../services/securityHeaders.js';

interface PageRequest {
  sectionId: string;
  title: string;
  htmlContent: string;
}

async function createOneNotePageHandler(
  request: HttpRequest,
  context: InvocationContext,
  auth: AuthResult,
): Promise<HttpResponseInit> {
  const userId = auth.userId;

  const body = (await request.json()) as PageRequest;
  if (!body.sectionId || !body.title) {
    return { status: 400, jsonBody: { error: 'sectionId and title are required' } };
  }
  assertOpaqueId(body.sectionId, 'sectionId');

  const accessToken = await getValidAccessTokenForSession(auth.session);
  const graph = createGraphClient(accessToken);

  // Check section-level deny first (admin UI stores section IDs directly),
  // then check parent notebook for notebook-level blocks.
  const tenantId = await getTenantId(userId);
  if (await isPathDenied(tenantId, userId, 'onenote', body.sectionId)) {
    return { status: 403, jsonBody: { error: 'Access restricted by deny list — section is blocked' } };
  }
  const notebookId = await resolveSectionNotebook(graph, body.sectionId);
  if (notebookId) {
    if (await isPathDenied(tenantId, userId, 'onenote', notebookId)) {
      return { status: 403, jsonBody: { error: 'Access restricted by deny list — parent notebook is blocked' } };
    }
  }

  // OneNote pages are created by POSTing HTML
  const html = `<!DOCTYPE html>
<html>
<head><title>${body.title}</title></head>
<body>${body.htmlContent ?? ''}</body>
</html>`;

  const result = await graph
    .api(`/me/onenote/sections/${body.sectionId}/pages`)
    .header('Content-Type', 'text/html')
    .post(html);

  return {
    status: 201,
    jsonBody: {
      id: result.id,
      title: result.title,
      self: result.self,
      contentUrl: result.contentUrl,
      createdDateTime: result.createdDateTime,
      status: 'created',
    },
  };
}

app.http('createOneNotePage', {
  methods: ['POST'],
  authLevel: 'anonymous',
  route: 'api/onenote/pages',
  handler: withSecurity(withPolicyEnforcement('onenote', createOneNotePageHandler)),
});
