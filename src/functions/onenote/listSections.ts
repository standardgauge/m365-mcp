import { app, HttpRequest, HttpResponseInit, InvocationContext } from '@azure/functions';
import { getValidAccessTokenForSession, getTenantId } from '../../services/tokenCache.js';
import { createGraphClient } from '../../services/graphClient.js';
import { filterDeniedPaths } from '../../services/denyList.js';
import { withPolicyEnforcement } from '../../services/policyEnforcement.js';
import { assertOpaqueId } from '../../services/opaqueId.js';
import type { AuthResult } from '../../services/authMiddleware.js';
import { withSecurity } from '../../services/securityHeaders.js';

interface SectionItem {
  id: string;
  displayName: string;
  pagesUrl: string;
}

async function listSectionsHandler(
  request: HttpRequest,
  context: InvocationContext,
  auth: AuthResult,
): Promise<HttpResponseInit> {
  const userId = auth.userId;
  const accessToken = await getValidAccessTokenForSession(auth.session);

  const notebookId = request.params['notebookId'];
  if (!notebookId) {
    return { status: 400, jsonBody: { error: 'Missing notebookId' } };
  }
  assertOpaqueId(notebookId, 'notebookId');

  const graph = createGraphClient(accessToken);

  const result = await graph
    .api(`/me/onenote/notebooks/${notebookId}/sections`)
    .select('id,displayName,pagesUrl')
    .top(100)
    .get();

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const allSections = (result.value ?? []).map((s: any) => ({
    id: s.id,
    displayName: s.displayName,
    pagesUrl: s.pagesUrl,
  }));

  // Post-filter sections against the deny list by section ID.
  // The admin UI stores section IDs directly for section-level blocks.
  const tenantId = await getTenantId(userId);
  const allowed = await filterDeniedPaths(tenantId, userId, 'onenote', allSections);
  const sections: SectionItem[] = allowed.map((s) => ({
    id: s.id!,
    displayName: (s as typeof allSections[number]).displayName,
    pagesUrl: (s as typeof allSections[number]).pagesUrl,
  }));

  return { status: 200, jsonBody: { sections, count: sections.length } };
}

app.http('listSections', {
  methods: ['GET'],
  authLevel: 'anonymous',
  route: 'api/onenote/notebooks/{notebookId}/sections',
  handler: withSecurity(withPolicyEnforcement('onenote', listSectionsHandler, {
    getDenyListPaths: (req) => {
      const notebookId = req.params['notebookId'];
      return notebookId ? [notebookId] : [];
    },
  })),
});
