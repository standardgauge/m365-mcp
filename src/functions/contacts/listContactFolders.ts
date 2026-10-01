import { app, HttpRequest, HttpResponseInit, InvocationContext } from '@azure/functions';
import { getValidAccessTokenForSession, getTenantId } from '../../services/tokenCache.js';
import { createGraphClient } from '../../services/graphClient.js';
import { filterDeniedPaths } from '../../services/denyList.js';
import { withPolicyEnforcement } from '../../services/policyEnforcement.js';
import type { AuthResult } from '../../services/authMiddleware.js';
import { withSecurity } from '../../services/securityHeaders.js';

interface ContactFolderItem {
  id: string;
  name: string;
  parentFolderId: string | null;
}

async function listContactFoldersHandler(
  request: HttpRequest,
  context: InvocationContext,
  auth: AuthResult,
): Promise<HttpResponseInit> {
  const userId = auth.userId;
  const accessToken = await getValidAccessTokenForSession(auth.session);

  const graph = createGraphClient(accessToken);

  const [childResult, rootCheck] = await Promise.all([
    graph.api('/me/contactFolders').select('id,displayName,parentFolderId').top(100).get(),
    graph.api('/me/contacts').select('id').top(1).get().catch(() => ({ value: [] })),
  ]);

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const children: ContactFolderItem[] = (childResult.value ?? []).map((f: any) => ({
    id: f.id,
    name: f.displayName,
    parentFolderId: f.parentFolderId ?? null,
  }));

  // /me/contactFolders only returns child folders — the root "Contacts" folder is never
  // included. Prepend a synthetic root entry whenever the root contains contacts OR when
  // there are no child folders (so the browser is never completely empty).
  const rootHasContacts = (rootCheck.value ?? []).length > 0;
  const folders: ContactFolderItem[] =
    rootHasContacts || children.length === 0
      ? [{ id: 'contacts-root', name: 'Contacts (Default)', parentFolderId: null }, ...children]
      : children;

  const allowed = await filterDeniedPaths(await getTenantId(userId), userId, 'contacts', folders);

  return { status: 200, jsonBody: { folders: allowed, count: allowed.length } };
}

app.http('listContactFolders', {
  methods: ['GET'],
  authLevel: 'anonymous',
  route: 'api/contacts/folders',
  handler: withSecurity(withPolicyEnforcement('contacts', listContactFoldersHandler)),
});
