import { app, HttpRequest, HttpResponseInit, InvocationContext } from '@azure/functions';
import { getValidAccessTokenForSession, getTenantId } from '../../services/tokenCache.js';
import { createGraphClient } from '../../services/graphClient.js';
import { resolveDenySubject } from '../../services/mailboxOwner.js';
import { filterDeniedPaths } from '../../services/denyList.js';
import { withPolicyEnforcement } from '../../services/policyEnforcement.js';
import { assertOpaqueId } from '../../services/opaqueId.js';
import type { AuthResult } from '../../services/authMiddleware.js';
import { withSecurity } from '../../services/securityHeaders.js';

interface MailFolder {
  id: string;
  name: string;
  /** We use the folder display name as the deny-list path for mail folders. */
  path: string;
  webUrl: string;
  totalItemCount: number;
  unreadItemCount: number;
  childFolderCount: number;
  wellKnownName: string | null;
}

async function listFoldersMailHandler(
  request: HttpRequest,
  context: InvocationContext,
  auth: AuthResult,
): Promise<HttpResponseInit> {
  const userId = auth.userId;
  const accessToken = await getValidAccessTokenForSession(auth.session);

  const mailboxId = request.query.get('mailboxId') ?? 'me';
  const parentFolderId = request.query.get('parentFolderId');

  if (mailboxId !== 'me') assertOpaqueId(mailboxId, 'mailboxId');
  if (parentFolderId) assertOpaqueId(parentFolderId, 'parentFolderId');

  const graph = createGraphClient(accessToken);
  const denySubject = await resolveDenySubject(graph, userId, mailboxId);

  const base = mailboxId === 'me' ? '/me' : `/users/${mailboxId}`;
  const apiPath = parentFolderId
    ? `${base}/mailFolders/${parentFolderId}/childFolders`
    : `${base}/mailFolders`;

  const result = await graph
    .api(apiPath)
    .select('id,displayName,totalItemCount,unreadItemCount,childFolderCount')
    .top(100)
    .get();

  const folders: MailFolder[] = (result.value ?? []).map(
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    (f: any) => ({
      id: f.id,
      name: f.displayName,
      // Use displayName as the path identifier — admin deny-list entries
      // should match on this field.
      path: f.displayName,
      // Mail folders don't have a web URL in the Graph API response; we
      // construct a reasonable placeholder using the folder ID.
      webUrl: f.id,
      totalItemCount: f.totalItemCount ?? 0,
      unreadItemCount: f.unreadItemCount ?? 0,
      childFolderCount: f.childFolderCount ?? 0,
      wellKnownName: f.wellKnownName ?? null,
    })
  );

  const allowed = await filterDeniedPaths(await getTenantId(userId), denySubject, 'mail', folders);

  return { status: 200, jsonBody: { folders: allowed, count: allowed.length } };
}

app.http('listFoldersMail', {
  methods: ['GET'],
  authLevel: 'anonymous',
  route: 'api/mail/folders',
  // Deny-list filtering is handled inside the handler via filterDeniedPaths
  // using folder display names (the deny list stores names, not Graph IDs).
  handler: withSecurity(withPolicyEnforcement('mail', listFoldersMailHandler)),
});
