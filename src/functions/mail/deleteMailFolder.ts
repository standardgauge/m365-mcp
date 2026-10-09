import { app, HttpRequest, HttpResponseInit, InvocationContext } from '@azure/functions';
import { getValidAccessTokenForSession } from '../../services/tokenCache.js';
import { createGraphClient } from '../../services/graphClient.js';
import { resolveDenySubject } from '../../services/mailboxOwner.js';
import { withPolicyEnforcement, checkDenyList } from '../../services/policyEnforcement.js';
import { resolveMailFolderName } from '../../services/containerResolver.js';
import { assertOpaqueId } from '../../services/opaqueId.js';
import type { AuthResult } from '../../services/authMiddleware.js';
import { withSecurity } from '../../services/securityHeaders.js';

/**
 * Delete a mail folder. Destructive: removes the folder AND all messages/subfolders
 * inside it. Guarded by an explicit `confirm=true` query parameter, and deliberately
 * NOT exposed as an MCP bridge tool (no `delete_mail_folder` tool) so an agent cannot
 * delete a folder in a single call. Available for direct/admin use only.
 */
async function deleteMailFolderHandler(
  request: HttpRequest,
  context: InvocationContext,
  auth: AuthResult,
): Promise<HttpResponseInit> {
  const userId = auth.userId;

  const folderId = request.params['folderId'];
  if (!folderId) {
    return { status: 400, jsonBody: { error: 'folderId path param is required' } };
  }
  assertOpaqueId(folderId, 'folderId');

  if (request.query.get('confirm') !== 'true') {
    return {
      status: 400,
      jsonBody: {
        error:
          'Deleting a mail folder is destructive (removes the folder and all of its contents). ' +
          'Pass confirm=true to proceed.',
      },
    };
  }

  const mailboxId = request.query.get('mailboxId');
  if (mailboxId && mailboxId !== 'me') assertOpaqueId(mailboxId, 'mailboxId');

  const accessToken = await getValidAccessTokenForSession(auth.session);
  const graph = createGraphClient(accessToken);
  const denySubject = await resolveDenySubject(graph, userId, mailboxId);

  const mailboxBase = mailboxId && mailboxId !== 'me'
    ? `/users/${mailboxId}`
    : '/me';

  // Deny-list check on the folder being deleted (resolve its display name).
  const folderName = await resolveMailFolderName(graph, folderId, mailboxBase);
  if (folderName) {
    const violation = await checkDenyList(userId, 'mail', folderName, undefined, denySubject);
    if (violation) {
      return { status: violation.status, jsonBody: { error: violation.error } };
    }
  }

  await graph.api(`${mailboxBase}/mailFolders/${folderId}`).delete();

  return {
    status: 200,
    jsonBody: {
      id: folderId,
      status: 'deleted',
    },
  };
}

app.http('deleteMailFolder', {
  methods: ['DELETE'],
  authLevel: 'anonymous',
  route: 'api/mail/folders/{folderId}',
  handler: withSecurity(withPolicyEnforcement('mail', deleteMailFolderHandler)),
});
