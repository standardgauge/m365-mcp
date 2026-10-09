import { app, HttpRequest, HttpResponseInit, InvocationContext } from '@azure/functions';
import { getValidAccessTokenForSession } from '../../services/tokenCache.js';
import { createGraphClient } from '../../services/graphClient.js';
import { resolveDenySubject } from '../../services/mailboxOwner.js';
import { withPolicyEnforcement, checkDenyList } from '../../services/policyEnforcement.js';
import { resolveMailFolderName } from '../../services/containerResolver.js';
import { assertOpaqueId } from '../../services/opaqueId.js';
import type { AuthResult } from '../../services/authMiddleware.js';
import { withSecurity } from '../../services/securityHeaders.js';

interface RenameMailFolderRequest {
  displayName: string;
  mailboxId?: string;
}

async function renameMailFolderHandler(
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

  const body = (await request.json()) as RenameMailFolderRequest;
  if (!body.displayName) {
    return { status: 400, jsonBody: { error: 'displayName is required' } };
  }
  if (body.mailboxId && body.mailboxId !== 'me') assertOpaqueId(body.mailboxId, 'mailboxId');

  const accessToken = await getValidAccessTokenForSession(auth.session);
  const graph = createGraphClient(accessToken);
  const denySubject = await resolveDenySubject(graph, userId, body.mailboxId);

  const mailboxBase = body.mailboxId && body.mailboxId !== 'me'
    ? `/users/${body.mailboxId}`
    : '/me';

  // Block renaming a folder that is itself deny-listed (resolve its current name) —
  // otherwise a rename could un-hide a folder the deny list intends to keep out of reach.
  const currentName = await resolveMailFolderName(graph, folderId, mailboxBase);
  if (currentName) {
    const sourceViolation = await checkDenyList(userId, 'mail', currentName, undefined, denySubject);
    if (sourceViolation) {
      return { status: sourceViolation.status, jsonBody: { error: sourceViolation.error } };
    }
  }

  // Block renaming TO a deny-listed name (parity with createMailFolder) — otherwise a
  // rename could shadow a denied name and route mail past the deny list.
  const targetViolation = await checkDenyList(userId, 'mail', body.displayName, undefined, denySubject);
  if (targetViolation) {
    return { status: targetViolation.status, jsonBody: { error: targetViolation.error } };
  }

  const result = await graph.api(`${mailboxBase}/mailFolders/${folderId}`).patch({
    displayName: body.displayName,
  });

  return {
    status: 200,
    jsonBody: {
      id: result.id,
      displayName: result.displayName,
      parentFolderId: result.parentFolderId,
      status: 'renamed',
    },
  };
}

app.http('renameMailFolder', {
  methods: ['PATCH'],
  authLevel: 'anonymous',
  route: 'api/mail/folders/{folderId}',
  handler: withSecurity(withPolicyEnforcement('mail', renameMailFolderHandler)),
});
