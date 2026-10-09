import { app, HttpRequest, HttpResponseInit, InvocationContext } from '@azure/functions';
import { getValidAccessTokenForSession } from '../../services/tokenCache.js';
import { createGraphClient } from '../../services/graphClient.js';
import { resolveDenySubject } from '../../services/mailboxOwner.js';
import { withPolicyEnforcement, checkDenyList } from '../../services/policyEnforcement.js';
import { resolveMailFolderName } from '../../services/containerResolver.js';
import { assertOpaqueId } from '../../services/opaqueId.js';
import type { AuthResult } from '../../services/authMiddleware.js';
import { withSecurity } from '../../services/securityHeaders.js';

interface MoveMailFolderRequest {
  destinationParentFolderId: string;
  mailboxId?: string;
}

async function moveMailFolderHandler(
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

  const body = (await request.json()) as MoveMailFolderRequest;
  if (!body.destinationParentFolderId) {
    return { status: 400, jsonBody: { error: 'destinationParentFolderId is required' } };
  }
  assertOpaqueId(body.destinationParentFolderId, 'destinationParentFolderId');
  if (body.mailboxId && body.mailboxId !== 'me') assertOpaqueId(body.mailboxId, 'mailboxId');

  const accessToken = await getValidAccessTokenForSession(auth.session);
  const graph = createGraphClient(accessToken);
  const denySubject = await resolveDenySubject(graph, userId, body.mailboxId);

  const mailboxBase = body.mailboxId && body.mailboxId !== 'me'
    ? `/users/${body.mailboxId}`
    : '/me';

  // Deny-list check on the folder being moved (resolve its display name).
  const movedName = await resolveMailFolderName(graph, folderId, mailboxBase);
  if (movedName) {
    const movedViolation = await checkDenyList(userId, 'mail', movedName, undefined, denySubject);
    if (movedViolation) {
      return { status: movedViolation.status, jsonBody: { error: movedViolation.error } };
    }
  }

  // Deny-list check on the destination parent folder.
  const destName = await resolveMailFolderName(graph, body.destinationParentFolderId, mailboxBase);
  if (destName) {
    const destViolation = await checkDenyList(userId, 'mail', destName, undefined, denySubject);
    if (destViolation) {
      return { status: destViolation.status, jsonBody: { error: 'Destination folder restricted by deny list' } };
    }
  }

  const result = await graph.api(`${mailboxBase}/mailFolders/${folderId}/move`).post({
    destinationId: body.destinationParentFolderId,
  });

  return {
    status: 200,
    jsonBody: {
      id: result.id,
      displayName: result.displayName,
      parentFolderId: result.parentFolderId,
      status: 'moved',
    },
  };
}

app.http('moveMailFolder', {
  methods: ['POST'],
  authLevel: 'anonymous',
  route: 'api/mail/folders/{folderId}/move',
  handler: withSecurity(withPolicyEnforcement('mail', moveMailFolderHandler)),
});
