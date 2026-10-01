import { app, HttpRequest, HttpResponseInit, InvocationContext } from '@azure/functions';
import { getValidAccessTokenForSession, getTenantId } from '../../services/tokenCache.js';
import { createGraphClient } from '../../services/graphClient.js';
import { isPathDenied } from '../../services/denyList.js';
import { withPolicyEnforcement } from '../../services/policyEnforcement.js';
import { assertOpaqueId } from '../../services/opaqueId.js';
import type { AuthResult } from '../../services/authMiddleware.js';
import { withSecurity } from '../../services/securityHeaders.js';

const TEXT_MIME_PREFIXES = ['text/', 'application/json', 'application/xml', 'application/javascript'];

function isTextMime(mimeType: string): boolean {
  return TEXT_MIME_PREFIXES.some((prefix) => mimeType.startsWith(prefix));
}

async function readOneDriveFileHandler(
  request: HttpRequest,
  context: InvocationContext,
  auth: AuthResult,
): Promise<HttpResponseInit> {
  const userId = auth.userId;
  const accessToken = await getValidAccessTokenForSession(auth.session);

  const itemId = request.params['itemId'];
  if (!itemId) {
    return { status: 400, jsonBody: { error: 'itemId path param is required' } };
  }
  assertOpaqueId(itemId, 'itemId');

  const graph = createGraphClient(accessToken);

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const metadata: any = await graph
    .api(`/me/drive/items/${itemId}`)
    .select('id,name,webUrl,size,file,parentReference,createdDateTime,lastModifiedDateTime')
    .get();

  const parentPath: string = metadata.parentReference?.path ?? '';
  const filePath = parentPath ? `${parentPath}/${metadata.name}` : `/${metadata.name}`;

  const denied = await isPathDenied(await getTenantId(userId), userId, 'onedrive', filePath);
  if (denied) {
    return {
      status: 403,
      jsonBody: { error: 'Access to this file is restricted by the deny list' },
    };
  }

  if (!metadata.file) {
    return { status: 400, jsonBody: { error: 'The specified item is not a file' } };
  }

  const stream = await graph.api(`/me/drive/items/${itemId}/content`).getStream();
  const chunks: Buffer[] = [];
  for await (const chunk of stream) {
    chunks.push(Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk as Uint8Array));
  }
  const fileBuffer = Buffer.concat(chunks);

  const mimeType: string = metadata.file.mimeType ?? 'application/octet-stream';
  const textContent = isTextMime(mimeType);

  return {
    status: 200,
    jsonBody: {
      id: metadata.id,
      name: metadata.name,
      webUrl: metadata.webUrl,
      mimeType,
      size: metadata.size,
      createdDateTime: metadata.createdDateTime,
      lastModifiedDateTime: metadata.lastModifiedDateTime,
      content: textContent ? fileBuffer.toString('utf-8') : fileBuffer.toString('base64'),
      encoding: textContent ? 'utf-8' : 'base64',
    },
  };
}

app.http('readOneDriveFile', {
  methods: ['GET'],
  authLevel: 'anonymous',
  route: 'api/onedrive/files/{itemId}',
  handler: withSecurity(withPolicyEnforcement('onedrive', readOneDriveFileHandler)),
});
