import { app, HttpRequest, HttpResponseInit, InvocationContext } from '@azure/functions';
import { getValidAccessTokenForSession, getTenantId } from '../../services/tokenCache.js';
import { createGraphClient } from '../../services/graphClient.js';
import { isPathDenied } from '../../services/denyList.js';
import { withPolicyEnforcement } from '../../services/policyEnforcement.js';
import { assertOpaqueId } from '../../services/opaqueId.js';
import type { AuthResult } from '../../services/authMiddleware.js';
import { withSecurity } from '../../services/securityHeaders.js';

// MIME types we treat as readable text and return as UTF-8 strings
const TEXT_MIME_PREFIXES = ['text/', 'application/json', 'application/xml', 'application/javascript'];

function isTextMime(mimeType: string): boolean {
  return TEXT_MIME_PREFIXES.some((prefix) => mimeType.startsWith(prefix));
}

async function readFileHandler(
  request: HttpRequest,
  context: InvocationContext,
  auth: AuthResult,
): Promise<HttpResponseInit> {
  const userId = auth.userId;

  const siteId = request.query.get('siteId');
  const itemId = request.params['itemId'];
  const driveId = request.query.get('driveId');

  if (!siteId || !itemId) {
    return {
      status: 400,
      jsonBody: { error: 'siteId query param and itemId path param are required' },
    };
  }
  assertOpaqueId(siteId, 'siteId');
  assertOpaqueId(itemId, 'itemId');
  if (driveId) assertOpaqueId(driveId, 'driveId');

  const accessToken = await getValidAccessTokenForSession(auth.session);
  const graph = createGraphClient(accessToken);

  // Fetch metadata first so we can check the deny list before downloading
  const metaPath = driveId
    ? `/sites/${siteId}/drives/${driveId}/items/${itemId}`
    : `/sites/${siteId}/drive/items/${itemId}`;

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const metadata: any = await graph
    .api(metaPath)
    .select('id,name,webUrl,size,file,parentReference,createdDateTime,lastModifiedDateTime')
    .get();

  // Build the deny-list path from parentReference
  const parentPath: string = metadata.parentReference?.path ?? '';
  const filePath = parentPath ? `${parentPath}/${metadata.name}` : `/${metadata.name}`;

  const denied = await isPathDenied(await getTenantId(userId), userId, 'sharepoint', filePath);
  if (denied) {
    return {
      status: 403,
      jsonBody: { error: 'Access to this file is restricted by the deny list' },
    };
  }

  if (!metadata.file) {
    return { status: 400, jsonBody: { error: 'The specified item is not a file' } };
  }

  // Download file content
  const contentPath = driveId
    ? `/sites/${siteId}/drives/${driveId}/items/${itemId}/content`
    : `/sites/${siteId}/drive/items/${itemId}/content`;

  const stream = await graph.api(contentPath).getStream();
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

app.http('readFile', {
  methods: ['GET'],
  authLevel: 'anonymous',
  route: 'api/sharepoint/files/{itemId}',
  handler: withSecurity(withPolicyEnforcement('sharepoint', readFileHandler, {
    getSiteId: (req) => req.query.get('siteId') ?? undefined,
  })),
});
