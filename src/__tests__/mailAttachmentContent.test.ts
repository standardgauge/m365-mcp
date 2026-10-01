/**
 * Attachment content for attached emails.
 *
 * `get_attachments` with an attachmentId used to return only metadata for a
 * `#microsoft.graph.itemAttachment` (an email attached to an email) because it read
 * `contentBytes`, which item attachments do not carry. These tests pin:
 *   - readMailAttachment: item attachments are read from `/$value` and returned as
 *     the raw message/rfc822 MIME; file attachments keep their existing shape;
 *     non-UTF-8 text falls back to base64; reference attachments carry no content.
 *   - The MCP get_attachments tool returns the forwarded thread end to end.
 */

import { jest } from '@jest/globals';
import { readFileSync } from 'fs';
import { join } from 'path';
import type { HttpRequest, InvocationContext } from '@azure/functions';

const mockAuthenticateRequest = jest.fn<() => Promise<unknown>>();
const mockGetValidAccessTokenForSession = jest.fn<() => Promise<string>>();
const mockGetTenantIdFromSession = jest.fn<() => string>();
const mockIsPathDenied = jest.fn<() => Promise<boolean>>();
const mockGetEnabledServices = jest.fn<() => Promise<string[]>>();
const mockGetUserServiceOverrides = jest.fn<() => Promise<string[]>>();

// Graph fake: path → resource for .get(), path → bytes for .getStream().
let graphGets: Record<string, unknown> = {};
let graphStreams: Record<string, Buffer> = {};
const requestedPaths: string[] = [];
function fakeApi(path: string) {
  requestedPaths.push(path);
  const req = {
    select: () => req,
    get: () => (path in graphGets ? Promise.resolve(graphGets[path]) : Promise.reject(new Error(`unexpected GET ${path}`))),
    getStream: () => {
      if (!(path in graphStreams)) return Promise.reject(new Error(`unexpected stream ${path}`));
      const bytes = graphStreams[path];
      // Deliver in two chunks to exercise reassembly.
      return Promise.resolve((async function* () {
        yield bytes.subarray(0, 10);
        yield bytes.subarray(10);
      })());
    },
  };
  return req;
}
const fakeGraph = { api: fakeApi };

jest.mock('../services/telemetry.js', () => ({}));
jest.mock('../services/authMiddleware.js', () => ({ authenticateRequest: () => mockAuthenticateRequest() }));
jest.mock('../services/tokenCache.js', () => ({
  getValidAccessTokenForSession: () => mockGetValidAccessTokenForSession(),
  getTenantIdFromSession: () => mockGetTenantIdFromSession(),
}));
jest.mock('../services/graphClient.js', () => ({ createGraphClient: () => fakeGraph }));
jest.mock('../services/denyList.js', () => ({
  filterDeniedPaths: (_t: unknown, _u: unknown, _s: unknown, items: unknown) => Promise.resolve(items),
  isPathDenied: () => mockIsPathDenied(),
}));
jest.mock('../services/containerResolver.js', () => ({
  resolveMailFolderName: () => Promise.resolve('Inbox'),
  resolveDefaultCalendarId: () => Promise.resolve(null),
  resolveContactParentFolder: () => Promise.resolve(null),
  resolveDefaultContactFolder: () => Promise.resolve(null),
  resolveSectionNotebook: () => Promise.resolve(null),
  normalizeContactFolderId: (_g: unknown, _u: unknown, id: string) => Promise.resolve(id),
}));
jest.mock('../services/sharepointFilter.js', () => ({ filterAndDisambiguateSites: (sites: unknown) => sites }));
jest.mock('../services/serviceSettings.js', () => ({
  getEnabledServices: () => mockGetEnabledServices(),
  getReadOnlyServices: () => Promise.resolve([]),
  getAllowedSites: () => Promise.resolve([]),
}));
jest.mock('../services/userServiceOverrides.js', () => ({ getUserServiceOverrides: () => mockGetUserServiceOverrides() }));
jest.mock('../services/userEmailSettings.js', () => ({
  getUserEmailSettings: () => Promise.resolve({ emailOutputMode: 'send' }),
  setUserEmailSettings: () => Promise.resolve(undefined),
}));
jest.mock('../services/userMailConfig.js', () => ({
  isMailIndexingDisabled: jest.fn<() => Promise<boolean>>().mockResolvedValue(false),
}));
jest.mock('../services/auditLog.js', () => ({ logAccess: jest.fn() }));
jest.mock('@azure/functions', () => ({ app: { http: jest.fn() } }));

import { app } from '@azure/functions';
import { readMailAttachment } from '../services/mailAttachmentContent.js';
import '../functions/mcp/mcpEndpoint.js';

interface HttpRegistration {
  handler: (req: HttpRequest, context: InvocationContext) => Promise<{ status: number; jsonBody?: unknown }>;
}
const httpMock = app.http as unknown as jest.Mock<(name: string, opts: HttpRegistration) => void>;
const registration = httpMock.mock.calls.find((call) => call[0] === 'mcpEndpoint');
if (!registration) throw new Error('mcpEndpoint handler was not registered');
const handler = registration[1].handler;

const SESSION = { userId: 'test-user', tenantId: 'test-tenant', accessToken: 't', sessionToken: 's' };

function callTool(name: string, args: Record<string, unknown>) {
  const req = {
    method: 'POST',
    headers: new Map<string, string>(),
    json: () => Promise.resolve({ jsonrpc: '2.0', id: 1, method: 'tools/call', params: { name, arguments: args } }),
  } as unknown as HttpRequest;
  const ctx = { error: jest.fn(), warn: jest.fn(), log: jest.fn() } as unknown as InvocationContext;
  return handler(req, ctx);
}

function toolResult(res: { jsonBody?: unknown }): { parsed: Record<string, unknown> | null; isError: boolean; text: string } {
  const body = res.jsonBody as { result?: { content?: Array<{ text: string }>; isError?: boolean } };
  const text = body.result?.content?.[0]?.text ?? '{}';
  const isError = body.result?.isError ?? false;
  return { parsed: isError ? null : JSON.parse(text), isError, text };
}

const EML = readFileSync(join(process.cwd(), 'src/__tests__/fixtures/forwarded-thread.eml'));
const ATT_PATH = '/me/messages/m1/attachments/att1';

const ITEM_ATTACHMENT = {
  '@odata.type': '#microsoft.graph.itemAttachment',
  id: 'att1',
  name: 'FW: Q3 vendor contract thread',
  contentType: 'message/rfc822',
  size: 4821,
  isInline: false,
};

beforeEach(() => {
  jest.clearAllMocks();
  graphGets = {};
  graphStreams = {};
  requestedPaths.length = 0;
  mockAuthenticateRequest.mockResolvedValue({ userId: 'test-user', session: SESSION });
  mockGetValidAccessTokenForSession.mockResolvedValue('access-token');
  mockGetTenantIdFromSession.mockReturnValue('test-tenant');
  mockIsPathDenied.mockResolvedValue(false);
  mockGetEnabledServices.mockResolvedValue(['mail']);
  mockGetUserServiceOverrides.mockResolvedValue([]);
});

describe('readMailAttachment', () => {
  it('returns the raw rfc822 MIME of an attached email from /$value', async () => {
    graphGets[ATT_PATH] = ITEM_ATTACHMENT;
    graphStreams[`${ATT_PATH}/$value`] = EML;

    const out = await readMailAttachment(fakeGraph as never, ATT_PATH);

    expect(requestedPaths).toContain(`${ATT_PATH}/$value`);
    expect(out).toMatchObject({ id: 'att1', name: ITEM_ATTACHMENT.name, contentType: 'message/rfc822', size: 4821, attachmentType: 'item', encoding: 'utf-8' });
    expect(out.content).toBe(EML.toString('utf-8'));
    expect(out.content).toContain('Subject: FW: Q3 vendor contract thread');
    // Non-ASCII body text survives the round trip.
    expect(out.content).toContain('café meeting moved to Friday');
    // Nested attachments of the forwarded message are part of the MIME.
    expect(out.content).toContain('filename="terms.txt"');
  });

  it('defaults a missing item-attachment contentType to message/rfc822', async () => {
    graphGets[ATT_PATH] = { ...ITEM_ATTACHMENT, contentType: null };
    graphStreams[`${ATT_PATH}/$value`] = EML;

    const out = await readMailAttachment(fakeGraph as never, ATT_PATH);
    expect(out.contentType).toBe('message/rfc822');
    expect(out.encoding).toBe('utf-8');
  });

  it('returns base64 when the MIME is not valid UTF-8, so no byte is lost', async () => {
    const latin1 = Buffer.concat([Buffer.from('Subject: caf'), Buffer.from([0xe9]), Buffer.from('\r\n\r\nbody')]);
    graphGets[ATT_PATH] = ITEM_ATTACHMENT;
    graphStreams[`${ATT_PATH}/$value`] = latin1;

    const out = await readMailAttachment(fakeGraph as never, ATT_PATH);
    expect(out.encoding).toBe('base64');
    expect(Buffer.from(out.content!, 'base64').equals(latin1)).toBe(true);
  });

  it('keeps the fileAttachment shape: text decoded, binary as base64, no /$value call', async () => {
    graphGets[ATT_PATH] = { '@odata.type': '#microsoft.graph.fileAttachment', id: 'att1', name: 'a.txt', contentType: 'text/plain', size: 5, contentBytes: Buffer.from('hello').toString('base64') };
    expect(await readMailAttachment(fakeGraph as never, ATT_PATH)).toEqual({ id: 'att1', name: 'a.txt', contentType: 'text/plain', size: 5, attachmentType: 'file', content: 'hello', encoding: 'utf-8' });

    graphGets[ATT_PATH] = { '@odata.type': '#microsoft.graph.fileAttachment', id: 'att1', name: 'a.pdf', contentType: 'application/pdf', size: 3, contentBytes: 'JVBE' };
    expect(await readMailAttachment(fakeGraph as never, ATT_PATH)).toMatchObject({ content: 'JVBE', encoding: 'base64', attachmentType: 'file' });

    expect(requestedPaths.some((p) => p.endsWith('/$value'))).toBe(false);
  });

  it('decodes a .eml file attachment (message/rfc822 fileAttachment) as text', async () => {
    graphGets[ATT_PATH] = { '@odata.type': '#microsoft.graph.fileAttachment', id: 'att1', name: 'thread.eml', contentType: 'message/rfc822', size: EML.length, contentBytes: EML.toString('base64') };
    const out = await readMailAttachment(fakeGraph as never, ATT_PATH);
    expect(out.encoding).toBe('utf-8');
    expect(out.content).toContain('Subject: FW: Q3 vendor contract thread');
  });

  it('returns metadata only for a reference attachment', async () => {
    graphGets[ATT_PATH] = { '@odata.type': '#microsoft.graph.referenceAttachment', id: 'att1', name: 'Deck', contentType: null, size: 0 };
    const out = await readMailAttachment(fakeGraph as never, ATT_PATH);
    expect(out).toEqual({ id: 'att1', name: 'Deck', contentType: 'application/octet-stream', size: 0, attachmentType: 'reference' });
  });
});

describe('MCP get_attachments on an attached email', () => {
  it('returns the forwarded thread content instead of metadata only', async () => {
    graphGets['/me/messages/m1'] = { parentFolderId: 'folder-1' };
    graphGets[ATT_PATH] = ITEM_ATTACHMENT;
    graphStreams[`${ATT_PATH}/$value`] = EML;

    const { parsed, isError, text } = toolResult(await callTool('get_attachments', { messageId: 'm1', attachmentId: 'att1' }));

    expect(isError).toBe(false);
    expect(text).not.toContain('unexpected');
    expect(parsed).toMatchObject({ id: 'att1', contentType: 'message/rfc822', attachmentType: 'item', encoding: 'utf-8' });
    expect(parsed!.content).toContain('The renewal terms changed on page 2');
  });

  it('still enforces the deny list before reading the attachment', async () => {
    graphGets['/me/messages/m1'] = { parentFolderId: 'folder-1' };
    graphGets[ATT_PATH] = ITEM_ATTACHMENT;
    graphStreams[`${ATT_PATH}/$value`] = EML;
    mockIsPathDenied.mockResolvedValue(true);

    const res = toolResult(await callTool('get_attachments', { messageId: 'm1', attachmentId: 'att1' }));
    expect(res.isError).toBe(true);
    expect(requestedPaths).not.toContain(`${ATT_PATH}/$value`);
  });
});
