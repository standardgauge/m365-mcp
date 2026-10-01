/**
 * Attachment support on the remote MCP mail tools.
 *
 * MCP clients call the JSON-RPC /api/mcp tools directly. These tests assert that
 * send_mail and create_draft advertise the `attachments` parameter in tools/list
 * and that the handlers forward attachments to Graph (inline for small files) and
 * reject malformed attachment input.
 */

import { jest } from '@jest/globals';
import type { HttpRequest, InvocationContext } from '@azure/functions';

const mockAuthenticateRequest = jest.fn<() => Promise<unknown>>();
const mockGetValidAccessTokenForSession = jest.fn<() => Promise<string>>();
const mockGetTenantIdFromSession = jest.fn<() => string>();
const mockIsPathDenied = jest.fn<() => Promise<boolean>>();
const mockGetEnabledServices = jest.fn<() => Promise<string[]>>();
const mockGetAllowedSites = jest.fn<() => Promise<Array<{ id: string; name: string }>>>();
const mockGetUserServiceOverrides = jest.fn<() => Promise<string[]>>();
const mockGetUserEmailSettings = jest.fn<() => Promise<{ emailOutputMode: string }>>();

const mockGraphPost = jest.fn<(payload: unknown) => Promise<unknown>>();
const mockGraphApi = jest.fn((_path: string) => ({ post: mockGraphPost }));
const mockCreateGraphClient = jest.fn(() => ({ api: mockGraphApi }));

jest.mock('../services/telemetry.js', () => ({}));
jest.mock('../services/authMiddleware.js', () => ({ authenticateRequest: () => mockAuthenticateRequest() }));
jest.mock('../services/tokenCache.js', () => ({
  getValidAccessTokenForSession: () => mockGetValidAccessTokenForSession(),
  getTenantIdFromSession: () => mockGetTenantIdFromSession(),
}));
jest.mock('../services/graphClient.js', () => ({ createGraphClient: () => mockCreateGraphClient() }));
jest.mock('../services/denyList.js', () => ({
  filterDeniedPaths: (_t: unknown, _u: unknown, _s: unknown, items: unknown) => Promise.resolve(items),
  isPathDenied: () => mockIsPathDenied(),
}));
jest.mock('../services/containerResolver.js', () => ({
  resolveMailFolderName: () => Promise.resolve(null),
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
  getAllowedSites: () => mockGetAllowedSites(),
}));
jest.mock('../services/userServiceOverrides.js', () => ({ getUserServiceOverrides: () => mockGetUserServiceOverrides() }));
jest.mock('../services/userEmailSettings.js', () => ({
  getUserEmailSettings: () => mockGetUserEmailSettings(),
  setUserEmailSettings: () => Promise.resolve(undefined),
}));
jest.mock('../services/userMailConfig.js', () => ({
  isMailIndexingDisabled: jest.fn<() => Promise<boolean>>().mockResolvedValue(false),
}));
jest.mock('../services/auditLog.js', () => ({ logAccess: jest.fn() }));
jest.mock('@azure/functions', () => ({ app: { http: jest.fn() } }));

import { app } from '@azure/functions';
import '../functions/mcp/mcpEndpoint.js';

interface HttpRegistration {
  handler: (req: HttpRequest, context: InvocationContext) => Promise<{ status: number; jsonBody?: unknown }>;
}
const httpMock = app.http as unknown as jest.Mock<(name: string, opts: HttpRegistration) => void>;
const registration = httpMock.mock.calls.find((call) => call[0] === 'mcpEndpoint');
if (!registration) throw new Error('mcpEndpoint handler was not registered');
const handler = registration[1].handler;

const TENANT = 'test-tenant';
const USER = 'test-user';
const SESSION = { userId: USER, tenantId: TENANT, accessToken: 't', sessionToken: 's' };
const AUTH = { userId: USER, session: SESSION };

function rpc(method: string, params: Record<string, unknown>): Promise<{ status: number; jsonBody?: unknown }> {
  const req = {
    method: 'POST',
    headers: new Map<string, string>(),
    json: () => Promise.resolve({ jsonrpc: '2.0', id: 1, method, params }),
  } as unknown as HttpRequest;
  const ctx = { error: jest.fn(), warn: jest.fn(), log: jest.fn() } as unknown as InvocationContext;
  return handler(req, ctx);
}

function toolResult(res: { jsonBody?: unknown }): { parsed: unknown; isError: boolean } {
  const body = res.jsonBody as { result?: { content?: Array<{ text: string }>; isError?: boolean } };
  const text = body.result?.content?.[0]?.text ?? '{}';
  const isError = body.result?.isError ?? false;
  return { parsed: isError ? null : JSON.parse(text), isError };
}

const B64 = Buffer.from('hello world').toString('base64');

beforeEach(() => {
  jest.clearAllMocks();
  mockAuthenticateRequest.mockResolvedValue(AUTH);
  mockGetValidAccessTokenForSession.mockResolvedValue('access-token');
  mockGetTenantIdFromSession.mockReturnValue(TENANT);
  mockIsPathDenied.mockResolvedValue(false);
  mockGetEnabledServices.mockResolvedValue(['mail']);
  mockGetAllowedSites.mockResolvedValue([]);
  mockGetUserServiceOverrides.mockResolvedValue([]);
});

describe('tools/list advertises attachments', () => {
  it('send_mail and create_draft carry an attachments array with object items', async () => {
    const res = await rpc('tools/list', {});
    const body = res.jsonBody as { result: { tools: Array<{ name: string; inputSchema: { properties: Record<string, { type: string; items?: { type: string; required?: string[] } }> } }> } };
    const byName = Object.fromEntries(body.result.tools.map((t) => [t.name, t]));

    for (const name of ['send_mail', 'create_draft']) {
      const att = byName[name].inputSchema.properties.attachments;
      expect(att).toBeDefined();
      expect(att.type).toBe('array');
      expect(att.items?.type).toBe('object');
      // No fixed required list: the fields depend on the source.
      expect(att.items?.required).toBeUndefined();
      const itemProps = Object.keys((att.items as unknown as { properties: Record<string, unknown> }).properties);
      expect(itemProps).toEqual(expect.arrayContaining(['name', 'contentType', 'content', 'driveItemId', 'siteId', 'itemId', 'driveId']));
    }
  });
});

describe('create_draft forwards attachments', () => {
  it('includes a small attachment inline on the /me/messages POST', async () => {
    mockGraphPost.mockResolvedValue({ id: 'draft-1', webLink: 'https://outlook/draft-1' });
    const res = await rpc('tools/call', {
      name: 'create_draft',
      arguments: { subject: 'S', body: 'B', to: ['a@x.com'], attachments: [{ name: 'a.txt', contentType: 'text/plain', content: B64 }] },
    });
    const { parsed, isError } = toolResult(res);

    expect(isError).toBe(false);
    expect(parsed).toMatchObject({ status: 'draft', id: 'draft-1' });
    expect(mockGraphApi).toHaveBeenCalledWith('/me/messages');
    const postedBody = mockGraphPost.mock.calls[0][0] as { attachments?: Array<{ '@odata.type': string; name: string }> };
    expect(postedBody.attachments).toHaveLength(1);
    expect(postedBody.attachments![0]).toMatchObject({ '@odata.type': '#microsoft.graph.fileAttachment', name: 'a.txt' });
  });

  it('rejects a malformed attachment (bad base64) as an error', async () => {
    const res = await rpc('tools/call', {
      name: 'create_draft',
      arguments: { subject: 'S', to: ['a@x.com'], attachments: [{ name: 'a.txt', content: 'not base64!!!' }] },
    });
    const { isError } = toolResult(res);
    expect(isError).toBe(true);
    expect(mockGraphPost).not.toHaveBeenCalled();
  });
});

describe('send_mail forwards attachments (send mode)', () => {
  it('includes a small attachment in the /me/sendMail message', async () => {
    mockGetUserEmailSettings.mockResolvedValue({ emailOutputMode: 'send' });
    mockGraphPost.mockResolvedValue(undefined);
    const res = await rpc('tools/call', {
      name: 'send_mail',
      arguments: { subject: 'S', body: 'B', to: ['a@x.com'], attachments: [{ name: 'a.txt', content: B64 }] },
    });
    const { parsed } = toolResult(res);

    expect(parsed).toMatchObject({ status: 'sent' });
    expect(mockGraphApi).toHaveBeenCalledWith('/me/sendMail');
    const postedBody = mockGraphPost.mock.calls[0][0] as { message: { attachments?: unknown[] }; saveToSentItems: boolean };
    expect(postedBody.message.attachments).toHaveLength(1);
    expect(postedBody.saveToSentItems).toBe(true);
  });
});
