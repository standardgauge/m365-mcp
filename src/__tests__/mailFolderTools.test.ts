/**
 * Mail folder MCP exposure tests.
 *
 * Covers the three new folder-management tools added to unblock mail-routing
 * automation in n8n ( / C&R v2): create_mail_folder, rename_mail_folder,
 * move_mail_folder.  (delete_mail_folder is deliberately NOT exposed as an MCP
 * tool — it is available only via direct REST with an explicit confirm=true
 * query param.)
 *
 * Two layers, mirroring calendarTools.test.ts:
 *   1. toolManifest — the new tools are registered in the TOOLS catalog with the
 *      correct REST endpoints, HTTP methods, and required-parameter lists.
 *   2. Remote MCP endpoint (/api/mcp JSON-RPC) — the tools surface in tools/list
 *      (gated by the "mail" enabledService) and are invokable via tools/call with
 *      deny-list enforcement intact.
 *
 * Schema-sync check: verifies that getManifest() and the live tools/list surface
 * agree on the input-schema shape (property names + types + required list) so a
 * client trusting the catalog won't send a payload the handler rejects.
 */

import { jest } from '@jest/globals';
import type { HttpRequest, InvocationContext } from '@azure/functions';

// ── Mock declarations ────────────────────────────────────────────────────────

const mockAuthenticateRequest = jest.fn<() => Promise<unknown>>();
const mockGetValidAccessTokenForSession = jest.fn<() => Promise<string>>();
const mockGetTenantIdFromSession = jest.fn<() => string>();
const mockIsPathDenied = jest.fn<(t: string, u: string, type: string, path: string) => Promise<boolean>>();
const mockGetEnabledServices = jest.fn<() => Promise<string[]>>();
const mockGetAllowedSites = jest.fn<() => Promise<Array<{ id: string; name: string }>>>();
const mockGetUserServiceOverrides = jest.fn<() => Promise<string[]>>();
const mockResolveMailFolderName = jest.fn<() => Promise<string | null>>();

const lastGraphCall: { path: string | null; postBody: unknown; patchBody: unknown } = {
  path: null, postBody: null, patchBody: null,
};
const mockGraphPost = jest.fn<(body?: unknown) => Promise<unknown>>();
const mockGraphPatch = jest.fn<(body?: unknown) => Promise<unknown>>();

function makeChain() {
  const chain: Record<string, unknown> = {};
  for (const m of ['select', 'top', 'orderby', 'filter', 'expand', 'header']) chain[m] = () => chain;
  chain.get = () => Promise.resolve({ value: [] });
  chain.post = (body?: unknown) => { lastGraphCall.postBody = body; return mockGraphPost(body); };
  chain.patch = (body?: unknown) => { lastGraphCall.patchBody = body; return mockGraphPatch(body); };
  chain.delete = () => Promise.resolve(undefined);
  return chain;
}
const mockCreateGraphClient = jest.fn(() => ({
  api: (path: string) => {
    lastGraphCall.path = path;
    return makeChain();
  },
}));

// telemetry.js has import-time side effects (patches console) — stub it out
jest.mock('../services/telemetry.js', () => ({}));

jest.mock('../services/authMiddleware.js', () => ({
  authenticateRequest: () => mockAuthenticateRequest(),
}));

jest.mock('../services/tokenCache.js', () => ({
  getValidAccessTokenForSession: () => mockGetValidAccessTokenForSession(),
  getTenantIdFromSession: () => mockGetTenantIdFromSession(),
}));

jest.mock('../services/graphClient.js', () => ({
  createGraphClient: () => mockCreateGraphClient(),
}));

jest.mock('../services/denyList.js', () => ({
  filterDeniedPaths: (_t: unknown, _u: unknown, _s: unknown, items: unknown) => Promise.resolve(items),
  isPathDenied: (...args: unknown[]) =>
    mockIsPathDenied(args[0] as string, args[1] as string, args[2] as string, args[3] as string),
}));

jest.mock('../services/containerResolver.js', () => ({
  resolveMailFolderName: () => mockResolveMailFolderName(),
  resolveDefaultCalendarId: () => Promise.resolve(null),
  resolveCalendarName: () => Promise.resolve(null),
  resolveContactParentFolder: () => Promise.resolve(null),
  resolveDefaultContactFolder: () => Promise.resolve(null),
  resolveSectionNotebook: () => Promise.resolve(null),
  normalizeContactFolderId: (_g: unknown, _u: unknown, id: string) => Promise.resolve(id),
}));

jest.mock('../services/sharepointFilter.js', () => ({
  filterAndDisambiguateSites: (sites: unknown) => sites,
}));

jest.mock('../services/serviceSettings.js', () => ({
  getEnabledServices: () => mockGetEnabledServices(),
  getReadOnlyServices: () => Promise.resolve([]),
  getAllowedSites: () => mockGetAllowedSites(),
}));

jest.mock('../services/userServiceOverrides.js', () => ({
  getUserServiceOverrides: () => mockGetUserServiceOverrides(),
}));

jest.mock('../services/userEmailSettings.js', () => ({
  getUserEmailSettings: () => Promise.resolve({ emailOutputMode: 'draft' }),
  setUserEmailSettings: () => Promise.resolve(undefined),
}));

jest.mock('../services/userMailConfig.js', () => ({
  isMailIndexingDisabled: jest.fn<() => Promise<boolean>>().mockResolvedValue(false),
}));

jest.mock('../services/auditLog.js', () => ({
  logAccess: jest.fn(),
}));

jest.mock('@azure/functions', () => ({
  app: { http: jest.fn() },
}));

// ── Import after mocks ──────────────────────────────────────────────────────

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
const AUTH = { userId: USER, session: { userId: USER, tenantId: TENANT, accessToken: 'fake', sessionToken: 'sess' } };

function rpc(method: string, params: Record<string, unknown>): Promise<{ status: number; jsonBody?: unknown }> {
  const req = {
    method: 'POST',
    headers: new Map<string, string>(),
    json: () => Promise.resolve({ jsonrpc: '2.0', id: 1, method, params }),
  } as unknown as HttpRequest;
  const ctx = { error: jest.fn(), warn: jest.fn(), log: jest.fn() } as unknown as InvocationContext;
  return handler(req, ctx);
}

const callTool = (name: string, args: Record<string, unknown>) => rpc('tools/call', { name, arguments: args });

function toolResult(res: { jsonBody?: unknown }): { parsed: unknown; isError: boolean; text: string } {
  const body = res.jsonBody as { result?: { content?: Array<{ text: string }>; isError?: boolean } };
  const text = body.result?.content?.[0]?.text ?? '{}';
  const isError = body.result?.isError ?? false;
  return { parsed: isError ? null : JSON.parse(text), isError, text };
}

const FOLDER_ID = 'AAMkFolderGuid001';
const NEW_FOLDER = { id: FOLDER_ID, displayName: '_Notifications', parentFolderId: 'msgroot' };

beforeEach(() => {
  jest.clearAllMocks();
  lastGraphCall.path = null;
  lastGraphCall.postBody = null;
  lastGraphCall.patchBody = null;
  mockAuthenticateRequest.mockResolvedValue(AUTH);
  mockGetValidAccessTokenForSession.mockResolvedValue('access-token');
  mockGetTenantIdFromSession.mockReturnValue(TENANT);
  mockIsPathDenied.mockResolvedValue(false);
  mockGetEnabledServices.mockResolvedValue(['mail', 'calendar', 'sharepoint']);
  mockGetAllowedSites.mockResolvedValue([]);
  mockGetUserServiceOverrides.mockResolvedValue([]);
  mockResolveMailFolderName.mockResolvedValue(null);
  mockGraphPost.mockResolvedValue(NEW_FOLDER);
  mockGraphPatch.mockResolvedValue({ ...NEW_FOLDER, displayName: 'github', status: 'renamed' });
});

// ─────────────────────────────────────────────────────────────────────────────
// toolManifest — new tools registered
// ─────────────────────────────────────────────────────────────────────────────

describe('toolManifest — mail folder tools', () => {
  const FOLDER_TOOLS = ['create_mail_folder', 'rename_mail_folder', 'move_mail_folder'];

  it('TOOLS contains all 3 mail folder tools', async () => {
    const { TOOLS } = await import('../mcp/toolManifest.js');
    const names = TOOLS.map((t) => t.name);
    for (const t of FOLDER_TOOLS) expect(names).toContain(t);
  });

  it('create_mail_folder has correct endpoint, method, and required params', async () => {
    const { TOOLS } = await import('../mcp/toolManifest.js');
    const tool = TOOLS.find((t) => t.name === 'create_mail_folder');
    expect(tool?.endpoint).toBe('/api/mail/folders');
    expect(tool?.method).toBe('POST');
    const required = (tool?.parameters ?? []).filter((p) => p.required).map((p) => p.name);
    expect(required).toEqual(['displayName']);
    const all = (tool?.parameters ?? []).map((p) => p.name);
    expect(all).toContain('parentFolderId');
    expect(all).toContain('mailboxId');
  });

  it('rename_mail_folder has correct endpoint, method, and required params', async () => {
    const { TOOLS } = await import('../mcp/toolManifest.js');
    const tool = TOOLS.find((t) => t.name === 'rename_mail_folder');
    expect(tool?.endpoint).toBe('/api/mail/folders/{folderId}');
    expect(tool?.method).toBe('PATCH');
    const required = (tool?.parameters ?? []).filter((p) => p.required).map((p) => p.name).sort();
    expect(required).toEqual(['displayName', 'folderId']);
  });

  it('move_mail_folder has correct endpoint, method, and required params', async () => {
    const { TOOLS } = await import('../mcp/toolManifest.js');
    const tool = TOOLS.find((t) => t.name === 'move_mail_folder');
    expect(tool?.endpoint).toBe('/api/mail/folders/{folderId}/move');
    expect(tool?.method).toBe('POST');
    const required = (tool?.parameters ?? []).filter((p) => p.required).map((p) => p.name).sort();
    expect(required).toEqual(['destinationParentFolderId', 'folderId']);
  });

  it('delete_mail_folder is NOT in TOOLS (admin-only REST endpoint)', async () => {
    const { TOOLS } = await import('../mcp/toolManifest.js');
    const names = TOOLS.map((t) => t.name);
    expect(names).not.toContain('delete_mail_folder');
  });

  it('getManifest() output includes the new folder tools', async () => {
    const { getManifest } = await import('../mcp/toolManifest.js');
    const manifest = getManifest() as { tools: Array<{ name: string }> };
    const names = manifest.tools.map((t) => t.name);
    for (const t of FOLDER_TOOLS) expect(names).toContain(t);
  });
});

// ─────────────────────────────────────────────────────────────────────────────
// Schema sync — getManifest() vs live tools/list
// ─────────────────────────────────────────────────────────────────────────────

describe('schema sync — manifest catalog vs live tools/list', () => {
  const FOLDER_TOOLS = ['create_mail_folder', 'rename_mail_folder', 'move_mail_folder'];

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  function shape(inputSchema: any) {
    const props: Record<string, unknown> = {};
    for (const [name, spec] of Object.entries(inputSchema.properties ?? {})) {
      const s = spec as Record<string, unknown>;
      props[name] = {
        type: s.type,
        ...(s.items ? { items: { type: (s.items as Record<string, unknown>).type } } : {}),
        ...(s.enum ? { enum: s.enum } : {}),
      };
    }
    return { type: inputSchema.type, required: [...(inputSchema.required ?? [])].sort(), props };
  }

  it('every new folder tool has the same input-schema shape in getManifest() and tools/list', async () => {
    const { getManifest } = await import('../mcp/toolManifest.js');
    const manifest = getManifest() as { tools: Array<{ name: string; inputSchema: unknown }> };
    const res = await rpc('tools/list', {});
    const body = res.jsonBody as { result: { tools: Array<{ name: string; inputSchema: unknown }> } };

    const manifestByName = Object.fromEntries(manifest.tools.map((t) => [t.name, t]));
    const liveByName = Object.fromEntries(body.result.tools.map((t) => [t.name, t]));

    for (const name of FOLDER_TOOLS) {
      expect(manifestByName[name]).toBeDefined();
      expect(liveByName[name]).toBeDefined();
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      expect(shape(manifestByName[name].inputSchema as any)).toEqual(shape(liveByName[name].inputSchema as any));
    }
  });
});

// ─────────────────────────────────────────────────────────────────────────────
// MCP tools/list — mail folder tools visible
// ─────────────────────────────────────────────────────────────────────────────

describe('MCP tools/list — mail folder tool exposure', () => {
  it('lists all 3 folder tools when the mail service is enabled', async () => {
    const res = await rpc('tools/list', {});
    const body = res.jsonBody as { result: { tools: Array<{ name: string }> } };
    const names = body.result.tools.map((t) => t.name);
    expect(names).toContain('create_mail_folder');
    expect(names).toContain('rename_mail_folder');
    expect(names).toContain('move_mail_folder');
    expect(names).not.toContain('delete_mail_folder');
  });

  it('hides folder tools when the mail service is disabled', async () => {
    mockGetEnabledServices.mockResolvedValue(['sharepoint', 'calendar']);
    const res = await rpc('tools/list', {});
    const body = res.jsonBody as { result: { tools: Array<{ name: string }> } };
    const names = body.result.tools.map((t) => t.name);
    expect(names).not.toContain('create_mail_folder');
    expect(names).not.toContain('rename_mail_folder');
    expect(names).not.toContain('move_mail_folder');
  });
});

// ─────────────────────────────────────────────────────────────────────────────
// MCP tools/call — create_mail_folder
// ─────────────────────────────────────────────────────────────────────────────

describe('MCP tools/call — create_mail_folder', () => {
  it('creates a root-level folder via POST /me/mailFolders', async () => {
    const res = await callTool('create_mail_folder', { displayName: '_Notifications' });
    const { parsed, isError } = toolResult(res);
    expect(isError).toBe(false);
    expect(parsed).toMatchObject({ id: FOLDER_ID, displayName: '_Notifications', status: 'created' });
    expect(lastGraphCall.path).toBe('/me/mailFolders');
    expect(lastGraphCall.postBody).toEqual({ displayName: '_Notifications' });
  });

  it('creates a child folder via POST /me/mailFolders/{parentId}/childFolders', async () => {
    await callTool('create_mail_folder', { displayName: 'Sub', parentFolderId: 'parent-123' });
    expect(lastGraphCall.path).toBe('/me/mailFolders/parent-123/childFolders');
  });

  it('routes to /users/{mailboxId} when a non-me mailboxId is provided', async () => {
    await callTool('create_mail_folder', { displayName: '_Shared', mailboxId: 'other@example.com' });
    expect(lastGraphCall.path).toBe('/users/other@example.com/mailFolders');
  });

  it('returns an error when displayName is missing', async () => {
    const res = await callTool('create_mail_folder', {});
    const { isError } = toolResult(res);
    expect(isError).toBe(true);
    expect(mockGraphPost).not.toHaveBeenCalled();
  });

  it('returns an error when the displayName is deny-listed', async () => {
    mockIsPathDenied.mockImplementation(async (_t, _u, _type, path) => path === '_Sensitive');
    const res = await callTool('create_mail_folder', { displayName: '_Sensitive' });
    const { isError } = toolResult(res);
    expect(isError).toBe(true);
    expect(mockGraphPost).not.toHaveBeenCalled();
  });

  it('rejects a path-injection parentFolderId', async () => {
    const res = await callTool('create_mail_folder', { displayName: 'X', parentFolderId: '../../users/victim' });
    const { isError } = toolResult(res);
    expect(isError).toBe(true);
    expect(mockGraphPost).not.toHaveBeenCalled();
  });

  it('is rejected when the mail service is not enabled', async () => {
    mockGetEnabledServices.mockResolvedValue(['sharepoint']);
    const res = await callTool('create_mail_folder', { displayName: '_Notifications' });
    const { isError } = toolResult(res);
    expect(isError).toBe(true);
    expect(mockGraphPost).not.toHaveBeenCalled();
  });
});

// ─────────────────────────────────────────────────────────────────────────────
// MCP tools/call — rename_mail_folder
// ─────────────────────────────────────────────────────────────────────────────

describe('MCP tools/call — rename_mail_folder', () => {
  it('renames a folder via PATCH /me/mailFolders/{folderId}', async () => {
    const res = await callTool('rename_mail_folder', { folderId: FOLDER_ID, displayName: 'github' });
    const { isError } = toolResult(res);
    expect(isError).toBe(false);
    expect(lastGraphCall.path).toBe(`/me/mailFolders/${FOLDER_ID}`);
    expect(lastGraphCall.patchBody).toEqual({ displayName: 'github' });
  });

  it('routes to /users/{mailboxId} when a non-me mailboxId is provided', async () => {
    await callTool('rename_mail_folder', { folderId: FOLDER_ID, displayName: 'github', mailboxId: 'other@example.com' });
    expect(lastGraphCall.path).toBe(`/users/other@example.com/mailFolders/${FOLDER_ID}`);
  });

  it('returns an error when the target name is deny-listed', async () => {
    mockIsPathDenied.mockImplementation(async (_t, _u, _type, path) => path === '_Sensitive');
    const res = await callTool('rename_mail_folder', { folderId: FOLDER_ID, displayName: '_Sensitive' });
    const { isError } = toolResult(res);
    expect(isError).toBe(true);
    expect(mockGraphPatch).not.toHaveBeenCalled();
  });

  it('returns an error when the source folder is deny-listed', async () => {
    mockResolveMailFolderName.mockResolvedValue('_Sensitive');
    mockIsPathDenied.mockImplementation(async (_t, _u, _type, path) => path === '_Sensitive');
    const res = await callTool('rename_mail_folder', { folderId: FOLDER_ID, displayName: 'NewName' });
    const { isError } = toolResult(res);
    expect(isError).toBe(true);
    expect(mockGraphPatch).not.toHaveBeenCalled();
  });

  it('rejects a path-injection folderId', async () => {
    const res = await callTool('rename_mail_folder', { folderId: '../../../users/victim', displayName: 'ok' });
    const { isError } = toolResult(res);
    expect(isError).toBe(true);
    expect(mockGraphPatch).not.toHaveBeenCalled();
  });
});

// ─────────────────────────────────────────────────────────────────────────────
// MCP tools/call — move_mail_folder
// ─────────────────────────────────────────────────────────────────────────────

describe('MCP tools/call — move_mail_folder', () => {
  it('moves a folder via POST /me/mailFolders/{folderId}/move', async () => {
    const res = await callTool('move_mail_folder', { folderId: FOLDER_ID, destinationParentFolderId: 'archive-id' });
    const { isError } = toolResult(res);
    expect(isError).toBe(false);
    expect(lastGraphCall.path).toBe(`/me/mailFolders/${FOLDER_ID}/move`);
    expect(lastGraphCall.postBody).toEqual({ destinationId: 'archive-id' });
  });

  it('routes to /users/{mailboxId} when a non-me mailboxId is provided', async () => {
    await callTool('move_mail_folder', { folderId: FOLDER_ID, destinationParentFolderId: 'archive-id', mailboxId: 'other@example.com' });
    expect(lastGraphCall.path).toBe(`/users/other@example.com/mailFolders/${FOLDER_ID}/move`);
  });

  it('returns an error when a folder is deny-listed', async () => {
    mockResolveMailFolderName.mockResolvedValue('_Sensitive');
    mockIsPathDenied.mockImplementation(async (_t, _u, _type, path) => path === '_Sensitive');
    const res = await callTool('move_mail_folder', { folderId: FOLDER_ID, destinationParentFolderId: 'archive-id' });
    const { isError } = toolResult(res);
    expect(isError).toBe(true);
    expect(mockGraphPost).not.toHaveBeenCalled();
  });

  it('rejects a path-injection destinationParentFolderId', async () => {
    const res = await callTool('move_mail_folder', { folderId: FOLDER_ID, destinationParentFolderId: '../../users/victim' });
    const { isError } = toolResult(res);
    expect(isError).toBe(true);
    expect(mockGraphPost).not.toHaveBeenCalled();
  });
});
