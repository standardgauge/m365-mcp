/**
 * Tests for email output mode enforcement on the remote MCP endpoint.
 *
 * MCP clients (Claude Desktop/Code, Cowork) call the JSON-RPC `/api/mcp` tools
 * directly, bypassing the REST `/api/mail/send` handler. Codex's review of PR #81
 * flagged that the MCP `send_mail` tool always delivered immediately, so a user
 * left in the new default draft mode could still have email go out through MCP.
 *
 * Covers:
 *   - send_mail in draft mode: saves to Drafts (/me/messages), returns queued_as_draft
 *   - send_mail in draft mode: respects the Drafts deny list
 *   - send_mail in send mode: delivers immediately (/me/sendMail), returns sent
 *   - get_email_output_mode: returns the calling user's stored mode
 *   - set_email_output_mode: persists 'draft' for the calling user
 *   - set_email_output_mode: refuses 'send' without enforcement, writes nothing,
 *     points at the web UI, and the dispatcher logs the refusal as denied
 *   - set_email_output_mode: rejects an invalid mode value
 *
 * Enforced draft mode:
 *   - set_email_output_mode: refuses when a tenant or user policy enforces draft,
 *     writes nothing, and the dispatcher logs the refusal as denied
 *   - get_email_output_mode: reports enforced / enforcedBy
 *   - send_draft in draft mode: refuses without telling the agent to flip the mode
 *   - send_mail honours the effective mode the service resolves (draft while enforced)
 */

import { jest } from '@jest/globals';
import type { HttpRequest, InvocationContext } from '@azure/functions';

// ── Mock declarations ────────────────────────────────────────────────────────

const mockAuthenticateRequest = jest.fn<() => Promise<unknown>>();
const mockGetValidAccessTokenForSession = jest.fn<() => Promise<string>>();
const mockGetTenantIdFromSession = jest.fn<() => string>();
const mockIsPathDenied = jest.fn<() => Promise<boolean>>();
const mockGetEnabledServices = jest.fn<() => Promise<string[]>>();
const mockGetAllowedSites = jest.fn<() => Promise<Array<{ id: string; name: string }>>>();
const mockGetUserServiceOverrides = jest.fn<() => Promise<string[]>>();
const mockGetUserEmailSettings = jest.fn<() => Promise<{ emailOutputMode: string; enforced?: boolean; enforcedBy?: 'tenant' | 'user' | null }>>();
const mockLogAccess = jest.fn<(entry: Record<string, unknown>) => void>();
const mockSetUserEmailSettings = jest.fn<(tenantId: string, userId: string, settings: { emailOutputMode: string }) => Promise<void>>();

const mockGraphPost = jest.fn<(payload: unknown) => Promise<unknown>>();
const mockGraphGet = jest.fn<() => Promise<unknown>>();
const mockGraphPatch = jest.fn<(payload: unknown) => Promise<unknown>>();
const mockGraphApi = jest.fn(() => ({ post: mockGraphPost, get: mockGraphGet, patch: mockGraphPatch, select: () => ({ get: mockGraphGet }) }));
const mockCreateGraphClient = jest.fn(() => ({ api: mockGraphApi }));

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
  getUserEmailSettings: () => mockGetUserEmailSettings(),
  setUserEmailSettings: (tenantId: unknown, userId: unknown, settings: unknown) =>
    mockSetUserEmailSettings(tenantId as string, userId as string, settings as { emailOutputMode: string }),
}));

jest.mock('../services/userMailConfig.js', () => ({
  isMailIndexingDisabled: jest.fn<(tenantId: string, userId: string) => Promise<boolean>>().mockResolvedValue(false),
}));

jest.mock('../services/auditLog.js', () => ({
  logAccess: (entry: unknown) => mockLogAccess(entry as Record<string, unknown>),
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

// ── Fixtures ──────────────────────────────────────────────────────────────────

const TENANT = 'test-tenant';
const USER = 'test-user';

const SESSION = {
  userId: USER,
  tenantId: TENANT,
  accessToken: 'fake-token',
  sessionToken: 'fake-session',
};
const AUTH = { userId: USER, session: SESSION };

function callTool(name: string, args: Record<string, unknown>): Promise<{ status: number; jsonBody?: unknown }> {
  const req = {
    method: 'POST',
    headers: new Map<string, string>(),
    json: () => Promise.resolve({ jsonrpc: '2.0', id: 1, method: 'tools/call', params: { name, arguments: args } }),
  } as unknown as HttpRequest;
  const ctx = { error: jest.fn(), warn: jest.fn(), log: jest.fn() } as unknown as InvocationContext;
  return handler(req, ctx);
}

/** Extract the tool result object from a JSON-RPC tools/call response. */
function toolResult(res: { jsonBody?: unknown }): { parsed: unknown; isError: boolean; text: string } {
  const body = res.jsonBody as { result?: { content?: Array<{ text: string }>; isError?: boolean } };
  const text = body.result?.content?.[0]?.text ?? '{}';
  const isError = body.result?.isError ?? false;
  // Success results are JSON-serialized; error results are a plain "Error: ..." string.
  return { parsed: isError ? null : JSON.parse(text), isError, text };
}

beforeEach(() => {
  jest.clearAllMocks();
  mockAuthenticateRequest.mockResolvedValue(AUTH);
  mockGetValidAccessTokenForSession.mockResolvedValue('access-token');
  mockGetTenantIdFromSession.mockReturnValue(TENANT);
  mockIsPathDenied.mockResolvedValue(false);
  mockGetEnabledServices.mockResolvedValue(['mail', 'sharepoint']);
  mockGetAllowedSites.mockResolvedValue([]);
  mockGetUserServiceOverrides.mockResolvedValue([]);
  mockSetUserEmailSettings.mockResolvedValue(undefined);
});

describe('MCP send_mail — draft mode (default)', () => {
  beforeEach(() => {
    mockGetUserEmailSettings.mockResolvedValue({ emailOutputMode: 'draft' });
    mockGraphPost.mockResolvedValue({ id: 'draft-9', webLink: 'https://outlook/draft-9' });
  });

  it('saves to Drafts (/me/messages) and returns queued_as_draft — does not send', async () => {
    const res = await callTool('send_mail', { subject: 'Hi', body: 'Body', to: ['a@x.com'] });
    const { parsed, isError } = toolResult(res);

    expect(isError).toBe(false);
    expect(parsed).toMatchObject({ status: 'queued_as_draft', draftId: 'draft-9', draftLink: 'https://outlook/draft-9' });
    expect(mockGraphApi).toHaveBeenCalledWith('/me/messages');
    expect(mockGraphApi).not.toHaveBeenCalledWith('/me/sendMail');
  });

  it('respects the Drafts deny list', async () => {
    mockIsPathDenied.mockResolvedValue(true);
    const res = await callTool('send_mail', { subject: 'Hi', body: 'Body', to: ['a@x.com'] });
    const { isError } = toolResult(res);

    expect(isError).toBe(true);
    expect(mockGraphPost).not.toHaveBeenCalled();
  });
});

describe('MCP send_mail — send mode', () => {
  beforeEach(() => {
    mockGetUserEmailSettings.mockResolvedValue({ emailOutputMode: 'send' });
    mockGraphPost.mockResolvedValue(undefined);
  });

  it('delivers immediately (/me/sendMail) and returns sent', async () => {
    const res = await callTool('send_mail', { subject: 'Hi', body: 'Body', to: ['a@x.com'] });
    const { parsed } = toolResult(res);

    expect(parsed).toMatchObject({ status: 'sent' });
    expect(mockGraphApi).toHaveBeenCalledWith('/me/sendMail');
    expect(mockGraphApi).not.toHaveBeenCalledWith('/me/messages');
  });
});

describe('MCP get_email_output_mode', () => {
  it("returns the calling user's stored mode", async () => {
    mockGetUserEmailSettings.mockResolvedValue({ emailOutputMode: 'send' });
    const res = await callTool('get_email_output_mode', {});
    const { parsed } = toolResult(res);

    expect(parsed).toEqual({ emailOutputMode: 'send', enforced: false, enforcedBy: null });
  });

  it('reports an enforced mode and which policy enforces it', async () => {
    mockGetUserEmailSettings.mockResolvedValue({ emailOutputMode: 'draft', enforced: true, enforcedBy: 'tenant' });
    const res = await callTool('get_email_output_mode', {});
    const { parsed } = toolResult(res);

    expect(parsed).toEqual({ emailOutputMode: 'draft', enforced: true, enforcedBy: 'tenant' });
  });
});

describe('MCP set_email_output_mode', () => {
  beforeEach(() => {
    mockGetUserEmailSettings.mockResolvedValue({ emailOutputMode: 'send', enforced: false, enforcedBy: null });
  });

  it("tightens the calling user's mode to draft and logs it as allowed", async () => {
    const res = await callTool('set_email_output_mode', { emailOutputMode: 'draft' });
    const { parsed } = toolResult(res);

    expect(parsed).toMatchObject({ status: 'updated', emailOutputMode: 'draft' });
    expect(mockSetUserEmailSettings).toHaveBeenCalledWith(TENANT, USER, { emailOutputMode: 'draft' });
    expect(mockLogAccess).toHaveBeenCalledWith(expect.objectContaining({ operation: 'set_email_output_mode', result: 'allowed' }));
  });

  it("refuses 'send' without any policy, writes nothing, points at the web UI, and logs denied", async () => {
    mockGetUserEmailSettings.mockResolvedValue({ emailOutputMode: 'draft', enforced: false, enforcedBy: null });

    const res = await callTool('set_email_output_mode', { emailOutputMode: 'send' });
    const { isError, text } = toolResult(res);

    expect(isError).toBe(true);
    expect(text).toContain('cannot switch to send mode');
    expect(text).toContain('web UI');
    // Not an administrator lock, so the message must not claim one.
    expect(text).not.toContain('enforced');
    expect(mockSetUserEmailSettings).not.toHaveBeenCalled();
    expect(mockLogAccess).toHaveBeenCalledWith(expect.objectContaining({
      tenantId: TENANT,
      userId: USER,
      operation: 'set_email_output_mode',
      result: 'denied',
      source: 'mcp',
      reason: expect.stringContaining('cannot switch to send mode'),
    }));
    expect(mockLogAccess).not.toHaveBeenCalledWith(expect.objectContaining({ result: 'allowed' }));
  });

  it("refuses 'send' even when the user is already in send mode", async () => {
    const res = await callTool('set_email_output_mode', { emailOutputMode: 'send' });
    const { isError } = toolResult(res);

    expect(isError).toBe(true);
    expect(mockSetUserEmailSettings).not.toHaveBeenCalled();
  });

  it('rejects an invalid mode value without writing', async () => {
    const res = await callTool('set_email_output_mode', { emailOutputMode: 'bogus' });
    const { isError } = toolResult(res);

    expect(isError).toBe(true);
    expect(mockSetUserEmailSettings).not.toHaveBeenCalled();
  });
});

describe('MCP set_email_output_mode — enforced draft mode', () => {
  it.each([
    ['tenant', 'for this tenant'],
    ['user', 'for your account'],
  ] as const)('refuses to switch to send under a %s policy, writes nothing, and logs denied', async (enforcedBy, scopeText) => {
    mockGetUserEmailSettings.mockResolvedValue({ emailOutputMode: 'draft', enforced: true, enforcedBy });

    const res = await callTool('set_email_output_mode', { emailOutputMode: 'send' });
    const { isError, text } = toolResult(res);

    expect(isError).toBe(true);
    expect(text).toContain('enforced to draft by your administrator');
    expect(text).toContain(scopeText);
    // The refusal must not coach the caller into a workaround.
    expect(text).not.toContain("set_email_output_mode('send')");
    expect(mockSetUserEmailSettings).not.toHaveBeenCalled();
    expect(mockLogAccess).toHaveBeenCalledWith(expect.objectContaining({
      tenantId: TENANT,
      userId: USER,
      operation: 'set_email_output_mode',
      result: 'denied',
      source: 'mcp',
      reason: expect.stringContaining('enforced to draft'),
    }));
  });

  it("refuses even a 'draft' write while enforced — the tool has no say in the mode", async () => {
    mockGetUserEmailSettings.mockResolvedValue({ emailOutputMode: 'draft', enforced: true, enforcedBy: 'user' });

    const res = await callTool('set_email_output_mode', { emailOutputMode: 'draft' });
    const { isError } = toolResult(res);

    expect(isError).toBe(true);
    expect(mockSetUserEmailSettings).not.toHaveBeenCalled();
  });

  it('still validates the value before consulting the policy', async () => {
    const res = await callTool('set_email_output_mode', { emailOutputMode: 'bogus' });
    const { isError } = toolResult(res);

    expect(isError).toBe(true);
    expect(mockGetUserEmailSettings).not.toHaveBeenCalled();
    expect(mockSetUserEmailSettings).not.toHaveBeenCalled();
  });

});

describe('MCP send paths — enforced draft mode', () => {
  it('send_mail saves to Drafts when the service resolves the effective mode to draft', async () => {
    // The service already folds the policy into emailOutputMode; send_mail only reads that field.
    mockGetUserEmailSettings.mockResolvedValue({ emailOutputMode: 'draft', enforced: true, enforcedBy: 'tenant' });
    mockGraphPost.mockResolvedValue({ id: 'draft-1', webLink: 'https://outlook/draft-1' });

    const res = await callTool('send_mail', { subject: 'Hi', body: 'Body', to: ['a@x.com'] });
    const { parsed, isError } = toolResult(res);

    expect(isError).toBe(false);
    expect(parsed).toMatchObject({ status: 'queued_as_draft', draftId: 'draft-1' });
    expect(mockGraphApi).not.toHaveBeenCalledWith('/me/sendMail');
  });

  it('send_draft refuses under enforcement and does not suggest switching the mode', async () => {
    mockGetUserEmailSettings.mockResolvedValue({ emailOutputMode: 'draft', enforced: true, enforcedBy: 'tenant' });
    mockGraphGet.mockResolvedValue({ isDraft: true, parentFolderId: null, subject: 'Hi', webLink: 'https://outlook/draft-2' });

    const res = await callTool('send_draft', { messageId: 'draft-2' });
    const { isError, text } = toolResult(res);

    expect(isError).toBe(true);
    expect(text).toContain('enforced by your administrator');
    expect(text).not.toContain("set_email_output_mode('send')");
    expect(mockGraphApi).not.toHaveBeenCalledWith('/me/messages/draft-2/send');
  });

  it('send_draft points an unenforced draft-mode user at the web UI, not set_email_output_mode', async () => {
    mockGetUserEmailSettings.mockResolvedValue({ emailOutputMode: 'draft', enforced: false, enforcedBy: null });
    mockGraphGet.mockResolvedValue({ isDraft: true, parentFolderId: null, subject: 'Hi', webLink: 'https://outlook/draft-3' });

    const res = await callTool('send_draft', { messageId: 'draft-3' });
    const { isError, text } = toolResult(res);

    expect(isError).toBe(true);
    expect(text).toContain('web UI');
    expect(text).not.toContain('set_email_output_mode');
  });
});
