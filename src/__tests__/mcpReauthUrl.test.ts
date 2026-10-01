/**
 * Regression test for the MCP re-authentication link (, codex review on
 * PR #150).
 *
 * The finding: mcpEndpoint built its login URL from FRONTEND_URL and stripped
 * only an exact trailing `/admin`, so a value like `https://host/admin/settings`
 * or `https://host/runtime` produced
 * `https://host/runtime/api/auth/login` — a path the Azure Functions host
 * intercepts before any of our code runs. The callback redirect and the deploy
 * smoke both knew such a value was bad; this entrypoint did not.
 *
 * These tests drive the real JSON-RPC dispatch with no session, which is the
 * exact code path that emits the message, and assert on the emitted text
 * rather than on a helper in isolation.
 */
import { jest } from '@jest/globals';
import type { HttpRequest, InvocationContext } from '@azure/functions';

const mockAuthenticateRequest = jest.fn<() => Promise<unknown>>();

jest.mock('../services/telemetry.js', () => ({}));
jest.mock('../services/authMiddleware.js', () => ({
  authenticateRequest: () => mockAuthenticateRequest(),
}));
jest.mock('../services/tokenCache.js', () => ({
  getValidAccessTokenForSession: () => Promise.resolve('access-token'),
  getTenantIdFromSession: () => 't',
}));
jest.mock('../services/graphClient.js', () => ({
  createGraphClient: () => ({ api: () => ({}) }),
}));
jest.mock('../services/denyList.js', () => ({
  filterDeniedPaths: (_t: unknown, _u: unknown, _s: unknown, items: unknown) => Promise.resolve(items),
  filterDeniedSearchHits: (_t: unknown, _u: unknown, items: unknown) => Promise.resolve(items),
  isPathDenied: () => Promise.resolve(false),
  canonicalizePath: (p: string) => (p.startsWith('/') ? p : `/${p}`).replace(/\/+$/, '').toLowerCase(),
}));
jest.mock('../services/calendarAccess.js', () => ({
  checkCalendarAccess: () => Promise.resolve(null),
}));
jest.mock('../services/containerResolver.js', () => ({
  resolveMailFolderName: () => Promise.resolve(null),
  resolveContactParentFolder: () => Promise.resolve(null),
  resolveDefaultContactFolder: () => Promise.resolve(null),
  resolveSectionNotebook: () => Promise.resolve(null),
  resolveDefaultCalendarId: () => Promise.resolve('cal'),
  resolveCalendarName: () => Promise.resolve('Calendar'),
  normalizeContactFolderId: (_g: unknown, _u: unknown, id: string) => Promise.resolve(id),
}));
jest.mock('../services/sharepointFilter.js', () => ({
  filterAndDisambiguateSites: (sites: unknown) => sites,
}));
jest.mock('../services/serviceSettings.js', () => ({
  getEnabledServices: () => Promise.resolve(['mail']),
  getReadOnlyServices: () => Promise.resolve([]),
  getAllowedSites: () => Promise.resolve([]),
}));
jest.mock('../services/userServiceOverrides.js', () => ({
  getUserServiceOverrides: () => Promise.resolve([]),
}));
jest.mock('../services/userEmailSettings.js', () => ({
  getUserEmailSettings: () => Promise.resolve({ emailOutputMode: 'draft' }),
  setUserEmailSettings: () => Promise.resolve(undefined),
}));
jest.mock('../services/userMailConfig.js', () => ({
  isMailIndexingDisabled: () => Promise.resolve(false),
}));
jest.mock('../services/mailboxTimeZone.js', () => ({
  resolveMailboxTimeZone: () => Promise.resolve('Pacific Standard Time'),
}));
jest.mock('../services/auditLog.js', () => ({ logAccess: jest.fn() }));
jest.mock('@azure/functions', () => ({ app: { http: jest.fn() } }));

import { app } from '@azure/functions';
import { reservedPrefixFor } from '../services/frontendUrl.js';
import '../functions/mcp/mcpEndpoint.js';

interface HttpRegistration {
  handler: (req: HttpRequest, context: InvocationContext) => Promise<{ status: number; jsonBody?: unknown }>;
}
const httpMock = app.http as unknown as jest.Mock<(name: string, opts: HttpRegistration) => void>;
const handler = httpMock.mock.calls.find((call) => call[0] === 'mcpEndpoint')![1].handler;

function rpc(method: string, params: Record<string, unknown>) {
  const req = {
    method: 'POST',
    headers: new Map<string, string>(),
    json: () => Promise.resolve({ jsonrpc: '2.0', id: 1, method, params }),
  } as unknown as HttpRequest;
  const ctx = { error: jest.fn(), warn: jest.fn(), log: jest.fn() } as unknown as InvocationContext;
  return handler(req, ctx);
}

/** Text of the `Session expired` result the unauthenticated path returns. */
async function reauthText(toolName: string): Promise<string> {
  const res = await rpc('tools/call', { name: toolName, arguments: {} });
  const body = res.jsonBody as { result?: { content?: Array<{ text: string }>; isError?: boolean } };
  expect(body.result?.isError).toBe(true);
  return body.result!.content![0].text;
}

const ORIGINAL_FRONTEND_URL = process.env.FRONTEND_URL;
const ORIGINAL_REDIRECT_URI = process.env.OAUTH_REDIRECT_URI;

let anyToolName: string;

beforeAll(async () => {
  // tools/list needs no session; take a real advertised name so the dispatch
  // reaches the `!auth` branch rather than the unknown-tool branch.
  mockAuthenticateRequest.mockResolvedValue(null);
  const res = await rpc('tools/list', {});
  const tools = (res.jsonBody as { result?: { tools?: Array<{ name: string }> } }).result?.tools ?? [];
  expect(tools.length).toBeGreaterThan(0);
  anyToolName = tools[0].name;
});

beforeEach(() => {
  mockAuthenticateRequest.mockReset();
  mockAuthenticateRequest.mockResolvedValue(null);
  jest.spyOn(console, 'error').mockImplementation(() => {});
});

afterEach(() => {
  jest.restoreAllMocks();
  if (ORIGINAL_FRONTEND_URL === undefined) delete process.env.FRONTEND_URL;
  else process.env.FRONTEND_URL = ORIGINAL_FRONTEND_URL;
  if (ORIGINAL_REDIRECT_URI === undefined) delete process.env.OAUTH_REDIRECT_URI;
  else process.env.OAUTH_REDIRECT_URI = ORIGINAL_REDIRECT_URI;
});

describe('MCP re-authentication link', () => {
  it.each([
    ['https://mcp.example.com/runtime'],
    ['https://mcp.example.com/runtime/webhooks'],
    ['https://mcp.example.com/admin'],
    ['https://mcp.example.com/admin/'],
    ['https://mcp.example.com/admin/settings'],
  ])('never points at a host-reserved path when FRONTEND_URL is %s', async (value) => {
    process.env.FRONTEND_URL = value;
    delete process.env.OAUTH_REDIRECT_URI;

    const text = await reauthText(anyToolName);

    expect(text).toContain('https://mcp.example.com/api/auth/login');
    expect(text).not.toContain('/runtime');
    expect(text).not.toContain('/admin');

    const url = text.split('re-authenticate at: ')[1].split(/\s/)[0];
    expect(reservedPrefixFor(url)).toBeNull();
  });

  it('keeps working for a correctly configured FRONTEND_URL', async () => {
    process.env.FRONTEND_URL = 'https://mcp.example.com';
    const text = await reauthText(anyToolName);
    expect(text).toContain('https://mcp.example.com/api/auth/login');
  });

  it('falls back to the OAUTH_REDIRECT_URI origin when FRONTEND_URL is unset', async () => {
    delete process.env.FRONTEND_URL;
    process.env.OAUTH_REDIRECT_URI = 'https://mcp.example.com/api/auth/callback';
    const text = await reauthText(anyToolName);
    expect(text).toContain('https://mcp.example.com/api/auth/login');
  });

  it('reads the value in force at message time, not at module load', async () => {
    process.env.FRONTEND_URL = 'https://first.example.com/admin';
    expect(await reauthText(anyToolName)).toContain('https://first.example.com/api/auth/login');

    process.env.FRONTEND_URL = 'https://second.example.com/runtime';
    expect(await reauthText(anyToolName)).toContain('https://second.example.com/api/auth/login');
  });
});
