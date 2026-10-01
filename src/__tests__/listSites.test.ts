/**
 * Regression tests for the listSites admin-bypass branch.
 *
 * The admin UI dropdown needs to show ALL sites a Global Admin can configure
 * for the allowlist — including sites not returned by Graph's /sites?search=*
 * (which only returns search-indexed sites). For non-admin callers, the strict
 * allowedSites filter must still apply.
 *
 * These tests verify that:
 *   1. Non-admin caller with allowedSites configured → sees only allowed sites
 *   2. Non-admin caller with empty allowedSites → sees all Graph results
 *   3. Admin caller with admin=true → sees union of Graph + allowedSites
 *   4. Caller passing admin=true but failing checkGlobalAdmin → no bypass
 *   5. Union dedupes by site ID (no duplicate entries)
 */

import { jest } from '@jest/globals';
import type { HttpRequest, InvocationContext } from '@azure/functions';
import type { AuthResult } from '../services/authMiddleware.js';

// ── Mock declarations ────────────────────────────────────────────────────────

interface GraphSiteResponse {
  value: Array<{
    id: string;
    name: string;
    displayName: string;
    webUrl: string;
    description: string | null;
    createdDateTime: string;
    lastModifiedDateTime: string;
  }>;
}

const mockAuthenticateRequest = jest.fn<(req: HttpRequest) => Promise<AuthResult | null>>();
const mockCheckGlobalAdmin = jest.fn<(userId: string, token?: string) => Promise<boolean>>();
const mockGetEnabledServices = jest.fn<(tenantId: string) => Promise<string[]>>();
const mockGetAllowedSites = jest.fn<(tenantId: string) => Promise<Array<{ id: string; name: string }>>>();
const mockIsPathDenied = jest.fn<() => Promise<boolean>>();
const mockGetTenantId = jest.fn<(userId: string) => Promise<string>>();
const mockGetValidAccessToken = jest.fn<(userId: string) => Promise<string>>();

// Graph client mock — Graph .api(...).select(...).get() returns sites.
const mockGraphGet = jest.fn<() => Promise<GraphSiteResponse>>();
const mockCreateGraphClient = jest.fn(() => ({
  api: () => ({
    select: () => ({
      get: mockGraphGet,
    }),
  }),
}));

// ── Wire up mocks before importing the module under test ─────────────────────

jest.mock('../services/authMiddleware.js', () => ({
  authenticateRequest: (req: unknown) => mockAuthenticateRequest(req as HttpRequest),
  checkGlobalAdmin: (userId: unknown, token?: unknown) =>
    mockCheckGlobalAdmin(userId as string, token as string | undefined),
}));

jest.mock('../services/serviceSettings.js', () => ({
  getEnabledServices: (...args: unknown[]) => mockGetEnabledServices(args[0] as string),
  getReadOnlyServices: () => Promise.resolve([]),
  getAllowedSites: (...args: unknown[]) => mockGetAllowedSites(args[0] as string),
}));

jest.mock('../services/userServiceOverrides.js', () => ({
  isServiceDisabledForUser: () => Promise.resolve(false),
}));

jest.mock('../services/denyList.js', () => ({
  isPathDenied: () => mockIsPathDenied(),
}));

jest.mock('../services/tokenCache.js', () => ({
  getTenantId: (...args: unknown[]) => mockGetTenantId(args[0] as string),
  getTenantIdFromSession: (session: { tenantId?: string }) => {
    if (!session?.tenantId) throw new Error('No tenantId in session');
    return session.tenantId;
  },
  getValidAccessToken: (...args: unknown[]) => mockGetValidAccessToken(args[0] as string),
  getValidAccessTokenForSession: (session: { userId?: string }) =>
    mockGetValidAccessToken(session?.userId as string),
}));

jest.mock('../services/graphClient.js', () => ({
  createGraphClient: (...args: unknown[]) => (mockCreateGraphClient as (...a: unknown[]) => unknown)(...args),
}));

jest.mock('../services/sharepointFilter.js', () => ({
  filterAndDisambiguateSites: (sites: unknown[]) => sites,
}));

// Stub out Azure Functions app.http so the import doesn't try to register a real handler.
jest.mock('@azure/functions', () => ({
  app: { http: jest.fn() },
}));

// ── Import the module under test (after mocks) ──────────────────────────────

// We need to import the file to get the registered handler. Since listSites
// registers via app.http, we can extract the handler from the mock call args.
import { app } from '@azure/functions';
import '../functions/sharepoint/listSites.js';

interface HttpRegistration {
  handler: (req: HttpRequest, context: InvocationContext) => Promise<{ status: number; jsonBody?: unknown }>;
}
const httpMock = app.http as unknown as jest.Mock<(name: string, opts: HttpRegistration) => void>;
const registration = httpMock.mock.calls.find((call) => call[0] === 'listSites');
if (!registration) throw new Error('listSites handler was not registered');
const wrappedHandler = registration[1].handler;

// ── Test fixtures ────────────────────────────────────────────────────────────

const TENANT = 'tenant-xyz';
const USER = 'user-abc';
const TOKEN = 'fake-graph-token';

const SITE_FROM_GRAPH = {
  id: 'site-1',
  name: 'Team Site',
  displayName: 'Team Site',
  webUrl: 'https://contoso.sharepoint.com',
  description: null,
  createdDateTime: '2024-01-01T00:00:00Z',
  lastModifiedDateTime: '2024-01-01T00:00:00Z',
};

const SITE_ONLY_IN_ALLOWLIST = {
  id: 'site-2',
  name: 'IT Hub',
};

function makeRequest(adminFlag?: boolean): HttpRequest {
  const params = new Map<string, string>();
  if (adminFlag) params.set('admin', 'true');
  return {
    query: { get: (k: string) => params.get(k) ?? null },
    headers: new Map<string, string>(),
  } as unknown as HttpRequest;
}

const fakeContext = {} as InvocationContext;

// ── Test setup ──────────────────────────────────────────────────────────────

beforeEach(() => {
  jest.clearAllMocks();
  mockAuthenticateRequest.mockResolvedValue({
    userId: USER,
    session: {
      userId: USER,
      homeAccountId: 'home-abc',
      displayName: 'Test',
      email: 'test@example.com',
      tenantId: TENANT,
      accessToken: TOKEN,
      expiresAt: Date.now() + 3_600_000,
      sessionToken: 'fake-session',
      sessionCreatedAt: Date.now(),
    },
  } as AuthResult);
  mockGetEnabledServices.mockResolvedValue(['mail', 'sharepoint']);
  mockIsPathDenied.mockResolvedValue(false);
  mockGetTenantId.mockResolvedValue(TENANT);
  mockGetValidAccessToken.mockResolvedValue(TOKEN);
  mockGraphGet.mockResolvedValue({ value: [SITE_FROM_GRAPH] });
});

// ── Tests ───────────────────────────────────────────────────────────────────

describe('listSites — non-admin path (allowedSites filter)', () => {
  it('returns only sites in allowedSites when allowedSites is non-empty', async () => {
    mockGetAllowedSites.mockResolvedValue([{ id: SITE_FROM_GRAPH.id, name: SITE_FROM_GRAPH.name }]);

    const res = await wrappedHandler(makeRequest(false), fakeContext);

    expect(res.status).toBe(200);
    const body = res.jsonBody as { sites: Array<{ id: string }>; count: number };
    expect(body.count).toBe(1);
    expect(body.sites.map((s) => s.id)).toEqual([SITE_FROM_GRAPH.id]);
    expect(mockCheckGlobalAdmin).not.toHaveBeenCalled();
  });

  it('filters out Graph sites that are not in allowedSites', async () => {
    const otherGraphSite = { ...SITE_FROM_GRAPH, id: 'site-not-allowed', displayName: 'Random Site' };
    mockGraphGet.mockResolvedValue({ value: [SITE_FROM_GRAPH, otherGraphSite] });
    mockGetAllowedSites.mockResolvedValue([{ id: SITE_FROM_GRAPH.id, name: SITE_FROM_GRAPH.name }]);

    const res = await wrappedHandler(makeRequest(false), fakeContext);

    const body = res.jsonBody as { sites: Array<{ id: string }> };
    expect(body.sites.map((s) => s.id)).toEqual([SITE_FROM_GRAPH.id]);
    expect(body.sites.find((s) => s.id === 'site-not-allowed')).toBeUndefined();
  });

  it('returns all Graph results when allowedSites is empty', async () => {
    mockGetAllowedSites.mockResolvedValue([]);

    const res = await wrappedHandler(makeRequest(false), fakeContext);

    const body = res.jsonBody as { sites: Array<{ id: string }>; count: number };
    expect(body.count).toBe(1);
    expect(body.sites.map((s) => s.id)).toEqual([SITE_FROM_GRAPH.id]);
  });
});

describe('listSites — admin bypass (admin=true + Global Admin)', () => {
  it('returns Graph results UNION allowedSites entries when caller is Global Admin', async () => {
    mockCheckGlobalAdmin.mockResolvedValue(true);
    mockGetAllowedSites.mockResolvedValue([
      { id: SITE_FROM_GRAPH.id, name: SITE_FROM_GRAPH.name }, // already in Graph
      SITE_ONLY_IN_ALLOWLIST, // only in allowlist, not in Graph
    ]);

    const res = await wrappedHandler(makeRequest(true), fakeContext);

    expect(res.status).toBe(200);
    const body = res.jsonBody as { sites: Array<{ id: string; displayName: string }>; count: number };
    expect(body.count).toBe(2);
    const ids = body.sites.map((s) => s.id);
    expect(ids).toContain(SITE_FROM_GRAPH.id);
    expect(ids).toContain(SITE_ONLY_IN_ALLOWLIST.id);
  });

  it('dedupes by site ID — sites in both Graph and allowedSites appear once', async () => {
    mockCheckGlobalAdmin.mockResolvedValue(true);
    mockGetAllowedSites.mockResolvedValue([
      { id: SITE_FROM_GRAPH.id, name: SITE_FROM_GRAPH.name }, // duplicate of Graph result
    ]);

    const res = await wrappedHandler(makeRequest(true), fakeContext);

    const body = res.jsonBody as { sites: Array<{ id: string }>; count: number };
    expect(body.count).toBe(1);
    expect(body.sites.filter((s) => s.id === SITE_FROM_GRAPH.id)).toHaveLength(1);
  });

  it('returns just Graph results when Global Admin has empty allowedSites', async () => {
    mockCheckGlobalAdmin.mockResolvedValue(true);
    mockGetAllowedSites.mockResolvedValue([]);

    const res = await wrappedHandler(makeRequest(true), fakeContext);

    const body = res.jsonBody as { sites: Array<{ id: string }>; count: number };
    expect(body.count).toBe(1);
    expect(body.sites.map((s) => s.id)).toEqual([SITE_FROM_GRAPH.id]);
  });
});

describe('listSites — admin bypass gate (security-critical)', () => {
  it('does NOT bypass filter when admin=true but checkGlobalAdmin returns false', async () => {
    mockCheckGlobalAdmin.mockResolvedValue(false);
    mockGetAllowedSites.mockResolvedValue([SITE_ONLY_IN_ALLOWLIST]);

    const res = await wrappedHandler(makeRequest(true), fakeContext);

    // Without Global Admin, the user falls through to the non-admin branch:
    // Graph returns [SITE_FROM_GRAPH], allowedSites contains only [SITE_ONLY_IN_ALLOWLIST],
    // intersection is empty.
    const body = res.jsonBody as { sites: Array<{ id: string }>; count: number };
    expect(body.count).toBe(0);
    expect(body.sites).toEqual([]);
    expect(mockCheckGlobalAdmin).toHaveBeenCalledWith(USER, TOKEN);
  });

  it('does NOT call checkGlobalAdmin when admin flag is absent (no wasted Graph call)', async () => {
    mockGetAllowedSites.mockResolvedValue([{ id: SITE_FROM_GRAPH.id, name: SITE_FROM_GRAPH.name }]);

    await wrappedHandler(makeRequest(false), fakeContext);

    expect(mockCheckGlobalAdmin).not.toHaveBeenCalled();
  });

  it('does NOT bypass filter when admin=true but caller fails Graph admin lookup (fail closed)', async () => {
    // checkGlobalAdmin returns false on any error (fails closed) — verify the
    // wrapper treats that the same as "not an admin".
    mockCheckGlobalAdmin.mockResolvedValue(false);
    mockGetAllowedSites.mockResolvedValue([SITE_ONLY_IN_ALLOWLIST]);

    const res = await wrappedHandler(makeRequest(true), fakeContext);

    const body = res.jsonBody as { sites: Array<{ id: string }> };
    // Sites NOT in Graph's response and NOT in the allowedSites intersection are excluded.
    expect(body.sites.find((s) => s.id === SITE_ONLY_IN_ALLOWLIST.id)).toBeUndefined();
  });
});
