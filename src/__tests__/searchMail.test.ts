/**
 * Tests for the searchMail HTTP handler.
 *
 * Verifies:
 *   1. Folder-scoped search always uses $filter directly (never $search) — prevents silent Archive degradation
 *   2. Missing query → 400
 *   3. Deny-listed folder → 403
 *   4. Global (no folderId) search uses $search
 *   5. Global search falls back to $filter when $search throws
 *   6. Both $search and $filter fail (global) → 502 with combined error
 *   7. Folder-scoped $filter failure → 502 with $filter error only
 *   8. $filter path respects $top
 *   9. OData injection in query is escaped in $filter
 *  10. $orderby is NOT sent to Graph in $filter path (avoids InefficientFilter)
 *  11. $filter results are sorted client-side by receivedDateTime desc
 *  12. Regression: degraded Archive $search returning ≤ maxResults unrelated messages never passes through
 *  13.: `to` / `participant` in a folder run a newest-first scan that matches recipients
 *  14.: a missing-criteria call is 400, a bad `since` is 400, and every 200 carries strategy metadata
 */

import { jest } from '@jest/globals';
import type { HttpRequest, InvocationContext } from '@azure/functions';
import type { AuthResult } from '../services/authMiddleware.js';

// ── Mock declarations ─────────────────────────────────────────────────────────

const mockAuthenticateRequest = jest.fn<(req: HttpRequest) => Promise<AuthResult | null>>();
const mockGetEnabledServices = jest.fn<(tenantId: string) => Promise<string[]>>();
const mockIsServiceDisabledForUser = jest.fn<() => Promise<boolean>>();
const mockIsPathDenied = jest.fn<() => Promise<boolean>>();
const mockGetTenantId = jest.fn<(userId: string) => Promise<string>>();
const mockGetValidAccessToken = jest.fn<(userId: string) => Promise<string>>();
const mockResolveMailFolderName = jest.fn<() => Promise<string | null>>();

// The graph builder chains: api(path).search/filter/select/top/get()
// We track what was actually called so tests can assert the fallback path was taken.
interface GraphCallRecord {
  path: string | null;
  searchQuery: string | null;
  filterExpr: string | null;
  topValue: number | null;
  orderbyValue: string | null;
}
const lastGraphCall: GraphCallRecord = { path: null, searchQuery: null, filterExpr: null, topValue: null, orderbyValue: null };

const mockGraphGet = jest.fn<() => Promise<{ value: unknown[] }>>();

const mockCreateGraphClient = jest.fn(() => ({
  api: (path: string) => {
    lastGraphCall.path = path;
    lastGraphCall.searchQuery = null;
    lastGraphCall.filterExpr = null;
    lastGraphCall.topValue = null;
    lastGraphCall.orderbyValue = null;
    const builder = {
      search: (q: string) => { lastGraphCall.searchQuery = q; return builder; },
      filter: (f: string) => { lastGraphCall.filterExpr = f; return builder; },
      select: (_: string) => builder,
      top: (n: number) => { lastGraphCall.topValue = n; return builder; },
      orderby: (o: string) => { lastGraphCall.orderbyValue = o; return builder; },
      get: () => mockGraphGet(),
    };
    return builder;
  },
}));

jest.mock('../services/authMiddleware.js', () => ({
  authenticateRequest: (req: unknown) => mockAuthenticateRequest(req as HttpRequest),
}));

jest.mock('../services/serviceSettings.js', () => ({
  getEnabledServices: (...args: unknown[]) => mockGetEnabledServices(args[0] as string),
  getReadOnlyServices: () => Promise.resolve([]),
  getAllowedSites: () => Promise.resolve([]),
}));

jest.mock('../services/userServiceOverrides.js', () => ({
  isServiceDisabledForUser: () => mockIsServiceDisabledForUser(),
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

jest.mock('../services/containerResolver.js', () => ({
  resolveMailFolderName: () => mockResolveMailFolderName(),
}));

jest.mock('../services/userMailConfig.js', () => ({
  isMailIndexingDisabled: jest.fn<(tenantId: string, userId: string) => Promise<boolean>>().mockResolvedValue(false),
}));

jest.mock('../services/auditLog.js', () => ({
  logAccess: jest.fn(),
}));

jest.mock('@azure/functions', () => ({
  app: { http: jest.fn() },
}));

import { app } from '@azure/functions';
import '../functions/mail/searchMail.js';

interface HttpRegistration {
  handler: (req: HttpRequest, context: InvocationContext) => Promise<{ status: number; jsonBody?: unknown }>;
}
const httpMock = app.http as unknown as jest.Mock<(name: string, opts: HttpRegistration) => void>;
const registration = httpMock.mock.calls.find((call) => call[0] === 'searchMail');
if (!registration) throw new Error('searchMail handler was not registered');
const wrappedHandler = registration[1].handler;

// ── Helpers ───────────────────────────────────────────────────────────────────

const TENANT = 'tenant-xyz';
const USER = 'user-abc';
const TOKEN = 'fake-graph-token';
const FOLDER_ID = 'archive-folder-id';

function makeRequest(params: Record<string, string>): HttpRequest {
  return {
    query: {
      get: (key: string) => params[key] ?? null,
    },
    headers: new Map<string, string>(),
  } as unknown as HttpRequest;
}

const fakeContext = { error: jest.fn() } as unknown as InvocationContext;

function makeMessage(id: string, subject: string) {
  return {
    id,
    subject,
    from: { emailAddress: { name: 'Sender', address: 'sender@example.com' } },
    toRecipients: [],
    receivedDateTime: '2024-01-01T00:00:00Z',
    bodyPreview: 'preview text',
    hasAttachments: false,
    parentFolderId: FOLDER_ID,
    isRead: false,
  };
}

beforeEach(() => {
  jest.clearAllMocks();
  lastGraphCall.path = null;
  lastGraphCall.searchQuery = null;
  lastGraphCall.filterExpr = null;
  lastGraphCall.topValue = null;
  lastGraphCall.orderbyValue = null;

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
  mockGetEnabledServices.mockResolvedValue(['mail']);
  mockIsServiceDisabledForUser.mockResolvedValue(false);
  mockIsPathDenied.mockResolvedValue(false);
  mockGetTenantId.mockResolvedValue(TENANT);
  mockGetValidAccessToken.mockResolvedValue(TOKEN);
  mockResolveMailFolderName.mockResolvedValue(null);
});

// ── Tests ─────────────────────────────────────────────────────────────────────

describe('searchMail', () => {
  it('uses $filter directly for folder-scoped searches — never calls $search', async () => {
    // This is the core fix: folder-scoped $search silently degrades on large folders
    // (e.g. Archive) on some Exchange Online tenants. We always use $filter for folder-scoped
    // queries to avoid this entirely.
    const messages = [makeMessage('msg-1', 'Amazon order'), makeMessage('msg-2', 'Amazon invoice')];
    mockGraphGet.mockResolvedValueOnce({ value: messages });

    const res = await wrappedHandler(
      makeRequest({ q: 'amazon', folderId: FOLDER_ID, maxResults: '5' }),
      fakeContext,
    );

    expect(res.status).toBe(200);
    // Must have used $filter, not $search
    expect(lastGraphCall.searchQuery).toBeNull();
    expect(lastGraphCall.filterExpr).toContain("contains(subject,'amazon')");
    expect((res.jsonBody as { count: number }).count).toBe(2);
  });

  it('regression: folder-scoped search cannot pass through degraded $search results of any size', async () => {
    // The old count-based detection (count > maxResults) was fragile: a degraded Archive page
    // with <= maxResults items (e.g. the observed 9-email fixed set with maxResults=25) would
    // pass through as valid results. With the new approach, $search is never called for
    // folder-scoped queries, so this scenario is impossible by construction.
    const filterMessages = [makeMessage('msg-amazon', 'Your Amazon order')];
    mockGraphGet.mockResolvedValueOnce({ value: filterMessages });

    const res = await wrappedHandler(
      // maxResults=25 (default); old detection would NOT have caught a 9-item degraded page
      makeRequest({ q: 'amazon', folderId: FOLDER_ID }),
      fakeContext,
    );

    expect(res.status).toBe(200);
    // Exactly one $filter call, zero $search calls
    expect(mockGraphGet).toHaveBeenCalledTimes(1);
    expect(lastGraphCall.searchQuery).toBeNull();
    expect(lastGraphCall.filterExpr).toContain("contains(subject,'amazon')");
    expect((res.jsonBody as { count: number }).count).toBe(1);
  });

  it('returns 400 when no search criterion is given', async () => {
    const res = await wrappedHandler(makeRequest({ folderId: FOLDER_ID }), fakeContext);
    expect(res.status).toBe(400);
    expect((res.jsonBody as { error: string }).error).toMatch(/At least one of q, participant, from, to/);
    expect(mockGraphGet).not.toHaveBeenCalled();
  });

  it('returns 400 for a malformed since without touching Graph', async () => {
    const res = await wrappedHandler(makeRequest({ q: 'amazon', since: 'yesterday' }), fakeContext);
    expect(res.status).toBe(400);
    expect((res.jsonBody as { error: string }).error).toMatch(/ISO-8601/);
    expect(mockGraphGet).not.toHaveBeenCalled();
  });

  it(': `to` on a folder-scoped search scans newest-first and matches the recipient address', async () => {
    const hit = { ...makeMessage('msg-hit', 'Re: intro'), toRecipients: [{ emailAddress: { name: 'Jordan', address: 'jdoe@fabrikam.com' } }] };
    const miss = { ...makeMessage('msg-miss', 'Re: intro'), toRecipients: [{ emailAddress: { name: 'Other', address: 'other@example.com' } }] };
    mockGraphGet.mockResolvedValueOnce({ value: [miss, hit] });

    const res = await wrappedHandler(
      makeRequest({ to: 'fabrikam.com', folderId: FOLDER_ID, since: '2023-06-01' }),
      fakeContext,
    );

    expect(res.status).toBe(200);
    expect(lastGraphCall.searchQuery).toBeNull(); // never $search in a folder
    expect(lastGraphCall.orderbyValue).toBe('receivedDateTime desc');
    expect(lastGraphCall.filterExpr).toBe('receivedDateTime ge 2023-06-01T00:00:00.000Z');
    const body = res.jsonBody as { results: Array<{ id: string }>; count: number; strategy: string; scanComplete: boolean; truncated: boolean };
    expect(body.results.map((r) => r.id)).toEqual(['msg-hit']);
    expect(body.count).toBe(1);
    expect(body.strategy).toBe('scan');
    expect(body.scanComplete).toBe(true);
    expect(body.truncated).toBe(false);
  });

  it(': a text-only folder search reports that recipients were not searched', async () => {
    mockGraphGet.mockResolvedValueOnce({ value: [] });

    const res = await wrappedHandler(makeRequest({ q: 'fabrikam', folderId: FOLDER_ID }), fakeContext);

    expect(res.status).toBe(200);
    const body = res.jsonBody as { count: number; strategy: string; searchedFields: string[]; notes: string[] };
    expect(body.count).toBe(0);
    expect(body.strategy).toBe('filter');
    expect(body.searchedFields).toEqual(['subject', 'from.address', 'from.name']);
    expect(body.notes[0]).toMatch(/Recipient addresses/);
  });

  it(': a mailbox-wide search reports relevance ordering and honours `participant` in KQL', async () => {
    mockGraphGet.mockResolvedValueOnce({ value: [makeMessage('m1', 'Intro')] });

    const res = await wrappedHandler(makeRequest({ participant: 'fabrikam.com' }), fakeContext);

    expect(res.status).toBe(200);
    expect(lastGraphCall.searchQuery).toBe('participants:"fabrikam.com"');
    const body = res.jsonBody as { strategy: string; ordering: string; limit: number; notes: string[] };
    expect(body.strategy).toBe('kql-search');
    expect(body.ordering).toBe('relevance');
    expect(body.limit).toBe(25);
    expect(body.notes[0]).toMatch(/relevance-ranked/);
  });

  it('returns 403 when folder is deny-listed', async () => {
    mockResolveMailFolderName.mockResolvedValue('Sensitive');
    mockIsPathDenied.mockResolvedValue(true);

    const res = await wrappedHandler(
      makeRequest({ q: 'amazon', folderId: FOLDER_ID }),
      fakeContext,
    );

    expect(res.status).toBe(403);
    expect(mockGraphGet).not.toHaveBeenCalled();
  });

  it('uses $search for global (non-folder-scoped) searches', async () => {
    const messages = [makeMessage('msg-1', 'Amazon order')];
    mockGraphGet.mockResolvedValueOnce({ value: messages });

    const res = await wrappedHandler(
      makeRequest({ q: 'amazon', maxResults: '5' }),
      fakeContext,
    );

    expect(res.status).toBe(200);
    expect(lastGraphCall.searchQuery).toBe('"amazon"');
    expect(lastGraphCall.filterExpr).toBeNull();
    expect((res.jsonBody as { count: number }).count).toBe(1);
  });

  it('falls back to $filter when $search throws on global search', async () => {
    const filterMessages = [makeMessage('msg-3', 'Amazon Prime receipt')];
    mockGraphGet
      .mockRejectedValueOnce(new Error('SearchQueryNotSupported'))
      .mockResolvedValueOnce({ value: filterMessages });

    const res = await wrappedHandler(
      makeRequest({ q: 'amazon', maxResults: '5' }),
      fakeContext,
    );

    expect(res.status).toBe(200);
    expect(lastGraphCall.searchQuery).toBeNull(); // second call used $filter
    expect(lastGraphCall.filterExpr).toContain("contains(subject,'amazon')");
    expect(lastGraphCall.filterExpr).toContain("contains(from/emailAddress/address,'amazon')");
    expect(lastGraphCall.filterExpr).toContain("contains(from/emailAddress/name,'amazon')");
    expect((res.jsonBody as { count: number }).count).toBe(1);
  });

  it('returns 502 when $search and $filter both fail on global search', async () => {
    mockGraphGet
      .mockRejectedValueOnce(new Error('Archive index unavailable'))
      .mockRejectedValueOnce(new Error('Filter also failed'));

    const res = await wrappedHandler(
      makeRequest({ q: 'amazon' }),
      fakeContext,
    );

    expect(res.status).toBe(502);
    expect((res.jsonBody as { error: string }).error).toContain('Archive index unavailable');
    expect((res.jsonBody as { error: string }).error).toContain('Filter also failed');
  });

  it('returns 502 with $filter-only error message when folder-scoped $filter fails', async () => {
    mockGraphGet.mockRejectedValueOnce(new Error('Filter failed on Archive'));

    const res = await wrappedHandler(
      makeRequest({ q: 'amazon', folderId: FOLDER_ID }),
      fakeContext,
    );

    expect(res.status).toBe(502);
    // Only one call (the $filter attempt); no $search was tried
    expect(mockGraphGet).toHaveBeenCalledTimes(1);
    const err = (res.jsonBody as { error: string }).error;
    expect(err).toContain('Filter failed on Archive');
    // Error message should NOT mention $search since it was never attempted
    expect(err).not.toContain('$search error');
  });

  it('escapes single quotes in query for $filter to prevent OData injection', async () => {
    const filterMessages = [makeMessage('msg-x', "O'Brien invoice")];
    mockGraphGet
      .mockRejectedValueOnce(new Error('SearchQueryNotSupported'))
      .mockResolvedValueOnce({ value: filterMessages });

    await wrappedHandler(
      makeRequest({ q: "O'Brien" }),
      fakeContext,
    );

    // Single quote must be doubled in OData string literals
    expect(lastGraphCall.filterExpr).toContain("contains(subject,'O''Brien')");
    expect(lastGraphCall.filterExpr).not.toContain("contains(subject,'O'Brien')");
  });

  it('honors $top in $filter path', async () => {
    mockGraphGet.mockResolvedValueOnce({ value: [makeMessage('m1', 'Amazon')] });

    await wrappedHandler(
      makeRequest({ q: 'amazon', folderId: FOLDER_ID, maxResults: '7' }),
      fakeContext,
    );

    expect(lastGraphCall.topValue).toBe(7);
  });

  it('does NOT pass $orderby to Graph in $filter path (avoids InefficientFilter)', async () => {
    mockGraphGet.mockResolvedValueOnce({ value: [makeMessage('m1', 'Amazon')] });

    await wrappedHandler(
      makeRequest({ q: 'amazon', folderId: FOLDER_ID }),
      fakeContext,
    );

    expect(lastGraphCall.orderbyValue).toBeNull();
  });

  it('sorts $filter results by receivedDateTime descending client-side', async () => {
    const older = { ...makeMessage('msg-old', 'Old Amazon'), receivedDateTime: '2023-01-01T00:00:00Z' };
    const newer = { ...makeMessage('msg-new', 'New Amazon'), receivedDateTime: '2024-06-01T00:00:00Z' };
    // Graph returns them oldest-first (unsorted)
    mockGraphGet.mockResolvedValueOnce({ value: [older, newer] });

    const res = await wrappedHandler(
      makeRequest({ q: 'amazon', folderId: FOLDER_ID, maxResults: '5' }),
      fakeContext,
    );

    expect(res.status).toBe(200);
    const results = (res.jsonBody as { results: Array<{ id: string }> }).results;
    expect(results[0].id).toBe('msg-new');
    expect(results[1].id).toBe('msg-old');
  });
});
