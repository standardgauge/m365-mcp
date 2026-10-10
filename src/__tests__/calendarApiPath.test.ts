/**
 * Handler-level tests for the calendar HTTP routes.
 *
 * Verifies that every calendar handler — including the new getEvent route —
 * preserves main's security posture:
 *   1. Graph API path routing (explicit calendarId → calendar-scoped path;
 *      omitted → mailbox-level /me/events, matching the existing handlers).
 *   2. assertOpaqueId rejects path-injection payloads with HTTP 400 before any
 *      Graph call.
 *   3. Deny-list enforcement (by calendar ID and, for getEvent, by calendar
 *      name) returns HTTP 403 and never reaches Graph.
 *   4. withSecurity wraps responses with security headers.
 *
 * All Graph, auth, and storage dependencies are mocked so the suite runs with
 * no cloud infrastructure.
 */

import { jest } from '@jest/globals';
import type { HttpRequest, InvocationContext } from '@azure/functions';
import type { AuthResult } from '../services/authMiddleware.js';

// ── Mock declarations ────────────────────────────────────────────────────────

const mockAuthenticateRequest = jest.fn<(req: HttpRequest) => Promise<AuthResult | null>>();
const mockGetEnabledServices = jest.fn<(tenantId: string) => Promise<string[]>>();
const mockIsServiceDisabledForUser = jest.fn<() => Promise<boolean>>();
const mockIsPathDenied = jest.fn<(t: string, u: string, type: string, path: string) => Promise<boolean>>();
const mockGetTenantId = jest.fn<(userId: string) => Promise<string>>();
const mockGetValidAccessTokenForSession = jest.fn<(session: unknown) => Promise<string>>();
const mockResolveDefaultCalendarId = jest.fn<() => Promise<string | null>>();
const mockResolveCalendarName = jest.fn<() => Promise<string>>();

const lastGraphCall: { path: string | null } = { path: null };
const mockGraphGet = jest.fn<() => Promise<unknown>>();
const mockGraphPost = jest.fn<() => Promise<unknown>>();
const mockGraphPatch = jest.fn<() => Promise<unknown>>();
const mockGraphDelete = jest.fn<() => Promise<unknown>>();

// Fully chainable Graph stub — every fluent method returns the same chain, and
// the terminal verbs resolve to the configured mock results.
function makeChain() {
  const chain: Record<string, unknown> = {};
  for (const m of ['select', 'top', 'orderby', 'filter', 'expand', 'header']) {
    chain[m] = () => chain;
  }
  chain.get = (...a: unknown[]) => mockGraphGet(...(a as []));
  chain.post = (...a: unknown[]) => mockGraphPost(...(a as []));
  chain.patch = (...a: unknown[]) => mockGraphPatch(...(a as []));
  chain.delete = (...a: unknown[]) => mockGraphDelete(...(a as []));
  return chain;
}
const mockCreateGraphClient = jest.fn(() => ({
  api: (path: string) => {
    lastGraphCall.path = path;
    return makeChain();
  },
}));

// ── Wire mocks ───────────────────────────────────────────────────────────────

// The outbound policy is covered in outboundPolicy.test.ts; here it allows
// everything unless a test makes it refuse.
const mockEnforceOutboundPolicy = jest.fn<() => Promise<void>>();
const mockEnforceEventUpdatePolicy = jest.fn<() => Promise<void>>();
jest.mock('../services/outboundPolicy.js', () => ({
  ...jest.requireActual<typeof import('../services/outboundPolicy.js')>('../services/outboundPolicy.js'),
  enforceOutboundPolicy: () => mockEnforceOutboundPolicy(),
  enforceEventUpdatePolicy: () => mockEnforceEventUpdatePolicy(),
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
  isPathDenied: (...args: unknown[]) =>
    mockIsPathDenied(args[0] as string, args[1] as string, args[2] as string, args[3] as string),
  filterDeniedPaths: (_t: string, _u: string, _s: string, items: unknown[]) => Promise.resolve(items),
}));

jest.mock('../services/tokenCache.js', () => ({
  getTenantId: (...args: unknown[]) => mockGetTenantId(args[0] as string),
  getTenantIdFromSession: (session: { tenantId?: string }) => {
    if (!session?.tenantId) throw new Error('No tenantId in session');
    return session.tenantId;
  },
  getValidAccessTokenForSession: (s: unknown) => mockGetValidAccessTokenForSession(s),
}));

jest.mock('../services/mailboxTimeZone.js', () => ({
  resolveMailboxTimeZone: jest.fn(async () => 'Pacific Standard Time'),
}));

jest.mock('../services/graphClient.js', () => ({
  createGraphClient: (...args: unknown[]) => (mockCreateGraphClient as (...a: unknown[]) => unknown)(...args),
}));

jest.mock('../services/containerResolver.js', () => ({
  resolveDefaultCalendarId: () => mockResolveDefaultCalendarId(),
  resolveCalendarName: () => mockResolveCalendarName(),
  clearResolverCaches: jest.fn(),
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

// ── Import handlers (triggers app.http registrations) ─────────────────────────

import { app } from '@azure/functions';
import { logAccess } from '../services/auditLog.js';
import { OutboundPolicyError } from '../services/outboundPolicy.js';
import '../functions/calendar/listEvents.js';
import '../functions/calendar/getEvent.js';
import '../functions/calendar/createEvent.js';
import '../functions/calendar/updateEvent.js';
import '../functions/calendar/deleteEvent.js';

interface HttpRegistration {
  handler: (req: HttpRequest, ctx: InvocationContext) => Promise<{ status: number; jsonBody?: unknown; headers?: Record<string, string> }>;
}
const httpMock = app.http as unknown as jest.Mock<(name: string, opts: HttpRegistration) => void>;

function getHandler(name: string) {
  const reg = httpMock.mock.calls.find((c) => c[0] === name);
  if (!reg) throw new Error(`Handler not registered: ${name}`);
  return reg[1].handler;
}

// ── Fixtures ─────────────────────────────────────────────────────────────────

const DEFAULT_CAL_ID = 'default-cal-guid-001';
const EXPLICIT_CAL_ID = 'AAMkExplicitCalGuid002';
const EVENT_ID = 'AAMkEventGuid003';
const TENANT = 'tenant-a';
const USER = 'user-example';
const TOKEN = 'fake-token';
const INJECT = 'legit-id/../../../users/victim';

const FAKE_AUTH: AuthResult = {
  userId: USER,
  session: {
    userId: USER,
    homeAccountId: 'home-abc',
    displayName: 'Test User',
    email: 'user@example.com',
    tenantId: TENANT,
    accessToken: TOKEN,
    expiresAt: Date.now() + 3_600_000,
    sessionToken: 'sess',
    sessionCreatedAt: Date.now(),
  },
} as AuthResult;

const ctx = { error: jest.fn() } as unknown as InvocationContext;

function makeReq(opts: {
  params?: Record<string, string>;
  query?: Record<string, string>;
  body?: unknown;
  headers?: Record<string, string>;
}): HttpRequest {
  return {
    json: async () => opts.body ?? {},
    params: opts.params ?? {},
    query: { get: (k: string) => (opts.query && k in opts.query ? opts.query[k] : null) },
    headers: new Map<string, string>(Object.entries(opts.headers ?? {})),
  } as unknown as HttpRequest;
}

beforeEach(() => {
  jest.clearAllMocks();
  lastGraphCall.path = null;
  mockAuthenticateRequest.mockResolvedValue(FAKE_AUTH);
  mockGetEnabledServices.mockResolvedValue(['calendar']);
  mockIsServiceDisabledForUser.mockResolvedValue(false);
  mockIsPathDenied.mockResolvedValue(false);
  mockGetTenantId.mockResolvedValue(TENANT);
  mockGetValidAccessTokenForSession.mockResolvedValue(TOKEN);
  mockResolveDefaultCalendarId.mockResolvedValue(DEFAULT_CAL_ID);
  mockResolveCalendarName.mockResolvedValue('My Calendar');
  mockGraphGet.mockResolvedValue({ id: EVENT_ID, subject: 'Test', start: {}, end: {}, value: [] });
  mockGraphPost.mockResolvedValue({ id: EVENT_ID, subject: 'New', start: {}, end: {}, webLink: 'http://x' });
  mockGraphPatch.mockResolvedValue({ id: EVENT_ID, subject: 'Test' });
  mockGraphDelete.mockResolvedValue(undefined);
  mockEnforceOutboundPolicy.mockResolvedValue(undefined);
  mockEnforceEventUpdatePolicy.mockResolvedValue(undefined);
});

// ─────────────────────────────────────────────────────────────────────────────
// getEvent — the new route
// ─────────────────────────────────────────────────────────────────────────────

describe('getEvent — Graph API path routing', () => {
  const handler = getHandler('getEvent');

  it('uses /me/calendars/{calendarId}/events/{eventId} when calendarId is explicit', async () => {
    await handler(makeReq({ params: { eventId: EVENT_ID }, query: { calendarId: EXPLICIT_CAL_ID } }), ctx);
    expect(lastGraphCall.path).toBe(`/me/calendars/${EXPLICIT_CAL_ID}/events/${EVENT_ID}`);
  });

  it('uses /me/events/{eventId} when calendarId is omitted', async () => {
    await handler(makeReq({ params: { eventId: EVENT_ID } }), ctx);
    expect(lastGraphCall.path).toBe(`/me/events/${EVENT_ID}`);
  });

  it('returns 200 with the event payload on success', async () => {
    const res = await handler(makeReq({ params: { eventId: EVENT_ID } }), ctx);
    expect(res.status).toBe(200);
    expect((res.jsonBody as { id: string }).id).toBe(EVENT_ID);
  });
});

describe('getEvent — opaque ID injection protection', () => {
  const handler = getHandler('getEvent');

  it('rejects an injected eventId with 400 before any Graph call', async () => {
    const res = await handler(makeReq({ params: { eventId: INJECT } }), ctx);
    expect(res.status).toBe(400);
    expect(lastGraphCall.path).toBeNull();
  });

  it('rejects an injected calendarId with 400 before any Graph call', async () => {
    const res = await handler(makeReq({ params: { eventId: EVENT_ID }, query: { calendarId: INJECT } }), ctx);
    expect(res.status).toBe(400);
    expect(lastGraphCall.path).toBeNull();
  });

  it('400s on a missing eventId path param', async () => {
    const res = await handler(makeReq({}), ctx);
    expect(res.status).toBe(400);
  });
});

describe('getEvent — deny-list enforcement (by ID and by name)', () => {
  const handler = getHandler('getEvent');

  it('returns 403 and skips Graph when the effective calendar ID is denied', async () => {
    mockIsPathDenied.mockImplementation(async (_t, _u, _type, path) => path === DEFAULT_CAL_ID);
    const res = await handler(makeReq({ params: { eventId: EVENT_ID } }), ctx);
    expect(res.status).toBe(403);
    expect(mockGraphGet).not.toHaveBeenCalled();
  });

  it('returns 403 when the calendar NAME is denied even though the ID is not', async () => {
    mockResolveCalendarName.mockResolvedValue('HR Sensitive Calendar');
    mockIsPathDenied.mockImplementation(async (_t, _u, _type, path) => path === 'HR Sensitive Calendar');
    const res = await handler(makeReq({ params: { eventId: EVENT_ID } }), ctx);
    expect(res.status).toBe(403);
    expect(mockGraphGet).not.toHaveBeenCalled();
  });

  it('allows the read when neither ID nor name is denied', async () => {
    const res = await handler(makeReq({ params: { eventId: EVENT_ID } }), ctx);
    expect(res.status).toBe(200);
    expect(mockGraphGet).toHaveBeenCalled();
  });
});

describe('getEvent — withSecurity headers', () => {
  const handler = getHandler('getEvent');

  it('attaches security headers to the response', async () => {
    const res = await handler(makeReq({ params: { eventId: EVENT_ID } }), ctx);
    expect(res.headers?.['X-Content-Type-Options']).toBe('nosniff');
    expect(res.headers?.['X-Frame-Options']).toBe('DENY');
  });
});

// ─────────────────────────────────────────────────────────────────────────────
// Existing handlers — path routing regression guard (main's behavior preserved)
// ─────────────────────────────────────────────────────────────────────────────

describe('listEvents — Graph API path', () => {
  const handler = getHandler('listEvents');

  it('uses /me/calendars/{calendarId}/events when calendarId is explicit', async () => {
    await handler(makeReq({ query: { calendarId: EXPLICIT_CAL_ID } }), ctx);
    expect(lastGraphCall.path).toBe(`/me/calendars/${EXPLICIT_CAL_ID}/events`);
  });

  it('uses /me/events when calendarId is omitted', async () => {
    await handler(makeReq({}), ctx);
    expect(lastGraphCall.path).toBe('/me/events');
  });

  it('rejects an injected calendarId with 400', async () => {
    const res = await handler(makeReq({ query: { calendarId: INJECT } }), ctx);
    expect(res.status).toBe(400);
    expect(lastGraphCall.path).toBeNull();
  });
});

describe('createEvent — Graph API path', () => {
  const handler = getHandler('createEvent');
  const body = { subject: 'Sync', start: '2026-07-01T10:00:00', end: '2026-07-01T11:00:00' };

  it('uses /me/calendars/{calendarId}/events when calendarId is explicit', async () => {
    await handler(makeReq({ body: { ...body, calendarId: EXPLICIT_CAL_ID } }), ctx);
    expect(lastGraphCall.path).toBe(`/me/calendars/${EXPLICIT_CAL_ID}/events`);
  });

  it('uses /me/events when calendarId is omitted', async () => {
    await handler(makeReq({ body }), ctx);
    expect(lastGraphCall.path).toBe('/me/events');
  });
});

describe('updateEvent — Graph API path', () => {
  const handler = getHandler('updateEvent');

  it('uses /me/calendars/{calendarId}/events/{eventId} when calendarId is explicit', async () => {
    await handler(makeReq({ params: { eventId: EVENT_ID }, query: { calendarId: EXPLICIT_CAL_ID }, body: { subject: 'x' } }), ctx);
    expect(lastGraphCall.path).toBe(`/me/calendars/${EXPLICIT_CAL_ID}/events/${EVENT_ID}`);
  });

  it('rejects an injected eventId with 400', async () => {
    const res = await handler(makeReq({ params: { eventId: INJECT }, body: { subject: 'x' } }), ctx);
    expect(res.status).toBe(400);
    expect(lastGraphCall.path).toBeNull();
  });
});

describe('createEvent / updateEvent — outbound-policy refusal audit row', () => {
  // A client can prepend its own entries to X-Forwarded-For; only the one the
  // ingress appended is trustworthy, and that is what the row must record.
  const spoofed = { 'x-forwarded-for': '198.51.100.66, 203.0.113.9' };
  const mockLogAccess = logAccess as unknown as jest.Mock;
  const createEvent = getHandler('createEvent');
  const updateEvent = getHandler('updateEvent');

  it('createEvent records the trusted client address, not the raw header', async () => {
    mockEnforceOutboundPolicy.mockRejectedValue(new OutboundPolicyError('external invitations are blocked'));
    const res = await createEvent(
      makeReq({ body: { subject: 'Sync', start: '2026-07-01T10:00:00', end: '2026-07-01T11:00:00' }, headers: spoofed }),
      ctx,
    );
    expect(res.status).toBe(403);
    expect(mockGraphPost).not.toHaveBeenCalled();
    expect(mockLogAccess).toHaveBeenCalledWith(
      expect.objectContaining({ operation: 'calendar.post', result: 'denied', ip: '203.0.113.9' }),
    );
  });

  it('updateEvent records the trusted client address, not the raw header', async () => {
    mockEnforceEventUpdatePolicy.mockRejectedValue(new OutboundPolicyError('external invitations are blocked'));
    const res = await updateEvent(
      makeReq({ params: { eventId: EVENT_ID }, body: { subject: 'x' }, headers: spoofed }),
      ctx,
    );
    expect(res.status).toBe(403);
    expect(mockGraphPatch).not.toHaveBeenCalled();
    expect(mockLogAccess).toHaveBeenCalledWith(
      expect.objectContaining({ operation: 'calendar.patch', result: 'denied', ip: '203.0.113.9' }),
    );
  });

  it('records no address when the selected entry is not an IP', async () => {
    mockEnforceOutboundPolicy.mockRejectedValue(new OutboundPolicyError('external invitations are blocked'));
    await createEvent(
      makeReq({
        body: { subject: 'Sync', start: '2026-07-01T10:00:00', end: '2026-07-01T11:00:00' },
        headers: { 'x-forwarded-for': '<script>' },
      }),
      ctx,
    );
    const row = mockLogAccess.mock.calls.map((c) => c[0] as { operation: string; ip?: string })
      .find((r) => r.operation === 'calendar.post');
    expect(row).toBeDefined();
    expect(row?.ip).toBeUndefined();
  });
});

describe('deleteEvent — Graph API path', () => {
  const handler = getHandler('deleteEvent');

  it('uses /me/calendars/{calendarId}/events/{eventId} when calendarId is explicit', async () => {
    await handler(makeReq({ params: { eventId: EVENT_ID }, query: { calendarId: EXPLICIT_CAL_ID } }), ctx);
    expect(lastGraphCall.path).toBe(`/me/calendars/${EXPLICIT_CAL_ID}/events/${EVENT_ID}`);
  });

  it('rejects an injected eventId with 400', async () => {
    const res = await handler(makeReq({ params: { eventId: INJECT } }), ctx);
    expect(res.status).toBe(400);
    expect(lastGraphCall.path).toBeNull();
  });
});
