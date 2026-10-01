/**
 * Tests for sendMail email output mode behavior.
 *
 * Covers:
 *   - draft mode: saves to Drafts, returns queued_as_draft status with draftId/draftLink
 *   - send mode: delivers immediately, returns sent status
 *   - draft mode: checks Drafts deny list (not Sent Items)
 *   - send mode: checks Sent Items deny list
 *   - validation: 400 when subject or to[] missing (both modes)
 *   - from: a proxy address of the mailbox is set as Graph `from`; anything else is 400
 */

import { jest } from '@jest/globals';
import type { HttpRequest, InvocationContext } from '@azure/functions';
import type { AuthResult } from '../services/authMiddleware.js';

const mockGetTenantIdFromSession = jest.fn<(session: unknown) => string>();
const mockGetValidAccessTokenForSession = jest.fn<(session: unknown) => Promise<string>>();
const mockGetUserEmailSettings = jest.fn<
  (tenantId: string, userId: string) => Promise<{ emailOutputMode: string }>
>();
const mockCheckDenyList = jest.fn<
  (userId: string, service: string, path: string, session?: unknown) => Promise<null | { status: number; error: string }>
>();
const mockGraphPost = jest.fn<(payload: unknown) => Promise<unknown>>();
const mockGraphGet = jest.fn<() => Promise<unknown>>();
// eslint-disable-next-line @typescript-eslint/no-explicit-any
const mockGraphApi = jest.fn((_path?: string): any => {
  const chain = { post: mockGraphPost, get: mockGraphGet, select: () => chain };
  return chain;
});
const mockCreateGraphClient = jest.fn<() => { api: typeof mockGraphApi }>(() => ({ api: mockGraphApi }));
const mockWithPolicyEnforcement = jest.fn(
  (_service: string, handler: (req: HttpRequest, ctx: InvocationContext, auth: AuthResult) => Promise<unknown>) =>
    handler,
);

jest.mock('../services/tokenCache.js', () => ({
  getValidAccessTokenForSession: (session: unknown) => mockGetValidAccessTokenForSession(session),
  getTenantIdFromSession: (session: unknown) => mockGetTenantIdFromSession(session),
}));

jest.mock('../services/graphClient.js', () => ({
  createGraphClient: () => mockCreateGraphClient(),
}));

jest.mock('../services/userEmailSettings.js', () => ({
  getUserEmailSettings: (tenantId: unknown, userId: unknown) =>
    mockGetUserEmailSettings(tenantId as string, userId as string),
}));

jest.mock('../services/policyEnforcement.js', () => ({
  withPolicyEnforcement: (service: unknown, handler: unknown) =>
    mockWithPolicyEnforcement(service as string, handler as (req: HttpRequest, ctx: InvocationContext, auth: AuthResult) => Promise<unknown>),
  checkDenyList: (...args: unknown[]) =>
    mockCheckDenyList(args[0] as string, args[1] as string, args[2] as string, args[3]),
}));

jest.mock('../services/securityHeaders.js', () => ({
  withSecurity: (handler: unknown) => handler,
}));

jest.mock('@azure/functions', () => ({
  app: { http: jest.fn() },
}));

import { app } from '@azure/functions';
import '../functions/mail/sendMail.js';

interface HttpRegistration {
  handler: (req: HttpRequest, context: InvocationContext, auth: AuthResult) => Promise<{ status: number; jsonBody?: unknown }>;
}
const httpMock = app.http as unknown as jest.Mock<(name: string, opts: HttpRegistration) => void>;
const registration = httpMock.mock.calls.find((call) => call[0] === 'sendMail');
if (!registration) throw new Error('sendMail handler was not registered');
const handler = registration[1].handler;

const TENANT = 'test-tenant';
const USER = 'test-user';

const AUTH: AuthResult = {
  userId: USER,
  session: {
    userId: USER,
    homeAccountId: 'home-user',
    displayName: 'User',
    email: 'user@example.com',
    tenantId: TENANT,
    accessToken: 'fake-token',
    expiresAt: Date.now() + 3_600_000,
    sessionToken: 'fake-session',
    sessionCreatedAt: Date.now(),
  },
};

function makeRequest(body: unknown): HttpRequest {
  return {
    method: 'POST',
    query: { get: () => null, has: () => false },
    headers: new Map<string, string>(),
    json: () => Promise.resolve(body),
  } as unknown as HttpRequest;
}

function makeContext(): InvocationContext {
  return { error: jest.fn(), warn: jest.fn(), log: jest.fn() } as unknown as InvocationContext;
}

beforeEach(() => {
  jest.clearAllMocks();
  mockGetTenantIdFromSession.mockReturnValue(TENANT);
  mockGetValidAccessTokenForSession.mockResolvedValue('access-token');
  mockCheckDenyList.mockResolvedValue(null);
  mockGraphGet.mockResolvedValue({
    mail: 'nate@example.com',
    proxyAddresses: ['SMTP:nate@example.com', 'smtp:nate@standardgauge.ai'],
  });
});

describe('sendMail — draft mode (default)', () => {
  beforeEach(() => {
    mockGetUserEmailSettings.mockResolvedValue({ emailOutputMode: 'draft' });
    mockGraphPost.mockResolvedValue({ id: 'draft-123', webLink: 'https://outlook.office.com/draft-123' });
  });

  it('saves to Drafts and returns queued_as_draft status', async () => {
    const res = await handler(
      makeRequest({ subject: 'Hello', body: 'World', to: ['alice@example.com'] }),
      makeContext(),
      AUTH,
    );

    expect(res.status).toBe(200);
    const body = res.jsonBody as { status: string; draftId: string; draftLink: string };
    expect(body.status).toBe('queued_as_draft');
    expect(body.draftId).toBe('draft-123');
    expect(body.draftLink).toBe('https://outlook.office.com/draft-123');
  });

  it('checks Drafts deny list (not Sent Items)', async () => {
    await handler(
      makeRequest({ subject: 'Hello', body: 'World', to: ['alice@example.com'] }),
      makeContext(),
      AUTH,
    );

    expect(mockCheckDenyList).toHaveBeenCalledWith(USER, 'mail', 'Drafts', AUTH.session);
    expect(mockCheckDenyList).not.toHaveBeenCalledWith(USER, 'mail', 'Sent Items', expect.anything());
  });

  it('posts to /me/messages (draft endpoint), not /me/sendMail', async () => {
    await handler(
      makeRequest({ subject: 'Test', body: 'Body', to: ['b@x.com'] }),
      makeContext(),
      AUTH,
    );

    expect(mockGraphApi).toHaveBeenCalledWith('/me/messages');
    expect(mockGraphApi).not.toHaveBeenCalledWith('/me/sendMail');
  });

  it('returns 403 when Drafts folder is deny-listed', async () => {
    mockCheckDenyList.mockResolvedValue({ status: 403, error: 'Access restricted by deny list' });

    const res = await handler(
      makeRequest({ subject: 'Hi', body: 'Yo', to: ['c@x.com'] }),
      makeContext(),
      AUTH,
    );

    expect(res.status).toBe(403);
    expect(mockGraphPost).not.toHaveBeenCalled();
  });

  it('includes cc/bcc in draft message', async () => {
    await handler(
      makeRequest({ subject: 'Hi', body: 'Msg', to: ['a@x.com'], cc: ['b@x.com'], bcc: ['c@x.com'] }),
      makeContext(),
      AUTH,
    );

    const posted = mockGraphPost.mock.calls[0][0] as {
      ccRecipients: Array<{ emailAddress: { address: string } }>;
      bccRecipients: Array<{ emailAddress: { address: string } }>;
    };
    expect(posted.ccRecipients[0].emailAddress.address).toBe('b@x.com');
    expect(posted.bccRecipients[0].emailAddress.address).toBe('c@x.com');
  });
});

describe('sendMail — send mode', () => {
  beforeEach(() => {
    mockGetUserEmailSettings.mockResolvedValue({ emailOutputMode: 'send' });
    mockGraphPost.mockResolvedValue(undefined);
  });

  it('sends immediately and returns sent status', async () => {
    const res = await handler(
      makeRequest({ subject: 'Hello', body: 'World', to: ['alice@example.com'] }),
      makeContext(),
      AUTH,
    );

    expect(res.status).toBe(200);
    const body = res.jsonBody as { status: string };
    expect(body.status).toBe('sent');
  });

  it('checks Sent Items deny list (not Drafts)', async () => {
    await handler(
      makeRequest({ subject: 'Hello', body: 'World', to: ['alice@example.com'] }),
      makeContext(),
      AUTH,
    );

    expect(mockCheckDenyList).toHaveBeenCalledWith(USER, 'mail', 'Sent Items', AUTH.session);
    expect(mockCheckDenyList).not.toHaveBeenCalledWith(USER, 'mail', 'Drafts', expect.anything());
  });

  it('posts to /me/sendMail', async () => {
    await handler(
      makeRequest({ subject: 'Test', body: 'Body', to: ['b@x.com'] }),
      makeContext(),
      AUTH,
    );

    expect(mockGraphApi).toHaveBeenCalledWith('/me/sendMail');
  });

  it('skips Sent Items deny check when saveToSentItems is false (inline attachments only)', async () => {
    await handler(
      makeRequest({ subject: 'Hi', body: 'Msg', to: ['a@x.com'], saveToSentItems: false }),
      makeContext(),
      AUTH,
    );

    expect(mockCheckDenyList).not.toHaveBeenCalled();
  });

  // A base64 payload whose decoded length exceeds INLINE_ATTACHMENT_LIMIT_BYTES (3 MB), forcing
  // the upload-session (draft + /messages/{id}/send) path. 4.2M base64 chars ≈ 3.15 MB decoded.
  const LARGE_ATTACHMENT = {
    name: 'big.bin',
    content: 'A'.repeat(4_200_000),
  };

  it('rejects saveToSentItems:false with a large (upload-session) attachment (400)', async () => {
    const res = await handler(
      makeRequest({
        subject: 'Hi',
        body: 'Msg',
        to: ['a@x.com'],
        saveToSentItems: false,
        attachments: [LARGE_ATTACHMENT],
      }),
      makeContext(),
      AUTH,
    );

    // The upload-session fallback always saves to Sent Items — honoring saveToSentItems:false is
    // impossible on Graph, so we must reject rather than silently save and skip the deny check.
    expect(res.status).toBe(400);
    expect((res.jsonBody as { error: string }).error).toMatch(/saveToSentItems:false is not supported/);
    expect(mockGraphPost).not.toHaveBeenCalled();
  });

  it('still checks Sent Items deny list for a large attachment even when saveToSentItems is false', async () => {
    mockCheckDenyList.mockResolvedValue({ status: 403, error: 'Access restricted by deny list' });

    // saveToSentItems omitted (defaults to true) → large attachment still saves to Sent Items,
    // so the deny list must be enforced. Contrast with the 400 case above where the flag is false.
    const res = await handler(
      makeRequest({
        subject: 'Hi',
        body: 'Msg',
        to: ['a@x.com'],
        attachments: [LARGE_ATTACHMENT],
      }),
      makeContext(),
      AUTH,
    );

    expect(mockCheckDenyList).toHaveBeenCalledWith(USER, 'mail', 'Sent Items', AUTH.session);
    expect(res.status).toBe(403);
    expect(mockGraphPost).not.toHaveBeenCalled();
  });
});

describe('sendMail — input validation', () => {
  beforeEach(() => {
    mockGetUserEmailSettings.mockResolvedValue({ emailOutputMode: 'draft' });
  });

  it('returns 400 when subject is missing', async () => {
    const res = await handler(
      makeRequest({ body: 'World', to: ['alice@example.com'] }),
      makeContext(),
      AUTH,
    );

    expect(res.status).toBe(400);
    expect(mockGetUserEmailSettings).not.toHaveBeenCalled();
  });

  it('returns 400 when to[] is empty', async () => {
    const res = await handler(
      makeRequest({ subject: 'Hi', body: 'World', to: [] }),
      makeContext(),
      AUTH,
    );

    expect(res.status).toBe(400);
    expect(mockGetUserEmailSettings).not.toHaveBeenCalled();
  });
});

describe('sendMail — from', () => {
  it.each(['draft', 'send'])('%s mode sets Graph from to a validated alias', async (mode) => {
    mockGetUserEmailSettings.mockResolvedValue({ emailOutputMode: mode });
    mockGraphPost.mockResolvedValue({ id: 'draft-123' });
    const res = await handler(
      makeRequest({ subject: 'Hi', body: 'World', to: ['alice@example.com'], from: 'nate@standardgauge.ai' }),
      makeContext(),
      AUTH,
    );
    expect(res.status).toBe(200);
    const posted = mockGraphPost.mock.calls[0][0] as { from?: unknown; message?: { from?: unknown } };
    const from = mode === 'send' ? posted.message?.from : posted.from;
    expect(from).toEqual({ emailAddress: { address: 'nate@standardgauge.ai' } });
  });

  it.each(['draft', 'send'])('%s mode returns 400 for an address that is not the mailbox\'s', async (mode) => {
    mockGetUserEmailSettings.mockResolvedValue({ emailOutputMode: mode });
    const res = await handler(
      makeRequest({ subject: 'Hi', body: 'World', to: ['alice@example.com'], from: 'ceo@competitor.example' }),
      makeContext(),
      AUTH,
    );
    expect(res.status).toBe(400);
    expect(mockGraphPost).not.toHaveBeenCalled();
  });
});
