/**
 * Send-as-alias tests: the `from` parameter on create_draft / send_mail / send_draft /
 * update_message / reply_to_message / reply_all_to_message / forward_message, and the reply
 * tools' default of the alias the original was addressed to.
 *
 * Two layers:
 *   1. mailFrom helpers — proxyAddresses parsing, validation, and the reply-default pick.
 *   2. Remote MCP endpoint — each tool sets Graph `from` only to a validated proxy address,
 *      rejects anything else before writing, and reports the From it used.
 *
 * All Graph, auth, and storage dependencies are mocked. Whether Exchange honors an alias From
 * on the wire is a live-tenant question these tests cannot answer; see
 * the send-as-alias notes.
 */

import { jest } from '@jest/globals';
import type { HttpRequest, InvocationContext } from '@azure/functions';

// ── Mock declarations ────────────────────────────────────────────────────────

const mockAuthenticateRequest = jest.fn<() => Promise<unknown>>();
const mockGetValidAccessTokenForSession = jest.fn<() => Promise<string>>();
const mockGetTenantIdFromSession = jest.fn<() => string>();
const mockIsPathDenied = jest.fn<(t: string, u: string, type: string, path: string) => Promise<boolean>>();
const mockGetEnabledServices = jest.fn<() => Promise<string[]>>();
const mockGetReadOnlyServices = jest.fn<() => Promise<string[]>>();
const mockGetAllowedSites = jest.fn<() => Promise<Array<{ id: string; name: string }>>>();
const mockGetUserServiceOverrides = jest.fn<() => Promise<string[]>>();
const mockResolveMailFolderName = jest.fn<() => Promise<string | null>>();

interface GraphCall {
  path: string;
  method: 'get' | 'post' | 'patch' | 'delete';
  body?: unknown;
}
const graphCalls: GraphCall[] = [];
const mockGraphGet = jest.fn<(path: string) => Promise<unknown>>();
const mockGraphPost = jest.fn<(path: string, body: unknown) => Promise<unknown>>();
const mockGraphPatch = jest.fn<(path: string, body: unknown) => Promise<unknown>>();
let emailOutputMode: 'draft' | 'send' = 'draft';

function makeChain(path: string) {
  const chain: Record<string, unknown> = {};
  for (const m of ['select', 'top', 'orderby', 'filter', 'expand', 'header', 'search']) chain[m] = () => chain;
  chain.get = async () => {
    graphCalls.push({ path, method: 'get' });
    return mockGraphGet(path);
  };
  chain.post = async (body: unknown) => {
    graphCalls.push({ path, method: 'post', body });
    return mockGraphPost(path, body);
  };
  chain.patch = async (body: unknown) => {
    graphCalls.push({ path, method: 'patch', body });
    return mockGraphPatch(path, body);
  };
  chain.delete = async () => {
    graphCalls.push({ path, method: 'delete' });
    return undefined;
  };
  return chain;
}
const mockCreateGraphClient = jest.fn(() => ({
  api: (path: string) => makeChain(path),
}));

// telemetry.js has import-time side effects (patches console) — stub it out
// Routing tests: the delegated-mailbox owner lookup (its own Graph call) is
// covered by mailboxOwner.test.ts and delegatedDenyList.test.ts.
jest.mock('../services/mailboxOwner.js', () => ({
  resolveDenySubject: (_g: unknown, callerId: string) => Promise.resolve(callerId),
}));

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
  filterDeniedSearchHits: (_t: unknown, _u: unknown, _s: unknown, items: unknown) => Promise.resolve(items),
  isPathDenied: (...args: unknown[]) =>
    mockIsPathDenied(args[0] as string, args[1] as string, args[2] as string, args[3] as string),
}));

jest.mock('../services/containerResolver.js', () => ({
  resolveMailFolderName: () => mockResolveMailFolderName(),
  resolveDefaultCalendarId: () => Promise.resolve('cal'),
  resolveCalendarName: () => Promise.resolve('Calendar'),
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
  getReadOnlyServices: () => mockGetReadOnlyServices(),
  getAllowedSites: () => mockGetAllowedSites(),
}));

jest.mock('../services/userServiceOverrides.js', () => ({
  getUserServiceOverrides: () => mockGetUserServiceOverrides(),
}));

jest.mock('../services/userEmailSettings.js', () => ({
  getUserEmailSettings: () => Promise.resolve({ emailOutputMode }),
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
import {
  parseMailboxAddresses,
  validateFromAddress,
  pickReplyFrom,
  FromAddressError,
} from '../services/mailFrom.js';

interface HttpRegistration {
  handler: (req: HttpRequest, context: InvocationContext) => Promise<{ status: number; jsonBody?: unknown }>;
}
const httpMock = app.http as unknown as jest.Mock<(name: string, opts: HttpRegistration) => void>;
const registration = httpMock.mock.calls.find((call) => call[0] === 'mcpEndpoint');
if (!registration) throw new Error('mcpEndpoint handler was not registered');
const handler = registration[1].handler;

const TENANT = 'test-tenant';
const USER = 'test-user';
const MESSAGE_ID = 'AAMkInboundMsgGuid001';
const DRAFT_ID = 'AAMkReplyDraftGuid002';
const AUTH = { userId: USER, session: { userId: USER, tenantId: TENANT, accessToken: 'fake', sessionToken: 'sess' } };

const PRIMARY = 'nate@example.com';
const ALIAS = 'nate@standardgauge.ai';
const MAILBOX = {
  mail: PRIMARY,
  proxyAddresses: [`SMTP:${PRIMARY}`, `smtp:${ALIAS}`, 'SIP:nate@example.com', 'X500:/o=ExchangeLabs/cn=nate'],
};
const CLIENT = 'ceo@acme.example';
const rcpt = (address: string) => ({ emailAddress: { address } });

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

function toolResult(res: { jsonBody?: unknown }): { parsed: any; isError: boolean; text: string } {
  const body = res.jsonBody as { result?: { content?: Array<{ text: string }>; isError?: boolean } };
  const text = body.result?.content?.[0]?.text ?? '{}';
  const isError = body.result?.isError ?? false;
  return { parsed: isError ? null : JSON.parse(text), isError, text };
}

const calls = (method: GraphCall['method']) => graphCalls.filter((c) => c.method === method);

/** Original message the reply tools read: sent by the client to our alias. */
let original: Record<string, unknown>;
let mailboxLookup: () => Promise<unknown>;

beforeEach(() => {
  jest.clearAllMocks();
  graphCalls.length = 0;
  emailOutputMode = 'draft';
  mockAuthenticateRequest.mockResolvedValue(AUTH);
  mockGetValidAccessTokenForSession.mockResolvedValue('access-token');
  mockGetTenantIdFromSession.mockReturnValue(TENANT);
  mockIsPathDenied.mockResolvedValue(false);
  mockGetEnabledServices.mockResolvedValue(['mail']);
  mockGetReadOnlyServices.mockResolvedValue([]);
  mockGetAllowedSites.mockResolvedValue([]);
  mockGetUserServiceOverrides.mockResolvedValue([]);
  mockResolveMailFolderName.mockResolvedValue('Inbox');
  original = {
    parentFolderId: 'inbox-folder-id',
    from: rcpt(CLIENT),
    toRecipients: [rcpt(ALIAS)],
    ccRecipients: [rcpt('partner@acme.example')],
  };
  mailboxLookup = () => Promise.resolve(MAILBOX);
  mockGraphGet.mockImplementation((path: string) => {
    if (path === '/me' || /^\/users\/[^/]+$/.test(path)) return mailboxLookup();
    if (path.endsWith(`/messages/${DRAFT_ID}`)) {
      return Promise.resolve({ isDraft: true, parentFolderId: 'drafts-id', subject: 'RE: SOW', from: rcpt(PRIMARY), toRecipients: [rcpt(CLIENT)], webLink: 'https://outlook/draft' });
    }
    return Promise.resolve(original);
  });
  mockGraphPost.mockImplementation((path: string) => {
    if (/\/(createReply|createReplyAll|createForward)$/.test(path) || path.endsWith('/messages')) {
      return Promise.resolve({ id: DRAFT_ID, subject: 'RE: SOW', from: rcpt(PRIMARY), toRecipients: [rcpt(CLIENT)], webLink: 'https://outlook/draft' });
    }
    return Promise.resolve({});
  });
  mockGraphPatch.mockImplementation((_path: string, body: unknown) => Promise.resolve({ id: DRAFT_ID, ...(body as object) }));
});

// ─────────────────────────────────────────────────────────────────────────────
// mailFrom helpers
// ─────────────────────────────────────────────────────────────────────────────

describe('mailFrom helpers', () => {
  const mailbox = parseMailboxAddresses(MAILBOX.proxyAddresses, MAILBOX.mail);

  it('parses SMTP proxy addresses, marks the uppercase SMTP: entry primary, ignores non-SMTP', () => {
    expect(mailbox.primary).toBe(PRIMARY);
    expect([...mailbox.addresses.values()].sort()).toEqual([ALIAS, PRIMARY].sort());
  });

  it('falls back to `mail` when proxyAddresses is absent', () => {
    const m = parseMailboxAddresses(undefined, PRIMARY);
    expect(m.primary).toBe(PRIMARY);
    expect([...m.addresses.values()]).toEqual([PRIMARY]);
  });

  it('validates case-insensitively and returns the address as Exchange spells it', () => {
    expect(validateFromAddress('  Nate@StandardGauge.AI ', mailbox)).toBe(ALIAS);
  });

  it.each([['someone@else.example'], [''], [42], [null]])('rejects %p', (bad) => {
    expect(() => validateFromAddress(bad, mailbox)).toThrow(FromAddressError);
  });

  it('pickReplyFrom: alias in To wins', () => {
    expect(pickReplyFrom({ from: rcpt(CLIENT), toRecipients: [rcpt(ALIAS)] }, mailbox)).toBe(ALIAS);
  });

  it('pickReplyFrom: alias in CC is found when To has none of ours', () => {
    expect(pickReplyFrom({ toRecipients: [rcpt('x@y.example')], ccRecipients: [rcpt(ALIAS)] }, mailbox)).toBe(ALIAS);
  });

  it('pickReplyFrom: our own sent message defaults to the From it went out as', () => {
    expect(pickReplyFrom({ from: rcpt(ALIAS), toRecipients: [rcpt(CLIENT)] }, mailbox)).toBe(ALIAS);
  });

  it('pickReplyFrom: primary match or no match leaves Graph\'s default (null)', () => {
    expect(pickReplyFrom({ toRecipients: [rcpt(PRIMARY), rcpt(ALIAS)] }, mailbox)).toBeNull();
    expect(pickReplyFrom({ toRecipients: [rcpt('list@group.example')] }, mailbox)).toBeNull();
  });
});

// ─────────────────────────────────────────────────────────────────────────────
// Reply tools — default and explicit From
// ─────────────────────────────────────────────────────────────────────────────

const REPLY_TOOLS = ['reply_to_message', 'reply_all_to_message', 'forward_message'] as const;
const replyArgs = (name: string, extra: Record<string, unknown> = {}) => ({
  messageId: MESSAGE_ID,
  comment: 'Thanks',
  ...(name === 'forward_message' ? { to: ['x@example.com'] } : {}),
  ...extra,
});

describe('reply/forward tools — From', () => {
  it.each(REPLY_TOOLS)('%s defaults From to the alias the original was addressed to', async (name) => {
    const { parsed, isError, text } = toolResult(await callTool(name, replyArgs(name)));
    expect(isError).toBe(false);
    expect(text).not.toContain('fromWarning');
    expect(calls('patch')).toEqual([
      { path: `/me/messages/${DRAFT_ID}`, method: 'patch', body: { from: rcpt(ALIAS) } },
    ]);
    expect(parsed).toMatchObject({ from: ALIAS, fromSource: 'original-recipient' });
  });

  it('leaves From alone when the original was addressed to the primary', async () => {
    original.toRecipients = [rcpt(PRIMARY)];
    const { parsed } = toolResult(await callTool('reply_to_message', replyArgs('reply_to_message')));
    expect(calls('patch')).toHaveLength(0);
    expect(parsed).toMatchObject({ from: PRIMARY, fromSource: 'mailbox-default' });
  });

  it('an explicit from overrides the default', async () => {
    original.toRecipients = [rcpt(PRIMARY)];
    const { parsed } = toolResult(await callTool('reply_to_message', replyArgs('reply_to_message', { from: ALIAS })));
    expect(calls('patch')[0].body).toEqual({ from: rcpt(ALIAS) });
    expect(parsed).toMatchObject({ from: ALIAS, fromSource: 'explicit' });
  });

  it.each(REPLY_TOOLS)('%s rejects a from outside the mailbox\'s proxy addresses without creating a draft', async (name) => {
    const { isError, text } = toolResult(await callTool(name, replyArgs(name, { from: 'ceo@competitor.example' })));
    expect(isError).toBe(true);
    expect(text).toContain('not an address of this mailbox');
    expect(calls('post')).toHaveLength(0);
    expect(calls('patch')).toHaveLength(0);
  });

  it('validates against the mailboxId mailbox, not the caller', async () => {
    const res = await callTool('reply_to_message', replyArgs('reply_to_message', { mailboxId: 'shared-mbx-id', from: ALIAS }));
    expect(toolResult(res).isError).toBe(false);
    expect(graphCalls.find((c) => c.path === '/users/shared-mbx-id')).toBeDefined();
    expect(calls('patch')[0].path).toBe(`/users/shared-mbx-id/messages/${DRAFT_ID}`);
  });

  it('falls back to the primary with a warning when the mailbox addresses cannot be read', async () => {
    mailboxLookup = () => Promise.reject(Object.assign(new Error('Insufficient privileges'), { statusCode: 403 }));
    const { parsed, isError } = toolResult(await callTool('reply_to_message', replyArgs('reply_to_message')));
    expect(isError).toBe(false);
    expect(calls('patch')).toHaveLength(0);
    expect(parsed.fromSource).toBe('mailbox-default');
    expect(parsed.fromWarning).toContain('Insufficient privileges');
  });

  it('an explicit from fails closed when the mailbox addresses cannot be read', async () => {
    mailboxLookup = () => Promise.reject(new Error('Insufficient privileges'));
    const { isError } = toolResult(await callTool('reply_to_message', replyArgs('reply_to_message', { from: ALIAS })));
    expect(isError).toBe(true);
    expect(calls('post')).toHaveLength(0);
  });

  it('deletes the draft when setting From fails, rather than leave one on the wrong address', async () => {
    mockGraphPatch.mockRejectedValue(new Error('ErrorSendAsDenied'));
    const { isError, text } = toolResult(await callTool('reply_to_message', replyArgs('reply_to_message')));
    expect(isError).toBe(true);
    expect(text).toContain('ErrorSendAsDenied');
    expect(calls('delete')).toEqual([{ path: `/me/messages/${DRAFT_ID}`, method: 'delete' }]);
  });
});

// ─────────────────────────────────────────────────────────────────────────────
// create_draft / send_mail / send_draft / update_message — explicit From
// ─────────────────────────────────────────────────────────────────────────────

describe('compose and dispatch tools — From', () => {
  const compose = { subject: 'SOW', body: 'Attached', to: [CLIENT] };

  it('create_draft sets Graph from on the draft POST', async () => {
    const { parsed, isError } = toolResult(await callTool('create_draft', { ...compose, from: ALIAS }));
    expect(isError).toBe(false);
    expect(calls('post')[0]).toMatchObject({ path: '/me/messages', body: { from: rcpt(ALIAS) } });
    expect(parsed.from).toBe(ALIAS);
  });

  it('create_draft without from sends no from (unchanged behavior)', async () => {
    await callTool('create_draft', compose);
    expect(calls('post')[0].body).not.toHaveProperty('from');
    expect(graphCalls.find((c) => c.path === '/me' && c.method === 'get')).toBeUndefined();
  });

  it('create_draft rejects a foreign from before writing', async () => {
    const { isError } = toolResult(await callTool('create_draft', { ...compose, from: 'boss@competitor.example' }));
    expect(isError).toBe(true);
    expect(calls('post')).toHaveLength(0);
  });

  it('send_mail in send mode puts from on the sendMail message', async () => {
    emailOutputMode = 'send';
    const { parsed } = toolResult(await callTool('send_mail', { ...compose, from: ALIAS }));
    expect(calls('post')[0].path).toBe('/me/sendMail');
    expect((calls('post')[0].body as { message: Record<string, unknown> }).message.from).toEqual(rcpt(ALIAS));
    expect(parsed).toMatchObject({ status: 'sent', from: ALIAS });
  });

  it('send_mail in draft mode puts from on the queued draft', async () => {
    const { parsed } = toolResult(await callTool('send_mail', { ...compose, from: ALIAS }));
    expect(calls('post')[0]).toMatchObject({ path: '/me/messages', body: { from: rcpt(ALIAS) } });
    expect(parsed).toMatchObject({ status: 'queued_as_draft', from: ALIAS });
  });

  it('send_mail rejects a foreign from before sending', async () => {
    emailOutputMode = 'send';
    const { isError } = toolResult(await callTool('send_mail', { ...compose, from: 'boss@competitor.example' }));
    expect(isError).toBe(true);
    expect(calls('post')).toHaveLength(0);
  });

  it('send_draft PATCHes from onto the draft before /send and reports it', async () => {
    emailOutputMode = 'send';
    const { parsed } = toolResult(await callTool('send_draft', { messageId: DRAFT_ID, from: ALIAS }));
    const writes = graphCalls.filter((c) => c.method !== 'get').map((c) => `${c.method} ${c.path}`);
    expect(writes).toEqual([`patch /me/messages/${DRAFT_ID}`, `post /me/messages/${DRAFT_ID}/send`]);
    expect(parsed.from).toBe(ALIAS);
  });

  it('send_draft without from reports the From the draft carries', async () => {
    emailOutputMode = 'send';
    const { parsed } = toolResult(await callTool('send_draft', { messageId: DRAFT_ID }));
    expect(calls('patch')).toHaveLength(0);
    expect(parsed.from).toBe(PRIMARY);
  });

  it('send_draft rejects a foreign from and does not send', async () => {
    emailOutputMode = 'send';
    const { isError } = toolResult(await callTool('send_draft', { messageId: DRAFT_ID, from: 'boss@competitor.example' }));
    expect(isError).toBe(true);
    expect(calls('post')).toHaveLength(0);
  });

  it('update_message accepts from alone and PATCHes it', async () => {
    const { isError } = toolResult(await callTool('update_message', { messageId: DRAFT_ID, from: ALIAS }));
    expect(isError).toBe(false);
    expect(calls('patch')[0].body).toEqual({ from: rcpt(ALIAS) });
  });
});
