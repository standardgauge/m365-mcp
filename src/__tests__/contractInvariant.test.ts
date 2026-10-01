/**
 * Connector-invariant contract tests.
 *
 * These assert the two systemic guarantees end-to-end through the native MCP
 * dispatch, not per-verb behaviour:
 *
 *   1. Silent input-dropping is impossible. Every tool the endpoint advertises
 *      rejects a parameter it does not declare, rather than accepting it,
 *      ignoring it, and returning success. Driving the assertion off the live
 *      tools/list output means a tool added later is covered automatically.
 *
 *   2. Silent truncation is impossible. Every list/search verb reports a numeric
 *      `limit` and a boolean `truncated`, and `truncated` is true exactly when
 *      the underlying collection extends past the returned page (Graph
 *      `@odata.nextLink`, or the search API's `moreResultsAvailable`).
 *
 * The Graph client is mocked so no network is touched; the collection it returns
 * — and whether it carries a continuation token — is set per-test.
 */
import { jest } from '@jest/globals';
import type { HttpRequest, InvocationContext } from '@azure/functions';

// Per-test control over what the mocked Graph returns.
let collection: { value: unknown[]; nextLink: boolean } = { value: [], nextLink: false };
let searchMore = false;
const graphCalls: Array<{ path: string; verb: string }> = [];

function collectionResponse(): Record<string, unknown> {
  const out: Record<string, unknown> = { value: collection.value };
  if (collection.nextLink) out['@odata.nextLink'] = 'https://graph.microsoft.com/v1.0/next-page';
  return out;
}

function searchResponse(): Record<string, unknown> {
  return {
    value: [
      {
        hitsContainers: [
          {
            hits: collection.value.map((r) => ({ resource: r })),
            moreResultsAvailable: searchMore,
          },
        ],
      },
    ],
  };
}

function makeChain(path: string) {
  const chain: Record<string, unknown> = {};
  for (const m of ['select', 'top', 'orderby', 'filter', 'expand', 'header', 'search']) chain[m] = () => chain;
  chain.get = () => {
    graphCalls.push({ path, verb: 'get' });
    return Promise.resolve(collectionResponse());
  };
  chain.post = () => {
    graphCalls.push({ path, verb: 'post' });
    return Promise.resolve(path.includes('/search/query') ? searchResponse() : { id: 'x' });
  };
  chain.patch = () => Promise.resolve({ id: 'x' });
  chain.delete = () => Promise.resolve(undefined);
  return chain;
}

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
  createGraphClient: () => ({ api: (path: string) => makeChain(path) }),
}));
jest.mock('../services/denyList.js', () => ({
  filterDeniedPaths: (_t: unknown, _u: unknown, _s: unknown, items: unknown) => Promise.resolve(items),
  filterDeniedSearchHits: (_t: unknown, _u: unknown, items: unknown) => Promise.resolve(items),
  isPathDenied: () => Promise.resolve(false),
  // sharepointSearch.siteRelativePathFromWebUrl canonicalizes the derived path.
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
  getEnabledServices: () => Promise.resolve(['mail', 'sharepoint', 'onedrive', 'calendar', 'onenote', 'contacts', 'teams']),
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
import '../functions/mcp/mcpEndpoint.js';

interface HttpRegistration {
  handler: (req: HttpRequest, context: InvocationContext) => Promise<{ status: number; jsonBody?: unknown }>;
}
const httpMock = app.http as unknown as jest.Mock<(name: string, opts: HttpRegistration) => void>;
const handler = httpMock.mock.calls.find((call) => call[0] === 'mcpEndpoint')![1].handler;

const AUTH = { userId: 'u', session: { userId: 'u', tenantId: 't', email: 'u@x.com', accessToken: 'fake', sessionToken: 'sess' } };

function rpc(method: string, params: Record<string, unknown>) {
  const req = {
    method: 'POST',
    headers: new Map<string, string>(),
    json: () => Promise.resolve({ jsonrpc: '2.0', id: 1, method, params }),
  } as unknown as HttpRequest;
  const ctx = { error: jest.fn(), warn: jest.fn(), log: jest.fn() } as unknown as InvocationContext;
  return handler(req, ctx);
}

function toolCall(name: string, args: Record<string, unknown>) {
  return rpc('tools/call', { name, arguments: args });
}

function result(res: { jsonBody?: unknown }): { parsed: Record<string, unknown> | null; isError: boolean; text: string } {
  const body = res.jsonBody as { result?: { content?: Array<{ text: string }>; isError?: boolean } };
  const text = body.result?.content?.[0]?.text ?? '{}';
  const isError = body.result?.isError ?? false;
  return { parsed: isError ? null : (JSON.parse(text) as Record<string, unknown>), isError, text };
}

async function listToolNames(): Promise<string[]> {
  const res = await rpc('tools/list', {});
  const body = (res.jsonBody as { result?: { tools?: Array<{ name: string }> } }).result;
  return (body?.tools ?? []).map((t) => t.name);
}

beforeEach(() => {
  jest.clearAllMocks();
  graphCalls.length = 0;
  collection = { value: [], nextLink: false };
  searchMore = false;
  mockAuthenticateRequest.mockResolvedValue(AUTH);
});

describe('Invariant 1 — no tool silently ignores an unsupported parameter', () => {
  it('every advertised tool rejects an undeclared argument', async () => {
    const names = await listToolNames();
    expect(names.length).toBeGreaterThan(30); // sanity: the surface is enumerated

    const accepted: string[] = [];
    for (const name of names) {
      const res = await toolCall(name, { __contract_unknown_arg__: 'x' });
      const { isError, text } = result(res);
      if (!(isError && /Unsupported parameter/.test(text))) accepted.push(`${name}: ${text.slice(0, 80)}`);
    }
    expect(accepted).toEqual([]);
  });

  it('a genuinely-declared argument is NOT rejected as unsupported', async () => {
    // list_events declares maxResults; passing it must not trip the guard.
    const { text } = result(await toolCall('list_events', { maxResults: 5 }));
    expect(text).not.toMatch(/Unsupported parameter/);
  });

  it('rejects an undeclared field nested inside a declared object (contact address subfield)', async () => {
    const { isError, text } = result(
      await toolCall('create_contact', { givenName: 'A', homeAddress: { street: '1', building: 'HQ' } }),
    );
    expect(isError).toBe(true);
    expect(text).toMatch(/Unsupported parameter/);
    expect(text).toContain('homeAddress.building');
  });

  it('rejects an undeclared field inside a declared array item (attachments)', async () => {
    const { isError, text } = result(
      await toolCall('send_mail', {
        subject: 's',
        body: 'b',
        to: ['x@y.com'],
        attachments: [{ name: 'a', content: 'x', retentionLabel: 'z' }],
      }),
    );
    expect(isError).toBe(true);
    expect(text).toMatch(/Unsupported parameter/);
    expect(text).toContain('attachments[0].retentionLabel');
  });

  it('rejects an undeclared field inside create_contacts_batch.contacts[]', async () => {
    const { isError, text } = result(
      await toolCall('create_contacts_batch', { contacts: [{ givenName: 'A', assistantName: 'B' }] }),
    );
    expect(isError).toBe(true);
    expect(text).toMatch(/Unsupported parameter/);
    expect(text).toContain('contacts[0].assistantName');
  });

  it('accepts fully-declared nested input (address subfields + attachment item)', async () => {
    const contactRes = result(
      await toolCall('create_contact', { givenName: 'A', homeAddress: { street: '1', city: 'SF' } }),
    );
    expect(contactRes.text).not.toMatch(/Unsupported parameter/);

    const mailRes = result(
      await toolCall('send_mail', {
        subject: 's',
        body: 'b',
        to: ['x@y.com'],
        attachments: [{ name: 'a', contentType: 'text/plain', content: 'x' }],
      }),
    );
    expect(mailRes.text).not.toMatch(/Unsupported parameter/);
  });
});

describe('Invariant 2 — list/search verbs surface truncation', () => {
  // (tool, args, key holding the rows). Each is exercised with and without a
  // server "more results" signal.
  const CASES: Array<{ tool: string; args: Record<string, unknown>; rowsKey: string }> = [
    { tool: 'search_mail', args: { q: 'x' }, rowsKey: 'results' },
    { tool: 'list_messages', args: {}, rowsKey: 'items' },
    { tool: 'search_contacts', args: { q: 'x' }, rowsKey: 'contacts' },
    { tool: 'search_sharepoint', args: { q: 'x' }, rowsKey: 'results' },
    { tool: 'list_events', args: {}, rowsKey: 'items' },
    { tool: 'list_calendars', args: {}, rowsKey: 'items' },
    { tool: 'list_sites', args: {}, rowsKey: 'items' },
    { tool: 'list_teams', args: {}, rowsKey: 'items' },
    { tool: 'list_folders_mail', args: {}, rowsKey: 'items' },
    { tool: 'list_onedrive', args: {}, rowsKey: 'items' },
    // get_attachments in list mode (attachmentId omitted) is a Graph collection
    // GET like the rest; the single-attachment download branch is not a list.
    { tool: 'get_attachments', args: { messageId: 'm1' }, rowsKey: 'attachments' },
  ];

  it.each(CASES)('$tool reports numeric limit + boolean truncated', async ({ tool, args, rowsKey }) => {
    collection = { value: [{ id: 'a', name: 'a', displayName: 'a', webUrl: 'https://example.sharepoint.com/x' }], nextLink: false };
    const { parsed, isError } = result(await toolCall(tool, args));
    expect(isError).toBe(false);
    expect(parsed).not.toBeNull();
    expect(typeof parsed!.limit).toBe('number');
    expect(typeof parsed!.truncated).toBe('boolean');
    expect(parsed!.truncated).toBe(false);
    expect(Array.isArray(parsed![rowsKey])).toBe(true);
  });

  it('search_mail reports truncated:true when Graph returns a nextLink', async () => {
    collection = { value: [{ id: 'm1' }], nextLink: true };
    const { parsed } = result(await toolCall('search_mail', { q: 'x' }));
    expect(parsed!.truncated).toBe(true);
  });

  it('list_calendars reports truncated:true when Graph returns a nextLink', async () => {
    collection = { value: [{ id: 'c1', name: 'A' }], nextLink: true };
    const { parsed } = result(await toolCall('list_calendars', {}));
    expect(parsed!.truncated).toBe(true);
  });

  it('get_attachments (list mode) reports truncated:true when Graph returns a nextLink', async () => {
    collection = { value: [{ id: 'att1', name: 'a.pdf', contentType: 'application/pdf', size: 10 }], nextLink: true };
    const { parsed } = result(await toolCall('get_attachments', { messageId: 'm1' }));
    expect(parsed!.truncated).toBe(true);
    expect(Array.isArray(parsed!.attachments)).toBe(true);
  });

  it('search_sharepoint reports truncated:true when the search index has more', async () => {
    collection = {
      value: [{ id: 'f1', name: 'x.docx', webUrl: 'https://example.sharepoint.com/Shared%20Documents/x.docx', parentReference: { siteId: 's' } }],
      nextLink: false,
    };
    searchMore = true;
    const { parsed } = result(await toolCall('search_sharepoint', { q: 'x' }));
    expect(parsed!.truncated).toBe(true);
  });

  it('echoes the applied cap: an over-cap maxResults comes back clamped in limit', async () => {
    collection = { value: [{ id: 'm1' }], nextLink: false };
    const { parsed } = result(await toolCall('search_mail', { q: 'x', maxResults: 9999 }));
    expect(parsed!.limit).toBe(100); // search_mail hard cap
  });
});
