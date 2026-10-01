/**
 * Contacts surface tests.
 *
 * Exercises the expanded contacts surface over the MCP JSON-RPC endpoint
 * (/api/mcp), the surface a bulk import (e.g. the Convent & Stuart Hall family
 * directory) actually drives:
 *   - create_contact / update_contact carry notes, category tags, all three
 *     postal addresses, and the secondary fields, and round-trip through
 *     search_contacts.
 *   - list_contact_folders + create_contact_folder expose folder ids.
 *   - create_contacts_batch creates many contacts via Graph $batch, chunked at
 *     the 20-op limit, and reports per-entry success/failure.
 *
 * Graph is mocked at the client boundary; assertions look at the request bodies
 * and paths the handlers build.
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

interface GraphCall { path: string; method: string; body?: unknown }
const graphCalls: GraphCall[] = [];

// Test-tunable Graph GET payloads.
let contactFoldersValue: unknown[] = [];
let contactsGetValue: unknown[] = [];

function makeChain(path: string) {
  const chain: Record<string, unknown> = {};
  for (const m of ['select', 'top', 'orderby', 'filter', 'expand', 'header']) chain[m] = () => chain;
  chain.get = () => {
    graphCalls.push({ path, method: 'GET' });
    if (path === '/me/contactFolders') return Promise.resolve({ value: contactFoldersValue });
    return Promise.resolve({ value: contactsGetValue });
  };
  chain.post = (body?: unknown) => {
    graphCalls.push({ path, method: 'POST', body });
    if (path === '/$batch') {
      const requests = ((body as { requests?: Array<{ id: string }> })?.requests) ?? [];
      return Promise.resolve({
        responses: requests.map((r) => ({ id: r.id, status: 201, body: { id: `new-${r.id}`, displayName: 'Created' } })),
      });
    }
    const b = (body ?? {}) as Record<string, unknown>;
    return Promise.resolve({ id: 'new-id', displayName: b.displayName ?? 'New', parentFolderId: b.parentFolderId ?? null });
  };
  chain.patch = (body?: unknown) => {
    graphCalls.push({ path, method: 'PATCH', body });
    return Promise.resolve({ id: 'cid', displayName: 'Updated' });
  };
  chain.delete = () => {
    graphCalls.push({ path, method: 'DELETE' });
    return Promise.resolve(undefined);
  };
  return chain;
}
const mockCreateGraphClient = jest.fn(() => ({ api: (path: string) => makeChain(path) }));

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
  resolveMailFolderName: () => Promise.resolve(null),
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

beforeEach(() => {
  jest.clearAllMocks();
  graphCalls.length = 0;
  contactFoldersValue = [];
  contactsGetValue = [];
  mockAuthenticateRequest.mockResolvedValue(AUTH);
  mockGetValidAccessTokenForSession.mockResolvedValue('access-token');
  mockGetTenantIdFromSession.mockReturnValue(TENANT);
  mockIsPathDenied.mockResolvedValue(false);
  mockGetEnabledServices.mockResolvedValue(['contacts', 'mail', 'calendar']);
  mockGetAllowedSites.mockResolvedValue([]);
  mockGetUserServiceOverrides.mockResolvedValue([]);
});

const FULL_CONTACT = {
  givenName: 'Jane',
  surname: 'Doe',
  middleName: 'Q',
  nickName: 'Janie',
  emailAddresses: ['jane@example.com', 'jdoe@work.com'],
  businessPhones: ['+1-555-0100'],
  homePhones: ['+1-555-0111'],
  mobilePhone: '+1-555-0199',
  companyName: 'Acme',
  jobTitle: 'Engineer',
  personalNotes: 'Met at the 26-27 orientation. Two kids in 3rd and 5th.',
  categories: ['CSH 26-27', 'Parent'],
  birthday: '1980-04-01T00:00:00Z',
  spouseName: 'John Doe',
  homeAddress: { street: '1 Home St', city: 'SF', state: 'CA', postalCode: '94110', countryOrRegion: 'USA' },
  businessAddress: { street: '2 Work Ave', city: 'SF', state: 'CA', postalCode: '94105' },
  otherAddress: { city: 'Oakland', state: 'CA' },
};

// ─────────────────────────────────────────────────────────────────────────────
// create_contact — full field set reaches Graph
// ─────────────────────────────────────────────────────────────────────────────

describe('create_contact — expanded fields', () => {
  it('sends notes, categories, all three addresses, and secondary fields to Graph', async () => {
    const res = await callTool('create_contact', { ...FULL_CONTACT });
    const { parsed, isError } = toolResult(res);
    expect(isError).toBe(false);
    expect((parsed as { status: string }).status).toBe('created');

    const post = graphCalls.find((c) => c.method === 'POST' && c.path === '/me/contacts');
    expect(post).toBeDefined();
    const b = post!.body as Record<string, unknown>;
    expect(b.personalNotes).toBe(FULL_CONTACT.personalNotes);
    expect(b.categories).toEqual(['CSH 26-27', 'Parent']);
    expect(b.middleName).toBe('Q');
    expect(b.nickName).toBe('Janie');
    expect(b.homePhones).toEqual(['+1-555-0111']);
    expect(b.spouseName).toBe('John Doe');
    expect(b.birthday).toBe('1980-04-01T00:00:00Z');
    expect(b.homeAddress).toEqual(FULL_CONTACT.homeAddress);
    expect(b.businessAddress).toEqual(FULL_CONTACT.businessAddress);
    expect(b.otherAddress).toEqual({ city: 'Oakland', state: 'CA' });
    // emailAddresses expanded to Graph {address,name} shape
    expect(b.emailAddresses).toEqual([
      { address: 'jane@example.com', name: 'jane@example.com' },
      { address: 'jdoe@work.com', name: 'jdoe@work.com' },
    ]);
  });

  it('routes to a folder when folderId is given, and treats contacts-root as the default folder', async () => {
    await callTool('create_contact', { givenName: 'A', folderId: 'folder-xyz' });
    expect(graphCalls.some((c) => c.path === '/me/contactFolders/folder-xyz/contacts')).toBe(true);

    graphCalls.length = 0;
    await callTool('create_contact', { givenName: 'B', folderId: 'contacts-root' });
    expect(graphCalls.some((c) => c.path === '/me/contacts')).toBe(true);
  });
});

// ─────────────────────────────────────────────────────────────────────────────
// update_contact — expanded fields, only-provided semantics
// ─────────────────────────────────────────────────────────────────────────────

describe('update_contact — expanded fields', () => {
  it('patches notes/categories/addresses and omits fields not provided', async () => {
    const res = await callTool('update_contact', {
      contactId: 'contact-1',
      personalNotes: 'Updated note',
      categories: ['VIP'],
      otherAddress: { city: 'Berkeley' },
    });
    expect(toolResult(res).isError).toBe(false);
    const patch = graphCalls.find((c) => c.method === 'PATCH');
    expect(patch?.path).toBe('/me/contacts/contact-1');
    const b = patch!.body as Record<string, unknown>;
    expect(b.personalNotes).toBe('Updated note');
    expect(b.categories).toEqual(['VIP']);
    expect(b.otherAddress).toEqual({ city: 'Berkeley' });
    // fields not supplied must not appear in the PATCH
    expect(b).not.toHaveProperty('givenName');
    expect(b).not.toHaveProperty('homeAddress');
  });
});

// ─────────────────────────────────────────────────────────────────────────────
// search_contacts — round-trips the full field set back out
// ─────────────────────────────────────────────────────────────────────────────

describe('search_contacts — round-trip', () => {
  it('returns notes, categories, and all three addresses', async () => {
    // Graph stores emailAddresses as {address,name} objects on read-back.
    contactsGetValue = [{
      id: 'c1',
      displayName: 'Jane Doe',
      parentFolderId: 'root',
      ...FULL_CONTACT,
      emailAddresses: [
        { address: 'jane@example.com', name: 'Jane' },
        { address: 'jdoe@work.com', name: 'Jane Doe' },
      ],
    }];
    const res = await callTool('search_contacts', { q: 'Jane' });
    const { parsed, isError } = toolResult(res);
    expect(isError).toBe(false);
    const contacts = (parsed as { contacts: Array<Record<string, unknown>> }).contacts;
    expect(contacts).toHaveLength(1);
    const c = contacts[0];
    expect(c.personalNotes).toBe(FULL_CONTACT.personalNotes);
    expect(c.categories).toEqual(['CSH 26-27', 'Parent']);
    expect(c.homeAddress).toEqual(FULL_CONTACT.homeAddress);
    expect(c.businessAddress).toEqual(FULL_CONTACT.businessAddress);
    expect(c.otherAddress).toEqual(FULL_CONTACT.otherAddress);
    expect(c.middleName).toBe('Q');
    expect(c.spouseName).toBe('John Doe');
    expect(c.emails).toEqual(['jane@example.com', 'jdoe@work.com']);
    // internal folder id must not leak
    expect(c).not.toHaveProperty('parentFolderId');
  });

  it('maps the synthetic contacts-root folderId to /me/contacts, not /me/contactFolders/contacts-root/contacts', async () => {
    contactsGetValue = [{ id: 'c1', displayName: 'Jane Doe', parentFolderId: 'root' }];
    const res = await callTool('search_contacts', { folderId: 'contacts-root' });
    expect(toolResult(res).isError).toBe(false);
    const gets = graphCalls.filter((c) => c.method === 'GET');
    expect(gets.some((c) => c.path === '/me/contacts')).toBe(true);
    expect(gets.some((c) => c.path === '/me/contactFolders/contacts-root/contacts')).toBe(false);
  });

  it('still routes an explicit non-root folderId to /me/contactFolders/{id}/contacts', async () => {
    contactsGetValue = [];
    const res = await callTool('search_contacts', { folderId: 'folder-xyz' });
    expect(toolResult(res).isError).toBe(false);
    const gets = graphCalls.filter((c) => c.method === 'GET');
    expect(gets.some((c) => c.path === '/me/contactFolders/folder-xyz/contacts')).toBe(true);
  });

  // Deny-list parity: routing contacts-root through contactsApiPath must not let it
  // slip past the up-front deny check. The default folder's deny key IS the synthetic
  // 'contacts-root' string (what resolveDefaultContactFolder/normalizeContactFolderId
  // return), so a deny-listed default folder must block search by its advertised id —
  // and must not reach Graph (no fail-open to /me/contacts).
  it('blocks search_contacts on a deny-listed contacts-root without hitting Graph', async () => {
    contactsGetValue = [{ id: 'c1', displayName: 'Jane Doe', parentFolderId: 'root' }];
    mockIsPathDenied.mockImplementation((_t, _u, _type, path) =>
      Promise.resolve(path === 'contacts-root'),
    );
    const res = await callTool('search_contacts', { folderId: 'contacts-root' });
    const r = toolResult(res);
    expect(r.isError).toBe(true);
    expect(r.text).toMatch(/deny list/i);
    // deny check keyed on the synthetic id, and no contacts collection was read
    expect(mockIsPathDenied).toHaveBeenCalledWith(TENANT, expect.any(String), 'contacts', 'contacts-root');
    const gets = graphCalls.filter((c) => c.method === 'GET');
    expect(gets.some((c) => c.path === '/me/contacts')).toBe(false);
  });
});

// ─────────────────────────────────────────────────────────────────────────────
// list_contact_folders + create_contact_folder
// ─────────────────────────────────────────────────────────────────────────────

describe('contact folder tools', () => {
  it('list_contact_folders returns children plus the synthetic default-folder id', async () => {
    contactFoldersValue = [{ id: 'f1', displayName: 'Family', parentFolderId: null }];
    contactsGetValue = [{ id: 'c1' }]; // root has contacts → synthetic root prepended
    const res = await callTool('list_contact_folders', {});
    const { parsed, isError } = toolResult(res);
    expect(isError).toBe(false);
    const out = parsed as { folders: Array<{ id: string; name: string }>; count: number };
    const ids = out.folders.map((f) => f.id);
    expect(ids).toContain('contacts-root');
    expect(ids).toContain('f1');
    expect(out.count).toBe(out.folders.length);
  });

  it('create_contact_folder posts a top-level folder', async () => {
    const res = await callTool('create_contact_folder', { displayName: 'CSH 26-27' });
    const { parsed, isError } = toolResult(res);
    expect(isError).toBe(false);
    expect((parsed as { status: string }).status).toBe('created');
    const post = graphCalls.find((c) => c.method === 'POST');
    expect(post?.path).toBe('/me/contactFolders');
    expect((post?.body as { displayName: string }).displayName).toBe('CSH 26-27');
  });

  it('create_contact_folder posts a child folder under a parent', async () => {
    await callTool('create_contact_folder', { displayName: 'Grade 3', parentFolderId: 'parent-1' });
    expect(graphCalls.some((c) => c.path === '/me/contactFolders/parent-1/childFolders')).toBe(true);
  });
});

// ─────────────────────────────────────────────────────────────────────────────
// create_contacts_batch — $batch, chunking, per-entry outcomes
// ─────────────────────────────────────────────────────────────────────────────

describe('create_contacts_batch', () => {
  it('creates every valid contact via a single $batch call', async () => {
    const contacts = [
      { givenName: 'A', personalNotes: 'n1', categories: ['x'] },
      { givenName: 'B', mobilePhone: '555' },
    ];
    const res = await callTool('create_contacts_batch', { contacts });
    const { parsed, isError } = toolResult(res);
    expect(isError).toBe(false);
    const out = parsed as { created: number; failed: number; results: Array<{ status: string }> };
    expect(out.created).toBe(2);
    expect(out.failed).toBe(0);
    const batchPosts = graphCalls.filter((c) => c.path === '/$batch');
    expect(batchPosts).toHaveLength(1);
    const requests = (batchPosts[0].body as { requests: Array<{ url: string; body: Record<string, unknown> }> }).requests;
    expect(requests[0].url).toBe('/me/contacts');
    expect(requests[0].body.personalNotes).toBe('n1');
  });

  it('chunks at 20 operations per $batch request', async () => {
    const contacts = Array.from({ length: 25 }, (_, i) => ({ givenName: `C${i}` }));
    const res = await callTool('create_contacts_batch', { contacts });
    const out = toolResult(res).parsed as { created: number };
    expect(out.created).toBe(25);
    const batchPosts = graphCalls.filter((c) => c.path === '/$batch');
    expect(batchPosts).toHaveLength(2);
    expect((batchPosts[0].body as { requests: unknown[] }).requests).toHaveLength(20);
    expect((batchPosts[1].body as { requests: unknown[] }).requests).toHaveLength(5);
  });

  it('rejects entries missing givenName client-side without a Graph call', async () => {
    const contacts = [{ givenName: 'Good' }, { surname: 'NoGiven' }];
    const res = await callTool('create_contacts_batch', { contacts });
    const out = toolResult(res).parsed as { created: number; failed: number; results: Array<{ index: number; status: string; error?: string }> };
    expect(out.created).toBe(1);
    expect(out.failed).toBe(1);
    const bad = out.results.find((r) => r.index === 1);
    expect(bad?.status).toBe('failed');
    expect(bad?.error).toMatch(/givenName/);
    // the single valid entry still goes out in one batch
    expect((graphCalls.find((c) => c.path === '/$batch')!.body as { requests: unknown[] }).requests).toHaveLength(1);
  });

  it('errors when contacts is empty', async () => {
    const res = await callTool('create_contacts_batch', { contacts: [] });
    expect(toolResult(res).isError).toBe(true);
  });
});

// ─────────────────────────────────────────────────────────────────────────────
// tools/list exposure — gated by the contacts enabledService
// ─────────────────────────────────────────────────────────────────────────────

describe('tools/list — contacts exposure', () => {
  const NEW_TOOLS = ['list_contact_folders', 'create_contact_folder', 'create_contacts_batch'];

  it('exposes the new contact tools when contacts is enabled', async () => {
    const res = await rpc('tools/list', {});
    const body = res.jsonBody as { result: { tools: Array<{ name: string }> } };
    const names = body.result.tools.map((t) => t.name);
    for (const t of NEW_TOOLS) expect(names).toContain(t);
  });

  it('hides all contact tools when contacts is not enabled', async () => {
    mockGetEnabledServices.mockResolvedValue(['mail']);
    const res = await rpc('tools/list', {});
    const body = res.jsonBody as { result: { tools: Array<{ name: string }> } };
    const names = body.result.tools.map((t) => t.name);
    for (const t of [...NEW_TOOLS, 'create_contact', 'search_contacts']) expect(names).not.toContain(t);
  });
});
