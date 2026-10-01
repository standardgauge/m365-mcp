/**
 * Attach from OneDrive / SharePoint by drive item ID.
 *
 * Drives the JSON-RPC /api/mcp tools end to end against a stubbed Graph client and asserts:
 *   - acceptance: a create_draft built from a OneDrive item carries an attachment whose SHA-256
 *     matches the bytes Graph served for that item, on both the inline (< 3 MB) path and the
 *     upload-session path;
 *   - SharePoint `{ siteId, itemId, driveId? }` reads from the right drive;
 *   - the source service's policy applies to the fetch: service enabled for the tenant and not
 *     disabled for the user, allowedSites, and the item path against that service's deny list;
 *   - folders are refused, and nothing is read when the mail-side deny check already failed;
 *   - send_mail takes the same references in both email output modes.
 */

import { jest } from '@jest/globals';
import { createHash } from 'node:crypto';
import { Readable } from 'node:stream';
import type { HttpRequest, InvocationContext } from '@azure/functions';

const mockAuthenticateRequest = jest.fn<() => Promise<unknown>>();
const mockIsPathDenied = jest.fn<(t: string, u: string, service: string, path: string) => Promise<boolean>>();
const mockGetEnabledServices = jest.fn<() => Promise<string[]>>();
const mockGetAllowedSites = jest.fn<() => Promise<Array<{ id: string; name: string }>>>();
const mockUserDisabled = jest.fn<() => string[]>();
const mockGetUserEmailSettings = jest.fn<() => Promise<{ emailOutputMode: string }>>();
const mockLogAccess = jest.fn<(entry: Record<string, unknown>) => void>();

interface DriveFile {
  name: string;
  mimeType?: string;
  parentPath: string;
  bytes?: Buffer;
  folder?: boolean;
}

// Graph stub: drive items are served by path; every POST is recorded.
const driveItems = new Map<string, DriveFile>();
const posts: Array<{ path: string; body: unknown }> = [];
const contentReads: string[] = [];

function graphApi(path: string) {
  const chain = {
    select: () => chain,
    get: async () => {
      const f = driveItems.get(path);
      if (!f) throw Object.assign(new Error('itemNotFound'), { statusCode: 404 });
      return {
        id: path.split('/').pop(),
        name: f.name,
        size: f.bytes?.length ?? 0,
        parentReference: { path: f.parentPath },
        ...(f.folder ? { folder: { childCount: 0 } } : { file: { mimeType: f.mimeType } }),
      };
    },
    getStream: async () => {
      const f = driveItems.get(path.replace(/\/content$/, ''));
      if (!f?.bytes) throw new Error('no content');
      contentReads.push(path);
      // Serve in odd-sized chunks so reassembly is actually exercised.
      const parts: Buffer[] = [];
      for (let i = 0; i < f.bytes.length; i += 7777) parts.push(f.bytes.subarray(i, i + 7777));
      return Readable.from(parts);
    },
    post: async (body: unknown) => {
      posts.push({ path, body });
      if (path.endsWith('/createUploadSession')) return { uploadUrl: 'https://upload.example/s1' };
      if (path.endsWith('/messages')) return { id: 'draft-1', webLink: 'https://outlook/draft-1' };
      return undefined;
    },
  };
  return chain;
}

jest.mock('../services/telemetry.js', () => ({}));
jest.mock('../services/authMiddleware.js', () => ({ authenticateRequest: () => mockAuthenticateRequest() }));
jest.mock('../services/tokenCache.js', () => ({
  getValidAccessTokenForSession: () => Promise.resolve('access-token'),
  getTenantIdFromSession: () => 'test-tenant',
  getTenantId: () => Promise.resolve('test-tenant'),
}));
jest.mock('../services/graphClient.js', () => ({ createGraphClient: () => ({ api: graphApi }) }));
jest.mock('../services/denyList.js', () => ({
  filterDeniedPaths: (_t: unknown, _u: unknown, _s: unknown, items: unknown) => Promise.resolve(items),
  isPathDenied: (t: string, u: string, s: string, p: string) => mockIsPathDenied(t, u, s, p),
}));
jest.mock('../services/containerResolver.js', () => ({
  resolveMailFolderName: () => Promise.resolve(null),
  resolveDefaultCalendarId: () => Promise.resolve(null),
  resolveContactParentFolder: () => Promise.resolve(null),
  resolveDefaultContactFolder: () => Promise.resolve(null),
  resolveSectionNotebook: () => Promise.resolve(null),
  normalizeContactFolderId: (_g: unknown, _u: unknown, id: string) => Promise.resolve(id),
}));
jest.mock('../services/sharepointFilter.js', () => ({ filterAndDisambiguateSites: (sites: unknown) => sites }));
jest.mock('../services/serviceSettings.js', () => ({
  getEnabledServices: () => mockGetEnabledServices(),
  getReadOnlyServices: () => Promise.resolve([]),
  getAllowedSites: () => mockGetAllowedSites(),
}));
jest.mock('../services/userServiceOverrides.js', () => ({
  getUserServiceOverrides: () => Promise.resolve(mockUserDisabled()),
  isServiceDisabledForUser: (_t: unknown, _u: unknown, service: string) => Promise.resolve(mockUserDisabled().includes(service)),
}));
jest.mock('../services/userEmailSettings.js', () => ({
  getUserEmailSettings: () => mockGetUserEmailSettings(),
  setUserEmailSettings: () => Promise.resolve(undefined),
}));
jest.mock('../services/userMailConfig.js', () => ({
  isMailIndexingDisabled: jest.fn<() => Promise<boolean>>().mockResolvedValue(false),
}));
jest.mock('../services/auditLog.js', () => ({ logAccess: (e: Record<string, unknown>) => mockLogAccess(e) }));
jest.mock('@azure/functions', () => ({ app: { http: jest.fn() } }));

import { app } from '@azure/functions';
import '../functions/mcp/mcpEndpoint.js';
import { INLINE_ATTACHMENT_LIMIT_BYTES } from '../services/mailAttachments.js';

interface HttpRegistration {
  handler: (req: HttpRequest, context: InvocationContext) => Promise<{ status: number; jsonBody?: unknown }>;
}
const httpMock = app.http as unknown as jest.Mock<(name: string, opts: HttpRegistration) => void>;
const registration = httpMock.mock.calls.find((call) => call[0] === 'mcpEndpoint');
if (!registration) throw new Error('mcpEndpoint handler was not registered');
const handler = registration[1].handler;

const SESSION = { userId: 'test-user', tenantId: 'test-tenant', email: 'u@x.com', accessToken: 't', sessionToken: 's' };

async function callTool(name: string, args: Record<string, unknown>): Promise<{ parsed: Record<string, unknown> | null; isError: boolean; text: string }> {
  const req = {
    method: 'POST',
    headers: new Map<string, string>(),
    json: () => Promise.resolve({ jsonrpc: '2.0', id: 1, method: 'tools/call', params: { name, arguments: args } }),
  } as unknown as HttpRequest;
  const ctx = { error: jest.fn(), warn: jest.fn(), log: jest.fn() } as unknown as InvocationContext;
  const res = await handler(req, ctx);
  const body = res.jsonBody as { result?: { content?: Array<{ text: string }>; isError?: boolean } };
  const text = body.result?.content?.[0]?.text ?? '';
  const isError = body.result?.isError ?? false;
  return { parsed: isError ? null : JSON.parse(text), isError, text };
}

const sha256 = (b: Buffer) => createHash('sha256').update(b).digest('hex');

/** Attachments on the draft POST, decoded back to bytes. */
function draftAttachments(): Array<{ name: string; contentType: string; bytes: Buffer }> {
  const draftPost = posts.find((p) => p.path === '/me/messages');
  const atts = ((draftPost?.body as { attachments?: Array<{ name: string; contentType: string; contentBytes: string }> })?.attachments) ?? [];
  return atts.map((a) => ({ name: a.name, contentType: a.contentType, bytes: Buffer.from(a.contentBytes, 'base64') }));
}

// A small "PDF" with every byte value, so any encoding slip changes the hash.
const PDF_BYTES = Buffer.concat([Buffer.from('%PDF-1.7\n'), Buffer.from(Array.from({ length: 4096 }, (_, i) => i % 256))]);

const realFetch = global.fetch;

beforeEach(() => {
  jest.clearAllMocks();
  driveItems.clear();
  posts.length = 0;
  contentReads.length = 0;
  mockAuthenticateRequest.mockResolvedValue({ userId: 'test-user', session: SESSION });
  mockIsPathDenied.mockResolvedValue(false);
  mockGetEnabledServices.mockResolvedValue(['mail', 'onedrive', 'sharepoint']);
  mockGetAllowedSites.mockResolvedValue([]);
  mockUserDisabled.mockReturnValue([]);
  mockGetUserEmailSettings.mockResolvedValue({ emailOutputMode: 'draft' });

  driveItems.set('/me/drive/items/OD1', {
    name: 'Proposal.pdf',
    mimeType: 'application/pdf',
    parentPath: '/drive/root:/Documents/Clients',
    bytes: PDF_BYTES,
  });
});

afterEach(() => {
  global.fetch = realFetch;
});

describe('acceptance: create_draft from a OneDrive item', () => {
  it('carries an attachment whose hash matches the Graph content', async () => {
    const { parsed, isError, text } = await callTool('create_draft', {
      subject: 'Proposal',
      to: ['client@example.com'],
      attachments: [{ driveItemId: 'OD1' }],
    });
    expect(text).not.toMatch(/^Error/);
    expect(isError).toBe(false);
    expect(parsed).toMatchObject({ id: 'draft-1', status: 'draft' });

    const [att] = draftAttachments();
    expect(att.name).toBe('Proposal.pdf');
    expect(att.contentType).toBe('application/pdf');
    expect(sha256(att.bytes)).toBe(sha256(PDF_BYTES));
    expect(contentReads).toEqual(['/me/drive/items/OD1/content']);
  });

  it('matches the hash on the upload-session path too (> 3 MB)', async () => {
    const big = Buffer.alloc(INLINE_ATTACHMENT_LIMIT_BYTES + 12345);
    for (let i = 0; i < big.length; i++) big[i] = (i * 31) % 251;
    driveItems.set('/me/drive/items/BIG', { name: 'Model.xlsx', parentPath: '/drive/root:/Documents', bytes: big });

    const received: Buffer[] = [];
    global.fetch = jest.fn(async (_url: unknown, init: unknown) => {
      received.push(Buffer.from((init as { body: Uint8Array }).body));
      return { ok: true, status: 200, text: async () => '' } as unknown as Response;
    }) as unknown as typeof fetch;

    const { isError } = await callTool('create_draft', { subject: 'S', to: ['a@x.com'], attachments: [{ driveItemId: 'BIG' }] });
    expect(isError).toBe(false);

    // Too big to ride inline: the draft POST has no attachments and the bytes went via the session.
    expect(draftAttachments()).toHaveLength(0);
    const session = posts.find((p) => p.path === '/me/messages/draft-1/attachments/createUploadSession');
    expect(session?.body).toMatchObject({ AttachmentItem: { name: 'Model.xlsx', size: big.length } });
    expect(sha256(Buffer.concat(received))).toBe(sha256(big));
  });

  it('honors name and contentType overrides and mixes with inline attachments in order', async () => {
    const { isError } = await callTool('create_draft', {
      subject: 'S',
      to: ['a@x.com'],
      attachments: [
        { name: 'note.txt', contentType: 'text/plain', content: Buffer.from('hi').toString('base64') },
        { driveItemId: 'OD1', name: 'Acme proposal.pdf', contentType: 'application/x-custom' },
      ],
    });
    expect(isError).toBe(false);
    const atts = draftAttachments();
    expect(atts.map((a) => a.name)).toEqual(['note.txt', 'Acme proposal.pdf']);
    expect(atts[1].contentType).toBe('application/x-custom');
    expect(sha256(atts[1].bytes)).toBe(sha256(PDF_BYTES));
  });

  it('audits the fetched file with its drive path', async () => {
    await callTool('create_draft', { subject: 'S', to: ['a@x.com'], attachments: [{ driveItemId: 'OD1' }] });
    expect(mockLogAccess).toHaveBeenCalledWith(
      expect.objectContaining({
        operation: 'create_draft.attachment',
        resource: 'onedrive:/drive/root:/Documents/Clients/Proposal.pdf',
        result: 'allowed',
        source: 'mcp',
      }),
    );
  });
});

describe('SharePoint { siteId, itemId }', () => {
  const SITE = 'contoso.sharepoint.com,g1,g2';

  beforeEach(() => {
    driveItems.set(`/sites/${SITE}/drive/items/SP1`, { name: 'Deck.pptx', mimeType: 'application/vnd.ms-powerpoint', parentPath: '/drives/b!d/root:/Shared', bytes: Buffer.from('deck-bytes') });
    driveItems.set(`/sites/${SITE}/drives/LIB2/items/SP2`, { name: 'Other.docx', parentPath: '/drives/LIB2/root:', bytes: Buffer.from('other-bytes') });
  });

  it('reads from the site default drive, or the named library when driveId is given', async () => {
    const { isError } = await callTool('create_draft', {
      subject: 'S',
      to: ['a@x.com'],
      attachments: [{ siteId: SITE, itemId: 'SP1' }, { siteId: SITE, itemId: 'SP2', driveId: 'LIB2' }],
    });
    expect(isError).toBe(false);
    const atts = draftAttachments();
    expect(atts.map((a) => a.bytes.toString())).toEqual(['deck-bytes', 'other-bytes']);
    expect(mockIsPathDenied).toHaveBeenCalledWith('test-tenant', 'test-user', 'sharepoint', '/drives/b!d/root:/Shared/Deck.pptx');
  });

  it('refuses a site outside allowedSites without reading it', async () => {
    mockGetAllowedSites.mockResolvedValue([{ id: 'other-site', name: 'Other' }]);
    const { isError, text } = await callTool('create_draft', { subject: 'S', to: ['a@x.com'], attachments: [{ siteId: SITE, itemId: 'SP1' }] });
    expect(isError).toBe(true);
    expect(text).toMatch(/attachments\[0\]: Site ".*" is not in the allowed sites list/);
    expect(contentReads).toHaveLength(0);
    expect(posts).toHaveLength(0);
  });

  it('refuses when SharePoint is not enabled for the tenant', async () => {
    mockGetEnabledServices.mockResolvedValue(['mail', 'onedrive']);
    const { isError, text } = await callTool('create_draft', { subject: 'S', to: ['a@x.com'], attachments: [{ siteId: SITE, itemId: 'SP1' }] });
    expect(isError).toBe(true);
    expect(text).toMatch(/Service "sharepoint" is not enabled/);
    expect(posts).toHaveLength(0);
  });
});

describe('source policy applies to the fetch', () => {
  it('refuses a deny-listed OneDrive path and sends nothing', async () => {
    mockIsPathDenied.mockImplementation(async (_t, _u, service, path) => service === 'onedrive' && path.includes('/Clients/'));
    const { isError, text } = await callTool('create_draft', { subject: 'S', to: ['a@x.com'], attachments: [{ driveItemId: 'OD1' }] });
    expect(isError).toBe(true);
    expect(text).toMatch(/attachments\[0\]: Access restricted by deny list/);
    expect(contentReads).toHaveLength(0);
    expect(posts).toHaveLength(0);
    expect(mockLogAccess).toHaveBeenCalledWith(expect.objectContaining({ result: 'denied', resource: 'onedrive:/drive/root:/Documents/Clients/Proposal.pdf' }));
  });

  it('refuses when OneDrive is disabled for the user, even though mail is enabled', async () => {
    mockUserDisabled.mockReturnValue(['onedrive']);
    const { isError, text } = await callTool('create_draft', { subject: 'S', to: ['a@x.com'], attachments: [{ driveItemId: 'OD1' }] });
    expect(isError).toBe(true);
    expect(text).toMatch(/Service "onedrive" is disabled for your account/);
    expect(posts).toHaveLength(0);
  });

  it('refuses a folder', async () => {
    driveItems.set('/me/drive/items/F1', { name: 'Clients', parentPath: '/drive/root:', folder: true });
    const { isError, text } = await callTool('create_draft', { subject: 'S', to: ['a@x.com'], attachments: [{ driveItemId: 'F1' }] });
    expect(isError).toBe(true);
    expect(text).toMatch(/"Clients" is a folder, not a file/);
    expect(posts).toHaveLength(0);
  });

  it('reads nothing when the Drafts folder is deny-listed', async () => {
    mockIsPathDenied.mockImplementation(async (_t, _u, service, path) => service === 'mail' && path === 'Drafts');
    const { isError } = await callTool('create_draft', { subject: 'S', to: ['a@x.com'], attachments: [{ driveItemId: 'OD1' }] });
    expect(isError).toBe(true);
    expect(contentReads).toHaveLength(0);
  });

  it('rejects a path-altering driveItemId before any Graph call', async () => {
    const { isError, text } = await callTool('create_draft', { subject: 'S', to: ['a@x.com'], attachments: [{ driveItemId: '../../users/ceo/drive/items/x' }] });
    expect(isError).toBe(true);
    expect(text).toMatch(/driveItemId contains characters/);
    expect(posts).toHaveLength(0);
  });

  it('rejects an undeclared field on an attachment item', async () => {
    const { isError, text } = await callTool('create_draft', { subject: 'S', to: ['a@x.com'], attachments: [{ driveItemId: 'OD1', path: '/etc/passwd' }] });
    expect(isError).toBe(true);
    expect(text).toMatch(/Unsupported parameter\(s\) for create_draft: attachments\[0\]\.path/);
    expect(contentReads).toHaveLength(0);
  });
});

describe('send_mail takes the same references', () => {
  it('draft mode: the queued draft carries the OneDrive file', async () => {
    const { parsed, isError } = await callTool('send_mail', { subject: 'S', to: ['a@x.com'], attachments: [{ driveItemId: 'OD1' }] });
    expect(isError).toBe(false);
    expect(parsed).toMatchObject({ status: 'queued_as_draft', draftId: 'draft-1' });
    expect(sha256(draftAttachments()[0].bytes)).toBe(sha256(PDF_BYTES));
  });

  it('send mode: the sent message carries the OneDrive file', async () => {
    mockGetUserEmailSettings.mockResolvedValue({ emailOutputMode: 'send' });
    const { parsed, isError } = await callTool('send_mail', { subject: 'S', to: ['a@x.com'], attachments: [{ driveItemId: 'OD1' }] });
    expect(isError).toBe(false);
    expect(parsed).toMatchObject({ status: 'sent' });
    const sent = posts.find((p) => p.path === '/me/sendMail')?.body as { message: { attachments: Array<{ contentBytes: string }> } };
    expect(sha256(Buffer.from(sent.message.attachments[0].contentBytes, 'base64'))).toBe(sha256(PDF_BYTES));
  });

  it('send mode: reads nothing when Sent Items is deny-listed', async () => {
    mockGetUserEmailSettings.mockResolvedValue({ emailOutputMode: 'send' });
    mockIsPathDenied.mockImplementation(async (_t, _u, service, path) => service === 'mail' && path === 'Sent Items');
    const { isError } = await callTool('send_mail', { subject: 'S', to: ['a@x.com'], attachments: [{ driveItemId: 'OD1' }] });
    expect(isError).toBe(true);
    expect(contentReads).toHaveLength(0);
    expect(posts).toHaveLength(0);
  });
});
