/**
 * Drive-item attachments on the REST mail routes, so the HTTP and MCP paths can't drift.
 *
 * Covers POST /api/mail/drafts and POST /api/mail/send with `{ driveItemId }`:
 *   - the draft carries the OneDrive bytes, hash for hash;
 *   - a policy refusal on the source file is a 403 and a folder is a 400, not a 500;
 *   - saveToSentItems:false is still refused when a drive item turns out to need an
 *     upload session, which is only knowable after the item is fetched.
 */

import { jest } from '@jest/globals';
import { createHash } from 'node:crypto';
import { Readable } from 'node:stream';
import type { HttpRequest, InvocationContext } from '@azure/functions';
import type { AuthResult } from '../services/authMiddleware.js';

const mockIsPathDenied = jest.fn<(t: string, u: string, service: string, path: string) => Promise<boolean>>();
const mockCheckServiceEnabled = jest.fn<(u: string, service: string) => Promise<null | { status: number; error: string }>>();
const mockGetUserEmailSettings = jest.fn<() => Promise<{ emailOutputMode: string }>>();

interface DriveFile { name: string; parentPath: string; bytes?: Buffer; folder?: boolean }
const driveItems = new Map<string, DriveFile>();
const posts: Array<{ path: string; body: unknown }> = [];

function graphApi(path: string) {
  const chain = {
    select: () => chain,
    get: async () => {
      const f = driveItems.get(path)!;
      return {
        name: f.name,
        size: f.bytes?.length ?? 0,
        parentReference: { path: f.parentPath },
        ...(f.folder ? { folder: {} } : { file: { mimeType: 'application/pdf' } }),
      };
    },
    getStream: async () => Readable.from([driveItems.get(path.replace(/\/content$/, ''))!.bytes!]),
    post: async (body: unknown) => {
      posts.push({ path, body });
      return path.endsWith('/messages') ? { id: 'draft-1', webLink: 'https://outlook/draft-1' } : undefined;
    },
  };
  return chain;
}

jest.mock('../services/tokenCache.js', () => ({
  getValidAccessTokenForSession: () => Promise.resolve('access-token'),
  getTenantIdFromSession: () => 'test-tenant',
}));
jest.mock('../services/graphClient.js', () => ({ createGraphClient: () => ({ api: graphApi }) }));
jest.mock('../services/userEmailSettings.js', () => ({ getUserEmailSettings: () => mockGetUserEmailSettings() }));
jest.mock('../services/denyList.js', () => ({
  isPathDenied: (t: string, u: string, s: string, p: string) => mockIsPathDenied(t, u, s, p),
}));
jest.mock('../services/auditLog.js', () => ({ logAccess: jest.fn() }));
jest.mock('../services/policyEnforcement.js', () => ({
  withPolicyEnforcement: (_service: unknown, handler: unknown) => handler,
  checkDenyList: () => Promise.resolve(null),
  checkServiceEnabled: (u: string, service: string) => mockCheckServiceEnabled(u, service),
  checkAllowedSite: () => Promise.resolve(null),
}));
jest.mock('../services/securityHeaders.js', () => ({ withSecurity: (handler: unknown) => handler }));
jest.mock('@azure/functions', () => ({ app: { http: jest.fn() } }));

import { app } from '@azure/functions';
import '../functions/mail/createDraft.js';
import '../functions/mail/sendMail.js';
import { INLINE_ATTACHMENT_LIMIT_BYTES } from '../services/mailAttachments.js';

type Handler = (req: HttpRequest, ctx: InvocationContext, auth: AuthResult) => Promise<{ status: number; jsonBody?: Record<string, unknown> }>;
const httpMock = app.http as unknown as jest.Mock<(name: string, opts: { handler: Handler }) => void>;
// Captured at load: beforeEach's clearAllMocks wipes the registration calls.
const routes = new Map(httpMock.mock.calls.map((c) => [c[0], c[1].handler]));
const route = (name: string): Handler => {
  const h = routes.get(name);
  if (!h) throw new Error(`${name} not registered`);
  return h;
};

const AUTH = { userId: 'test-user', session: { userId: 'test-user', tenantId: 'test-tenant', email: 'u@x.com' } } as unknown as AuthResult;
const ctx = { error: jest.fn(), log: jest.fn() } as unknown as InvocationContext;
const call = (name: string, body: Record<string, unknown>) =>
  route(name)({ json: () => Promise.resolve(body) } as unknown as HttpRequest, ctx, AUTH);
const sha256 = (b: Buffer) => createHash('sha256').update(b).digest('hex');

const PDF = Buffer.from('%PDF-1.7 rest-route bytes \x00\x01\x02\xff');

beforeEach(() => {
  jest.clearAllMocks();
  driveItems.clear();
  posts.length = 0;
  mockIsPathDenied.mockResolvedValue(false);
  mockCheckServiceEnabled.mockResolvedValue(null);
  mockGetUserEmailSettings.mockResolvedValue({ emailOutputMode: 'send' });
  driveItems.set('/me/drive/items/OD1', { name: 'Proposal.pdf', parentPath: '/drive/root:/Documents', bytes: PDF });
});

describe('POST /api/mail/drafts', () => {
  it('attaches the OneDrive item with a matching hash', async () => {
    const res = await call('createDraft', { subject: 'S', to: ['a@x.com'], attachments: [{ driveItemId: 'OD1' }] });
    expect(res.status).toBe(201);
    const draft = posts.find((p) => p.path === '/me/messages')!.body as { attachments: Array<{ name: string; contentBytes: string }> };
    expect(draft.attachments[0].name).toBe('Proposal.pdf');
    expect(sha256(Buffer.from(draft.attachments[0].contentBytes, 'base64'))).toBe(sha256(PDF));
  });

  it('returns 403 when the source file is deny-listed', async () => {
    mockIsPathDenied.mockResolvedValue(true);
    const res = await call('createDraft', { subject: 'S', to: ['a@x.com'], attachments: [{ driveItemId: 'OD1' }] });
    expect(res.status).toBe(403);
    expect(res.jsonBody?.error).toMatch(/deny list/);
    expect(posts).toHaveLength(0);
  });

  it('returns 403 when OneDrive is not enabled', async () => {
    mockCheckServiceEnabled.mockResolvedValue({ status: 403, error: 'Service "onedrive" is not enabled for this tenant' });
    const res = await call('createDraft', { subject: 'S', to: ['a@x.com'], attachments: [{ driveItemId: 'OD1' }] });
    expect(res.status).toBe(403);
    expect(mockCheckServiceEnabled).toHaveBeenCalledWith('test-user', 'onedrive');
  });

  it('returns 400 for a folder and for a malformed reference', async () => {
    driveItems.set('/me/drive/items/F1', { name: 'Docs', parentPath: '/drive/root:', folder: true });
    expect((await call('createDraft', { subject: 'S', to: ['a@x.com'], attachments: [{ driveItemId: 'F1' }] })).status).toBe(400);
    expect((await call('createDraft', { subject: 'S', to: ['a@x.com'], attachments: [{ driveItemId: 'a/b' }] })).status).toBe(400);
  });
});

describe('POST /api/mail/send', () => {
  it('sends with the OneDrive item attached', async () => {
    const res = await call('sendMail', { subject: 'S', to: ['a@x.com'], attachments: [{ driveItemId: 'OD1' }] });
    expect(res.status).toBe(200);
    const sent = posts.find((p) => p.path === '/me/sendMail')!.body as { message: { attachments: Array<{ contentBytes: string }> } };
    expect(sha256(Buffer.from(sent.message.attachments[0].contentBytes, 'base64'))).toBe(sha256(PDF));
  });

  it('refuses saveToSentItems:false when a drive item needs an upload session', async () => {
    driveItems.set('/me/drive/items/BIG', { name: 'big.bin', parentPath: '/drive/root:', bytes: Buffer.alloc(INLINE_ATTACHMENT_LIMIT_BYTES + 1, 1) });
    const res = await call('sendMail', { subject: 'S', to: ['a@x.com'], saveToSentItems: false, attachments: [{ driveItemId: 'BIG' }] });
    expect(res.status).toBe(400);
    expect(res.jsonBody?.error).toMatch(/saveToSentItems:false is not supported/);
    expect(posts).toHaveLength(0);
  });
});
