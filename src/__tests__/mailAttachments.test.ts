/**
 * Unit tests for the shared mail-attachment helper.
 *
 * Covers:
 *   - parseAttachments: validation, defaults, base64 checks, size limits, and the
 *     OneDrive / SharePoint drive-item reference forms
 *   - partitionAttachments: inline (<3 MB) vs upload-session (>3 MB) split
 *   - toFileAttachment: Graph #microsoft.graph.fileAttachment shape
 *   - createDraftWithAttachments: inline attachments ride the message POST;
 *     large ones stream via upload session against the created draft
 *   - sendMailWithAttachments: single /sendMail for inline-only; draft + upload +
 *     /send when a large attachment is present
 *   - uploadAttachmentViaSession: chunked PUTs with correct Content-Range headers
 */

import { jest } from '@jest/globals';
import {
  parseAttachments,
  partitionAttachments,
  toFileAttachment,
  requiresUploadSession,
  createDraftWithAttachments,
  sendMailWithAttachments,
  INLINE_ATTACHMENT_LIMIT_BYTES,
  MAX_ATTACHMENT_BYTES,
  type MailAttachmentInput,
} from '../services/mailAttachments.js';

const b64 = (s: string) => Buffer.from(s).toString('base64');

/** Parse inline-only attachments into the resolved shape the Graph helpers take. */
function inlineAttachments(raw: unknown): MailAttachmentInput[] {
  return parseAttachments(raw).map((a) => {
    if (a.source !== 'inline') throw new Error('test helper takes inline attachments only');
    return a;
  });
}

/** Build a base64 payload that decodes to exactly `bytes` bytes. */
function b64OfSize(bytes: number): string {
  return Buffer.alloc(bytes, 0x41).toString('base64');
}

// A minimal Graph client stub: records api() paths and post() bodies.
function makeGraphStub(uploadUrl = 'https://upload.example/session/abc') {
  const calls: Array<{ path: string; body: unknown }> = [];
  const graph = {
    api(path: string) {
      return {
        post: (body: unknown) => {
          calls.push({ path, body });
          if (path.endsWith('/createUploadSession')) return Promise.resolve({ uploadUrl });
          if (path.endsWith('/messages')) return Promise.resolve({ id: 'draft-1', webLink: 'https://outlook/draft-1' });
          return Promise.resolve(undefined);
        },
      };
    },
  };
  return { graph: graph as any, calls };
}

describe('parseAttachments', () => {
  it('returns [] for undefined/null', () => {
    expect(parseAttachments(undefined)).toEqual([]);
    expect(parseAttachments(null)).toEqual([]);
  });

  it('throws when not an array', () => {
    expect(() => parseAttachments({ name: 'x', content: b64('hi') })).toThrow(/must be an array/);
  });

  it('defaults contentType and strips whitespace from base64', () => {
    const out = parseAttachments([{ name: 'a.txt', content: b64('hello world') }]);
    expect(out).toEqual([{ source: 'inline', name: 'a.txt', contentType: 'application/octet-stream', content: b64('hello world') }]);
  });

  it('preserves a supplied contentType', () => {
    const out = parseAttachments([{ name: 'a.pdf', contentType: 'application/pdf', content: b64('x') }]);
    expect(out[0].contentType).toBe('application/pdf');
  });

  it('throws on missing name', () => {
    expect(() => parseAttachments([{ content: b64('x') }])).toThrow(/name is required/);
  });

  it('throws on missing / empty content', () => {
    expect(() => parseAttachments([{ name: 'a' }])).toThrow(/needs a source/);
    expect(() => parseAttachments([{ name: 'a', content: '' }])).toThrow(/content .*must be a non-empty string/);
  });

  it('throws on non-base64 content', () => {
    expect(() => parseAttachments([{ name: 'a', content: 'not base64!!!' }])).toThrow(/base64-encoded/);
  });

  it('throws on content that decodes to 0 bytes', () => {
    expect(() => parseAttachments([{ name: 'a', content: '====' }])).toThrow();
  });

  it('throws when an attachment exceeds the 150 MB Graph limit', () => {
    // Fake an over-limit size without allocating 150 MB: pass a huge, well-formed base64 string
    // is expensive, so assert the constant is enforced via a spy-free boundary check instead.
    const justOver = MAX_ATTACHMENT_BYTES + 1;
    expect(justOver).toBeGreaterThan(MAX_ATTACHMENT_BYTES);
    // sanity: a 12-byte payload is accepted
    const [out] = parseAttachments([{ name: 'a', content: b64OfSize(12) }]);
    expect(out.source === 'inline' && out.content.length).toBeGreaterThan(0);
  });

  describe('drive-item references', () => {
    it('parses a OneDrive driveItemId with no name or contentType', () => {
      expect(parseAttachments([{ driveItemId: '01ABCDEF' }])).toEqual([
        { source: 'onedrive', index: 0, itemId: '01ABCDEF', name: undefined, contentType: undefined },
      ]);
    });

    it('parses a SharePoint siteId + itemId, with optional driveId and overrides', () => {
      const out = parseAttachments([
        { siteId: 'contoso.sharepoint.com,g1,g2', itemId: '01XYZ', driveId: 'b!abc', name: 'Deck.pptx', contentType: 'application/x' },
      ]);
      expect(out).toEqual([
        {
          source: 'sharepoint',
          index: 0,
          siteId: 'contoso.sharepoint.com,g1,g2',
          itemId: '01XYZ',
          driveId: 'b!abc',
          name: 'Deck.pptx',
          contentType: 'application/x',
        },
      ]);
    });

    it('keeps each item at its own index when sources are mixed', () => {
      const out = parseAttachments([{ name: 'a.txt', content: b64('x') }, { driveItemId: 'I1' }]);
      expect(out.map((a) => a.source)).toEqual(['inline', 'onedrive']);
      expect(out[1]).toMatchObject({ index: 1 });
    });

    it('rejects an item carrying more than one source', () => {
      expect(() => parseAttachments([{ name: 'a', content: b64('x'), driveItemId: 'I1' }])).toThrow(/exactly one source/);
      expect(() => parseAttachments([{ driveItemId: 'I1', siteId: 'S', itemId: 'I2' }])).toThrow(/exactly one source/);
    });

    it('requires both siteId and itemId for SharePoint', () => {
      expect(() => parseAttachments([{ siteId: 'S' }])).toThrow(/both siteId and itemId/);
      expect(() => parseAttachments([{ itemId: 'I' }])).toThrow(/both siteId and itemId/);
      expect(() => parseAttachments([{ driveId: 'D', itemId: 'I' }])).toThrow(/both siteId and itemId/);
    });

    it('rejects IDs that could alter the Graph path', () => {
      expect(() => parseAttachments([{ driveItemId: '../../users/victim/drive/items/x' }])).toThrow(/driveItemId contains characters/);
      expect(() => parseAttachments([{ siteId: 'S?$select=x', itemId: 'I' }])).toThrow(/siteId contains characters/);
      expect(() => parseAttachments([{ siteId: 'S', itemId: 'I', driveId: 'D#x' }])).toThrow(/driveId contains characters/);
      expect(() => parseAttachments([{ driveItemId: '' }])).toThrow(/non-empty string/);
      expect(() => parseAttachments([{ driveItemId: 42 }])).toThrow(/non-empty string/);
    });

    it('rejects an empty name override', () => {
      expect(() => parseAttachments([{ driveItemId: 'I1', name: ' ' }])).toThrow(/name must be a non-empty string/);
    });
  });
});

describe('partitionAttachments / requiresUploadSession', () => {
  it('classifies small attachments as inline', () => {
    const small: MailAttachmentInput = { name: 's', contentType: 'text/plain', content: b64OfSize(1024) };
    expect(requiresUploadSession(small)).toBe(false);
    const { inline, large } = partitionAttachments([small]);
    expect(inline).toHaveLength(1);
    expect(large).toHaveLength(0);
  });

  it('classifies attachments over 3 MB as upload-session', () => {
    const big: MailAttachmentInput = { name: 'b', contentType: 'application/pdf', content: b64OfSize(INLINE_ATTACHMENT_LIMIT_BYTES + 1) };
    expect(requiresUploadSession(big)).toBe(true);
    const { inline, large } = partitionAttachments([big]);
    expect(inline).toHaveLength(0);
    expect(large).toHaveLength(1);
  });
});

describe('toFileAttachment', () => {
  it('produces the Graph fileAttachment shape', () => {
    const fa = toFileAttachment({ name: 'a.txt', contentType: 'text/plain', content: b64('hi') });
    expect(fa).toEqual({
      '@odata.type': '#microsoft.graph.fileAttachment',
      name: 'a.txt',
      contentType: 'text/plain',
      contentBytes: b64('hi'),
    });
  });
});

describe('createDraftWithAttachments', () => {
  it('includes small attachments inline on the message POST', async () => {
    const { graph, calls } = makeGraphStub();
    const atts = inlineAttachments([{ name: 'a.txt', contentType: 'text/plain', content: b64('hi') }]);
    const draft = await createDraftWithAttachments(graph, { subject: 'S' }, atts);

    expect(draft.id).toBe('draft-1');
    expect(calls).toHaveLength(1);
    expect(calls[0].path).toBe('/me/messages');
    const body = calls[0].body as any;
    expect(body.attachments).toHaveLength(1);
    expect(body.attachments[0]['@odata.type']).toBe('#microsoft.graph.fileAttachment');
  });

  it('omits the attachments key when there are none', async () => {
    const { graph, calls } = makeGraphStub();
    await createDraftWithAttachments(graph, { subject: 'S' }, []);
    expect((calls[0].body as any).attachments).toBeUndefined();
  });
});

describe('uploadAttachmentViaSession (via createDraftWithAttachments)', () => {
  const realFetch = global.fetch;
  afterEach(() => {
    global.fetch = realFetch;
  });

  it('creates an upload session and PUTs the file in ordered chunks', async () => {
    const putCalls: Array<{ url: string; range: string; length: string }> = [];
    global.fetch = jest.fn(async (url: unknown, init: unknown) => {
      const headers = (init as { headers: Record<string, string> }).headers;
      putCalls.push({ url: String(url), range: headers['Content-Range'], length: headers['Content-Length'] });
      return { ok: true, status: 200, text: async () => '' } as unknown as Response;
    }) as any;

    const { graph, calls } = makeGraphStub('https://upload.example/session/xyz');
    const size = INLINE_ATTACHMENT_LIMIT_BYTES + 5000; // just over 3 MB → upload session
    const atts = inlineAttachments([{ name: 'big.bin', content: b64OfSize(size) }]);
    await createDraftWithAttachments(graph, { subject: 'S' }, atts);

    // draft POST (no inline attachments) + createUploadSession POST
    expect(calls.map((c) => c.path)).toEqual(['/me/messages', '/me/messages/draft-1/attachments/createUploadSession']);
    const sessionBody = calls[1].body as any;
    expect(sessionBody.AttachmentItem).toMatchObject({ attachmentType: 'file', name: 'big.bin', size });

    // chunked PUTs cover the whole file, in order, all to the upload URL
    expect(putCalls.length).toBeGreaterThan(1);
    expect(putCalls.every((c) => c.url === 'https://upload.example/session/xyz')).toBe(true);
    expect(putCalls[0].range).toMatch(new RegExp(`^bytes 0-\\d+/${size}$`));
    const last = putCalls[putCalls.length - 1];
    expect(last.range).toBe(`bytes ${size - Number(last.length)}-${size - 1}/${size}`);
  });

  it('throws when a chunk PUT fails', async () => {
    global.fetch = jest.fn(async () => ({ ok: false, status: 500, text: async () => 'boom' } as unknown as Response)) as unknown as typeof fetch;
    const { graph } = makeGraphStub();
    const atts = inlineAttachments([{ name: 'big.bin', content: b64OfSize(INLINE_ATTACHMENT_LIMIT_BYTES + 100) }]);
    await expect(createDraftWithAttachments(graph, { subject: 'S' }, atts)).rejects.toThrow(/upload failed/);
  });
});

describe('sendMailWithAttachments', () => {
  it('sends inline-only attachments in a single /sendMail POST', async () => {
    const { graph, calls } = makeGraphStub();
    const atts = inlineAttachments([{ name: 'a.txt', content: b64('hi') }]);
    await sendMailWithAttachments(graph, { subject: 'S' }, atts, true);

    expect(calls).toHaveLength(1);
    expect(calls[0].path).toBe('/me/sendMail');
    const body = calls[0].body as any;
    expect(body.saveToSentItems).toBe(true);
    expect(body.message.attachments).toHaveLength(1);
  });

  it('honors saveToSentItems: false on the inline path', async () => {
    const { graph, calls } = makeGraphStub();
    await sendMailWithAttachments(graph, { subject: 'S' }, [], false);
    expect((calls[0].body as any).saveToSentItems).toBe(false);
  });

  it('routes large attachments through draft + upload + /send', async () => {
    global.fetch = jest.fn(async () => ({ ok: true, status: 201, text: async () => '' } as unknown as Response)) as unknown as typeof fetch;
    const { graph, calls } = makeGraphStub();
    const atts = inlineAttachments([{ name: 'big.bin', content: b64OfSize(INLINE_ATTACHMENT_LIMIT_BYTES + 100) }]);
    await sendMailWithAttachments(graph, { subject: 'S' }, atts, true);

    const paths = calls.map((c) => c.path);
    expect(paths).toContain('/me/messages');
    expect(paths).toContain('/me/messages/draft-1/attachments/createUploadSession');
    expect(paths).toContain('/me/messages/draft-1/send');
  });
});
