/**
 * Unit tests for services/mailSearch.ts.
 *
 * The defect: search_mail returned plausible empty results for queries it could
 * never match. A counterparty's address in Sent Items matched nothing (the
 * folder-scoped $filter covered subject and sender only), and `q="a"` read as a
 * newest-N list when it was a relevance-ranked sample. These tests pin:
 *
 *   1. The KQL builder turns address criteria into property restrictions and
 *      widens an address-shaped `q` to participants.
 *   2. The client-side matcher matches recipients (To / Cc / Bcc), not just
 *      subject and sender.
 *   3. Routing: folder + address criterion → newest-first scan; folder + bare text
 *      → the $filter (never $search); mailbox-wide → KQL $search.
 *   4. A scan reports how far it looked, and stops on budget / nextLink loops.
 *   5. listMessages is $orderby-driven and deterministic.
 *   6. Every value the caller supplies is sanitised before reaching KQL / OData.
 */
import {
  buildMailKql,
  assertNoKqlProperties,
  isAddressShaped,
  isBareDomain,
  KQL_RESULT_CAP,
  listMessages,
  messageMatches,
  parseSince,
  searchMail,
  toMessageSummary,
  SCAN_BUDGET,
  type GraphMessage,
} from '../services/mailSearch.js';

// ── Fake Graph client ─────────────────────────────────────────────────────────

interface Call {
  path: string;
  search?: string;
  filter?: string;
  orderby?: string;
  select?: string;
  top?: number;
}

type Page = { value?: GraphMessage[]; '@odata.nextLink'?: string };

/** Records every request; `respond` decides what each `get()` returns (or throws). */
function fakeGraph(respond: (call: Call, index: number) => Page | Error) {
  const calls: Call[] = [];
  const graph = {
    api: (path: string) => {
      const call: Call = { path };
      calls.push(call);
      const b = {
        search: (v: string) => { call.search = v; return b; },
        filter: (v: string) => { call.filter = v; return b; },
        orderby: (v: string) => { call.orderby = v; return b; },
        select: (v: string) => { call.select = v; return b; },
        top: (v: number) => { call.top = v; return b; },
        get: async () => {
          const r = respond(call, calls.length - 1);
          if (r instanceof Error) throw r;
          return r;
        },
      };
      return b;
    },
  };
  return { graph, calls };
}

function msg(id: string, over: Partial<GraphMessage> = {}): GraphMessage {
  return {
    id,
    subject: `Subject ${id}`,
    from: { emailAddress: { name: 'Test User', address: 'nate@example.com' } },
    toRecipients: [{ emailAddress: { name: 'Jordan Doe', address: 'jdoe@fabrikam.com' } }],
    receivedDateTime: '2026-09-23T17:05:12Z',
    bodyPreview: 'preview',
    parentFolderId: 'sent',
    ...over,
  };
}

const BASE = { base: '/me', maxResults: 25 };

// ── isAddressShaped ───────────────────────────────────────────────────────────

describe('isAddressShaped', () => {
  it.each([
    'jdoe@fabrikam.com',
    '@fabrikam.com',
    'jdoe@',
    'fabrikam.com',
    'mail.example.co.uk',
    '  fabrikam.com  ',
  ])('is true for %s', (q) => expect(isAddressShaped(q)).toBe(true));

  it.each(['fabrikam', 'acme', 'Q3 roadmap', 'a', '', 'foo bar@baz.com', 'v1.2', 'example.c0m', '.com'])(
    'is false for %j (a bare word cannot be told from a subject word)',
    (q) => expect(isAddressShaped(q)).toBe(false),
  );
});

// ── parseSince ────────────────────────────────────────────────────────────────

describe('parseSince', () => {
  it('accepts an ISO date and an ISO date-time', () => {
    expect(parseSince('2026-09-15').toISOString()).toBe('2026-09-15T00:00:00.000Z');
    expect(parseSince('2026-09-15T10:00:00Z').getTime()).toBe(Date.parse('2026-09-15T10:00:00Z'));
  });

  it.each(['yesterday', '15/09/2026', '2026-13-45', 'now'])('rejects %s', (v) => {
    expect(() => parseSince(v)).toThrow(/ISO-8601/);
  });
});

// ── buildMailKql ──────────────────────────────────────────────────────────────

describe('buildMailKql', () => {
  it('phrase-wraps free text (unchanged F9 behaviour)', () => {
    expect(buildMailKql({ q: 'budget' })).toBe('"budget"');
  });

  it('widens an address-shaped q to participants:', () => {
    expect(buildMailKql({ q: 'jdoe@fabrikam.com' })).toBe(
      '("jdoe@fabrikam.com" OR participants:"jdoe@fabrikam.com")',
    );
    expect(buildMailKql({ q: 'fabrikam.com' })).toBe('("fabrikam.com" OR participants:"fabrikam.com")');
  });

  it('does not widen a bare word', () => {
    expect(buildMailKql({ q: 'fabrikam' })).toBe('"fabrikam"');
  });

  it('maps participant / from / to / since to KQL property restrictions', () => {
    expect(buildMailKql({ participant: 'fabrikam.com' })).toBe('participants:"fabrikam.com"');
    expect(buildMailKql({ from: 'alice@example.com' })).toBe('from:"alice@example.com"');
    expect(buildMailKql({ to: 'bob@example.com' })).toBe('recipients:"bob@example.com"');
    expect(buildMailKql({ q: 'roadmap', to: 'bob@example.com', since: '2026-09-15' })).toBe(
      '"roadmap" AND recipients:"bob@example.com" AND received>=2026-09-15',
    );
  });

  it('strips embedded double quotes from every value so no value can inject operators', () => {
    expect(buildMailKql({ q: 'x" OR "boss' })).toBe('"x  OR  boss"');
    expect(buildMailKql({ to: 'a" OR subject:"secret' })).toBe('recipients:"a  OR subject: secret"');
  });

  it('requires at least one criterion and points at list_messages', () => {
    expect(() => buildMailKql({})).toThrow(/list_messages/);
    expect(() => buildMailKql({ q: '   ', since: '2026-09-15' })).toThrow(/at least one of/);
  });
});

// ── messageMatches ────────────────────────────────────────────────────────────

describe('messageMatches', () => {
  const m = msg('m1', {
    ccRecipients: [{ emailAddress: { name: 'Riley Clark', address: 'riley.clark@northwindtraders.com' } }],
    bccRecipients: [{ emailAddress: { name: '', address: 'asmith@adventureworks.com' } }],
  });

  it('matches a recipient address fragment via participant (the miss)', () => {
    expect(messageMatches(m, { participant: 'fabrikam' })).toBe(true);
    expect(messageMatches(m, { participant: 'JDOE@FABRIKAM.COM' })).toBe(true);
    expect(messageMatches(m, { participant: 'nobody' })).toBe(false);
  });

  it('matches Cc, Bcc and display names through to', () => {
    expect(messageMatches(m, { to: 'northwindtraders' })).toBe(true); // Cc
    expect(messageMatches(m, { to: 'adventureworks' })).toBe(true); // Bcc
    expect(messageMatches(m, { to: 'clark' })).toBe(true); // display name
    expect(messageMatches(m, { to: 'example' })).toBe(false); // sender is not a recipient
  });

  it('from matches the sender only', () => {
    expect(messageMatches(m, { from: 'example.com' })).toBe(true);
    expect(messageMatches(m, { from: 'fabrikam' })).toBe(false);
  });

  it('text in a scan matches subject, preview and every address', () => {
    expect(messageMatches(m, { q: 'subject m1' })).toBe(true);
    expect(messageMatches(m, { q: 'preview' })).toBe(true);
    expect(messageMatches(m, { q: 'adventureworks' })).toBe(true);
    expect(messageMatches(m, { q: 'unrelated' })).toBe(false);
  });

  it('since compares as an instant, not as a string (Graph omits fractional seconds)', () => {
    const at = msg('x', { receivedDateTime: '2026-09-15T00:00:00Z' });
    expect(messageMatches(at, { q: 'subject', since: '2026-09-15' })).toBe(true);
    expect(messageMatches(at, { q: 'subject', since: '2026-09-15T00:00:01Z' })).toBe(false);
    expect(messageMatches(msg('y', { receivedDateTime: undefined }), { q: 'subject', since: '2026-09-15' })).toBe(false);
  });

  it('all criteria must hold together', () => {
    expect(messageMatches(m, { q: 'subject', to: 'fabrikam', from: 'example' })).toBe(true);
    expect(messageMatches(m, { q: 'subject', to: 'fabrikam', from: 'nobody' })).toBe(false);
  });
});

// ── searchMail routing ────────────────────────────────────────────────────────

describe('searchMail — folder-scoped', () => {
  it('bare text uses the subject/sender $filter, never $search, and says what it did not search', async () => {
    const { graph, calls } = fakeGraph(() => ({ value: [msg('m1')] }));
    const out = await searchMail(graph, { ...BASE, folderId: 'sent', q: "O'Brien" });
    expect(calls).toHaveLength(1);
    expect(calls[0].search).toBeUndefined();
    expect(calls[0].filter).toContain("contains(subject,'O''Brien')");
    expect(calls[0].orderby).toBeUndefined();
    expect(out.strategy).toBe('filter');
    expect(out.searchedFields).toEqual(['subject', 'from.address', 'from.name']);
    expect(out.notes[0]).toMatch(/Recipient addresses .* were not searched/);
  });

  it('an address criterion runs a newest-first scan and matches the recipient', async () => {
    const pages: Page[] = [{ value: [msg('older-unrelated', { toRecipients: [{ emailAddress: { address: 'x@y.com' } }] }), msg('hit')] }];
    const { graph, calls } = fakeGraph((_c, i) => pages[i]);
    const out = await searchMail(graph, { ...BASE, folderId: 'sent', to: 'jdoe@fabrikam.com' });
    expect(calls[0].path).toBe('/me/mailFolders/sent/messages');
    expect(calls[0].search).toBeUndefined();
    expect(calls[0].orderby).toBe('receivedDateTime desc');
    expect(calls[0].select).toContain('bccRecipients');
    expect(out.strategy).toBe('scan');
    expect(out.messages.map((m) => m.id)).toEqual(['hit']);
    expect(out.scanned).toBe(2);
    expect(out.scanComplete).toBe(true);
    expect(out.moreAvailable).toBe(false);
    expect(out.notes).toEqual([]);
  });

  it('an address-shaped q also takes the scan path', async () => {
    const { graph, calls } = fakeGraph(() => ({ value: [msg('hit')] }));
    const out = await searchMail(graph, { ...BASE, folderId: 'sent', q: 'fabrikam.com' });
    expect(calls[0].orderby).toBe('receivedDateTime desc');
    expect(out.strategy).toBe('scan');
    expect(out.messages.map((m) => m.id)).toEqual(['hit']);
  });

  it('since is pushed to Graph as the $filter that licenses $orderby', async () => {
    const { graph, calls } = fakeGraph(() => ({ value: [] }));
    await searchMail(graph, { ...BASE, folderId: 'sent', q: 'roadmap', since: '2026-09-15' });
    expect(calls[0].filter).toBe('receivedDateTime ge 2026-09-15T00:00:00.000Z');
    expect(calls[0].orderby).toBe('receivedDateTime desc');
  });

  it('follows nextLink pages until it holds maxResults + 1 matches, then reports more', async () => {
    const page = (i: number, more: boolean): Page => ({
      value: [msg(`p${i}-a`), msg(`p${i}-b`)],
      ...(more ? { '@odata.nextLink': `https://graph/next/${i}` } : {}),
    });
    const { graph, calls } = fakeGraph((_c, i) => page(i, true));
    const out = await searchMail(graph, { ...BASE, maxResults: 3, folderId: 'sent', to: 'fabrikam' });
    // Two pages give 4 matches ≥ want(3)+1.
    expect(calls).toHaveLength(2);
    expect(calls[1].path).toBe('https://graph/next/0');
    expect(out.messages).toHaveLength(4);
    expect(out.moreAvailable).toBe(true);
    expect(out.scanComplete).toBe(false);
  });

  it('stops at the scan budget and says how far back it looked', async () => {
    let n = 0;
    const { graph, calls } = fakeGraph(() => ({
      value: Array.from({ length: 100 }, () => msg(`m${n++}`, {
        receivedDateTime: `2026-0${1 + (n % 8)}-01T00:00:00Z`,
        toRecipients: [{ emailAddress: { address: 'someone@else.com' } }],
      })),
      '@odata.nextLink': `https://graph/next/${n}`,
    }));
    const out = await searchMail(graph, { ...BASE, folderId: 'sent', to: 'never-matches' });
    expect(calls).toHaveLength(SCAN_BUDGET / 100);
    expect(out.messages).toEqual([]);
    expect(out.scanned).toBe(SCAN_BUDGET);
    expect(out.scanComplete).toBe(false);
    expect(out.moreAvailable).toBe(true);
    expect(out.scanHorizon).toMatch(/^2026-/);
    expect(out.notes[0]).toMatch(new RegExp(`Scanned the ${SCAN_BUDGET} newest messages`));
    expect(out.notes[0]).toMatch(/not proof of absence/);
  });

  it('stops on a repeated nextLink', async () => {
    const { graph, calls } = fakeGraph(() => ({ value: [msg('x', { toRecipients: [] })], '@odata.nextLink': 'https://graph/same' }));
    const out = await searchMail(graph, { ...BASE, folderId: 'sent', to: 'nobody' });
    expect(calls).toHaveLength(2);
    expect(out.scanComplete).toBe(false);
  });

  it('a Graph failure surfaces as an error, never as an empty result', async () => {
    const { graph } = fakeGraph(() => new Error('MailboxNotEnabledForRESTAPI'));
    await expect(searchMail(graph, { ...BASE, folderId: 'sent', to: 'x' })).rejects.toThrow(/Mail search failed\. scan error: MailboxNotEnabledForRESTAPI/);
    await expect(searchMail(graph, { ...BASE, folderId: 'sent', q: 'x' })).rejects.toThrow(/Mail search failed\. \$filter error/);
  });
});

describe('searchMail — mailbox-wide', () => {
  it('uses KQL $search with property restrictions and reports relevance ordering', async () => {
    const { graph, calls } = fakeGraph(() => ({ value: [msg('m1')], '@odata.nextLink': 'https://graph/more' }));
    const out = await searchMail(graph, { ...BASE, q: 'roadmap', to: 'fabrikam.com' });
    expect(calls).toHaveLength(1);
    expect(calls[0].path).toBe('/me/messages');
    expect(calls[0].search).toBe('"roadmap" AND recipients:"fabrikam.com"');
    expect(calls[0].top).toBe(25);
    expect(out.strategy).toBe('kql-search');
    expect(out.ordering).toBe('relevance');
    expect(out.searchedFields).toContain('recipients');
    expect(out.moreAvailable).toBe(true);
    expect(out.notes[0]).toMatch(/relevance-ranked, not newest-first/);
  });

  it('re-applies since client-side even though it is in the KQL', async () => {
    const { graph } = fakeGraph(() => ({
      value: [msg('new', { receivedDateTime: '2026-09-20T00:00:00Z' }), msg('old', { receivedDateTime: '2026-01-01T00:00:00Z' })],
    }));
    const out = await searchMail(graph, { ...BASE, q: 'roadmap', since: '2026-09-15' });
    expect(out.messages.map((m) => m.id)).toEqual(['new']);
  });

  it('falls back to $filter for a text-only query when $search throws', async () => {
    const { graph, calls } = fakeGraph((_c, i) => (i === 0 ? new Error('SearchQueryNotSupported') : { value: [msg('m1')] }));
    const out = await searchMail(graph, { ...BASE, q: 'roadmap' });
    expect(calls).toHaveLength(2);
    expect(calls[1].filter).toContain("contains(subject,'roadmap')");
    expect(out.strategy).toBe('filter');
    expect(out.notes[0]).toMatch(/\$search failed \(SearchQueryNotSupported\); fell back to \$filter/);
  });

  it('falls back to a scan for an address query when $search throws', async () => {
    const { graph, calls } = fakeGraph((_c, i) => (i === 0 ? new Error('boom') : { value: [msg('m1')] }));
    const out = await searchMail(graph, { ...BASE, to: 'fabrikam' });
    expect(calls).toHaveLength(2);
    expect(calls[1].orderby).toBe('receivedDateTime desc');
    expect(out.strategy).toBe('scan');
    expect(out.notes[0]).toMatch(/\$search failed \(boom\); fell back to a scan/);
  });

  it('reports both errors when $search and the fallback fail', async () => {
    const { graph } = fakeGraph((_c, i) => new Error(i === 0 ? 'first' : 'second'));
    await expect(searchMail(graph, { ...BASE, q: 'x' })).rejects.toThrow(/\$search error: first\. fallback error: second/);
  });

  it('rejects a call with no criteria before touching Graph', async () => {
    const { graph, calls } = fakeGraph(() => ({ value: [] }));
    await expect(searchMail(graph, { ...BASE })).rejects.toThrow(/at least one of q, participant, from, or to/);
    expect(calls).toHaveLength(0);
  });
});

// ──: the silent mailbox-wide cases ─────────────────────────────────────

describe('assertNoKqlProperties', () => {
  it.each([
    ['received>=2026-01-01', /pass `since`/i],
    ['roadmap received:2026-09', /pass `since`/i],
    ['sent>2026-01-01', /pass `since`/i],
    ['participants:exampleequity.com', /pass `participant`, `from` or `to`/i],
    ['from:alice', /pass `participant`, `from` or `to`/i],
    ['(to:bob)', /pass `participant`, `from` or `to`/i],
    ['subject:roadmap', /only the words to match/i],
  ])('rejects %j instead of phrase-searching it', (q, hint) => {
    expect(() => assertNoKqlProperties(q)).toThrow(/KQL property syntax/);
    expect(() => assertNoKqlProperties(q)).toThrow(hint);
  });

  it.each(['Q3 roadmap', 'Re budget', 'ratio 3:1', 'jdoe@fabrikam.com', 'tomorrow', 'fromage'])(
    'accepts plain text %j',
    (q) => expect(() => assertNoKqlProperties(q)).not.toThrow(),
  );

  it('rejects on every route before touching Graph, folder-scoped included', async () => {
    const { graph, calls } = fakeGraph(() => ({ value: [] }));
    await expect(searchMail(graph, { ...BASE, q: 'received>=2026-01-01' })).rejects.toThrow(/since/);
    await expect(searchMail(graph, { ...BASE, folderId: 'sent', q: 'participants:example.com' })).rejects.toThrow(/participant/);
    expect(calls).toHaveLength(0);
  });
});

describe('isBareDomain', () => {
  it.each(['fabrikam.com', '@fabrikam.com', ' mail.example.co.uk '])('is true for %j', (v) =>
    expect(isBareDomain(v)).toBe(true));
  it.each(['jdoe@fabrikam.com', 'jdoe@', 'fabrikam'])('is false for %j', (v) =>
    expect(isBareDomain(v)).toBe(false));
});

describe('searchMail — mailbox-wide,', () => {
  const old = (id: string) => msg(id, { receivedDateTime: '2025-01-01T00:00:00Z' });
  const fresh = (id: string) => msg(id, { receivedDateTime: '2026-09-20T00:00:00Z' });
  const many = (n: number, f: (id: string) => GraphMessage, prefix: string) =>
    Array.from({ length: n }, (_, i) => f(`${prefix}${i}`));

  it('re-checks an empty KQL answer to an address criterion with a scan (participants:<domain> returns [])', async () => {
    const { graph, calls } = fakeGraph((_c, i) => (i === 0 ? { value: [] } : { value: [msg('m1')] }));
    const out = await searchMail(graph, { ...BASE, participant: 'fabrikam.com' });
    expect(calls).toHaveLength(2);
    expect(calls[0].search).toBe('participants:"fabrikam.com"');
    expect(calls[1].search).toBeUndefined();
    expect(calls[1].orderby).toBe('receivedDateTime desc');
    expect(out.strategy).toBe('scan');
    expect(out.messages.map((m) => m.id)).toEqual(['m1']);
    expect(out.scanComplete).toBe(true);
    expect(out.notes[0]).toMatch(/KQL \$search returned no matches.*bare domain/);
  });

  it('does not re-check an empty text-only KQL result', async () => {
    const { graph, calls } = fakeGraph(() => ({ value: [] }));
    const out = await searchMail(graph, { ...BASE, q: 'roadmap' });
    expect(calls).toHaveLength(1);
    expect(out.strategy).toBe('kql-search');
    expect(out.messages).toEqual([]);
  });

  it('a failing confirming scan is an error, not an empty result', async () => {
    const { graph } = fakeGraph((_c, i) => (i === 0 ? { value: [] } : new Error('scan boom')));
    await expect(searchMail(graph, { ...BASE, to: 'fabrikam.com' })).rejects.toThrow(/confirming scan failed: scan boom/);
  });

  it('with since, pages past out-of-window hits instead of returning one filtered page', async () => {
    const { graph, calls } = fakeGraph((_c, i) =>
      i === 0
        ? { value: many(100, old, 'o'), '@odata.nextLink': 'https://graph/p2' }
        : { value: [fresh('n1'), fresh('n2')] });
    const out = await searchMail(graph, { ...BASE, q: 'roadmap', since: '2026-09-15' });
    expect(calls).toHaveLength(2);
    expect(calls[0].top).toBe(100);
    expect(calls[1].path).toBe('https://graph/p2');
    expect(out.messages.map((m) => m.id)).toEqual(['n1', 'n2']);
    expect(out.moreAvailable).toBe(false);
  });

  it('stops paging once it holds maxResults + 1 in-window matches', async () => {
    const { graph, calls } = fakeGraph(() => ({ value: many(30, fresh, 'n'), '@odata.nextLink': 'https://graph/more' }));
    const out = await searchMail(graph, { ...BASE, q: 'roadmap', since: '2026-09-15' });
    expect(calls).toHaveLength(1);
    expect(out.moreAvailable).toBe(true);
  });

  it('reports the index cap instead of presenting it as the end of the mailbox (~275 hits, no nextLink)', async () => {
    const { graph } = fakeGraph((_c, i) =>
      i < 2
        ? { value: many(100, old, `o${i}-`), '@odata.nextLink': `https://graph/p${i + 2}` }
        : { value: [...many(73, old, 'o2-'), fresh('n1'), fresh('n2')] });
    const out = await searchMail(graph, { ...BASE, q: 'roadmap', since: '2026-09-15' });
    expect(out.messages.map((m) => m.id)).toEqual(['n1', 'n2']);
    expect(275).toBeGreaterThanOrEqual(KQL_RESULT_CAP);
    expect(out.moreAvailable).toBe(true);
    expect(out.notes.join(' ')).toMatch(/stopped after 275 hits.*not the end of the mailbox/);
    expect(out.notes.join(' ')).toMatch(/273 of 275 hits fell before since/);
  });

  it('stops at the page budget, reports more, and says hits were dropped by since', async () => {
    let n = 0;
    const { graph, calls } = fakeGraph(() => ({ value: many(100, old, `o${n}-`), '@odata.nextLink': `https://graph/p${++n}` }));
    const out = await searchMail(graph, { ...BASE, q: 'roadmap', since: '2026-09-15' });
    expect(calls).toHaveLength(5);
    expect(out.messages).toEqual([]);
    expect(out.moreAvailable).toBe(true);
    expect(out.notes.join(' ')).toMatch(/500 of 500 hits fell before since/);
  });

  it('a short result without since is complete and carries no cap note', async () => {
    const { graph } = fakeGraph(() => ({ value: [msg('m1')] }));
    const out = await searchMail(graph, { ...BASE, q: 'roadmap' });
    expect(out.moreAvailable).toBe(false);
    expect(out.notes).toHaveLength(1);
  });

  it('warns that a domain-shaped q may miss address-only matches', async () => {
    const { graph } = fakeGraph(() => ({ value: [msg('m1')] }));
    const out = await searchMail(graph, { ...BASE, q: '@fabrikam.com' });
    expect(out.notes.join(' ')).toMatch(/does not reliably match a bare domain/);
    const full = await searchMail(fakeGraph(() => ({ value: [msg('m1')] })).graph, { ...BASE, q: 'jdoe@fabrikam.com' });
    expect(full.notes.join(' ')).not.toMatch(/bare domain/);
  });
});

// ── listMessages ──────────────────────────────────────────────────────────────

describe('listMessages', () => {
  it('enumerates newest-first with $orderby, no $search, and surfaces truncation', async () => {
    const { graph, calls } = fakeGraph(() => ({ value: [msg('a'), msg('b')], '@odata.nextLink': 'https://graph/more' }));
    const out = await listMessages(graph, { base: '/me', folderId: 'sent', maxResults: 2 });
    expect(calls).toHaveLength(1);
    expect(calls[0].path).toBe('/me/mailFolders/sent/messages');
    expect(calls[0].search).toBeUndefined();
    expect(calls[0].filter).toBeUndefined();
    expect(calls[0].orderby).toBe('receivedDateTime desc');
    expect(calls[0].top).toBe(2);
    expect(out.strategy).toBe('list');
    expect(out.ordering).toBe('newest-first');
    expect(out.moreAvailable).toBe(true);
  });

  it('applies since as a $filter and lists the whole mailbox without a folder', async () => {
    const { graph, calls } = fakeGraph(() => ({ value: [] }));
    const out = await listMessages(graph, { base: '/users/u1', since: '2026-09-15', maxResults: 100 });
    expect(calls[0].path).toBe('/users/u1/messages');
    expect(calls[0].filter).toBe('receivedDateTime ge 2026-09-15T00:00:00.000Z');
    expect(out.moreAvailable).toBe(false);
  });

  it('rejects a malformed since before touching Graph', async () => {
    const { graph, calls } = fakeGraph(() => ({ value: [] }));
    await expect(listMessages(graph, { base: '/me', since: 'last week', maxResults: 5 })).rejects.toThrow(/ISO-8601/);
    expect(calls).toHaveLength(0);
  });
});

// ── toMessageSummary ──────────────────────────────────────────────────────────

describe('toMessageSummary', () => {
  it('includes cc and sentDateTime and tolerates missing fields', () => {
    const s = toMessageSummary(msg('m1', {
      ccRecipients: [{ emailAddress: { address: 'cc@example.com' } }],
      sentDateTime: '2026-09-23T17:05:00Z',
      from: null,
      subject: undefined,
    }));
    expect(s).toMatchObject({
      id: 'm1',
      subject: null,
      from: null,
      to: [{ name: 'Jordan Doe', address: 'jdoe@fabrikam.com' }],
      cc: [{ name: '', address: 'cc@example.com' }],
      sentDateTime: '2026-09-23T17:05:00Z',
      preview: 'preview',
      hasAttachments: false,
      isRead: false,
      folderId: 'sent',
    });
  });
});
