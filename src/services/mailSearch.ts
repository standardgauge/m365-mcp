/**
 * Mail search and enumeration shared by the `search_mail` / `list_messages` MCP
 * tools and the `/api/mail/search` REST route.
 *
 * The defect this module exists to fix: `search_mail` returned plausible empty
 * results for queries it could never match, with no error.
 *
 *   - A folder-scoped search ran `$filter` over subject and sender only. In Sent
 *     Items the sender is always the mailbox owner, so a counterparty's address
 *     could not match anything there. Measured 2026-09-28: `fabrikam`,
 *     `adventureworks`, `clark` all returned 0 against 1, 3 and 3 real sends.
 *   - A global search wraps `q` in a KQL phrase (F9), so the KQL
 *     property syntax the tool description advertised (`from:alice`) was searched
 *     as literal text.
 *   - `$search` is relevance-ranked. `q="a"` with `maxResults=N` came back as a
 *     sample spanning a year, which reads as "the newest N" and is not.
 *
 * So callers now say what they mean. `participant` / `from` / `to` are explicit
 * address criteria that this module turns into KQL property restrictions (global)
 * or a newest-first client-side scan (folder-scoped, where `$search` is unsafe per
 *). An address- or domain-shaped `q` is widened to also match participants.
 * `listMessages` is the deterministic "newest N in a folder" path. Every outcome
 * reports which fields were searched, how results are ordered, and, for a scan,
 * how far back it looked, so an empty result is never mistaken for proof of absence.
 */
import { sanitizeKqlPhrase } from './kqlSearch.js';

/** Minimal Graph client surface this module needs (the real client satisfies it). */
// eslint-disable-next-line @typescript-eslint/no-explicit-any
export type GraphLike = { api: (path: string) => any };

export interface GraphEmailAddress { name?: string; address?: string }
export interface GraphRecipient { emailAddress?: GraphEmailAddress }
export interface GraphMessage {
  id: string;
  subject?: string | null;
  from?: GraphRecipient | null;
  toRecipients?: GraphRecipient[];
  ccRecipients?: GraphRecipient[];
  bccRecipients?: GraphRecipient[];
  receivedDateTime?: string;
  sentDateTime?: string;
  bodyPreview?: string;
  hasAttachments?: boolean;
  isRead?: boolean;
  parentFolderId?: string;
}

export interface MailSearchCriteria {
  /** Free text. Phrase-searched; widened to participants when address-shaped. */
  q?: string;
  /** Matches the sender or any To / Cc / Bcc recipient. */
  participant?: string;
  /** Matches the sender. */
  from?: string;
  /** Matches any To / Cc / Bcc recipient. */
  to?: string;
  /** ISO-8601 date or date-time; only messages received on or after it. */
  since?: string;
}

export interface MailQueryOutcome {
  messages: GraphMessage[];
  /** More matches exist than were returned, or the scan stopped before the end. */
  moreAvailable: boolean;
  strategy: 'kql-search' | 'filter' | 'scan' | 'list';
  ordering: 'relevance' | 'newest-first';
  /** Message fields the query could have matched against. */
  searchedFields: string[];
  /** Scan only: messages examined. */
  scanned?: number;
  /** Scan only: receivedDateTime of the oldest message examined. */
  scanHorizon?: string | null;
  /** Scan only: true when every message in scope was examined. */
  scanComplete?: boolean;
  /** Human-readable caveats a caller should read before trusting an empty result. */
  notes: string[];
}

export const MESSAGE_SELECT_FIELDS =
  'id,subject,from,toRecipients,ccRecipients,receivedDateTime,sentDateTime,bodyPreview,hasAttachments,parentFolderId,isRead';
// bccRecipients is only populated on messages the mailbox sent, which is exactly
// the Sent Items case, so the scan selects it for matching.
const SCAN_SELECT_FIELDS = `${MESSAGE_SELECT_FIELDS},bccRecipients`;

const SCAN_PAGE_SIZE = 100;
/** Hard bound on messages a folder scan examines before reporting an incomplete horizon. */
export const SCAN_BUDGET = 1000;
const SCAN_MAX_PAGES = Math.ceil(SCAN_BUDGET / SCAN_PAGE_SIZE);

const EMAIL_SHAPED = /^[^\s@"]*@[^\s@"]+$|^[^\s@"]+@$/;
const DNS_LABEL = /^[a-z0-9-]+$/i;
const TLD = /^[a-z]{2,}$/i;

/** `fabrikam.com`, `mail.example.co.uk`: two or more DNS labels ending in an alphabetic TLD. */
function isDomainShaped(t: string): boolean {
  const labels = t.split('.');
  if (labels.length < 2) return false;
  return labels.every((l) => DNS_LABEL.test(l)) && TLD.test(labels[labels.length - 1]);
}

/**
 * True when `q` is a single email-address- or domain-shaped token
 * (`jdoe@fabrikam.com`, `@fabrikam.com`, `fabrikam.com`).
 * A bare word like `fabrikam` cannot be told apart from a subject word,
 * which is why `participant` exists as an explicit parameter.
 */
export function isAddressShaped(q: string): boolean {
  const t = q.trim();
  if (!t || /\s/.test(t)) return false;
  return EMAIL_SHAPED.test(t) || isDomainShaped(t);
}

/** Normalise `since` to a UTC ISO timestamp, or throw a caller-facing error. */
export function parseSince(since: string): Date {
  const d = new Date(since);
  if (!/^\d{4}-\d{2}-\d{2}/.test(since.trim()) || Number.isNaN(d.getTime())) {
    throw new Error(`since must be an ISO-8601 date or date-time (e.g. "2026-09-15"), got "${since}"`);
  }
  return d;
}

function clean(v: string | undefined): string | undefined {
  const t = v?.trim();
  return t ? t : undefined;
}

/** Normalised criteria plus the derived address-widening flag. */
interface Plan {
  text?: string;
  widenText: boolean;
  participant?: string;
  from?: string;
  to?: string;
  since?: Date;
}

function plan(c: MailSearchCriteria): Plan {
  const text = clean(c.q);
  const p: Plan = {
    text,
    widenText: Boolean(text && isAddressShaped(text)),
    participant: clean(c.participant),
    from: clean(c.from),
    to: clean(c.to),
  };
  if (clean(c.since)) p.since = parseSince(c.since!.trim());
  if (!p.text && !p.participant && !p.from && !p.to) {
    throw new Error(
      'search_mail needs at least one of q, participant, from, or to. ' +
      'To list the newest messages in a folder, use list_messages.',
    );
  }
  return p;
}

function hasAddressCriteria(p: Plan): boolean {
  return Boolean(p.participant || p.from || p.to || p.widenText);
}

const kqlValue = (v: string) => `"${sanitizeKqlPhrase(v).trim()}"`;

/**
 * Build the KQL for a global `$search`. Every caller-supplied value goes inside a
 * double-quoted phrase with embedded quotes stripped (F9), so no
 * value can inject operators or property restrictions of its own. Only the
 * property names and the AND / OR joins come from this function.
 */
export function buildMailKql(c: MailSearchCriteria): string {
  return kqlFromPlan(plan(c));
}

function kqlFromPlan(p: Plan): string {
  const parts: string[] = [];
  if (p.text) {
    const phrase = kqlValue(p.text);
    parts.push(p.widenText ? `(${phrase} OR participants:${phrase})` : phrase);
  }
  if (p.participant) parts.push(`participants:${kqlValue(p.participant)}`);
  if (p.from) parts.push(`from:${kqlValue(p.from)}`);
  if (p.to) parts.push(`recipients:${kqlValue(p.to)}`);
  if (p.since) parts.push(`received>=${p.since.toISOString().slice(0, 10)}`);
  return parts.join(' AND ');
}

/** Epoch ms of `receivedDateTime`, or -Infinity when missing/unparseable so a `since` bound excludes it. */
function receivedAt(m: GraphMessage): number {
  const t = m.receivedDateTime ? Date.parse(m.receivedDateTime) : NaN;
  return Number.isNaN(t) ? Number.NEGATIVE_INFINITY : t;
}

function addrMatches(r: GraphRecipient | null | undefined, needle: string): boolean {
  const e = r?.emailAddress;
  if (!e) return false;
  return (e.address ?? '').toLowerCase().includes(needle) || (e.name ?? '').toLowerCase().includes(needle);
}

function recipients(m: GraphMessage): GraphRecipient[] {
  return [...(m.toRecipients ?? []), ...(m.ccRecipients ?? []), ...(m.bccRecipients ?? [])];
}

/** Client-side matcher used by the folder scan. Case-insensitive substring semantics. */
export function messageMatches(m: GraphMessage, c: MailSearchCriteria): boolean {
  return matchesPlan(m, plan(c));
}

function matchesPlan(m: GraphMessage, p: Plan): boolean {
  if (p.since && receivedAt(m) < p.since.getTime()) return false;
  const matchFrom = (n: string) => addrMatches(m.from, n);
  const matchTo = (n: string) => recipients(m).some((r) => addrMatches(r, n));
  const matchAny = (n: string) => matchFrom(n) || matchTo(n);
  if (p.participant && !matchAny(p.participant.toLowerCase())) return false;
  if (p.from && !matchFrom(p.from.toLowerCase())) return false;
  if (p.to && !matchTo(p.to.toLowerCase())) return false;
  if (p.text) {
    const n = p.text.toLowerCase();
    const inText = (m.subject ?? '').toLowerCase().includes(n) || (m.bodyPreview ?? '').toLowerCase().includes(n);
    // In a scan every address field is in hand, so text matches them too. That is
    // a strict superset of the legacy subject/sender $filter it replaces.
    if (!inText && !matchAny(n)) return false;
  }
  return true;
}

interface ScanResult { matches: GraphMessage[]; scanned: number; horizon: string | null; complete: boolean }

/**
 * Walk a message collection newest-first, keeping messages that satisfy
 * `predicate`, until `want` matches are held, the collection ends, or the scan
 * budget is spent. Only Graph-issued `@odata.nextLink` URLs are followed, and a
 * repeated link stops the walk (same guard as `graphPaging.collectPage`).
 */
async function scanNewestFirst(
  graph: GraphLike,
  apiPath: string,
  since: Date | undefined,
  select: string,
  want: number,
  predicate: (m: GraphMessage) => boolean,
): Promise<ScanResult> {
  let req = graph.api(apiPath);
  // $orderby needs its property in $filter first or Graph returns InefficientFilter;
  // receivedDateTime is both, so this pairing is the supported shape.
  if (since) req = req.filter(`receivedDateTime ge ${since.toISOString()}`);
  let page = await req.orderby('receivedDateTime desc').select(select).top(SCAN_PAGE_SIZE).get() as {
    value?: GraphMessage[]; '@odata.nextLink'?: string;
  };

  const matches: GraphMessage[] = [];
  const visited = new Set<string>();
  let scanned = 0;
  let horizon: string | null = null;
  let pages = 1;
  for (;;) {
    for (const m of page.value ?? []) {
      if (scanned >= SCAN_BUDGET) break;
      scanned++;
      if (m.receivedDateTime) horizon = m.receivedDateTime;
      if (predicate(m)) matches.push(m);
    }
    const next = page['@odata.nextLink'];
    const exhausted = !next;
    if (exhausted) return { matches, scanned, horizon, complete: true };
    if (matches.length >= want || scanned >= SCAN_BUDGET || pages >= SCAN_MAX_PAGES || visited.has(next)) {
      return { matches, scanned, horizon, complete: false };
    }
    visited.add(next);
    page = await graph.api(next).get();
    pages++;
  }
}

function scanNotes(scan: ScanResult, since: Date | undefined): string[] {
  if (scan.complete) return [];
  return [
    `Scanned the ${scan.scanned} newest messages in scope${scan.horizon ? `, back to ${scan.horizon}` : ''}. ` +
    'Older messages were not examined, so a missing message is not proof of absence. ' +
    (since ? 'Narrow `since` or `folderId` to reach further.' : 'Pass `since` to bound the scan to a window.'),
  ];
}

function sortNewestFirst(values: GraphMessage[]): GraphMessage[] {
  return values.sort((a, b) => (b.receivedDateTime ?? '').localeCompare(a.receivedDateTime ?? ''));
}

const TEXT_ONLY_FILTER_FIELDS = ['subject', 'from.address', 'from.name'];
const SCAN_FIELDS = ['subject', 'bodyPreview', 'from', 'toRecipients', 'ccRecipients', 'bccRecipients'];
const KQL_TEXT_FIELDS = ['subject', 'body', 'from', 'attachments (Graph default $search fields)'];

export interface SearchMailOptions extends MailSearchCriteria {
  /** Graph path prefix: `/me` or `/users/{id}`. */
  base: string;
  folderId?: string;
  maxResults: number;
}

/**
 * Run `search_mail`. Routing:
 *
 *   - Global (no `folderId`): KQL `$search`, relevance-ranked. Address criteria
 *     become `participants:` / `from:` / `recipients:` restrictions. If `$search`
 *     throws, a text-only query falls back to the legacy `$filter` and an address
 *     query falls back to the scan.
 *   - Folder-scoped, text-only, no `since`: the legacy subject/sender `$filter`
 *     over the whole folder ( — never `$search` inside a folder).
 *   - Folder-scoped with any address criterion, an address-shaped `q`, or `since`:
 *     a newest-first scan with client-side matching over sender and all recipients.
 */
export async function searchMail(graph: GraphLike, opts: SearchMailOptions): Promise<MailQueryOutcome> {
  const p = plan(opts);
  const apiPath = opts.folderId ? `${opts.base}/mailFolders/${opts.folderId}/messages` : `${opts.base}/messages`;
  const want = opts.maxResults;

  const runScan = async (extraNotes: string[] = []): Promise<MailQueryOutcome> => {
    const scan = await scanNewestFirst(graph, apiPath, p.since, SCAN_SELECT_FIELDS, want + 1, (m) => matchesPlan(m, p));
    return {
      messages: scan.matches,
      moreAvailable: scan.matches.length > want || !scan.complete,
      strategy: 'scan',
      ordering: 'newest-first',
      searchedFields: SCAN_FIELDS,
      scanned: scan.scanned,
      scanHorizon: scan.horizon,
      scanComplete: scan.complete,
      notes: [...extraNotes, ...scanNotes(scan, p.since)],
    };
  };

  const runTextFilter = async (): Promise<MailQueryOutcome> => {
    const escaped = p.text!.replace(/'/g, "''"); // OData single-quote escape
    const res = await graph
      .api(apiPath)
      .filter(
        `contains(subject,'${escaped}') or ` +
        `contains(from/emailAddress/address,'${escaped}') or ` +
        `contains(from/emailAddress/name,'${escaped}')`,
      )
      .select(MESSAGE_SELECT_FIELDS)
      .top(want)
      .get() as { value?: GraphMessage[]; '@odata.nextLink'?: string };
    // Sort client-side: Graph's $filter + $orderby on messages requires the orderby
    // property to also appear in $filter, or it returns InefficientFilter.
    return {
      messages: sortNewestFirst(res.value ?? []),
      moreAvailable: Boolean(res['@odata.nextLink']),
      strategy: 'filter',
      ordering: 'newest-first',
      searchedFields: TEXT_ONLY_FILTER_FIELDS,
      notes: [
        'Matched subject and sender only. Recipient addresses and body text were not searched; ' +
        'pass `participant` or `to` to match who a message was sent to.',
      ],
    };
  };

  if (opts.folderId) {
    if (!hasAddressCriteria(p) && !p.since) {
      try {
        return await runTextFilter();
      } catch (err) {
        throw new Error(`Mail search failed. $filter error: ${errMsg(err)}`);
      }
    }
    try {
      return await runScan();
    } catch (err) {
      throw new Error(`Mail search failed. scan error: ${errMsg(err)}`);
    }
  }

  let searchError = '';
  try {
    const res = await graph.api(apiPath).search(kqlFromPlan(p)).select(MESSAGE_SELECT_FIELDS).top(want).get() as {
      value?: GraphMessage[]; '@odata.nextLink'?: string;
    };
    const fields = [...KQL_TEXT_FIELDS];
    if (p.widenText || p.participant) fields.push('participants');
    if (p.to) fields.push('recipients');
    // `received>=` is in the KQL, but a search index that ignored it would return
    // older messages with HTTP 200. Re-apply the bound here so it always holds.
    const values = (res.value ?? []).filter((m) => !p.since || receivedAt(m) >= p.since!.getTime());
    return {
      messages: values,
      moreAvailable: Boolean(res['@odata.nextLink']),
      strategy: 'kql-search',
      ordering: 'relevance',
      searchedFields: fields,
      notes: [
        'Results are relevance-ranked, not newest-first; the first N are not the most recent N. ' +
        'Use list_messages to enumerate newest messages.',
      ],
    };
  } catch (err) {
    searchError = errMsg(err);
  }

  try {
    if (hasAddressCriteria(p) || p.since) return await runScan([`$search failed (${searchError}); fell back to a scan.`]);
    const out = await runTextFilter();
    out.notes.unshift(`$search failed (${searchError}); fell back to $filter.`);
    return out;
  } catch (err) {
    throw new Error(`Mail search failed. $search error: ${searchError}. fallback error: ${errMsg(err)}`);
  }
}

export interface ListMessagesOptions {
  base: string;
  folderId?: string;
  since?: string;
  maxResults: number;
}

/**
 * Deterministic "newest N" enumeration: `$orderby=receivedDateTime desc`, with an
 * optional `receivedDateTime ge` bound. This is the path for "what was sent since
 * X" — the question `q="a"` was being used to approximate.
 */
export async function listMessages(graph: GraphLike, opts: ListMessagesOptions): Promise<MailQueryOutcome> {
  const since = clean(opts.since) ? parseSince(opts.since!.trim()) : undefined;
  const apiPath = opts.folderId ? `${opts.base}/mailFolders/${opts.folderId}/messages` : `${opts.base}/messages`;
  let req = graph.api(apiPath);
  if (since) req = req.filter(`receivedDateTime ge ${since.toISOString()}`);
  const res = await req.orderby('receivedDateTime desc').select(MESSAGE_SELECT_FIELDS).top(opts.maxResults).get() as {
    value?: GraphMessage[]; '@odata.nextLink'?: string;
  };
  return {
    messages: res.value ?? [],
    moreAvailable: Boolean(res['@odata.nextLink']),
    strategy: 'list',
    ordering: 'newest-first',
    searchedFields: [],
    notes: [],
  };
}

function errMsg(err: unknown): string {
  return err instanceof Error ? err.message : String(err);
}

/** Public message shape returned by both surfaces. */
export interface MessageSummary {
  id: string;
  subject: string | null;
  from: { name: string; address: string } | null;
  to: Array<{ name: string; address: string }>;
  cc: Array<{ name: string; address: string }>;
  receivedDateTime: string;
  sentDateTime?: string;
  preview: string;
  hasAttachments: boolean;
  isRead: boolean;
  folderId: string;
}

const addr = (r: GraphRecipient) => ({ name: r.emailAddress?.name ?? '', address: r.emailAddress?.address ?? '' });

export function toMessageSummary(m: GraphMessage): MessageSummary {
  return {
    id: m.id,
    subject: m.subject ?? null,
    from: m.from?.emailAddress ? addr(m.from) : null,
    to: (m.toRecipients ?? []).map(addr),
    cc: (m.ccRecipients ?? []).map(addr),
    receivedDateTime: m.receivedDateTime ?? '',
    ...(m.sentDateTime ? { sentDateTime: m.sentDateTime } : {}),
    preview: m.bodyPreview ?? '',
    hasAttachments: m.hasAttachments ?? false,
    isRead: m.isRead ?? false,
    folderId: m.parentFolderId ?? '',
  };
}

/** Outcome metadata both surfaces attach to the response body. */
export function outcomeMeta(o: MailQueryOutcome): Record<string, unknown> {
  return {
    strategy: o.strategy,
    ordering: o.ordering,
    searchedFields: o.searchedFields,
    ...(o.scanned !== undefined ? { scanned: o.scanned, scanHorizon: o.scanHorizon, scanComplete: o.scanComplete } : {}),
    ...(o.notes.length ? { notes: o.notes } : {}),
  };
}
