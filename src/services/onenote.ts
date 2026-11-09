/**
 * OneNote page helpers for the onenote MCP tools.
 *
 * The deny list keys OneNote on notebooks and sections: an entry naming a
 * notebook's or a section's ID or display name blocks every page in it. The
 * admin UI records notebook blocks by name, so an ID-only check would miss them.
 * Every page tool resolves its target to the section and notebook that hold it,
 * and checks both by ID and by name, before any page content is read or written.
 * A target whose section or notebook can't be resolved is refused.
 *
 * Graph v1.0 has no full-text search over OneNote pages for work and school
 * accounts. search_onenote_pages matches page titles with an OData `contains`
 * filter, which the delegated Notes.* permissions allow.
 */
import type { Client } from '@microsoft/microsoft-graph-client';
import { filterDeniedPaths, isPathDenied, type DenySubject } from './denyList.js';
import { encodeGraphId } from './opaqueId.js';
import { ToolCallError } from './toolCallError.js';

// eslint-disable-next-line @typescript-eslint/no-explicit-any
type GraphObject = Record<string, any>;

/** Graph's page-collection ceiling for one request. */
export const ONENOTE_PAGE_LIMIT = 100;

/**
 * Largest page body read_onenote_page returns. Images and attachments come back
 * as resource URLs rather than inline bytes, so a page's HTML is normally small;
 * this only stops a pathological page from being held in memory whole.
 */
export const MAX_ONENOTE_PAGE_BYTES = 5 * 1024 * 1024;

/** The commands Graph's PATCH /pages/{id}/content accepts. */
export const ONENOTE_UPDATE_ACTIONS = ['append', 'insert', 'prepend', 'replace'] as const;
export const ONENOTE_UPDATE_POSITIONS = ['after', 'before'] as const;

/** Fields read for every page, and the parent expansion the deny check needs. */
export const PAGE_SELECT = 'id,title,createdDateTime,lastModifiedDateTime';
export const PAGE_PARENT_EXPAND = 'parentSection($select=id,displayName),parentNotebook($select=id,displayName)';

interface Container { id?: string; displayName?: string }

export interface PageSummary {
  id: string;
  title: string | null;
  createdDateTime: string | null;
  lastModifiedDateTime: string | null;
  sectionId: string | null;
  sectionName: string | null;
  notebookId: string | null;
  notebookName: string | null;
}

export function toPageSummary(p: GraphObject): PageSummary {
  return {
    id: p.id,
    title: p.title ?? null,
    createdDateTime: p.createdDateTime ?? null,
    lastModifiedDateTime: p.lastModifiedDateTime ?? null,
    sectionId: p.parentSection?.id ?? null,
    sectionName: p.parentSection?.displayName ?? null,
    notebookId: p.parentNotebook?.id ?? null,
    notebookName: p.parentNotebook?.displayName ?? null,
  };
}

/**
 * True unless both the section and the notebook are present and neither is
 * deny-listed by ID or name. A missing parent counts as denied (fail closed).
 */
async function containersDenied(
  tenantId: string,
  subject: DenySubject,
  section: Container | undefined,
  notebook: Container | undefined,
): Promise<boolean> {
  if (!section?.id || !notebook?.id) return true;
  const items = [
    { id: section.id, name: section.displayName },
    { id: notebook.id, name: notebook.displayName },
  ];
  const kept = await filterDeniedPaths(tenantId, subject, 'onenote', items);
  return kept.length !== items.length;
}

/**
 * Refuse when the section, or the notebook holding it, is deny-listed. The
 * caller's section ID is checked before Graph is called, then the section is
 * read to check its name and its notebook.
 */
export async function assertSectionAllowed(graph: Client, tenantId: string, subject: DenySubject, sectionId: string): Promise<GraphObject> {
  if (await isPathDenied(tenantId, subject, 'onenote', sectionId)) {
    throw new ToolCallError('Access restricted by deny list — section is blocked');
  }
  const section: GraphObject = await graph
    .api(`/me/onenote/sections/${encodeGraphId(sectionId, 'sectionId')}`)
    .select('id,displayName')
    .expand('parentNotebook($select=id,displayName)')
    .get();
  if (await containersDenied(tenantId, subject, { id: sectionId, displayName: section?.displayName }, section?.parentNotebook)) {
    throw new ToolCallError('Access restricted by deny list — the section or its notebook is blocked');
  }
  return section;
}

/**
 * Read a page's metadata and refuse when its section or notebook is
 * deny-listed. Returns the page summary for the caller to echo.
 */
export async function assertPageAllowed(graph: Client, tenantId: string, subject: DenySubject, pageId: string): Promise<PageSummary> {
  const page: GraphObject = await graph
    .api(`/me/onenote/pages/${encodeGraphId(pageId, 'pageId')}`)
    .select(PAGE_SELECT)
    .expand(PAGE_PARENT_EXPAND)
    .get();
  if (await containersDenied(tenantId, subject, page?.parentSection, page?.parentNotebook)) {
    throw new ToolCallError('Access restricted by deny list — the page\'s section or notebook is blocked');
  }
  return toPageSummary({ ...page, id: page?.id ?? pageId });
}

/**
 * Drop pages whose section or notebook is deny-listed, or whose parents Graph
 * didn't return. Each distinct container is checked once.
 */
export async function filterDeniedPages(tenantId: string, subject: DenySubject, pages: GraphObject[]): Promise<GraphObject[]> {
  const containers = new Map<string, { id: string; name?: string }>();
  for (const p of pages) {
    for (const c of [p.parentSection, p.parentNotebook] as Array<Container | undefined>) {
      if (c?.id && !containers.has(c.id)) containers.set(c.id, { id: c.id, name: c.displayName });
    }
  }
  const allowed = new Set((await filterDeniedPaths(tenantId, subject, 'onenote', [...containers.values()])).map((c) => c.id));
  return pages.filter((p) => allowed.has(p.parentSection?.id) && allowed.has(p.parentNotebook?.id));
}

/** Quote a string as an OData literal: wrap in single quotes, double any inside. */
export function odataString(value: string): string {
  return `'${value.replace(/'/g, "''")}'`;
}

/** An optional `skip` argument: a non-negative integer, or 0 when absent. */
export function skipArg(value: unknown): number {
  if (value === undefined || value === null) return 0;
  if (typeof value !== 'number' || !Number.isInteger(value) || value < 0) {
    throw new ToolCallError('skip must be a non-negative integer');
  }
  return value;
}

/**
 * The next `skip` for a page listing. Graph applies `skip` to the raw
 * collection and the deny list filters afterwards, so the cursor advances by
 * the rows Graph returned, not by the rows the caller saw. Advancing by the
 * filtered count would re-read part of the same page, and a page that was
 * entirely denied would point the caller back at the same `skip` forever.
 */
export function pageContinuation(skip: number, rawCount: number, truncated: boolean): { nextSkip?: number } {
  return truncated ? { nextSkip: skip + rawCount } : {};
}

const GENERATED_ID_RE = /^[A-Za-z0-9]+:\{[0-9A-Fa-f-]+\}\{\d+\}$/;
const DATA_ID_RE = /^#[A-Za-z0-9_.:-]{1,200}$/;

/**
 * Validate an update target: `body`, `title`, `#<data-id>`, or a generated
 * element ID such as `div:{33f8a242-7c33-4bb2-90c5-8425a68cc5bf}{40}` as
 * returned by read_onenote_page in html format.
 */
export function updateTarget(value: unknown): string {
  if (typeof value !== 'string' || value === '') throw new ToolCallError('target is required');
  if (value === 'body' || value === 'title' || DATA_ID_RE.test(value) || GENERATED_ID_RE.test(value)) return value;
  throw new ToolCallError(
    "target must be 'body', 'title', '#<data-id>', or an element id from read_onenote_page (for example 'div:{guid}{40}')",
  );
}

/**
 * Read a page's HTML through a stream, refusing past MAX_ONENOTE_PAGE_BYTES.
 */
export async function readPageHtml(graph: Client, pageId: string, includeIds: boolean, limit = MAX_ONENOTE_PAGE_BYTES): Promise<string> {
  const path = `/me/onenote/pages/${encodeGraphId(pageId, 'pageId')}/content${includeIds ? '?includeIDs=true' : ''}`;
  const stream = (await graph.api(path).getStream()) as AsyncIterable<unknown>;
  const chunks: Buffer[] = [];
  let total = 0;
  for await (const chunk of stream) {
    const buf = Buffer.isBuffer(chunk) ? chunk : typeof chunk === 'string' ? Buffer.from(chunk) : Buffer.from(chunk as Uint8Array);
    total += buf.length;
    if (total > limit) throw new ToolCallError(`The page is over the ${limit}-byte limit for reading a OneNote page`);
    chunks.push(buf);
  }
  return Buffer.concat(chunks).toString('utf8');
}

const BLOCK_TAGS = new Set(['p', 'div', 'br', 'li', 'tr', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'table', 'ul', 'ol', 'title']);
const SKIP_TAGS = new Set(['script', 'style', 'head']);
const NAMED_ENTITIES = new Map<string, string>([
  ['amp', '&'], ['lt', '<'], ['gt', '>'], ['quot', '"'], ['apos', "'"], ['nbsp', ' '],
]);

function decodeEntities(text: string): string {
  return text.replace(/&(#x[0-9a-fA-F]+|#\d+|[a-zA-Z]+);/g, (whole, ref: string) => {
    if (ref.startsWith('#x') || ref.startsWith('#X')) {
      const cp = parseInt(ref.slice(2), 16);
      return Number.isFinite(cp) && cp <= 0x10ffff ? String.fromCodePoint(cp) : whole;
    }
    if (ref.startsWith('#')) {
      const cp = parseInt(ref.slice(1), 10);
      return Number.isFinite(cp) && cp <= 0x10ffff ? String.fromCodePoint(cp) : whole;
    }
    return NAMED_ENTITIES.get(ref.toLowerCase()) ?? whole;
  });
}

/**
 * Plain text from a page's HTML: tags dropped, block elements on their own
 * lines with blank lines removed, list items bulleted, entities decoded. The page title (from <title>)
 * comes first. Written as a scanner rather than tag-stripping regexes so a
 * malformed tag can't leave half of itself behind.
 */
export function onenoteHtmlToText(html: string): string {
  let out = '';
  let titleText = '';
  let skipDepth = 0;
  let inTitle = false;
  let i = 0;
  while (i < html.length) {
    const ch = html.charAt(i);
    if (ch !== '<') {
      const next = html.indexOf('<', i);
      const end = next === -1 ? html.length : next;
      const run = html.slice(i, end);
      if (inTitle) titleText += run;
      else if (skipDepth === 0) out += run.replace(/\s+/g, ' ');
      i = end;
      continue;
    }
    const close = html.indexOf('>', i + 1);
    if (close === -1) break; // unterminated tag: drop the remainder
    const raw = html.slice(i + 1, close).trim();
    i = close + 1;
    if (raw.startsWith('!') || raw.startsWith('?')) continue; // doctype, comment, PI
    const isEnd = raw.startsWith('/');
    const name = (isEnd ? raw.slice(1) : raw).split(/[\s/>]/, 1)[0].toLowerCase();
    if (name === 'title') {
      inTitle = !isEnd;
      continue;
    }
    if (SKIP_TAGS.has(name)) {
      if (isEnd) skipDepth = Math.max(0, skipDepth - 1);
      else if (!raw.endsWith('/')) skipDepth++;
      continue;
    }
    if (skipDepth > 0) continue;
    if (name === 'li' && !isEnd) out += '\n- ';
    else if (name === 'td' || name === 'th') out += isEnd ? '\t' : '';
    else if (BLOCK_TAGS.has(name)) out += '\n';
  }
  const body = decodeEntities(out)
    .split('\n')
    .map((line) => line.replace(/[ \t]+$/g, '').replace(/^ +/, ''))
    .filter((line) => line !== '')
    .join('\n');
  const title = decodeEntities(titleText).replace(/\s+/g, ' ').trim();
  return title ? (body ? `${title}\n\n${body}` : title) : body;
}
