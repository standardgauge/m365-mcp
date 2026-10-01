/**
 * Unit tests for offset-based Graph pagination.
 *
 * `collectPage` walks `@odata.nextLink` pages internally (SharePoint collections
 * do not support `$skip`) until the requested `[offset, offset + limit)` window
 * is buffered, then returns that slice plus continuation metadata.
 */

import { collectPage, type GraphPage } from '../services/graphPaging.js';

/**
 * Build a fake Graph collection of `total` sequential items paged at `pageSize`,
 * and return `{ fetchFirst, fetchNext, calls }` wired the way a handler wires the
 * real Graph client. `calls` records how many network round trips were made.
 */
function fakeCollection(total: number, pageSize: number) {
  const items = Array.from({ length: total }, (_, i) => ({ id: `item-${i}` }));
  const pageAt = (start: number): GraphPage<{ id: string }> => {
    const slice = items.slice(start, start + pageSize);
    const next = start + pageSize < total ? `https://graph/next?$skiptoken=${start + pageSize}` : undefined;
    return { value: slice, ...(next ? { '@odata.nextLink': next } : {}) };
  };
  const calls: string[] = [];
  return {
    calls,
    fetchFirst: () => { calls.push('first'); return Promise.resolve(pageAt(0)); },
    fetchNext: (link: string) => {
      calls.push(link);
      const start = Number(new URL(link).searchParams.get('$skiptoken'));
      return Promise.resolve(pageAt(start));
    },
  };
}

describe('collectPage', () => {
  it('returns the first window from a single page', async () => {
    const c = fakeCollection(9, 200);
    const res = await collectPage(c.fetchFirst, c.fetchNext, 0, 100);
    expect(res.items.map((i) => i.id)).toEqual(Array.from({ length: 9 }, (_, i) => `item-${i}`));
    expect(res.hasMore).toBe(false);
    expect(res.nextOffset).toBe(9);
    expect(c.calls).toEqual(['first']); // one round trip
  });

  it('reports hasMore when the collection extends past the window', async () => {
    const c = fakeCollection(500, 200);
    const res = await collectPage(c.fetchFirst, c.fetchNext, 0, 50);
    expect(res.items).toHaveLength(50);
    expect(res.items[0].id).toBe('item-0');
    expect(res.hasMore).toBe(true);
    expect(res.nextOffset).toBe(50);
  });

  it('walks nextLink pages to reach an offset deep past the first page', async () => {
    const c = fakeCollection(5000, 200);
    const res = await collectPage(c.fetchFirst, c.fetchNext, 3000, 25);
    expect(res.items[0].id).toBe('item-3000');
    expect(res.items).toHaveLength(25);
    expect(res.items[24].id).toBe('item-3024');
    expect(res.hasMore).toBe(true);
    expect(res.nextOffset).toBe(3025);
    // Reached offset 3000 by walking 200-item pages: 1 first + 15 nextLinks.
    expect(c.calls.length).toBe(16);
  });

  it('returns a short final window and hasMore=false at the end of the collection', async () => {
    const c = fakeCollection(210, 200);
    const res = await collectPage(c.fetchFirst, c.fetchNext, 200, 50);
    expect(res.items.map((i) => i.id)).toEqual(['item-200', 'item-201', 'item-202', 'item-203', 'item-204', 'item-205', 'item-206', 'item-207', 'item-208', 'item-209']);
    expect(res.hasMore).toBe(false);
    expect(res.nextOffset).toBe(210);
  });

  it('returns an empty window when the offset is past the end', async () => {
    const c = fakeCollection(30, 200);
    const res = await collectPage(c.fetchFirst, c.fetchNext, 100, 25);
    expect(res.items).toEqual([]);
    expect(res.hasMore).toBe(false);
    expect(res.nextOffset).toBe(100);
  });

  it('halts on a cyclic nextLink rather than looping forever', async () => {
    let n = 0;
    const fetchFirst = () => Promise.resolve<GraphPage<{ id: string }>>({ value: [{ id: 'a' }], '@odata.nextLink': 'https://graph/loop' });
    const fetchNext = () => {
      n++;
      // Always returns the same nextLink — a cycle.
      return Promise.resolve<GraphPage<{ id: string }>>({ value: [{ id: `b${n}` }], '@odata.nextLink': 'https://graph/loop' });
    };
    const res = await collectPage(fetchFirst, fetchNext, 0, 1000);
    // Followed the loop exactly once (then detected the repeat and stopped).
    expect(n).toBe(1);
    expect(res.items.map((i) => i.id)).toEqual(['a', 'b1']);
  });
});
