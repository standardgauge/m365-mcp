/**
 * Unit tests for the truncation-aware result envelope.
 */
import { resolveMaxResults, toEnvelope, graphCollectionHasMore } from '../services/resultEnvelope.js';

describe('resolveMaxResults', () => {
  it('uses the default when nothing is requested', () => {
    expect(resolveMaxResults(undefined, 25, 50)).toBe(25);
  });
  it('honors a valid request under the cap', () => {
    expect(resolveMaxResults(10, 25, 50)).toBe(10);
  });
  it('clamps a request above the cap to the cap', () => {
    expect(resolveMaxResults(500, 25, 50)).toBe(50);
  });
  it('clamps the default to the cap when the default exceeds it', () => {
    expect(resolveMaxResults(undefined, 100, 50)).toBe(50);
  });
  it('falls back to the default for junk / non-positive input', () => {
    expect(resolveMaxResults('abc', 25, 50)).toBe(25);
    expect(resolveMaxResults(0, 25, 50)).toBe(25);
    expect(resolveMaxResults(-4, 25, 50)).toBe(25);
  });
  it('coerces a numeric string', () => {
    expect(resolveMaxResults('12', 25, 50)).toBe(12);
  });
});

describe('toEnvelope', () => {
  it('reports not-truncated when under the limit and no server signal', () => {
    const env = toEnvelope([1, 2, 3], 25, false);
    expect(env).toEqual({ items: [1, 2, 3], count: 3, limit: 25, truncated: false });
  });

  it('reports truncated when the server says more are available', () => {
    const env = toEnvelope([1, 2, 3], 25, true);
    expect(env.truncated).toBe(true);
    expect(env.count).toBe(3);
  });

  it('reports truncated and slices when more rows than the limit were collected', () => {
    const env = toEnvelope([1, 2, 3, 4, 5], 3, false);
    expect(env.items).toEqual([1, 2, 3]);
    expect(env.count).toBe(3);
    expect(env.truncated).toBe(true);
  });

  it('echoes the applied limit', () => {
    expect(toEnvelope([], 42, false).limit).toBe(42);
  });
});

describe('graphCollectionHasMore', () => {
  it('is true when @odata.nextLink is present', () => {
    expect(graphCollectionHasMore({ value: [], '@odata.nextLink': 'https://graph/next' })).toBe(true);
  });
  it('is false when absent', () => {
    expect(graphCollectionHasMore({ value: [] })).toBe(false);
  });
  it('is false for a non-object', () => {
    expect(graphCollectionHasMore(undefined)).toBe(false);
    expect(graphCollectionHasMore(null)).toBe(false);
  });
});
