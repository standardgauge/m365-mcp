/**
 * resolveDenySubject: whose per-user deny lists apply to an operation on a
 * given mailbox (threat model §9.5).
 */
import type { Client } from '@microsoft/microsoft-graph-client';
import { resolveDenySubject, clearMailboxOwnerCache } from '../services/mailboxOwner.js';

const CALLER = 'caller-oid';

function graphReturning(result: unknown) {
  const paths: string[] = [];
  const graph = {
    api: (path: string) => {
      paths.push(path);
      const chain = {
        select: () => chain,
        get: () => (result instanceof Error ? Promise.reject(result) : Promise.resolve(result)),
      };
      return chain;
    },
  } as unknown as Client;
  return { graph, paths };
}

beforeEach(() => clearMailboxOwnerCache());

describe('resolveDenySubject', () => {
  it.each([undefined, '', 'me'])('is the caller alone for their own mailbox (%p), with no lookup', async (mailboxId) => {
    const { graph, paths } = graphReturning({ id: 'x' });
    expect(await resolveDenySubject(graph, CALLER, mailboxId)).toBe(CALLER);
    expect(paths).toHaveLength(0);
  });

  it("is the caller alone when mailboxId is the caller's own object ID", async () => {
    const { graph, paths } = graphReturning({ id: 'x' });
    expect(await resolveDenySubject(graph, CALLER, CALLER.toUpperCase())).toBe(CALLER);
    expect(paths).toHaveLength(0);
  });

  it('is the caller and the owner for another mailbox, looked up by UPN', async () => {
    const { graph, paths } = graphReturning({ id: 'owner-oid' });
    expect(await resolveDenySubject(graph, CALLER, 'owner@example.com')).toEqual([CALLER, 'owner-oid']);
    expect(paths).toEqual(['/users/owner%40example.com']);
  });

  it("is the caller alone when the UPN turns out to be the caller's", async () => {
    const { graph } = graphReturning({ id: CALLER });
    expect(await resolveDenySubject(graph, CALLER, 'me@example.com')).toBe(CALLER);
  });

  it('caches the owner lookup', async () => {
    const { graph, paths } = graphReturning({ id: 'owner-oid' });
    await resolveDenySubject(graph, CALLER, 'owner@example.com');
    await resolveDenySubject(graph, CALLER, 'OWNER@example.com');
    expect(paths).toHaveLength(1);
  });

  it('throws when the owner cannot be looked up', async () => {
    const { graph } = graphReturning(new Error('Resource not found'));
    await expect(resolveDenySubject(graph, CALLER, 'owner@example.com')).rejects.toThrow('Resource not found');
  });

  it('throws when the directory returns no ID', async () => {
    const { graph } = graphReturning({});
    await expect(resolveDenySubject(graph, CALLER, 'owner@example.com')).rejects.toThrow(/owner of that mailbox/);
  });

  it('does not cache a failed lookup', async () => {
    await expect(resolveDenySubject(graphReturning({}).graph, CALLER, 'owner@example.com')).rejects.toThrow();
    expect(await resolveDenySubject(graphReturning({ id: 'owner-oid' }).graph, CALLER, 'owner@example.com'))
      .toEqual([CALLER, 'owner-oid']);
  });
});
