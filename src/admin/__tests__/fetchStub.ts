import { vi } from 'vitest';

export interface Call { method: string; url: string; body: unknown }
type Reply = { status?: number; body?: unknown };
type Handler = (call: Call) => Reply | undefined;

/**
 * Replace global fetch with a router over `handler`. Every request is recorded
 * in the returned `calls`, and a request the handler does not answer fails the
 * test with a 599 so an unexpected call is visible rather than silently empty.
 */
export function stubFetch(handler: Handler): Call[] {
  const calls: Call[] = [];
  vi.stubGlobal('fetch', vi.fn((input: string, init: RequestInit = {}) => {
    const call: Call = {
      method: init.method ?? 'GET',
      url: input,
      body: typeof init.body === 'string' ? JSON.parse(init.body) as unknown : undefined,
    };
    calls.push(call);
    const reply = handler(call) ?? { status: 599, body: { error: `unexpected ${call.method} ${call.url}` } };
    const status = reply.status ?? 200;
    return Promise.resolve({
      ok: status >= 200 && status < 300,
      status,
      json: () => Promise.resolve(reply.body ?? {}),
    } as Response);
  }));
  return calls;
}
