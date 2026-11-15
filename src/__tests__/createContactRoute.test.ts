/**
 * REST POST /api/contacts — the create route refuses a contact it could not
 * give a display name, the same as create_contact over MCP.
 *
 * The security and policy wrappers are stubbed to pass the request straight
 * to the handler; Graph is mocked at the client boundary.
 */

import { jest } from '@jest/globals';
import type { HttpRequest, InvocationContext } from '@azure/functions';

interface GraphCall { path: string; body?: unknown }
const graphCalls: GraphCall[] = [];
const AUTH = { userId: 'test-user', session: { userId: 'test-user', tenantId: 'test-tenant', accessToken: 'fake', sessionToken: 'sess' } };

jest.mock('../services/telemetry.js', () => ({}));

jest.mock('../services/securityHeaders.js', () => ({
  withSecurity: (h: unknown) => h,
}));

jest.mock('../services/policyEnforcement.js', () => ({
  withPolicyEnforcement:
    (_service: string, h: (req: unknown, ctx: unknown, auth: unknown) => unknown) =>
    (req: unknown, ctx: unknown) =>
      h(req, ctx, AUTH),
  checkDenyList: () => Promise.resolve(null),
}));

jest.mock('../services/tokenCache.js', () => ({
  getValidAccessTokenForSession: () => Promise.resolve('access-token'),
  getTenantId: () => Promise.resolve('test-tenant'),
}));

jest.mock('../services/graphClient.js', () => ({
  createGraphClient: () => ({
    api: (path: string) => ({
      post: (body: Record<string, unknown>) => {
        graphCalls.push({ path, body });
        return Promise.resolve({ id: 'new-id', displayName: body.displayName });
      },
    }),
  }),
}));

jest.mock('../services/denyList.js', () => ({
  isPathDenied: () => Promise.resolve(false),
}));

jest.mock('../services/containerResolver.js', () => ({
  resolveDefaultContactFolder: () => Promise.resolve(null),
}));

jest.mock('@azure/functions', () => ({
  app: { http: jest.fn() },
}));

import { app } from '@azure/functions';
import '../functions/contacts/createContact.js';

type Handler = (req: HttpRequest, ctx: InvocationContext) => Promise<{ status: number; jsonBody?: unknown }>;
const httpMock = app.http as unknown as jest.Mock<(name: string, opts: { handler: Handler }) => void>;
const registration = httpMock.mock.calls.find((call) => call[0] === 'createContact');
if (!registration) throw new Error('createContact handler was not registered');
const handler = registration[1].handler;

function post(body: Record<string, unknown>) {
  const req = { method: 'POST', json: () => Promise.resolve(body) } as unknown as HttpRequest;
  return handler(req, { error: jest.fn(), warn: jest.fn(), log: jest.fn() } as unknown as InvocationContext);
}

beforeEach(() => {
  graphCalls.length = 0;
});

describe('POST /api/contacts — display name', () => {
  it('returns 400 and sends nothing when every name field is blank', async () => {
    const res = await post({ givenName: '   ', surname: ' ', companyName: '', displayName: '  ' });
    expect(res.status).toBe(400);
    expect((res.jsonBody as { error: string }).error).toMatch(/needs a name/);
    expect(graphCalls).toHaveLength(0);
  });

  it('composes the display name on a plain create', async () => {
    const res = await post({ givenName: 'Jane', surname: 'Doe' });
    expect(res.status).toBe(201);
    expect(graphCalls[0].body).toMatchObject({ displayName: 'Jane Doe' });
  });
});
