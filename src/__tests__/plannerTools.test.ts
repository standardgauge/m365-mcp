/**
 * Planner MCP tools: plans, buckets and tasks.
 *
 * Driven through the real /api/mcp JSON-RPC dispatch so service gating,
 * read-only mode, opaque-ID validation and the deny list are exercised the way
 * a client reaches them. Graph is a route table: each test sets what a path
 * returns and reads back the calls made, including the If-Match header every
 * Planner update and delete needs.
 */
import { jest } from '@jest/globals';
import type { HttpRequest, InvocationContext } from '@azure/functions';

// ── Graph route table ────────────────────────────────────────────────────────

interface GraphCall { verb: string; path: string; body?: unknown; headers: Record<string, string> }
const graphCalls: GraphCall[] = [];
let routes: Record<string, unknown> = {};

function routeResult(verb: string, path: string): Promise<unknown> {
  const key = `${verb} ${path}`;
  if (!(key in routes)) return Promise.reject(Object.assign(new Error(`unrouted ${key}`), { statusCode: 404 }));
  const value = routes[key];
  return value instanceof Error ? Promise.reject(value) : Promise.resolve(value);
}

function makeChain(path: string) {
  const headers: Record<string, string> = {};
  const chain: Record<string, unknown> = {};
  for (const m of ['select', 'top', 'orderby', 'filter', 'expand']) chain[m] = () => chain;
  chain.header = (k: string, v: string) => {
    headers[k] = v;
    return chain;
  };
  for (const verb of ['get', 'post', 'patch', 'delete']) {
    chain[verb] = (body?: unknown) => {
      graphCalls.push({ verb, path, body, headers: { ...headers } });
      return routeResult(verb, path);
    };
  }
  return chain;
}

// ── Deny list: an entry matches a plan by ID or by title ────────────────────

let deniedPlans: string[] = [];
const matches = (v: unknown) => typeof v === 'string' && deniedPlans.includes(v);

const mockGetEnabledServices = jest.fn<() => Promise<string[]>>();
const mockGetReadOnlyServices = jest.fn<() => Promise<string[]>>();

jest.mock('../services/telemetry.js', () => ({}));
jest.mock('../services/authMiddleware.js', () => ({
  authenticateRequest: () => Promise.resolve({ userId: 'u1', session: { userId: 'u1', tenantId: 't1', accessToken: 'x', sessionToken: 's' } }),
}));
jest.mock('../services/tokenCache.js', () => ({
  getValidAccessTokenForSession: () => Promise.resolve('access-token'),
  getTenantIdFromSession: () => 't1',
}));
jest.mock('../services/graphClient.js', () => ({
  createGraphClient: () => ({ api: (path: string) => makeChain(path) }),
}));
jest.mock('../services/denyList.js', () => ({
  isPathDenied: () => Promise.resolve(false),
  filterDeniedSearchHits: (_t: unknown, _u: unknown, items: unknown) => Promise.resolve(items),
  filterDeniedPaths: (_t: unknown, _u: unknown, type: string, items: Array<{ id?: string; name?: string }>) =>
    Promise.resolve(type === 'planner' ? items.filter((i) => !matches(i.id) && !matches(i.name)) : items),
}));
jest.mock('../services/serviceSettings.js', () => ({
  getEnabledServices: () => mockGetEnabledServices(),
  getReadOnlyServices: () => mockGetReadOnlyServices(),
  getAllowedSites: () => Promise.resolve([]),
}));
jest.mock('../services/userServiceOverrides.js', () => ({
  getUserServiceOverrides: () => Promise.resolve([]),
}));
jest.mock('../services/userEmailSettings.js', () => ({
  getUserEmailSettings: () => Promise.resolve({ emailOutputMode: 'draft' }),
  setUserEmailSettings: () => Promise.resolve(undefined),
}));
jest.mock('../services/userMailConfig.js', () => ({
  isMailIndexingDisabled: () => Promise.resolve(false),
}));
jest.mock('../services/auditLog.js', () => ({ logAccess: jest.fn() }));
jest.mock('@azure/functions', () => ({ app: { http: jest.fn() } }));

import { app } from '@azure/functions';
import '../functions/mcp/mcpEndpoint.js';

interface HttpRegistration {
  handler: (req: HttpRequest, context: InvocationContext) => Promise<{ status: number; jsonBody?: unknown }>;
}
const httpMock = app.http as unknown as jest.Mock<(name: string, opts: HttpRegistration) => void>;
const registration = httpMock.mock.calls.find((call) => call[0] === 'mcpEndpoint');
if (!registration) throw new Error('mcpEndpoint handler was not registered');
const handler = registration[1].handler;

function rpc(method: string, params: Record<string, unknown>) {
  const payload = { jsonrpc: '2.0', id: 1, method, params };
  const req = {
    method: 'POST',
    headers: new Map<string, string>(),
    json: () => Promise.resolve(payload),
    text: () => Promise.resolve(JSON.stringify(payload)),
  } as unknown as HttpRequest;
  const ctx = { error: jest.fn(), warn: jest.fn(), log: jest.fn() } as unknown as InvocationContext;
  return handler(req, ctx);
}

async function call(name: string, args: Record<string, unknown>): Promise<{ parsed: any; isError: boolean; text: string }> {
  const res = await rpc('tools/call', { name, arguments: args });
  const body = res.jsonBody as { result?: { content?: Array<{ text: string }>; isError?: boolean } };
  const text = body.result?.content?.[0]?.text ?? '';
  const isError = body.result?.isError ?? false;
  return { parsed: isError ? null : JSON.parse(text), isError, text };
}

async function listedTools(): Promise<string[]> {
  const res = await rpc('tools/list', {});
  return (res.jsonBody as { result: { tools: Array<{ name: string }> } }).result.tools.map((t) => t.name);
}

const writes = () => graphCalls.filter((c) => c.verb !== 'get');

const PLAN = { id: 'plan-1', title: 'Launch', container: { containerId: 'group-1', type: 'group' }, createdDateTime: '2026-11-01T00:00:00Z' };
const TASK = {
  '@odata.etag': 'W/"task-etag"',
  id: 'task-1',
  title: 'Draft brief',
  planId: 'plan-1',
  bucketId: 'bucket-1',
  percentComplete: 50,
  priority: 3,
  dueDateTime: '2026-11-10T12:00:00Z',
  assignments: { 'aaaaaaaa-0000-0000-0000-000000000001': {}, 'aaaaaaaa-0000-0000-0000-000000000002': {} },
  hasDescription: true,
  checklistItemCount: 1,
};
const DETAILS = {
  '@odata.etag': 'W/"details-etag"',
  description: 'Two pages',
  checklist: { c1: { title: 'Outline', isChecked: true } },
};

const PLANNER_TOOLS = [
  'list_planner_plans', 'list_planner_buckets', 'list_planner_tasks', 'get_planner_task',
  'create_planner_plan', 'create_planner_bucket', 'update_planner_bucket',
  'create_planner_task', 'update_planner_task', 'delete_planner_task',
];
const PLANNER_WRITES = [
  'create_planner_plan', 'create_planner_bucket', 'update_planner_bucket',
  'create_planner_task', 'update_planner_task', 'delete_planner_task',
];

beforeEach(() => {
  graphCalls.length = 0;
  deniedPlans = [];
  mockGetEnabledServices.mockResolvedValue(['mail', 'planner']);
  mockGetReadOnlyServices.mockResolvedValue([]);
  routes = {
    'get /planner/plans/plan-1': PLAN,
    'get /planner/tasks/task-1': TASK,
    'get /planner/tasks/task-1/details': DETAILS,
    'get /planner/buckets/bucket-1': { '@odata.etag': 'W/"bucket-etag"', id: 'bucket-1', name: 'To do', planId: 'plan-1' },
  };
});

describe('tools/list', () => {
  it('lists the Planner tools when the planner service is enabled', async () => {
    const names = await listedTools();
    for (const t of PLANNER_TOOLS) expect(names).toContain(t);
  });

  it('hides every Planner tool when the service is not enabled', async () => {
    mockGetEnabledServices.mockResolvedValue(['mail', 'teams']);
    const names = await listedTools();
    for (const t of PLANNER_TOOLS) expect(names).not.toContain(t);
  });

  it('hides the write tools, and only those, when planner is read-only', async () => {
    mockGetReadOnlyServices.mockResolvedValue(['planner']);
    const names = await listedTools();
    for (const t of PLANNER_WRITES) expect(names).not.toContain(t);
    for (const t of PLANNER_TOOLS.filter((x) => !PLANNER_WRITES.includes(x))) expect(names).toContain(t);
  });

  it('advertises the same input-schema shape in getManifest() and tools/list', async () => {
    const { getManifest } = await import('../mcp/toolManifest.js');
    const manifest = getManifest() as { tools: Array<{ name: string; inputSchema: any }> };
    const res = await rpc('tools/list', {});
    const live = (res.jsonBody as { result: { tools: Array<{ name: string; inputSchema: any }> } }).result.tools;
    const shape = (schema: any) => ({
      required: [...(schema.required ?? [])].sort(),
      props: Object.fromEntries(
        Object.entries(schema.properties ?? {}).map(([k, v]: [string, any]) => [k, { type: v.type, items: v.items?.type, enum: v.enum }]),
      ),
    });
    for (const name of PLANNER_TOOLS) {
      const m = manifest.tools.find((t) => t.name === name);
      const l = live.find((t) => t.name === name);
      expect(m).toBeDefined();
      expect(l).toBeDefined();
      expect(shape(m!.inputSchema)).toEqual(shape(l!.inputSchema));
    }
  });
});

describe('read-only mode at tools/call', () => {
  it('refuses a Planner write without touching Graph', async () => {
    mockGetReadOnlyServices.mockResolvedValue(['planner']);
    const r = await call('delete_planner_task', { taskId: 'task-1' });
    expect(r.isError).toBe(true);
    expect(r.text).toMatch(/read-only/);
    expect(graphCalls).toEqual([]);
  });
});

describe('list_planner_plans', () => {
  it('lists the plans shared with the user by default', async () => {
    routes['get /me/planner/plans'] = { value: [PLAN] };
    const r = await call('list_planner_plans', {});
    expect(r.isError).toBe(false);
    expect(r.parsed.items).toEqual([
      { id: 'plan-1', title: 'Launch', containerId: 'group-1', containerType: 'group', createdDateTime: '2026-11-01T00:00:00Z' },
    ]);
    expect(r.parsed).toMatchObject({ count: 1, limit: 100, truncated: false });
  });

  it("lists a group's plans when groupId is given", async () => {
    routes['get /groups/group-1/planner/plans'] = { value: [PLAN] };
    const r = await call('list_planner_plans', { groupId: 'group-1' });
    expect(r.parsed.count).toBe(1);
  });

  it('drops a plan the deny list names by title', async () => {
    deniedPlans = ['Launch'];
    routes['get /me/planner/plans'] = { value: [PLAN, { id: 'plan-2', title: 'Ops' }] };
    const r = await call('list_planner_plans', {});
    expect(r.parsed.items.map((p: { id: string }) => p.id)).toEqual(['plan-2']);
  });

  it('reports truncated when Graph has another page', async () => {
    routes['get /me/planner/plans'] = { value: [PLAN], '@odata.nextLink': 'https://graph.microsoft.com/v1.0/next' };
    routes['get https://graph.microsoft.com/v1.0/next'] = { value: [] };
    const r = await call('list_planner_plans', { maxResults: 1 });
    expect(r.parsed.truncated).toBe(true);
  });
});

describe('list_planner_buckets', () => {
  it('returns buckets in board order', async () => {
    routes['get /planner/plans/plan-1/buckets'] = {
      value: [
        { id: 'b2', name: 'Done', planId: 'plan-1', orderHint: '8585' },
        { id: 'b1', name: 'To do', planId: 'plan-1', orderHint: '8584' },
      ],
    };
    const r = await call('list_planner_buckets', { planId: 'plan-1' });
    expect(r.parsed.items.map((b: { id: string }) => b.id)).toEqual(['b1', 'b2']);
  });

  it('refuses a denied plan before listing it', async () => {
    deniedPlans = ['plan-1'];
    const r = await call('list_planner_buckets', { planId: 'plan-1' });
    expect(r.isError).toBe(true);
    expect(r.text).toMatch(/deny list/);
    expect(graphCalls.map((c) => c.path)).not.toContain('/planner/plans/plan-1/buckets');
  });
});

describe('list_planner_tasks', () => {
  it('maps status, priority and assignees', async () => {
    routes['get /planner/plans/plan-1/tasks'] = { value: [TASK] };
    const r = await call('list_planner_tasks', { planId: 'plan-1' });
    expect(r.parsed.items[0]).toMatchObject({
      id: 'task-1',
      status: 'inProgress',
      percentComplete: 50,
      priority: 'important',
      dueDateTime: '2026-11-10T12:00:00Z',
      assigneeIds: ['aaaaaaaa-0000-0000-0000-000000000001', 'aaaaaaaa-0000-0000-0000-000000000002'],
    });
  });

  it('resolves a bucket to its plan and checks the plan against the deny list', async () => {
    deniedPlans = ['Launch'];
    const r = await call('list_planner_tasks', { bucketId: 'bucket-1' });
    expect(r.isError).toBe(true);
    expect(r.text).toMatch(/deny list/);
  });

  it('lists a bucket when its plan is allowed', async () => {
    routes['get /planner/buckets/bucket-1/tasks'] = { value: [TASK] };
    const r = await call('list_planner_tasks', { bucketId: 'bucket-1' });
    expect(r.parsed.count).toBe(1);
  });

  it('refuses planId and bucketId together', async () => {
    const r = await call('list_planner_tasks', { planId: 'plan-1', bucketId: 'bucket-1' });
    expect(r.isError).toBe(true);
    expect(r.text).toMatch(/not both/);
    expect(graphCalls).toEqual([]);
  });

  it('leaves out completed tasks when includeCompleted is false', async () => {
    routes['get /planner/plans/plan-1/tasks'] = { value: [TASK, { ...TASK, id: 'task-2', percentComplete: 100 }] };
    const r = await call('list_planner_tasks', { planId: 'plan-1', includeCompleted: false });
    expect(r.parsed.items.map((t: { id: string }) => t.id)).toEqual(['task-1']);
  });

  it("drops tasks in denied or unreadable plans from the user's own task list", async () => {
    deniedPlans = ['Launch'];
    routes['get /me/planner/tasks'] = {
      value: [TASK, { ...TASK, id: 'task-2', planId: 'plan-2' }, { ...TASK, id: 'task-3', planId: 'plan-gone' }],
    };
    routes['get /planner/plans/plan-2'] = { id: 'plan-2', title: 'Ops' };
    const r = await call('list_planner_tasks', {});
    expect(r.parsed.items.map((t: { id: string }) => t.id)).toEqual(['task-2']);
  });
});

describe('get_planner_task', () => {
  it('merges the task with its description and checklist', async () => {
    const r = await call('get_planner_task', { taskId: 'task-1' });
    expect(r.parsed).toMatchObject({
      id: 'task-1',
      description: 'Two pages',
      checklist: [{ id: 'c1', title: 'Outline', isChecked: true }],
    });
  });

  it('refuses a task in a denied plan without reading its details', async () => {
    deniedPlans = ['plan-1'];
    const r = await call('get_planner_task', { taskId: 'task-1' });
    expect(r.isError).toBe(true);
    expect(graphCalls.map((c) => c.path)).not.toContain('/planner/tasks/task-1/details');
  });
});

describe('create_planner_plan', () => {
  it('creates the plan in the group container', async () => {
    routes['post /planner/plans'] = PLAN;
    const r = await call('create_planner_plan', { groupId: 'group-1', title: 'Launch' });
    expect(r.parsed).toMatchObject({ id: 'plan-1', status: 'created' });
    expect(writes()).toEqual([
      expect.objectContaining({
        verb: 'post',
        path: '/planner/plans',
        body: { container: { url: 'https://graph.microsoft.com/v1.0/groups/group-1' }, title: 'Launch' },
      }),
    ]);
  });

  it('refuses a title the deny list blocks', async () => {
    deniedPlans = ['Board'];
    const r = await call('create_planner_plan', { groupId: 'group-1', title: 'Board' });
    expect(r.isError).toBe(true);
    expect(writes()).toEqual([]);
  });
});

describe('buckets', () => {
  it('creates a bucket first on the board', async () => {
    routes['post /planner/buckets'] = { id: 'b9', name: 'Later', planId: 'plan-1', orderHint: ' !' };
    const r = await call('create_planner_bucket', { planId: 'plan-1', name: 'Later' });
    expect(r.parsed).toMatchObject({ id: 'b9', status: 'created' });
    expect(writes()[0].body).toEqual({ name: 'Later', planId: 'plan-1', orderHint: ' !' });
  });

  it('renames a bucket with the ETag it just read', async () => {
    routes['patch /planner/buckets/bucket-1'] = undefined;
    const r = await call('update_planner_bucket', { bucketId: 'bucket-1', name: 'Doing' });
    expect(r.parsed).toMatchObject({ id: 'bucket-1', name: 'Doing', status: 'updated' });
    expect(writes()).toEqual([
      expect.objectContaining({ verb: 'patch', body: { name: 'Doing' }, headers: { 'If-Match': 'W/"bucket-etag"' } }),
    ]);
  });
});

describe('create_planner_task', () => {
  beforeEach(() => {
    routes['post /planner/tasks'] = { ...TASK, id: 'task-9', title: 'Ship', assignments: {} };
    routes['get /planner/tasks/task-9/details'] = { '@odata.etag': 'W/"new-details"', description: '' };
    routes['patch /planner/tasks/task-9/details'] = undefined;
    routes['get /users/pat%40fabrikam.com'] = { id: 'BBBBBBBB-0000-0000-0000-000000000009' };
  });

  it('maps the friendly fields onto the Graph task body', async () => {
    const r = await call('create_planner_task', {
      planId: 'plan-1',
      title: 'Ship',
      bucketId: 'bucket-1',
      dueDateTime: '2026-11-20',
      status: 'notStarted',
      priority: 'urgent',
      assigneeIds: ['pat@fabrikam.com', 'aaaaaaaa-0000-0000-0000-000000000001'],
    });
    expect(r.isError).toBe(false);
    expect(r.parsed).toMatchObject({ id: 'task-9', status: 'created' });
    const post = writes().find((c) => c.path === '/planner/tasks')!;
    expect(post.body).toEqual({
      planId: 'plan-1',
      title: 'Ship',
      bucketId: 'bucket-1',
      dueDateTime: '2026-11-20T12:00:00.000Z',
      percentComplete: 0,
      priority: 1,
      assignments: {
        'bbbbbbbb-0000-0000-0000-000000000009': { '@odata.type': '#microsoft.graph.plannerAssignment', orderHint: ' !' },
        'aaaaaaaa-0000-0000-0000-000000000001': { '@odata.type': '#microsoft.graph.plannerAssignment', orderHint: ' !' },
      },
    });
  });

  it('writes the description to the details resource with its own ETag', async () => {
    await call('create_planner_task', { planId: 'plan-1', title: 'Ship', description: 'Notes here' });
    expect(writes().map((c) => [c.verb, c.path])).toEqual([
      ['post', '/planner/tasks'],
      ['patch', '/planner/tasks/task-9/details'],
    ]);
    expect(writes()[1]).toMatchObject({ body: { description: 'Notes here' }, headers: { 'If-Match': 'W/"new-details"' } });
  });

  it('reports a description that failed to save instead of claiming success', async () => {
    routes['patch /planner/tasks/task-9/details'] = new Error('boom');
    const r = await call('create_planner_task', { planId: 'plan-1', title: 'Ship', description: 'Notes here' });
    expect(r.isError).toBe(false);
    expect(r.parsed.descriptionWarning).toMatch(/description was not saved/);
  });

  it('refuses a denied plan before writing', async () => {
    deniedPlans = ['plan-1'];
    const r = await call('create_planner_task', { planId: 'plan-1', title: 'Ship' });
    expect(r.isError).toBe(true);
    expect(writes()).toEqual([]);
  });

  it('refuses a bad status, date or assignee before writing', async () => {
    for (const args of [
      { status: 'done' },
      { priority: 'high' },
      { dueDateTime: 'next tuesday' },
      { assigneeIds: ['not-an-id'] },
    ]) {
      const r = await call('create_planner_task', { planId: 'plan-1', title: 'Ship', ...args });
      expect(r.isError).toBe(true);
    }
    expect(writes()).toEqual([]);
  });

  it('rejects a path-shaped planId before any Graph call', async () => {
    const r = await call('create_planner_task', { planId: '../../users/x', title: 'Ship' });
    expect(r.isError).toBe(true);
    expect(graphCalls).toEqual([]);
  });
});

describe('update_planner_task', () => {
  beforeEach(() => {
    routes['patch /planner/tasks/task-1'] = { ...TASK, percentComplete: 100 };
    routes['patch /planner/tasks/task-1/details'] = undefined;
  });

  it('PATCHes only the passed fields with the ETag it just read', async () => {
    const r = await call('update_planner_task', { taskId: 'task-1', status: 'completed', dueDateTime: '' });
    expect(r.parsed).toMatchObject({ id: 'task-1', status: 'updated' });
    expect(writes()).toEqual([
      expect.objectContaining({
        verb: 'patch',
        path: '/planner/tasks/task-1',
        body: { percentComplete: 100, dueDateTime: null },
        headers: { 'If-Match': 'W/"task-etag"', Prefer: 'return=representation' },
      }),
    ]);
  });

  it('replaces the assignee set, nulling whoever was dropped', async () => {
    await call('update_planner_task', {
      taskId: 'task-1',
      assigneeIds: ['aaaaaaaa-0000-0000-0000-000000000002', 'aaaaaaaa-0000-0000-0000-000000000003'],
    });
    expect(writes()[0].body).toEqual({
      assignments: {
        'aaaaaaaa-0000-0000-0000-000000000001': null,
        'aaaaaaaa-0000-0000-0000-000000000003': { '@odata.type': '#microsoft.graph.plannerAssignment', orderHint: ' !' },
      },
    });
  });

  it('changes only the description through the details resource', async () => {
    await call('update_planner_task', { taskId: 'task-1', description: 'Three pages' });
    expect(writes()).toEqual([
      expect.objectContaining({ path: '/planner/tasks/task-1/details', body: { description: 'Three pages' }, headers: { 'If-Match': 'W/"details-etag"' } }),
    ]);
  });

  it('refuses a call that changes nothing', async () => {
    const r = await call('update_planner_task', { taskId: 'task-1' });
    expect(r.isError).toBe(true);
    expect(r.text).toMatch(/Nothing to update/);
    expect(graphCalls).toEqual([]);
  });

  it('refuses a task in a denied plan', async () => {
    deniedPlans = ['Launch'];
    const r = await call('update_planner_task', { taskId: 'task-1', title: 'x' });
    expect(r.isError).toBe(true);
    expect(writes()).toEqual([]);
  });
});

describe('delete_planner_task', () => {
  it('deletes with the ETag it just read', async () => {
    routes['delete /planner/tasks/task-1'] = undefined;
    const r = await call('delete_planner_task', { taskId: 'task-1' });
    expect(r.parsed).toEqual({ id: 'task-1', planId: 'plan-1', status: 'deleted' });
    expect(writes()).toEqual([expect.objectContaining({ verb: 'delete', headers: { 'If-Match': 'W/"task-etag"' } })]);
  });

  it('refuses a task in a denied plan', async () => {
    deniedPlans = ['plan-1'];
    const r = await call('delete_planner_task', { taskId: 'task-1' });
    expect(r.isError).toBe(true);
    expect(writes()).toEqual([]);
  });
});
