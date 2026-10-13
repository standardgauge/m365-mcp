/**
 * Post-deploy functional smoke (scripts/post-deploy-smoke.js).
 *
 * The probe runs only against a live instance, in the deploy workflow, so a
 * mistake in it would otherwise surface as a red deploy on every tracking fork
 * or, worse, as a check that passes for the wrong reason. Two things keep it
 * honest here:
 *
 *   1. Its tables agree with the server. The probe carries its own copy of the
 *      write-tool list, the key-refused list and a read and a write per
 *      service; each is checked against mcpEndpoint.ts, and every write probe
 *      is checked to stop at argument validation even if the read-only gate
 *      were gone.
 *   2. It fails on each regression it exists to catch. runSmoke is driven
 *      against a fake instance that behaves correctly, then against one
 *      variant per regression.
 */
import { readFileSync } from 'fs';
import { join } from 'path';
import { assertOpaqueIds } from '../services/opaqueId.js';

type Check = { name: string; ok: boolean | 'skip'; detail: string };
type Result = { checks: Check[]; failed: number };
type Probe = { tool: string; args: (now?: Date) => Record<string, unknown> };
type Smoke = {
  SERVICES: string[];
  READ_PROBES: Record<string, Probe>;
  WRITE_PROBES: Record<string, { tool: string; args: Record<string, unknown> }>;
  WRITE_TOOLS: string[];
  KEY_REFUSED_TOOLS: string[];
  REFUSED_ID: string;
  parseServices: (v?: string) => string[] | null;
  runSmoke: (o: {
    baseUrl: string; key: string; expectedSha?: string; expectServices?: string[] | null;
    fetchImpl: FakeFetch; now?: Date;
  }) => Promise<Result>;
  report: (r: Result, o: { baseUrl: string; expectedSha?: string }) => string;
  main: (argv: string[], env: Record<string, string | undefined>) => Promise<number>;
};

const smoke = require('../../scripts/post-deploy-smoke.js') as Smoke;

const root = join(__dirname, '..', '..');
const endpoint = readFileSync(join(root, 'src', 'functions', 'mcp', 'mcpEndpoint.ts'), 'utf8');

/** The quoted names inside `const <name> ... = new Set([ ... ]);`. */
function setLiteral(name: string): string[] {
  const m = endpoint.match(new RegExp(`const ${name}[^=]*= new Set\\(\\[([\\s\\S]*?)\\]\\)`));
  if (!m) throw new Error(`${name} not found in mcpEndpoint.ts`);
  return [...m[1].replace(/\/\/.*$/gm, '').matchAll(/'([a-z_]+)'/g)].map((x) => x[1]);
}

/** TOOL_SERVICE_MAP as a Map of tool -> service. */
function toolServiceMap(): Map<string, string> {
  const m = endpoint.match(/const TOOL_SERVICE_MAP[^=]*= \{([\s\S]*?)\n\};/);
  if (!m) throw new Error('TOOL_SERVICE_MAP not found in mcpEndpoint.ts');
  return new Map([...m[1].matchAll(/([a-z_]+): '([a-z]+)'/g)].map((x) => [x[1], x[2]]));
}

/** The property names a tool's inputSchema declares. */
function schemaProperties(tool: string): string[] {
  const start = endpoint.indexOf(`name: '${tool}'`);
  if (start < 0) throw new Error(`${tool} not defined in mcpEndpoint.ts`);
  const slice = endpoint.slice(start, endpoint.indexOf('handler:', start));
  const props = slice.match(/properties: \{([\s\S]*?)\}\s*,\s*\n?\s*required/);
  if (!props) return [];
  return [...props[1].matchAll(/(?:^|[\s{,])([A-Za-z]+): prop\(/g)].map((x) => x[1]);
}

describe('post-deploy smoke tables agree with the server', () => {
  const services = toolServiceMap();
  const writeTools = setLiteral('WRITE_TOOLS');

  it('knows every service the server maps a tool to', () => {
    expect([...new Set(services.values())].sort()).toEqual([...smoke.SERVICES].sort());
  });

  it('carries the server WRITE_TOOLS exactly', () => {
    expect([...smoke.WRITE_TOOLS].sort()).toEqual([...writeTools].sort());
  });

  it('carries the server KEY_REFUSED_TOOLS exactly', () => {
    expect([...smoke.KEY_REFUSED_TOOLS].sort()).toEqual(setLiteral('KEY_REFUSED_TOOLS').sort());
  });

  it.each(smoke.SERVICES)('%s: the read probe is a read tool of that service, called with declared arguments', (svc) => {
    const { tool, args } = smoke.READ_PROBES[svc];
    expect(services.get(tool)).toBe(svc);
    expect(writeTools).not.toContain(tool);
    const declared = schemaProperties(tool);
    for (const arg of Object.keys(args(new Date('2026-01-15T12:00:00Z')))) expect(declared).toContain(arg);
  });

  it.each(smoke.SERVICES)('%s: the write probe is a write tool of that service that cannot reach Graph', (svc) => {
    const { tool, args } = smoke.WRITE_PROBES[svc];
    expect(services.get(tool)).toBe(svc);
    expect(writeTools).toContain(tool);
    const declared = schemaProperties(tool);
    for (const arg of Object.keys(args)) expect(declared).toContain(arg);
    // The same validation tools/call runs after the read-only gate. Rejecting
    // here means a broken gate still stops before any Graph call.
    const opaque = new Set(Object.keys(args).filter((a) => args[a] === smoke.REFUSED_ID));
    expect(opaque.size).toBeGreaterThan(0);
    expect(() => assertOpaqueIds(args as Record<string, string>, opaque)).toThrow();
  });

  it('dates the calendar read as a one-day window from today', () => {
    expect(smoke.READ_PROBES.calendar.args(new Date('2026-03-31T23:00:00Z')))
      .toEqual({ startDateTime: '2026-03-31', endDateTime: '2026-04-01' });
  });
});

// ── A fake instance ──

const BASE = 'https://mcp.example.com';
const KEY = 'm365k_fake-smoke-key-value';
const SHA = 'abc1234def';

type FakeResponse = { status: number; headers: { get: (h: string) => string | null }; text: () => Promise<string> };
type FakeFetch = (url: string, init?: { method?: string; headers?: Record<string, string>; body?: string }) => Promise<FakeResponse>;
type Call = { method: string; path: string; rpc?: { method: string; params?: { name?: string; arguments?: Record<string, unknown> } } };

interface Instance {
  held: string[];
  sha: string;
  healthType: string;
  listExtra: string[];
  readError: Record<string, string>;
  writeRefusal: (tool: string) => string | null;
  outsideRefusal: (svc: string) => string | null;
  mailSettings: number;
  admin: number;
  logout: number;
}

function correct(): Instance {
  return {
    held: ['mail', 'calendar', 'sharepoint'],
    sha: SHA,
    healthType: 'application/json; charset=utf-8',
    listExtra: [],
    readError: {},
    writeRefusal: (tool) => `Service "${tool}" is read-only for this agent key; write operations are disabled.`,
    outsideRefusal: (svc) => `Service "${svc}" is not permitted for this agent key. The key's owner can mint a key with wider scope in the web UI.`,
    mailSettings: 403,
    admin: 401,
    logout: 403,
  };
}

function fake(inst: Instance): { fetchImpl: FakeFetch; calls: Call[] } {
  const services = toolServiceMap();
  const writes = new Set(smoke.WRITE_TOOLS);
  const calls: Call[] = [];
  const respond = (status: number, body: unknown, type = 'application/json') => ({
    status,
    headers: { get: (h: string) => (h.toLowerCase() === 'content-type' ? type : null) },
    text: async () => (typeof body === 'string' ? body : JSON.stringify(body)),
  });
  const toolResult = (id: number, text: string, isError = false) =>
    respond(200, { jsonrpc: '2.0', id, result: { content: [{ type: 'text', text }], ...(isError ? { isError: true } : {}) } });

  const fetchImpl: FakeFetch = async (url, init = {}) => {
    const path = new URL(url).pathname;
    const method = init.method ?? 'GET';
    const headers = init.headers ?? {};
    const call: Call = { method, path };
    calls.push(call);
    const keyed = headers.authorization === `Bearer ${KEY}`;

    if (path === '/health') return respond(200, { status: 'ok', sha: inst.sha }, inst.healthType);
    if (path === '/api/mail/settings') {
      return respond(inst.mailSettings, inst.mailSettings === 403 ? { error: 'Email settings cannot be changed with an agent key' } : { emailOutputMode: 'draft' });
    }
    if (path === '/api/manage/agent-keys') return respond(inst.admin, inst.admin === 401 ? { error: 'Authentication required' } : []);
    if (path === '/api/auth/logout') {
      expect(headers.origin).toBe(BASE);
      return respond(inst.logout, inst.logout === 403 ? { error: 'Forbidden' } : '', inst.logout === 403 ? 'application/json' : 'text/plain');
    }
    if (path !== '/api/mcp' || !keyed) return respond(404, 'not found', 'text/plain');

    const msg = JSON.parse(init.body ?? '{}');
    call.rpc = msg;
    if (msg.method === 'initialize') return respond(200, { jsonrpc: '2.0', id: msg.id, result: { protocolVersion: '2024-11-05' } });
    if (msg.method === 'tools/list') {
      const names = [...services.keys()].filter((t) => inst.held.includes(services.get(t)!) && !writes.has(t) && t !== 'set_email_output_mode');
      return respond(200, { jsonrpc: '2.0', id: msg.id, result: { tools: [...names, ...inst.listExtra].map((name) => ({ name })) } });
    }
    if (msg.method === 'tools/call') {
      const tool = msg.params.name as string;
      const svc = services.get(tool)!;
      if (!inst.held.includes(svc)) {
        const r = inst.outsideRefusal(svc);
        return r ? toolResult(msg.id, r, true) : toolResult(msg.id, '{"items":[]}');
      }
      if (writes.has(tool)) {
        const r = inst.writeRefusal(tool);
        return r ? toolResult(msg.id, r, true) : toolResult(msg.id, 'Bad request: messageId contains characters not allowed in an ID', true);
      }
      if (inst.readError[tool]) return toolResult(msg.id, `Error: ${inst.readError[tool]}`, true);
      return toolResult(msg.id, JSON.stringify({ items: [], count: 0 }));
    }
    return respond(200, { jsonrpc: '2.0', id: msg.id, error: { code: -32601, message: 'Unknown method' } });
  };
  return { fetchImpl, calls };
}

async function run(mutate: (i: Instance) => void = () => {}, opts: { expectServices?: string[] } = {}) {
  const inst = correct();
  mutate(inst);
  const { fetchImpl, calls } = fake(inst);
  const result = await smoke.runSmoke({
    baseUrl: `${BASE}/`, key: KEY, expectedSha: SHA, expectServices: opts.expectServices ?? null, fetchImpl,
    now: new Date('2026-01-15T12:00:00Z'),
  });
  const failedNames = result.checks.filter((c) => c.ok === false).map((c) => c.name);
  const toolCalls = calls.filter((c) => c.rpc?.method === 'tools/call').map((c) => c.rpc!.params!);
  return { result, failedNames, calls, toolCalls };
}

describe('post-deploy smoke against a correct instance', () => {
  it('passes every check', async () => {
    const { result } = await run();
    expect(result.checks.filter((c) => c.ok === false)).toEqual([]);
    expect(result.failed).toBe(0);
  });

  it('makes one real read per held service, with its probe arguments', async () => {
    const { toolCalls } = await run();
    for (const svc of ['mail', 'calendar', 'sharepoint']) {
      const { tool, args } = smoke.READ_PROBES[svc];
      expect(toolCalls).toContainEqual({ name: tool, arguments: args(new Date('2026-01-15T12:00:00Z')) });
    }
  });

  it('proves a write refused on every held service, and every outside service refused', async () => {
    const { result } = await run();
    const passed = result.checks.filter((c) => c.ok === true).map((c) => c.name);
    for (const svc of ['mail', 'calendar', 'sharepoint']) expect(passed).toContain(`write ${smoke.WRITE_PROBES[svc].tool} refused`);
    for (const svc of ['onedrive', 'onenote', 'contacts', 'teams']) expect(passed).toContain(`${svc} outside the key refused`);
    for (const name of ['POST /api/mail/settings refused', 'admin API refused', 'logout refused']) expect(passed).toContain(name);
  });

  it('calls logout last', async () => {
    const { calls } = await run();
    expect(calls[calls.length - 1].path).toBe('/api/auth/logout');
  });

  it('notes, rather than fails, a key that holds every service', async () => {
    const { result } = await run((i) => { i.held = [...smoke.SERVICES]; });
    expect(result.failed).toBe(0);
    expect(result.checks).toContainEqual(expect.objectContaining({ name: 'service outside the key refused', ok: 'skip' }));
  });

  it('never puts the key in a check or the report', async () => {
    const { result } = await run((i) => { i.readError.list_sites = 'boom'; });
    expect(JSON.stringify(result)).not.toContain(KEY);
    expect(smoke.report(result, { baseUrl: BASE, expectedSha: SHA })).not.toContain(KEY);
  });
});

describe('post-deploy smoke fails on each regression it exists to catch', () => {
  it('a stale sha on /health', async () => {
    expect((await run((i) => { i.sha = '0000000'; })).failedNames).toContain('/health');
  });

  it('/health answered by the SPA catch-all', async () => {
    expect((await run((i) => { i.healthType = 'text/html'; })).failedNames).toContain('/health');
  });

  it('a live read failing, as the list_teams query option and the consent gap did', async () => {
    const { failedNames, result } = await run((i) => { i.readError.list_events = 'Access is denied. Check credentials and try again.'; });
    expect(failedNames).toEqual(['read list_events']);
    expect(result.checks.find((c) => c.name === 'read list_events')!.detail).toContain('Access is denied');
  });

  it('a key the server did not recognise, and stops before calling any tool', async () => {
    const { failedNames, toolCalls } = await run((i) => { i.listExtra = ['set_email_output_mode']; });
    expect(failedNames).toEqual(['key recognised']);
    expect(toolCalls).toEqual([]);
  });

  it('a write tool in tools/list, and calls no write tool', async () => {
    const { failedNames, toolCalls, result } = await run((i) => { i.listExtra = ['delete_message']; });
    expect(failedNames).toContain('no write tool listed');
    expect(toolCalls.map((c) => c.name).filter((n) => smoke.WRITE_TOOLS.includes(n!))).toEqual([]);
    expect(result.checks).toContainEqual(expect.objectContaining({ name: 'write delete_message refused', ok: 'skip' }));
  });

  it('a write that gets past the read-only gate', async () => {
    const { failedNames } = await run((i) => { i.writeRefusal = (t) => (t === 'delete_event' ? null : correct().writeRefusal(t)); });
    expect(failedNames).toEqual(['write delete_event refused']);
  });

  it('a service outside the key that answers', async () => {
    const { failedNames } = await run((i) => { i.outsideRefusal = (s) => (s === 'teams' ? null : correct().outsideRefusal(s)); });
    expect(failedNames).toEqual(['teams outside the key refused']);
  });

  it('POST /api/mail/settings accepted', async () => {
    expect((await run((i) => { i.mailSettings = 200; })).failedNames).toEqual(['POST /api/mail/settings refused']);
  });

  it('the admin API answering a key', async () => {
    expect((await run((i) => { i.admin = 200; })).failedNames).toEqual(['admin API refused']);
  });

  it('logout signing the owner out', async () => {
    expect((await run((i) => { i.logout = 302; })).failedNames).toEqual(['logout refused']);
  });

  it('a service missing from tools/list when SMOKE_EXPECT_SERVICES names it', async () => {
    const { failedNames, result } = await run((i) => { i.held = ['mail', 'calendar']; }, { expectServices: ['mail', 'calendar', 'teams'] });
    expect(failedNames).toEqual(['services as expected']);
    expect(result.checks.find((c) => c.name === 'services as expected')!.detail).toBe('missing: teams');
  });
});

describe('post-deploy smoke entry point', () => {
  let log: jest.SpyInstance;
  beforeEach(() => { log = jest.spyOn(console, 'log').mockImplementation(() => {}); });
  afterEach(() => log.mockRestore());

  it('skips with a notice when SMOKE_AGENT_KEY is absent', async () => {
    expect(await smoke.main([BASE, SHA], {})).toBe(0);
    expect(log.mock.calls.flat().join('\n')).toMatch(/^::notice::SMOKE_AGENT_KEY is not set/m);
  });

  it('refuses a missing base URL and an unknown expected service', async () => {
    const err = jest.spyOn(console, 'error').mockImplementation(() => {});
    expect(await smoke.main([], { SMOKE_AGENT_KEY: KEY })).toBe(2);
    expect(await smoke.main([BASE], { SMOKE_AGENT_KEY: KEY, SMOKE_EXPECT_SERVICES: 'mail,yammer' })).toBe(2);
    err.mockRestore();
  });

  it('parses SMOKE_EXPECT_SERVICES', () => {
    expect(smoke.parseServices(' Mail, calendar ,')).toEqual(['mail', 'calendar']);
    expect(smoke.parseServices('')).toBeNull();
  });
});

describe('deploy workflow runs the smoke', () => {
  const wf = readFileSync(join(root, '.github', 'workflows', 'deploy.yml'), 'utf8');
  const job = wf.slice(wf.indexOf('\n  functional-smoke:'));

  it('as its own job after the deploy, with the key from a repository secret', () => {
    expect(job).toMatch(/\n {4}needs: build-push-deploy\n/);
    expect(job).toContain('SMOKE_AGENT_KEY: ${{ secrets.SMOKE_AGENT_KEY }}');
    expect(job).toContain('node scripts/post-deploy-smoke.js "https://${FQDN}" "$EXPECTED_SHA"');
    expect(job).toContain('EXPECTED_SHA: ${{ env.DEPLOY_SHA }}');
    expect(wf).toContain('fqdn: ${{ steps.smoke.outputs.fqdn }}');
  });

  it('opens an issue only when the probe itself failed', () => {
    expect(job).toContain("if: failure() && steps.probe.outcome == 'failure'");
    expect(job).toMatch(/permissions:\n\s+contents: read\n\s+issues: write/);
  });
});
