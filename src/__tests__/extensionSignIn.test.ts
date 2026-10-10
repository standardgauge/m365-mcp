/**
 * The desktop extension connects stdio before it signs in.
 *
 * A launch that had to sign in used to block in authenticate() before reading
 * stdin. An MCP client kills a server that has not connected within its connect
 * timeout (30s in Claude Code), which killed the polling verifier too, so the
 * sign-in page it had opened could never finish and every new session opened
 * another.
 *
 * These tests run the generated server/index.js as a real process against a
 * fake server, with stub keyring and browser commands on PATH, and check:
 *   - initialize answers at once and a launch opens no browser on its own
 *   - a tool call without a session answers sign-in-required and starts the
 *     sign-in in the background, which then completes without a restart
 *   - a relaunch adopts the sign-in in flight instead of opening another page
 */

import { jest } from '@jest/globals';
import { spawn, ChildProcess } from 'child_process';
import * as crypto from 'crypto';
import * as fs from 'fs';
import * as http from 'http';
import * as os from 'os';
import * as path from 'path';
import * as readline from 'readline';
import type { AddressInfo } from 'net';
import type { HttpRequest, InvocationContext } from '@azure/functions';

const httpMock = jest.fn();

jest.mock('@azure/functions', () => ({
  app: { http: httpMock },
}));

process.env.MCP_INSTANCE_NAME = 'Test M365 MCP';
// Update signing derives its key from the HMAC key.
process.env.MCP_SESSION_HMAC_KEY = crypto.randomBytes(32).toString('hex');

import '../functions/install/installEndpoint.js';

const SLUG = 'test-m365';
const POSIX = process.platform !== 'win32';
const describePosix = POSIX ? describe : describe.skip;

jest.setTimeout(30000);

type Handler = (req: HttpRequest, ctx: InvocationContext) => Promise<{ jsonBody?: unknown }>;

function getHandler(name: string): Handler {
  const reg = httpMock.mock.calls.find((call) => call[0] === name);
  if (!reg) throw new Error(`Handler '${name}' not registered`);
  return (reg[1] as { handler: Handler }).handler;
}

async function renderServerJs(origin: string): Promise<string> {
  // The served origin comes from OAUTH_REDIRECT_URI, never request headers
  // (AC-427), so point it at the fake server for this render.
  process.env.OAUTH_REDIRECT_URI = new URL('/api/auth/callback', origin).toString();
  const req = {
    headers: new Map<string, string>(),
    query: { get: () => null, has: () => false },
  } as unknown as HttpRequest;
  const res = await getHandler('extensionUpdate')(req, {} as InvocationContext);
  return (res.jsonBody as { files: Record<string, string> }).files['server/index.js'];
}

function confirmationCode(verifier: string): string {
  const challenge = crypto.createHash('sha256').update(verifier).digest('hex');
  const hex = crypto.createHash('sha256').update('m365-mcp-install-confirm:' + challenge).digest('hex').slice(0, 8).toUpperCase();
  return hex.slice(0, 4) + '-' + hex.slice(4);
}

// ── Fake server ─────────────────────────────────────────────────────────────

interface FakeServer {
  origin: string;
  poll: { status: number; token?: string };
  polledVerifiers: string[];
  mcpCalls: Array<{ auth: string | undefined; method: string }>;
  close: () => Promise<void>;
}

async function startFakeServer(): Promise<FakeServer> {
  const state = {
    poll: { status: 202 } as { status: number; token?: string },
    polledVerifiers: [] as string[],
    mcpCalls: [] as Array<{ auth: string | undefined; method: string }>,
  };
  const server = http.createServer((req, res) => {
    const url = new URL(req.url ?? '/', 'http://localhost');
    let body = '';
    req.on('data', (c) => (body += c));
    req.on('end', () => {
      const send = (status: number, json: unknown) => {
        res.writeHead(status, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify(json));
      };
      if (url.pathname === '/api/auth/install-poll') {
        state.polledVerifiers.push(url.searchParams.get('nonce_verifier') ?? '');
        if (state.poll.status === 200) {
          send(200, { sessionToken: state.poll.token, email: 'adele@fabrikam.com', displayName: 'Adele' });
        } else {
          send(state.poll.status, {});
        }
        return;
      }
      if (url.pathname === '/api/mcp') {
        const msg = JSON.parse(body);
        const auth = req.headers.authorization;
        state.mcpCalls.push({ auth, method: msg.method });
        if (msg.method === 'initialize') {
          send(200, { jsonrpc: '2.0', id: msg.id, result: { protocolVersion: '2024-11-05', capabilities: { tools: { listChanged: false } }, serverInfo: { name: 'fake' } } });
        } else if (msg.method === 'tools/list') {
          send(200, { jsonrpc: '2.0', id: msg.id, result: { tools: [] } });
        } else if (auth === 'Bearer good') {
          send(200, { jsonrpc: '2.0', id: msg.id, result: { content: [{ type: 'text', text: 'ok' }] } });
        } else {
          send(200, { jsonrpc: '2.0', id: msg.id, result: { content: [{ type: 'text', text: 'Session expired. Please re-authenticate at: x' }], isError: true } });
        }
        return;
      }
      send(404, {});
    });
  });
  await new Promise<void>((r) => server.listen(0, '127.0.0.1', r));
  const { port } = server.address() as AddressInfo;
  return {
    origin: `http://127.0.0.1:${port}`,
    get poll() { return state.poll; },
    set poll(v) { state.poll = v; },
    polledVerifiers: state.polledVerifiers,
    mcpCalls: state.mcpCalls,
    close: () => new Promise<void>((r) => { server.closeAllConnections(); server.close(() => r()); }),
  };
}

// ── Extension process ───────────────────────────────────────────────────────

type Msg = any;

class Extension {
  readonly messages: Msg[] = [];
  private nextId = 1;
  constructor(readonly proc: ChildProcess) {
    readline.createInterface({ input: proc.stdout! }).on('line', (line) => {
      try { this.messages.push(JSON.parse(line)); } catch { /* not JSON */ }
    });
  }

  async request(method: string, params: Msg = {}): Promise<Msg> {
    const id = this.nextId++;
    this.proc.stdin!.write(JSON.stringify({ jsonrpc: '2.0', id, method, params }) + '\n');
    return this.waitFor((m) => m.id === id, 5000);
  }

  async waitFor(pred: (m: Msg) => boolean, ms: number): Promise<Msg> {
    const deadline = Date.now() + ms;
    for (;;) {
      const found = this.messages.find(pred);
      if (found) return found;
      if (Date.now() > deadline) throw new Error('timed out waiting for a message');
      await sleep(25);
    }
  }
}

const sleep = (ms: number) => new Promise((r) => setTimeout(r, ms));

async function until(pred: () => boolean, ms: number): Promise<void> {
  const deadline = Date.now() + ms;
  while (!pred()) {
    if (Date.now() > deadline) throw new Error('condition not met in time');
    await sleep(25);
  }
}

let home: string;
let server: FakeServer;
let scriptPath: string;
const procs: ChildProcess[] = [];

function stub(dir: string, name: string, body: string) {
  fs.writeFileSync(path.join(dir, name), `#!/bin/sh\n${body}\n`, { mode: 0o755 });
}

function launch(): Extension {
  const proc = spawn(process.execPath, [scriptPath], {
    env: { HOME: home, PATH: `${path.join(home, 'bin')}${path.delimiter}${process.env.PATH}` },
    stdio: ['pipe', 'pipe', 'pipe'],
  });
  proc.stderr!.resume();
  procs.push(proc);
  return new Extension(proc);
}

const pendingFile = () => path.join(home, `.${SLUG}-signin.json`);
const tokenFile = () => path.join(home, `.${SLUG}-token`);
const openedLog = () => path.join(home, 'opened.log');
const opened = () => (fs.existsSync(openedLog()) ? fs.readFileSync(openedLog(), 'utf8').trim().split('\n').filter(Boolean) : []);
const resultText = (m: Msg) => m.result.content.map((c: { text: string }) => c.text).join(' ');

describePosix('desktop extension sign-in (generated server/index.js)', () => {
  beforeEach(async () => {
    home = fs.realpathSync(fs.mkdtempSync(path.join(os.tmpdir(), 'ext-home-')));
    const bin = path.join(home, 'bin');
    fs.mkdirSync(bin);
    // No keyring on either platform, so the token lives in the 0600 file; the
    // browser openers only record what they were asked to open.
    stub(bin, 'secret-tool', 'exit 1');
    stub(bin, 'security', 'exit 44');
    stub(bin, 'xdg-open', `echo "$1" >> "${openedLog()}"`);
    stub(bin, 'open', `echo "$1" >> "${openedLog()}"`);
    server = await startFakeServer();
    scriptPath = path.join(home, 'index.js');
    fs.writeFileSync(scriptPath, await renderServerJs(server.origin));
  });

  afterEach(async () => {
    for (const p of procs.splice(0)) p.kill('SIGKILL');
    await server.close();
    fs.rmSync(home, { recursive: true, force: true });
  });

  it('answers initialize at once without a session and opens no browser on launch', async () => {
    const ext = launch();
    const started = Date.now();
    const init = await ext.request('initialize', { protocolVersion: '2024-11-05' });
    expect(Date.now() - started).toBeLessThan(5000);
    expect(init.result.capabilities.tools.listChanged).toBe(true);

    const list = await ext.request('tools/list');
    expect(list.result.tools).toEqual([]);
    expect(server.mcpCalls.every((c) => c.auth === undefined)).toBe(true);

    await sleep(2500);
    expect(opened()).toEqual([]);
    expect(fs.existsSync(pendingFile())).toBe(false);
    expect(server.polledVerifiers).toEqual([]);
  });

  it('a tool call without a session answers sign-in-required and the sign-in completes in the background', async () => {
    const ext = launch();
    await ext.request('initialize');

    const first = await ext.request('tools/call', { name: 'list_folders_mail', arguments: {} });
    expect(first.result.isError).toBe(true);
    expect(resultText(first)).toMatch(/Not signed in .*browser page has opened/);
    expect(opened()).toHaveLength(1);

    const pending = JSON.parse(fs.readFileSync(pendingFile(), 'utf8'));
    expect(fs.statSync(pendingFile()).mode & 0o777).toBe(0o600);
    expect(resultText(first)).toContain(confirmationCode(pending.verifier));
    expect(server.mcpCalls.filter((c) => c.method === 'tools/call')).toHaveLength(0);

    // A second call while the sign-in is open does not open another page.
    const second = await ext.request('tools/call', { name: 'list_folders_mail', arguments: {} });
    expect(resultText(second)).toContain(confirmationCode(pending.verifier));
    expect(opened()).toHaveLength(1);

    await until(() => server.polledVerifiers.includes(pending.verifier), 8000);
    server.poll = { status: 200, token: 'good' };
    await ext.waitFor((m) => m.method === 'notifications/tools/list_changed', 8000);

    const after = await ext.request('tools/call', { name: 'list_folders_mail', arguments: {} });
    expect(after.result.isError).toBeUndefined();
    expect(resultText(after)).toBe('ok');
    expect(server.mcpCalls.at(-1)).toEqual({ auth: 'Bearer good', method: 'tools/call' });
    expect(JSON.parse(fs.readFileSync(tokenFile(), 'utf8')).token).toBe('good');
    expect(fs.existsSync(pendingFile())).toBe(false);
  });

  it('a relaunch after the client killed the process adopts the sign-in instead of opening another page', async () => {
    const first = launch();
    await first.request('tools/call', { name: 'list_folders_mail', arguments: {} });
    expect(opened()).toHaveLength(1);
    const { verifier } = JSON.parse(fs.readFileSync(pendingFile(), 'utf8'));
    first.proc.kill('SIGKILL');

    const second = launch();
    const res = await second.request('tools/call', { name: 'list_folders_mail', arguments: {} });
    expect(resultText(res)).toMatch(/already open in your browser/);
    expect(resultText(res)).toContain(confirmationCode(verifier));
    expect(opened()).toHaveLength(1);

    server.polledVerifiers.length = 0;
    await until(() => server.polledVerifiers.length > 0, 8000);
    expect(new Set(server.polledVerifiers)).toEqual(new Set([verifier]));

    server.poll = { status: 200, token: 'good' };
    await until(() => fs.existsSync(tokenFile()), 8000);
    const after = await second.request('tools/call', { name: 'list_folders_mail', arguments: {} });
    expect(resultText(after)).toBe('ok');
  });

  it('a sign-in older than the login window is replaced, not adopted', async () => {
    const stale = 'a'.repeat(32);
    fs.writeFileSync(pendingFile(), JSON.stringify({ verifier: stale, startedAt: Date.now() - 11 * 60 * 1000 }), { mode: 0o600 });

    const ext = launch();
    const res = await ext.request('tools/call', { name: 'list_folders_mail', arguments: {} });
    expect(resultText(res)).toMatch(/browser page has opened/);
    expect(opened()).toHaveLength(1);
    const { verifier } = JSON.parse(fs.readFileSync(pendingFile(), 'utf8'));
    expect(verifier).not.toBe(stale);
  });

  it('a session the server refuses is dropped and the refused call starts sign-in', async () => {
    fs.writeFileSync(tokenFile(), JSON.stringify({ token: 'dead' }), { mode: 0o600 });
    const ext = launch();

    const res = await ext.request('tools/call', { name: 'list_folders_mail', arguments: {} });
    expect(server.mcpCalls.at(-1)).toEqual({ auth: 'Bearer dead', method: 'tools/call' });
    expect(resultText(res)).toMatch(/Not signed in/);
    expect(fs.existsSync(tokenFile())).toBe(false);
    expect(opened()).toHaveLength(1);
  });

  it('picks up a session another process stored while this one was polling', async () => {
    const ext = launch();
    await ext.request('initialize');
    await ext.request('tools/call', { name: 'list_folders_mail', arguments: {} });

    // The server reads a consumed verifier as pending, so only the store can
    // tell this process that a sibling collected the session.
    fs.writeFileSync(tokenFile(), JSON.stringify({ token: 'good' }), { mode: 0o600 });
    await ext.waitFor((m) => m.method === 'notifications/tools/list_changed', 8000);
    const after = await ext.request('tools/call', { name: 'list_folders_mail', arguments: {} });
    expect(resultText(after)).toBe('ok');
    expect(fs.existsSync(pendingFile())).toBe(false);
  });

  it('an expired sign-in clears, and the next tool call opens a fresh one', async () => {
    server.poll = { status: 410 };
    const ext = launch();
    await ext.request('tools/call', { name: 'list_folders_mail', arguments: {} });
    const { verifier } = JSON.parse(fs.readFileSync(pendingFile(), 'utf8'));
    await until(() => !fs.existsSync(pendingFile()), 8000);

    server.poll = { status: 202 };
    const res = await ext.request('tools/call', { name: 'list_folders_mail', arguments: {} });
    expect(resultText(res)).toMatch(/browser page has opened/);
    expect(opened()).toHaveLength(2);
    expect(JSON.parse(fs.readFileSync(pendingFile(), 'utf8')).verifier).not.toBe(verifier);
  });
});
