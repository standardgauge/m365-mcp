/**
 * Desktop extension auto-update hardening (threat model section 4.3, G10).
 *
 *   - Signing: /api/extension-update carries a payload signed by this
 *     instance's key, and the extension only applies a payload that verifies
 *     against the public key baked into the code it was served.
 *   - Containment: every payload path must resolve inside the extension
 *     directory, checked for the whole set before anything is written.
 *   - Fixed origin: the URL baked into served code comes from configuration,
 *     never from Host or X-Forwarded-Host.
 *
 * The client-side checks run against src/install/extension-update.js, the
 * module inlined verbatim into the generated server/index.js.
 */

import { jest } from '@jest/globals';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import * as vm from 'vm';
import { generateKeyPairSync, randomBytes, sign } from 'crypto';
import type { HttpRequest, InvocationContext } from '@azure/functions';

const httpMock = jest.fn();
jest.mock('@azure/functions', () => ({ app: { http: httpMock } }));

const REDIRECT_URI = 'https://mcp.example.com/api/auth/callback';
process.env.MCP_INSTANCE_NAME = 'Test M365 MCP';
process.env.OAUTH_REDIRECT_URI = REDIRECT_URI;
process.env.MCP_SESSION_HMAC_KEY = randomBytes(32).toString('hex');

import { configuredPublicOrigin } from '../functions/install/installEndpoint.js';

interface Signed { payload: string; signature: string }
interface UpdatePayload { version: string; files: Record<string, string> }
interface UpdateModule {
  isNewerVersion: (remote: string, local: string) => boolean;
  verifySignedUpdate: (signed: unknown, publicKeyB64: string, localVersion: string) => UpdatePayload;
  resolveUpdatePath: (extDir: string, filePath: string) => string;
  applyUpdate: (extDir: string, files: Record<string, unknown>) => void;
}
// eslint-disable-next-line @typescript-eslint/no-require-imports
const update = require('../install/extension-update.js') as UpdateModule;

type Handler = (req: HttpRequest, ctx: InvocationContext) => Promise<{
  status: number;
  body?: string | Buffer;
  jsonBody?: unknown;
}>;

function handler(name: string): Handler {
  const reg = httpMock.mock.calls.find((call) => call[0] === name);
  if (!reg) throw new Error(`Handler '${name}' not registered`);
  return (reg[1] as { handler: Handler }).handler;
}

/** A request whose headers name a host the attacker controls. */
function forgedRequest(): HttpRequest {
  return {
    headers: new Map<string, string>([
      ['host', 'evil.example.net'],
      ['x-forwarded-proto', 'http'],
      ['x-forwarded-host', 'evil.example.net'],
    ]),
    query: { get: () => null, has: () => false },
  } as unknown as HttpRequest;
}

function ctx(): InvocationContext {
  return { error: jest.fn(), warn: jest.fn(), log: jest.fn() } as unknown as InvocationContext;
}

async function fetchUpdate(): Promise<{ version: string; files: Record<string, string>; signed: Signed }> {
  const res = await handler('extensionUpdate')(forgedRequest(), ctx());
  expect(res.status).toBe(200);
  return res.jsonBody as { version: string; files: Record<string, string>; signed: Signed };
}

function bakedPublicKey(serverJs: string): string {
  const m = serverJs.match(/const UPDATE_PUBLIC_KEY = '([A-Za-z0-9+/=]+)';/);
  if (!m) throw new Error('UPDATE_PUBLIC_KEY not found in served code');
  return m[1];
}

function tmpDir(): string {
  return fs.mkdtempSync(path.join(os.tmpdir(), 'ext-update-'));
}

// ── Signing ──────────────────────────────────────────────────────────────────

describe('update signing', () => {
  it('signs the update with the key whose public half the served code carries', async () => {
    const body = await fetchUpdate();
    const key = bakedPublicKey(body.files['server/index.js']);
    const payload = update.verifySignedUpdate(body.signed, key, '2.9.0');
    expect(payload.version).toBe(body.version);
    expect(payload.files).toEqual(body.files);
  });

  it('the bundle and the update carry the same key, stable across requests', async () => {
    const a = bakedPublicKey((await fetchUpdate()).files['server/index.js']);
    const b = bakedPublicKey((await fetchUpdate()).files['server/index.js']);
    expect(a).toBe(b);
  });

  it('refuses a payload altered after signing', async () => {
    const body = await fetchUpdate();
    const key = bakedPublicKey(body.files['server/index.js']);
    const tampered = JSON.parse(body.signed.payload) as UpdatePayload;
    tampered.files['server/index.js'] = 'require("child_process").exec("curl evil")';
    expect(() =>
      update.verifySignedUpdate({ ...body.signed, payload: JSON.stringify(tampered) }, key, '2.9.0'),
    ).toThrow(/signature does not verify/);
  });

  it('refuses a payload signed by any other key', async () => {
    const body = await fetchUpdate();
    const key = bakedPublicKey(body.files['server/index.js']);
    const other = generateKeyPairSync('ed25519').privateKey;
    const forged = JSON.stringify({ version: '99.0.0', files: { 'server/index.js': 'evil()' } });
    const signature = sign(null, Buffer.from(forged), other).toString('base64');
    expect(() => update.verifySignedUpdate({ payload: forged, signature }, key, '2.10.0'))
      .toThrow(/signature does not verify/);
  });

  it('refuses an unsigned update, which is all a pre-2.10 response carries', async () => {
    const body = await fetchUpdate();
    const key = bakedPublicKey(body.files['server/index.js']);
    expect(() => update.verifySignedUpdate(undefined, key, '2.9.0')).toThrow(/not signed/);
    expect(() => update.verifySignedUpdate({ payload: body.signed.payload }, key, '2.9.0'))
      .toThrow(/not signed/);
  });

  it('refuses a validly signed payload that is not newer, so old ones cannot be replayed', async () => {
    const body = await fetchUpdate();
    const key = bakedPublicKey(body.files['server/index.js']);
    expect(() => update.verifySignedUpdate(body.signed, key, body.version)).toThrow(/not newer/);
    expect(() => update.verifySignedUpdate(body.signed, key, '99.0.0')).toThrow(/not newer/);
  });

  it('served server/index.js only applies the signed block', async () => {
    const serverJs = (await fetchUpdate()).files['server/index.js'];
    expect(serverJs).toContain('extensionUpdate.verifySignedUpdate(update.signed, UPDATE_PUBLIC_KEY, EXTENSION_VERSION)');
    expect(serverJs).toContain('extensionUpdate.applyUpdate(');
    expect(serverJs).not.toContain('Object.entries(update.files)');
  });

  it('served server/index.js inlines the tested module verbatim and parses', async () => {
    const serverJs = (await fetchUpdate()).files['server/index.js'];
    const src = fs.readFileSync(path.join(process.cwd(), 'src', 'install', 'extension-update.js'), 'utf-8');
    expect(serverJs).toContain(src);
    expect(() => new vm.Script(serverJs.replace(/^#!.*\n/, ''))).not.toThrow();
  });
});

// ── Path containment ─────────────────────────────────────────────────────────

describe('update path containment', () => {
  const extDir = path.resolve(os.tmpdir(), 'ext-root');

  it.each([
    '../outside.js',
    'server/../../outside.js',
    '..',
    '/etc/passwd',
    'C:\\Windows\\evil.js',
    'C:evil.js',
    '\\\\server\\share\\evil.js',
    '',
    'server/\0.js',
  ])('refuses %j', (p) => {
    expect(() => update.resolveUpdatePath(extDir, p)).toThrow();
  });

  it.each(['manifest.json', 'server/index.js', './package.json'])('accepts %j', (p) => {
    expect(update.resolveUpdatePath(extDir, p).startsWith(extDir + path.sep)).toBe(true);
  });

  it('writes nothing when any one path in the payload escapes', () => {
    const dir = tmpDir();
    fs.writeFileSync(path.join(dir, 'manifest.json'), 'old');
    expect(() =>
      update.applyUpdate(dir, { 'manifest.json': 'new', '../escaped.js': 'evil' }),
    ).toThrow(/escapes/);
    expect(fs.readFileSync(path.join(dir, 'manifest.json'), 'utf8')).toBe('old');
    expect(fs.existsSync(path.join(dir, '..', 'escaped.js'))).toBe(false);
  });

  it('refuses a write that a symlinked directory would redirect outside', () => {
    const dir = tmpDir();
    const outside = tmpDir();
    fs.symlinkSync(outside, path.join(dir, 'server'), 'dir');
    expect(() => update.applyUpdate(dir, { 'server/index.js': 'evil' })).toThrow(/via a link/);
    expect(fs.readdirSync(outside)).toEqual([]);
  });

  it('writes nothing when a symlink escape follows a valid file', () => {
    const dir = tmpDir();
    const outside = tmpDir();
    fs.writeFileSync(path.join(dir, 'manifest.json'), 'old');
    fs.symlinkSync(outside, path.join(dir, 'server'), 'dir');
    expect(() =>
      update.applyUpdate(dir, { 'manifest.json': 'new', 'server/index.js': 'evil' }),
    ).toThrow(/via a link/);
    expect(fs.readFileSync(path.join(dir, 'manifest.json'), 'utf8')).toBe('old');
    expect(fs.readdirSync(dir).sort()).toEqual(['manifest.json', 'server']);
    expect(fs.readdirSync(outside)).toEqual([]);
  });

  it('writes nothing when a later path runs through a broken link or a file', () => {
    const dir = tmpDir();
    fs.writeFileSync(path.join(dir, 'manifest.json'), 'old');
    fs.symlinkSync(path.join(dir, 'missing'), path.join(dir, 'server'), 'dir');
    fs.writeFileSync(path.join(dir, 'lib'), 'a file');
    expect(() =>
      update.applyUpdate(dir, { 'manifest.json': 'new', 'server/index.js': 'evil' }),
    ).toThrow(/broken link/);
    expect(() =>
      update.applyUpdate(dir, { 'manifest.json': 'new', 'lib/x.js': 'evil' }),
    ).toThrow(/through a file/);
    expect(fs.readFileSync(path.join(dir, 'manifest.json'), 'utf8')).toBe('old');
  });

  it.each([
    [{ 'manifest.json': 'new', server: 'not a dir', 'server/index.js': 'new' }, /needs a directory/],
    [{ 'manifest.json': 'new', 'server/index.js': 'new', server: 'not a dir' }, /needs a directory/],
    [{ 'manifest.json': 'new', 'Server': 'x', 'server/index.js': 'new' }, /needs a directory/],
    [{ 'manifest.json': 'new', './x.js': 'a', 'x.js': 'b' }, /same path twice/],
    [{ 'manifest.json': 'new', 'X.js': 'a', 'x.js': 'b' }, /same path twice/],
    [{ 'manifest.json': 'new', 'x.js': 'a', 'x.js.tmp': 'b' }, /temp name/],
    [{ 'manifest.json': 'new', 'x.js.tmp': 'b', 'x.js': 'a' }, /temp name/],
    [{ 'manifest.json': 'new', 'x.tmp/index.js': 'a', x: 'b' }, /temp name/],
    [{ 'manifest.json': 'new', x: 'b', 'x.tmp/index.js': 'a' }, /temp name/],
    [{ 'manifest.json': 'new', 'X.TMP/a/b.js': 'a', x: 'b' }, /temp name/],
  ])('writes nothing when payload entries collide with each other (%#)', (files, err) => {
    const dir = tmpDir();
    fs.writeFileSync(path.join(dir, 'manifest.json'), 'old');
    expect(() => update.applyUpdate(dir, files)).toThrow(err);
    expect(fs.readdirSync(dir)).toEqual(['manifest.json']);
    expect(fs.readFileSync(path.join(dir, 'manifest.json'), 'utf8')).toBe('old');
  });

  it('writes nothing when a later target or its temp name is an existing directory', () => {
    const dir = tmpDir();
    fs.writeFileSync(path.join(dir, 'manifest.json'), 'old');
    fs.mkdirSync(path.join(dir, 'server'));
    fs.mkdirSync(path.join(dir, 'lib.js.tmp'));
    expect(() => update.applyUpdate(dir, { 'manifest.json': 'new', server: 'x' })).toThrow(/existing directory/);
    expect(() => update.applyUpdate(dir, { 'manifest.json': 'new', 'lib.js': 'x' })).toThrow(/existing directory/);
    expect(fs.readFileSync(path.join(dir, 'manifest.json'), 'utf8')).toBe('old');
  });

  it('does not follow a link planted at the temp name', () => {
    const dir = tmpDir();
    const outside = tmpDir();
    const target = path.join(outside, 'victim');
    fs.writeFileSync(target, 'untouched');
    fs.symlinkSync(target, path.join(dir, 'manifest.json.tmp'));
    update.applyUpdate(dir, { 'manifest.json': 'new' });
    expect(fs.readFileSync(target, 'utf8')).toBe('untouched');
    expect(fs.readFileSync(path.join(dir, 'manifest.json'), 'utf8')).toBe('new');
  });

  it('refuses non-text content', () => {
    const dir = tmpDir();
    expect(() => update.applyUpdate(dir, { 'manifest.json': { a: 1 } })).toThrow(/not text/);
  });

  it('applies a well-formed payload inside the extension directory', async () => {
    const dir = tmpDir();
    const body = await fetchUpdate();
    update.applyUpdate(dir, body.files);
    expect(fs.readFileSync(path.join(dir, 'server', 'index.js'), 'utf8')).toBe(body.files['server/index.js']);
    expect(fs.readFileSync(path.join(dir, 'manifest.json'), 'utf8')).toBe(body.files['manifest.json']);
  });
});

// ── Fixed origin ─────────────────────────────────────────────────────────────

describe('served origin comes from configuration', () => {
  it('derives the origin from OAUTH_REDIRECT_URI', () => {
    expect(configuredPublicOrigin(REDIRECT_URI)).toBe('https://mcp.example.com');
    expect(configuredPublicOrigin('http://localhost:7071/api/auth/callback')).toBe('http://localhost:7071');
  });

  it('refuses to guess when it is unset or not an absolute http(s) URL', () => {
    expect(() => configuredPublicOrigin('')).toThrow(/not set/);
    expect(() => configuredPublicOrigin('/api/auth/callback')).toThrow(/absolute/);
    expect(() => configuredPublicOrigin('javascript:alert(1)')).toThrow(/http/);
  });

  it.each([
    'installLanding',
    'installBundle',
    'extensionUpdate',
    'installSh',
    'installPs1',
  ])('%s ignores Host and X-Forwarded-Host', async (name) => {
    const res = await handler(name)(forgedRequest(), ctx());
    expect(res.status).toBe(200);
    const text = res.body !== undefined
      ? Buffer.isBuffer(res.body) ? res.body.toString('latin1') : String(res.body)
      : JSON.stringify(res.jsonBody);
    expect(text).not.toContain('evil.example.net');
    if (name !== 'installBundle') expect(text).toContain('https://mcp.example.com');
  });

  it('serves no code when the origin is not configured', async () => {
    const saved = process.env.OAUTH_REDIRECT_URI;
    delete process.env.OAUTH_REDIRECT_URI;
    try {
      for (const name of ['installLanding', 'installBundle', 'extensionUpdate', 'installSh', 'installPs1']) {
        const res = await handler(name)(forgedRequest(), ctx());
        expect(res.status).toBe(500);
      }
    } finally {
      process.env.OAUTH_REDIRECT_URI = saved;
    }
  });
});
