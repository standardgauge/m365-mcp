/**
 * Tests for the /install endpoints and .mcpb bundle generation.
 *
 * Covers:
 *   - GET /api/extension-version — returns correct version
 *   - GET /api/extension-update — returns files map with correct structure
 *   - GET /install.mcpb — generates valid zip with 3 expected files
 *   - GET /install — returns HTML landing page
 *   - GET /install.ps1 — returns rendered PowerShell installer
 *   - GET /install.sh — returns rendered shell installer
 *   - GET /install/m365-mcp-shim.js — serves the stdio shim both installers use
 *   - isNewerVersion logic (the module inlined into the generated extension)
 *   - Generated server.js contains auto-update code
 *
 * Update signing, path containment and the configured origin (threat model
 * G10) are in extensionUpdate.test.ts.
 */

import { jest } from '@jest/globals';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import * as vm from 'vm';
import type { HttpRequest, InvocationContext } from '@azure/functions';
import * as zlib from 'zlib';
import { randomBytes } from 'crypto';

// ── Mock Azure Functions app ─────────────────────────────────────────────────

const httpMock = jest.fn();

jest.mock('@azure/functions', () => ({
  app: { http: httpMock },
}));

// Set env vars before import
process.env.MCP_INSTANCE_NAME = 'Test M365 MCP';
// The served origin comes from here, not from request headers.
process.env.OAUTH_REDIRECT_URI = 'https://mcp.example.com/api/auth/callback';
// Update signing derives its key from the HMAC key.
process.env.MCP_SESSION_HMAC_KEY = randomBytes(32).toString('hex');

// ── Import after mocks ──────────────────────────────────────────────────────

import '../functions/install/installEndpoint.js';

// Extract registered handlers
interface HttpRegistration {
  handler: (req: HttpRequest, context: InvocationContext) => Promise<{
    status: number;
    body?: string | Buffer;
    jsonBody?: unknown;
    headers?: Record<string, string>;
  }>;
}

function getHandler(name: string): HttpRegistration['handler'] {
  const reg = httpMock.mock.calls.find((call) => call[0] === name);
  if (!reg) throw new Error(`Handler '${name}' not registered`);
  return (reg[1] as HttpRegistration).handler;
}

const installLanding = getHandler('installLanding');
const installBundle = getHandler('installBundle');
const extensionVersion = getHandler('extensionVersion');
const extensionUpdate = getHandler('extensionUpdate');
const installPs1 = getHandler('installPs1');
const installSh = getHandler('installSh');
const installShim = getHandler('installShim');

// ── Test helpers ────────────────────────────────────────────────────────────

function makeRequest(): HttpRequest {
  return {
    headers: new Map<string, string>([
      ['host', 'mcp.example.com'],
      ['x-forwarded-proto', 'https'],
      ['x-forwarded-host', 'mcp.example.com'],
    ]),
    query: { get: () => null, has: () => false },
  } as unknown as HttpRequest;
}

function makeContext(): InvocationContext {
  return {
    error: jest.fn(),
    warn: jest.fn(),
    log: jest.fn(),
  } as unknown as InvocationContext;
}

// Minimal ZIP parser — extract file names and contents from a zip buffer
function parseZipEntries(buf: Buffer): Array<{ name: string; content: string }> {
  const entries: Array<{ name: string; content: string }> = [];
  let offset = 0;

  while (offset < buf.length - 4) {
    const sig = buf.readUInt32LE(offset);
    if (sig !== 0x04034b50) break; // not a local file header

    const compressionMethod = buf.readUInt16LE(offset + 8);
    const compressedSize = buf.readUInt32LE(offset + 18);
    const _uncompressedSize = buf.readUInt32LE(offset + 22);
    const nameLen = buf.readUInt16LE(offset + 26);
    const extraLen = buf.readUInt16LE(offset + 28);

    const nameStart = offset + 30;
    const name = buf.subarray(nameStart, nameStart + nameLen).toString('utf-8');
    const dataStart = nameStart + nameLen + extraLen;
    const compressedData = buf.subarray(dataStart, dataStart + compressedSize);

    let content: string;
    if (compressionMethod === 8) {
      content = zlib.inflateRawSync(compressedData).toString('utf-8');
    } else {
      content = compressedData.toString('utf-8');
    }

    entries.push({ name, content });
    offset = dataStart + compressedSize;
  }

  return entries;
}

// ── isNewerVersion: the real one, from the module inlined into server/index.js ──

// eslint-disable-next-line @typescript-eslint/no-require-imports
const { isNewerVersion } = require('../install/extension-update.js') as {
  isNewerVersion: (remote: string, local: string) => boolean;
};

// ── Tests ────────────────────────────────────────────────────────────────────

describe('GET /api/extension-version', () => {
  it('returns the current extension version', async () => {
    const res = await extensionVersion(makeRequest(), makeContext());
    expect(res.status).toBe(200);
    const body = res.jsonBody as { version: string };
    expect(body.version).toMatch(/^\d+\.\d+\.\d+$/);
  });
});

describe('GET /api/extension-update', () => {
  it('returns all three extension files', async () => {
    const res = await extensionUpdate(makeRequest(), makeContext());
    expect(res.status).toBe(200);
    const body = res.jsonBody as { version: string; files: Record<string, string> };
    expect(body.version).toMatch(/^\d+\.\d+\.\d+$/);
    expect(body.files).toHaveProperty(['manifest.json']);
    expect(body.files).toHaveProperty(['package.json']);
    expect(body.files).toHaveProperty(['server/index.js']);
  });

  it('manifest.json contains valid JSON with correct slug', async () => {
    const res = await extensionUpdate(makeRequest(), makeContext());
    const body = res.jsonBody as { files: Record<string, string> };
    const manifest = JSON.parse(body.files['manifest.json']);
    expect(manifest.name).toBe('test-m365');
    expect(manifest.display_name).toBe('Test M365 MCP');
    expect(manifest.server.entry_point).toBe('server/index.js');
  });

  // The .mcpb manifest is a machine-readable license surface shipped to users,
  // so it has to track the project license rather than drift from it. It carried
  // a stale `UNLICENSED` through the AGPL relicensing; assert against
  // package.json so the next relicensing can only be done in one place.
  it('manifest.json declares the same license as package.json', async () => {
    const pkg = JSON.parse(
      fs.readFileSync(path.join(__dirname, '..', '..', 'package.json'), 'utf8'),
    ) as { license?: string };
    expect(pkg.license).toBeTruthy();

    const res = await extensionUpdate(makeRequest(), makeContext());
    const body = res.jsonBody as { files: Record<string, string> };
    const manifest = JSON.parse(body.files['manifest.json']);
    expect(manifest.license).toBe(pkg.license);
  });

  it('server/index.js contains the correct MCP_URL', async () => {
    const res = await extensionUpdate(makeRequest(), makeContext());
    const body = res.jsonBody as { files: Record<string, string> };
    expect(body.files['server/index.js']).toContain("const MCP_URL = 'https://mcp.example.com'");
  });

  it('server/index.js contains auto-update code', async () => {
    const res = await extensionUpdate(makeRequest(), makeContext());
    const body = res.jsonBody as { files: Record<string, string> };
    const serverJs = body.files['server/index.js'];
    expect(serverJs).toContain('async function autoUpdate()');
    expect(serverJs).toContain('isNewerVersion');
    expect(serverJs).toContain('withTimeout');
    expect(serverJs).toContain('/api/extension-version');
  });

  it('server/index.js authenticate() uses the PKCE nonce flow', async () => {
    const res = await extensionUpdate(makeRequest(), makeContext());
    const body = res.jsonBody as { files: Record<string, string> };
    const serverJs = body.files['server/index.js'];
    // The browser URL must carry only SHA256(verifier); the raw verifier is
    // sent solely to install-poll. A pre-PKCE bundle sent a raw 32-char nonce
    // to both, which the hardened server now 400s.
    expect(serverJs).toContain("crypto.createHash('sha256').update(verifier).digest('hex')");
    expect(serverJs).toContain("'/api/auth/login?install_nonce=' + challenge");
    expect(serverJs).toContain("'/api/auth/install-poll?nonce_verifier=' + verifier");
    // Guard against regressing to the broken flow.
    expect(serverJs).not.toContain("install-poll?nonce=' + nonce");
  });

  it('server/index.js verifyToken() makes an authenticated tools/call and reads the body', async () => {
    const res = await extensionUpdate(makeRequest(), makeContext());
    const body = res.jsonBody as { files: Record<string, string> };
    const serverJs = body.files['server/index.js'];
    // /api/mcp always returns HTTP 200 (auth failure is an isError body, not a
    // 401), and unauthenticated tools/list also returns 200 — so the old
    // 'tools/list + status === 200' check could never detect a dead token.
    expect(serverJs).toContain("method: 'tools/call'");
    expect(serverJs).toContain("name: 'list_folders_mail'");
    expect(serverJs).toContain('/re-authenticate|session expired/i');
    // The verify path must no longer rely on tools/list.
    expect(serverJs).not.toContain("method: 'tools/list', params: {} }),\n      { 'Authorization'");
  });

  it('server/index.js encrypts the token at rest on every platform', async () => {
    const res = await extensionUpdate(makeRequest(), makeContext());
    const body = res.jsonBody as { files: Record<string, string> };
    const serverJs = body.files['server/index.js'];

    // macOS Keychain (existing)
    expect(serverJs).toContain('find-generic-password');
    expect(serverJs).toContain('add-generic-password');

    // Windows DPAPI via PowerShell
    expect(serverJs).toContain('ConvertTo-SecureString');
    expect(serverJs).toContain('ConvertFrom-SecureString');
    expect(serverJs).toContain("enc: 'dpapi'");
    expect(serverJs).toContain('function dpapiEncrypt');
    expect(serverJs).toContain('function dpapiDecrypt');

    // Linux libsecret via secret-tool
    expect(serverJs).toContain('secret-tool');
    expect(serverJs).toContain("'store'");
    expect(serverJs).toContain("'lookup'");
    expect(serverJs).toContain("'clear'");
  });

  it('server/index.js keeps a 0600-file fallback', async () => {
    const res = await extensionUpdate(makeRequest(), makeContext());
    const body = res.jsonBody as { files: Record<string, string> };
    const serverJs = body.files['server/index.js'];
    expect(serverJs).toContain('mode: 0o600');
  });

  it('server/index.js never passes the token via DPAPI/secret-tool argv', async () => {
    const res = await extensionUpdate(makeRequest(), makeContext());
    const body = res.jsonBody as { files: Record<string, string> };
    const serverJs = body.files['server/index.js'];
    // Both secret paths read the secret from stdin (input: token / ReadToEnd),
    // never as a command-line argument that would leak via the process list.
    expect(serverJs).toContain('input: token');
    expect(serverJs).toContain('[Console]::In.ReadToEnd()');
  });
});

// ── Keychain service fallback in the generated extension ───────────────────
//
// Runs the generated server/index.js token functions against a fake keychain
// (security / secret-tool stubbed through child_process), so the fallback is
// tested by behaviour rather than by the strings it contains.

type FakeStore = Map<string, string>;

interface TokenFns {
  loadToken: () => string | null;
  saveToken: (token: string) => void;
  deleteToken: () => void;
  KEYCHAIN_SERVICE: string;
  LEGACY_KEYCHAIN_SERVICE: string;
}

async function renderedServerJs(legacyPrefix?: string): Promise<string> {
  if (legacyPrefix === undefined) delete process.env.EXTENSION_LEGACY_KEYCHAIN_PREFIX;
  else process.env.EXTENSION_LEGACY_KEYCHAIN_PREFIX = legacyPrefix;
  try {
    const res = await extensionUpdate(makeRequest(), makeContext());
    return (res.jsonBody as { files: Record<string, string> }).files['server/index.js'];
  } finally {
    delete process.env.EXTENSION_LEGACY_KEYCHAIN_PREFIX;
  }
}

function loadTokenFns(serverJs: string, platform: string, store: FakeStore, stderr: string[]): TokenFns {
  const home = fs.mkdtempSync(path.join(os.tmpdir(), 'ext-home-'));
  const flag = (args: string[], name: string) => args[args.indexOf(name) + 1];
  const execFileSync = (cmd: string, args: string[], opts: { input?: string } = {}) => {
    if (cmd === 'security') {
      const service = flag(args, '-s');
      if (args[0] === 'find-generic-password') {
        if (!store.has(service)) throw new Error('not found');
        return store.get(service) + '\n';
      }
      if (args[0] === 'add-generic-password') { store.set(service, flag(args, '-w')); return ''; }
      if (args[0] === 'delete-generic-password') {
        if (!store.delete(service)) throw new Error('not found');
        return '';
      }
    }
    if (cmd === 'secret-tool') {
      const service = flag(args, 'service');
      if (args[0] === 'lookup') {
        if (!store.has(service)) throw new Error('not found');
        return store.get(service) + '\n';
      }
      if (args[0] === 'store') { store.set(service, opts.input ?? ''); return ''; }
      if (args[0] === 'clear') { store.delete(service); return ''; }
    }
    throw new Error(`unexpected command ${cmd} ${args.join(' ')}`);
  };
  const modules: Record<string, unknown> = {
    child_process: { execFileSync },
    fs,
    path,
    os: { ...os, homedir: () => home },
    https: {}, http: {}, crypto: {}, readline: {},
  };
  // Drop the trailing main() call so loading the file runs nothing.
  const body = serverJs.slice(serverJs.indexOf('\n'), serverJs.lastIndexOf('\nmain()'));
  const sandbox = {
    require: (name: string) => modules[name],
    process: { platform, env: {}, stderr: { write: (line: string) => { stderr.push(line); } } },
    console,
    setTimeout,
    exports: {} as TokenFns,
  };
  vm.runInNewContext(
    body + '\nexports.loadToken = loadToken; exports.saveToken = saveToken; exports.deleteToken = deleteToken;' +
      '\nexports.KEYCHAIN_SERVICE = KEYCHAIN_SERVICE; exports.LEGACY_KEYCHAIN_SERVICE = LEGACY_KEYCHAIN_SERVICE;',
    sandbox,
  );
  return sandbox.exports;
}

describe('generated extension: legacy keychain service fallback', () => {
  const PREFIX = 'com.fabrikam.';

  it('embeds no legacy service when EXTENSION_LEGACY_KEYCHAIN_PREFIX is unset', async () => {
    const fns = loadTokenFns(await renderedServerJs(), 'darwin', new Map(), []);
    expect(fns.LEGACY_KEYCHAIN_SERVICE).toBe('');
    expect(fns.loadToken()).toBeNull();
  });

  it('carries the legacy service inside the signed payload 2.11.0+ clients apply', async () => {
    process.env.EXTENSION_LEGACY_KEYCHAIN_PREFIX = PREFIX;
    try {
      const res = await extensionUpdate(makeRequest(), makeContext());
      const body = res.jsonBody as { files: Record<string, string>; signed: { payload: string } };
      const signedFiles = (JSON.parse(body.signed.payload) as { files: Record<string, string> }).files;
      expect(signedFiles['server/index.js']).toBe(body.files['server/index.js']);
      expect(signedFiles['server/index.js']).toContain(`const LEGACY_KEYCHAIN_SERVICE = "${PREFIX}`);
    } finally {
      delete process.env.EXTENSION_LEGACY_KEYCHAIN_PREFIX;
    }
  });

  it('ignores a prefix outside the reverse-DNS charset', async () => {
    const serverJs = await renderedServerJs("x'; process.exit(1); '");
    expect(serverJs).toContain('const LEGACY_KEYCHAIN_SERVICE = "";');
  });

  it.each(['darwin', 'linux'])('%s: prefers the current service when both exist', async (platform) => {
    const serverJs = await renderedServerJs(PREFIX);
    const store: FakeStore = new Map();
    const stderr: string[] = [];
    const fns = loadTokenFns(serverJs, platform, store, stderr);
    expect(fns.LEGACY_KEYCHAIN_SERVICE).toBe(PREFIX + fns.KEYCHAIN_SERVICE.split('.').pop());
    store.set(fns.KEYCHAIN_SERVICE, 'current-token');
    store.set(fns.LEGACY_KEYCHAIN_SERVICE, 'legacy-token');
    expect(fns.loadToken()).toBe('current-token');
    expect(stderr).toEqual([]);
  });

  it.each(['darwin', 'linux'])('%s: moves a legacy-only session to the current service', async (platform) => {
    const store: FakeStore = new Map();
    const stderr: string[] = [];
    const fns = loadTokenFns(await renderedServerJs(PREFIX), platform, store, stderr);
    store.set(fns.LEGACY_KEYCHAIN_SERVICE, 'legacy-token');

    expect(fns.loadToken()).toBe('legacy-token');
    expect(store.get(fns.KEYCHAIN_SERVICE)).toBe('legacy-token');
    // The old entry stays so an older version still finds it.
    expect(store.get(fns.LEGACY_KEYCHAIN_SERVICE)).toBe('legacy-token');
    expect(stderr.join('')).toContain('Moved the saved session');
    expect(stderr.join('')).not.toContain('legacy-token');
  });

  it('returns null with nothing under either service', async () => {
    const fns = loadTokenFns(await renderedServerJs(PREFIX), 'darwin', new Map(), []);
    expect(fns.loadToken()).toBeNull();
  });

  it.each(['darwin', 'linux'])('%s: deleteToken removes both entries', async (platform) => {
    const store: FakeStore = new Map();
    const fns = loadTokenFns(await renderedServerJs(PREFIX), platform, store, []);
    store.set(fns.KEYCHAIN_SERVICE, 'dead-token');
    store.set(fns.LEGACY_KEYCHAIN_SERVICE, 'dead-token');
    fns.deleteToken();
    expect(store.size).toBe(0);
    // A dead session is not resurrected from the old name on the next launch.
    expect(fns.loadToken()).toBeNull();
  });
});

describe('GET /install.mcpb', () => {
  it('returns a valid zip with 3 files', async () => {
    const res = await installBundle(makeRequest(), makeContext());
    expect(res.status).toBe(200);
    expect(res.headers?.['Content-Type']).toBe('application/octet-stream');
    expect(res.headers?.['Content-Disposition']).toContain('.mcpb');

    const entries = parseZipEntries(res.body as Buffer);
    const names = entries.map((e) => e.name).sort();
    expect(names).toEqual(['manifest.json', 'package.json', 'server/index.js']);
  });

  it('zip manifest.json has correct structure', async () => {
    const res = await installBundle(makeRequest(), makeContext());
    const entries = parseZipEntries(res.body as Buffer);
    const manifest = JSON.parse(entries.find((e) => e.name === 'manifest.json')!.content);
    expect(manifest.manifest_version).toBe('0.2');
    expect(manifest.server.type).toBe('node');
  });

  it('zip server/index.js contains MCP_URL', async () => {
    const res = await installBundle(makeRequest(), makeContext());
    const entries = parseZipEntries(res.body as Buffer);
    const serverJs = entries.find((e) => e.name === 'server/index.js')!.content;
    expect(serverJs).toContain('https://mcp.example.com');
  });
});

describe('GET /install', () => {
  it('returns HTML landing page', async () => {
    const res = await installLanding(makeRequest(), makeContext());
    expect(res.status).toBe(200);
    expect(res.headers?.['Content-Type']).toContain('text/html');
    expect(res.body as string).toContain('Test M365 MCP');
    expect(res.body as string).toContain('install.mcpb');
  });

  it('landing page contains setup instructions', async () => {
    const res = await installLanding(makeRequest(), makeContext());
    expect(res.body as string).toContain('Desktop App');
    expect(res.body as string).toContain('Extensions');
  });
});

describe('GET /install.ps1', () => {
  it('template is present in dist/install after build (packaging verification)', () => {
    // When dist/ exists (i.e. after `npm run build`), the postbuild step must have
    // copied the template to dist/install/.  This is the path that runs in Docker
    // (process.cwd() = /home/site/wwwroot; only dist/ is present in the image).
    // Running this test suite *after* build in CI catches the packaging gap before deploy.
    const distDir = path.join(process.cwd(), 'dist');
    if (!fs.existsSync(distDir)) return; // pre-build dev run — src/install fallback covers it
    const distTemplate = path.join(distDir, 'install', 'install-mcp.ps1.template');
    expect(fs.existsSync(distTemplate)).toBe(true);
  });

  it('returns a PowerShell script with substituted URL and MCP name', async () => {
    const res = await installPs1(makeRequest(), makeContext());
    expect(res.status).toBe(200);
    expect(res.headers?.['Content-Type']).toContain('text/plain');
    const script = res.body as string;
    expect(script).toContain('https://mcp.example.com');
    expect(script).toContain('test-m365');
    expect(script).not.toContain('{{MCP_URL}}');
    expect(script).not.toContain('{{MCP_NAME}}');
    expect(script).not.toContain('{{INSTANCE_NAME}}');
  });

  it('PS1 template writes all config files with UTF-8 no-BOM', () => {
    const templatePath = path.join(process.cwd(), 'src', 'install', 'install-mcp.ps1.template');
    const template = fs.readFileSync(templatePath, 'utf-8');
    // All JSON writes must use WriteAllText with UTF8Encoding($false) — no BOM
    expect(template).toContain('[System.IO.File]::WriteAllText');
    expect(template).toContain('System.Text.UTF8Encoding $false');
    // Must NOT use Set-Content or Out-File to write the desktop config (produces BOM in PS 5.1)
    expect(template).not.toMatch(/\|\s*Set-Content\s+\$DESKTOP_CONFIG/);
    expect(template).not.toMatch(/Out-File\s+\$DESKTOP_CONFIG/);
  });

  it('PS1 template reads config files with explicit UTF-8 encoding', () => {
    const templatePath = path.join(process.cwd(), 'src', 'install', 'install-mcp.ps1.template');
    const template = fs.readFileSync(templatePath, 'utf-8');
    // Get-Content must use -Encoding UTF8 so PS 5.1 reads UTF-8 files correctly
    expect(template).toMatch(/Get-Content \$ConfigPath -Raw -Encoding UTF8/);
  });

  it('PS1 template uses -Force when adding mcpServers to handle null property', () => {
    const templatePath = path.join(process.cwd(), 'src', 'install', 'install-mcp.ps1.template');
    const template = fs.readFileSync(templatePath, 'utf-8');
    // -Force is required: if "mcpServers": null exists in JSON, ConvertFrom-Json creates the
    // property with a null value. Add-Member without -Force throws on existing properties.
    expect(template).toMatch(/Add-Member.*-Name 'mcpServers'.*-Force/);
  });
});

describe('isNewerVersion (semver comparison)', () => {
  it('detects newer major version', () => {
    expect(isNewerVersion('3.0.0', '2.6.1')).toBe(true);
  });

  it('detects newer minor version', () => {
    expect(isNewerVersion('2.7.0', '2.6.1')).toBe(true);
  });

  it('detects newer patch version', () => {
    expect(isNewerVersion('2.6.2', '2.6.1')).toBe(true);
  });

  it('returns false for same version', () => {
    expect(isNewerVersion('2.6.1', '2.6.1')).toBe(false);
  });

  it('returns false for older major version (no downgrade)', () => {
    expect(isNewerVersion('1.0.0', '2.6.1')).toBe(false);
  });

  it('returns false for older minor version (no downgrade)', () => {
    expect(isNewerVersion('2.5.0', '2.6.1')).toBe(false);
  });

  it('returns false for older patch version (no downgrade)', () => {
    expect(isNewerVersion('2.6.0', '2.6.1')).toBe(false);
  });

  it('handles mismatched segment lengths', () => {
    expect(isNewerVersion('2.7', '2.6.1')).toBe(true);
    expect(isNewerVersion('2.6.1', '2.7')).toBe(false);
  });
});

// ──: first-party stdio shim replaces supergateway ────────────────────

describe('GET /install/m365-mcp-shim.js', () => {
  it('is registered at the route the installers download from', () => {
    const reg = httpMock.mock.calls.find((call) => call[0] === 'installShim');
    expect((reg?.[1] as { route: string }).route).toBe('install/m365-mcp-shim.js');
  });

  it('serves the shim source verbatim as JavaScript', async () => {
    const res = await installShim(makeRequest(), makeContext());
    expect(res.status).toBe(200);
    expect(res.headers?.['Content-Type']).toContain('application/javascript');
    const src = fs.readFileSync(path.join(process.cwd(), 'src', 'install', 'm365-mcp-shim.js'), 'utf-8');
    expect(res.body).toBe(src);
  });

  it('shim is present in dist/install after build (packaging verification)', () => {
    const distDir = path.join(process.cwd(), 'dist');
    if (!fs.existsSync(distDir)) return;
    expect(fs.existsSync(path.join(distDir, 'install', 'm365-mcp-shim.js'))).toBe(true);
  });
});

describe('GET /install.sh', () => {
  it('returns the shell installer with URL and MCP name substituted', async () => {
    const res = await installSh(makeRequest(), makeContext());
    expect(res.status).toBe(200);
    expect(res.headers?.['Content-Type']).toContain('text/plain');
    const script = res.body as string;
    expect(script).toContain('MCP_URL="https://mcp.example.com"');
    expect(script).toContain('MCP_NAME="test-m365"');
    expect(script).not.toMatch(/\{\{[A-Z_]+\}\}/);
  });
});

describe('installers configure the shim, not supergateway', () => {
  it.each([
    ['install.sh', () => installSh(makeRequest(), makeContext())],
    ['install.ps1', () => installPs1(makeRequest(), makeContext())],
  ])('%s', async (_name, render) => {
    const script = (await render()).body as string;
    expect(script).not.toMatch(/supergateway|npx/i);
    expect(script).toContain('https://mcp.example.com/install/m365-mcp-shim.js');
    expect(script).toContain('--streamableHttp');
    // An existing env block (M365_MCP_ATTACH_ROOTS) survives a re-install.
    expect(script).toMatch(/env/);
  });
});

describe('installers keep the session token out of config files and /tmp', () => {
  it.each([
    ['install.sh', () => installSh(makeRequest(), makeContext())],
    ['install.ps1', () => installPs1(makeRequest(), makeContext())],
  ])('%s hands the token to the shim credential store', async (_name, render) => {
    const script = (await render()).body as string;
    expect(script).toContain('--store-token');
    expect(script).toContain('--token-store');
    // The config entry no longer carries a bearer header.
    expect(script).not.toMatch(/Authorization:Bearer/);
  });

  it('install.sh writes the poll response to a private mktemp file, removed on exit', async () => {
    const script = (await installSh(makeRequest(), makeContext())).body as string;
    expect(script).not.toContain('/tmp/install-poll-resp.json');
    expect(script).toMatch(/POLL_RESP=\$\(umask 077 && mktemp /);
    expect(script).toContain(`trap 'rm -f "$POLL_RESP"' EXIT`);
    // The connection test reads the header from stdin, not argv.
    expect(script).toContain('curl -s -K -');
  });
});

