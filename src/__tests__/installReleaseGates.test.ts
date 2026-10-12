/**
 * Installer release gates.
 *
 * The install routes are public by design and serve code that goes on to hold
 * a bearer token on every user's machine, which makes them the surface with
 * the widest blast radius this server has. Three properties keep that safe
 * (threat model 3.5 and 4.3), and each already had a test. This file is what
 * makes them release gates rather than properties that happen to hold:
 *
 *   1. Every route the install module registers serves URLs taken from
 *      OAUTH_REDIRECT_URI and nothing from the request's Host, X-Forwarded-*
 *      or Forwarded headers. Routes are read back from the registrations, so
 *      a route added later is covered the day it lands, and the source is
 *      checked as well, so a header read that does not yet change the served
 *      text still fails.
 *   2. The update payload is signed with Ed25519, the key baked into the
 *      served extension verifies it (checked with Node's crypto, not the
 *      module under test), the extension verifies before it applies, and
 *      without the signing key the instance serves nothing rather than an
 *      unsigned update. What the image ships under dist/install is what these
 *      tests exercised, byte for byte.
 *   3. The config entries the installers write hold no token. The shipped
 *      config-writing code of each installer is run with a token in scope
 *      against a config that already carries one from an earlier installer,
 *      and neither the result nor its backup may contain either.
 *
 * The checks live in scripts/installer-release-gates.js so the CI
 * container-boot job runs the same ones against the image as built.
 */

import { jest } from '@jest/globals';
import * as fs from 'fs';
import * as path from 'path';
import { randomBytes } from 'crypto';
import type { HttpRequest, InvocationContext } from '@azure/functions';

const httpMock = jest.fn();
jest.mock('@azure/functions', () => ({ app: { http: httpMock } }));

const ORIGIN = 'https://mcp.example.com';
process.env.MCP_INSTANCE_NAME = 'Test M365 MCP';
process.env.OAUTH_REDIRECT_URI = `${ORIGIN}/api/auth/callback`;
process.env.MCP_SESSION_HMAC_KEY = randomBytes(32).toString('hex');

import '../functions/install/installEndpoint.js';

interface Gates {
  FORGED_HOST: string;
  ROUTES: string[];
  ROUTES_NAMING_ORIGIN: string[];
  zipEntries: (buf: Buffer) => Array<{ name: string; content: string }>;
  originProblems: (text: string, expected: string) => string[];
  signingProblems: (input: { update: unknown; version: string; bundleServerJs?: string }) => string[];
  installShProblems: (script: string) => string[];
  installPs1Problems: (script: string) => string[];
  manifestProblems: (manifest: unknown) => string[];
  runShConfigWriter: (script: string, opts: { origin: string }) => string[];
  runPs1ConfigWriter: (script: string, opts: { origin: string }) => string[];
  pwshAvailable: () => boolean;
}
const gates = require('../../scripts/installer-release-gates.js') as Gates;

const root = path.join(__dirname, '..', '..');
const read = (p: string): string => fs.readFileSync(path.join(root, p), 'utf8');

interface Response {
  status: number;
  body?: string | Buffer;
  jsonBody?: unknown;
}
type Handler = (req: HttpRequest, ctx: InvocationContext) => Promise<Response>;
interface Registration {
  name: string;
  route: string;
  handler: Handler;
}
interface UpdateResponse {
  version: string;
  files: Record<string, string>;
  signed: { payload: string; signature: string };
}

// Every route the install module registers, read back from app.http.
const registrations: Registration[] = httpMock.mock.calls.map((call) => {
  const opts = call[1] as { route: string; handler: Handler };
  return { name: call[0] as string, route: opts.route, handler: opts.handler };
});

function handler(name: string): Handler {
  const reg = registrations.find((r) => r.name === name);
  if (!reg) throw new Error(`Handler '${name}' not registered`);
  return reg.handler;
}

/** A request whose every host-bearing header names a host the attacker controls. */
function forgedRequest(route: string): HttpRequest {
  const host = gates.FORGED_HOST;
  return {
    url: `http://${host}/${route}`,
    headers: new Map<string, string>([
      ['host', host],
      ['x-forwarded-host', host],
      ['x-forwarded-proto', 'http'],
      ['x-forwarded-server', host],
      ['x-original-host', host],
      ['forwarded', `host=${host};proto=http`],
    ]),
    query: { get: () => null, has: () => false },
  } as unknown as HttpRequest;
}

function ctx(): InvocationContext {
  return { error: jest.fn(), warn: jest.fn(), log: jest.fn() } as unknown as InvocationContext;
}

/** Everything a response puts in front of a client, as text. */
function servedText(res: Response): string {
  if (Buffer.isBuffer(res.body)) return gates.zipEntries(res.body).map((e) => e.content).join('\n');
  if (typeof res.body === 'string') return res.body;
  return JSON.stringify(res.jsonBody ?? '');
}

async function serve(reg: Registration): Promise<Response> {
  return reg.handler(forgedRequest(reg.route), ctx());
}

async function withoutEnv<T>(name: string, fn: () => Promise<T>): Promise<T> {
  const saved = process.env[name];
  delete process.env[name];
  try {
    return await fn();
  } finally {
    process.env[name] = saved;
  }
}

// ── Gate 1: origin ──────────────────────────────────────────────────────────

describe('gate 1: served URLs come from OAUTH_REDIRECT_URI, never from a request header', () => {
  it('the install module registers the routes the live gate fetches', () => {
    expect(registrations.length).toBeGreaterThan(0);
    for (const route of gates.ROUTES) expect(registrations.map((r) => `/${r.route}`)).toContain(route);
  });

  it.each(registrations.map((r) => [r.name, r.route, r] as const))(
    '%s (/%s) serves only the configured origin under forged Host, X-Forwarded-* and Forwarded',
    async (_name, route, reg) => {
      const res = await serve(reg);
      expect(res.status).toBe(200);
      const text = servedText(res);
      expect(gates.originProblems(text, ORIGIN)).toEqual([]);
      if (gates.ROUTES_NAMING_ORIGIN.includes(`/${route}`)) expect(text).toContain(ORIGIN);
    },
  );

  it('with no configured origin, no route falls back to a header: nothing served names any origin', async () => {
    await withoutEnv('OAUTH_REDIRECT_URI', async () => {
      for (const reg of registrations) {
        const text = servedText(await serve(reg));
        expect(text).not.toContain(gates.FORGED_HOST);
        expect(text).not.toMatch(/https?:\/\/[A-Za-z0-9.-]+/);
      }
    });
  });

  it('the install module reads the origin from configuration and no host header from the request', () => {
    const src = read('src/functions/install/installEndpoint.ts');
    expect(src).toMatch(/redirectUri: string \| undefined = process\.env\.OAUTH_REDIRECT_URI/);
    expect(src).not.toMatch(/headers\.get\(\s*['"`](host|x-forwarded-host|x-forwarded-proto|x-forwarded-server|x-original-host|forwarded)['"`]/i);
    expect(src).not.toMatch(/\b(request|req)\.url\b/);
  });

  it('the installer templates take the server URL from the rendered placeholder and nothing else', () => {
    expect(read('src/install/install-mcp.sh.template')).toMatch(/^MCP_URL="\{\{MCP_URL\}\}"$/m);
    expect(read('src/install/install-mcp.ps1.template')).toMatch(/^\$MCP_URL = "\{\{MCP_URL\}\}"$/m);
  });
});

// ── Gate 2: signing ─────────────────────────────────────────────────────────

describe('gate 2: the update payload is signed and the served extension carries the verifying key', () => {
  let update: UpdateResponse;
  let version: string;
  let bundleServerJs: string;

  beforeAll(async () => {
    update = (await handler('extensionUpdate')(forgedRequest('api/extension-update'), ctx())).jsonBody as UpdateResponse;
    version = ((await handler('extensionVersion')(forgedRequest('api/extension-version'), ctx())).jsonBody as { version: string }).version;
    const bundle = (await handler('installBundle')(forgedRequest('install.mcpb'), ctx())).body as Buffer;
    bundleServerJs = gates.zipEntries(bundle).find((e) => e.name === 'server/index.js')!.content;
  });

  it('signs with Ed25519, the baked-in key verifies it, and the extension applies only the signed block', () => {
    expect(gates.signingProblems({ update, version, bundleServerJs })).toEqual([]);
  });

  it('without the signing key the instance serves nothing, never an unsigned update', async () => {
    await withoutEnv('MCP_SESSION_HMAC_KEY', async () => {
      for (const name of ['extensionUpdate', 'installBundle']) {
        const res = await handler(name)(forgedRequest(''), ctx());
        expect(res.status).toBe(500);
        expect(servedText(res)).not.toContain('"files"');
      }
    });
  });

  it('the served extension inlines the shipped update module verbatim', () => {
    const shipped = read('src/install/extension-update.js');
    expect(update.files['server/index.js']).toContain(shipped);
    expect(bundleServerJs).toContain(shipped);
  });

  it('what the image ships under dist/install is byte for byte what these tests exercised', () => {
    // The image carries dist/ only (Dockerfile), and readInstallFile prefers
    // dist/install when it exists. CI builds before it tests; a developer
    // running before a build exercises src/install, which is the same text.
    const dist = path.join(root, 'dist', 'install');
    if (!fs.existsSync(dist)) {
      expect(process.env.CI).toBeUndefined();
      return;
    }
    for (const file of fs.readdirSync(path.join(root, 'src', 'install'))) {
      expect(fs.readFileSync(path.join(dist, file))).toEqual(fs.readFileSync(path.join(root, 'src', 'install', file)));
    }
  });
});

// ── Gate 3: no token in config ──────────────────────────────────────────────

describe('gate 3: the config the installers write holds no token', () => {
  let sh: string;
  let ps1: string;
  let manifest: unknown;

  beforeAll(async () => {
    sh = (await handler('installSh')(forgedRequest('install.sh'), ctx())).body as string;
    ps1 = (await handler('installPs1')(forgedRequest('install.ps1'), ctx())).body as string;
    const bundle = (await handler('installBundle')(forgedRequest('install.mcpb'), ctx())).body as Buffer;
    manifest = JSON.parse(gates.zipEntries(bundle).find((e) => e.name === 'manifest.json')!.content);
  });

  it('install.sh clears the token before its config step, which never names it', () => {
    expect(gates.installShProblems(sh)).toEqual([]);
  });

  it('install.sh: the shipped config writer, run with a token in scope, writes none into the config or its backup', () => {
    expect(gates.runShConfigWriter(sh, { origin: ORIGIN })).toEqual([]);
  });

  it('install.ps1 writes the config from a function that never sees the token', () => {
    expect(gates.installPs1Problems(ps1)).toEqual([]);
  });

  const pwsh = gates.pwshAvailable();

  it('PowerShell is present where the gate has to run', () => {
    // ubuntu-latest ships pwsh; a developer box may not, and then the test
    // below is skipped rather than failed.
    if (process.env.CI) expect(pwsh).toBe(true);
  });

  (pwsh ? it : it.skip)('install.ps1: the shipped config writer, run with a token in scope, writes none into the config or its backup', () => {
    expect(gates.runPs1ConfigWriter(ps1, { origin: ORIGIN })).toEqual([]);
  });

  it('the extension manifest starts the bridge with no token in its arguments or environment', () => {
    expect(gates.manifestProblems(manifest)).toEqual([]);
  });
});
