#!/usr/bin/env node
/**
 * installer-release-gates.js: release gates on the installer surface.
 *
 * The install routes are public by design and serve code that goes on to hold
 * a bearer token on every user's machine, which makes them the surface with
 * the widest blast radius this server has. Three properties keep that safe.
 * Each is held by the code today; this file is what turns them from
 * properties that happen to hold into checks a release has to pass.
 *
 *   1. Origin. Every URL in served code comes from OAUTH_REDIRECT_URI, never
 *      from the request's Host, X-Forwarded-* or Forwarded headers, so whoever
 *      can set a header cannot choose the server, and so the update source,
 *      of every client installed from that response.
 *   2. Signing. /api/extension-update carries a payload signed with Ed25519;
 *      the served extension bakes in the verifying key, verifies before it
 *      applies, and never applies the unsigned block.
 *   3. No token in config. The entries the installers write into the MCP
 *      client configs name the credential store and hold no token, and the
 *      backup they take of an earlier config has the earlier token removed.
 *      Checked by running the shipped config-writing code, with a token in
 *      scope, against a config that already carries one.
 *
 * The checks are functions over served text so one implementation runs in two
 * places. src/__tests__/installReleaseGates.test.ts feeds them the route
 * handlers' output on every `npm test`. main() below feeds them a running
 * instance: the CI container-boot job runs it against the image as built, and
 * an operator can run it against a deployment (go-live checklist item 9):
 *
 *   node scripts/installer-release-gates.js http://localhost:8080 https://mcp.example.com
 *
 * Every request carries forged Host, X-Forwarded-Host, X-Forwarded-Proto and
 * Forwarded headers. Exit 0 when every gate passes, 1 otherwise, 2 on usage.
 * bash and python3 run the shell installer's config writer; pwsh runs the
 * PowerShell one and is required when CI is set.
 */
'use strict';

const { createPublicKey, randomBytes, verify } = require('crypto');
const { spawnSync } = require('child_process');
const fs = require('fs');
const http = require('http');
const https = require('https');
const os = require('os');
const path = require('path');
const zlib = require('zlib');

/** The host every request forges; it must never appear in anything served. */
const FORGED_HOST = 'evil.example.net';

/** Origins the installers may link to besides the instance: prerequisite install pages. */
const DOC_LINKS = ['https://nodejs.org', 'https://learn.microsoft.com'];

/** Every route the install module serves. */
const ROUTES = [
  '/install',
  '/install.mcpb',
  '/install.sh',
  '/install.ps1',
  '/install/m365-mcp-shim.js',
  '/api/extension-version',
  '/api/extension-update',
];

/** The routes whose served text must name the configured origin. */
const ROUTES_NAMING_ORIGIN = ['/install', '/install.mcpb', '/install.sh', '/install.ps1', '/api/extension-update'];

/** The env block an earlier install left on the entry; a re-install keeps it. */
const ATTACH_ROOTS = 'default:/srv/share';

// ── Served text ─────────────────────────────────────────────────────────────

/** Entries of a zip built by the bundle generator: local headers, deflate or stored. */
function zipEntries(buf) {
  const entries = [];
  let offset = 0;
  while (offset + 30 <= buf.length && buf.readUInt32LE(offset) === 0x04034b50) {
    const method = buf.readUInt16LE(offset + 8);
    const compressedSize = buf.readUInt32LE(offset + 18);
    const nameLen = buf.readUInt16LE(offset + 26);
    const extraLen = buf.readUInt16LE(offset + 28);
    const name = buf.subarray(offset + 30, offset + 30 + nameLen).toString('utf8');
    const dataStart = offset + 30 + nameLen + extraLen;
    const data = buf.subarray(dataStart, dataStart + compressedSize);
    entries.push({ name, content: (method === 8 ? zlib.inflateRawSync(data) : data).toString('utf8') });
    offset = dataStart + compressedSize;
  }
  return entries;
}

/** Every absolute http(s) URL in `text`, trailing punctuation dropped. */
function urlsIn(text) {
  return [...text.matchAll(/https?:\/\/[A-Za-z0-9.-]+(?::\d+)?[^\s"'<>()\\`]*/g)]
    .map((m) => m[0].replace(/[.,;:]+$/, ''));
}

// ── Gate 1: origin ──────────────────────────────────────────────────────────

/**
 * Problems with the origins a served text names: the forged host, or any
 * origin that is neither `expected` nor a documented prerequisite link.
 */
function originProblems(text, expected) {
  const problems = [];
  if (text.includes(FORGED_HOST)) problems.push(`names the forged host ${FORGED_HOST}`);
  const foreign = new Set();
  for (const url of urlsIn(text)) {
    let origin;
    try {
      origin = new URL(url).origin;
    } catch {
      foreign.add(url);
      continue;
    }
    if (origin !== expected && !DOC_LINKS.includes(origin)) foreign.add(origin);
  }
  if (foreign.size > 0) {
    problems.push(`names origins other than ${expected} and the documented links: ${[...foreign].join(', ')}`);
  }
  return problems;
}

// ── Gate 2: signing ─────────────────────────────────────────────────────────

/** The base64 SPKI key the served extension verifies updates against. */
function bakedPublicKey(serverJs) {
  const m = typeof serverJs === 'string' && serverJs.match(/^const UPDATE_PUBLIC_KEY = '([A-Za-z0-9+/=]+)';$/m);
  return m ? m[1] : null;
}

function safeVerify(key, data, signature) {
  try {
    return verify(null, data, key, signature);
  } catch {
    return false;
  }
}

/** The served extension must verify the signed block before it applies, and apply nothing else. */
function appliesOnlySignedProblems(serverJs, label) {
  const problems = [];
  const verifyAt = serverJs.indexOf('extensionUpdate.verifySignedUpdate(update.signed, UPDATE_PUBLIC_KEY, EXTENSION_VERSION)');
  const applyAt = serverJs.indexOf('extensionUpdate.applyUpdate(');
  if (verifyAt < 0) problems.push(`${label} server/index.js does not verify the signed block`);
  if (applyAt < 0) problems.push(`${label} server/index.js never applies an update`);
  else if (verifyAt >= 0 && applyAt < verifyAt) problems.push(`${label} server/index.js applies before it verifies`);
  if (/applyUpdate\([^)]*update\.files/.test(serverJs)) problems.push(`${label} server/index.js applies the unsigned files`);
  if (!/function verifySignedUpdate\(/.test(serverJs)) problems.push(`${label} server/index.js does not inline the verifier`);
  return problems;
}

/**
 * Problems with the update response: `update` is the /api/extension-update
 * body, `version` the /api/extension-version one, `bundleServerJs` the
 * server/index.js inside /install.mcpb (optional). The signature is checked
 * with Node's own crypto, not the module the extension inlines.
 */
function signingProblems({ update, version, bundleServerJs }) {
  const signed = update && update.signed;
  if (!signed || typeof signed.payload !== 'string' || typeof signed.signature !== 'string') {
    return ['the update response carries no signed block'];
  }
  const serverJs = update.files && update.files['server/index.js'];
  if (typeof serverJs !== 'string') return ['the update response carries no server/index.js'];
  const keyB64 = bakedPublicKey(serverJs);
  if (!keyB64) return ['the served extension bakes in no UPDATE_PUBLIC_KEY'];
  let key;
  try {
    key = createPublicKey({ key: Buffer.from(keyB64, 'base64'), format: 'der', type: 'spki' });
  } catch (err) {
    return [`UPDATE_PUBLIC_KEY is not a public key: ${err.message}`];
  }

  const problems = [];
  if (key.asymmetricKeyType !== 'ed25519') problems.push(`UPDATE_PUBLIC_KEY is ${key.asymmetricKeyType}, not ed25519`);
  const signature = Buffer.from(signed.signature, 'base64');
  if (signature.length !== 64) problems.push(`the signature is ${signature.length} bytes, not an Ed25519 signature`);
  const payloadBytes = Buffer.from(signed.payload, 'utf8');
  if (!safeVerify(key, payloadBytes, signature)) problems.push('the signature does not verify against the baked-in key');
  // Not vacuous: one changed byte and the same check must fail.
  const tampered = Buffer.from(payloadBytes);
  tampered[tampered.length - 2] ^= 0x01;
  if (safeVerify(key, tampered, signature)) problems.push('the signature still verifies over a changed payload');

  let payload;
  try {
    payload = JSON.parse(signed.payload);
  } catch {
    return [...problems, 'the signed payload is not JSON'];
  }
  if (payload.version !== version) problems.push(`the signed version ${payload.version} is not the advertised ${version}`);
  if (update.version !== version) problems.push(`the unsigned version ${update.version} is not the advertised ${version}`);
  if (JSON.stringify(payload.files) !== JSON.stringify(update.files)) problems.push('the signed files differ from the unsigned files');
  if (bakedPublicKey(payload.files && payload.files['server/index.js']) !== keyB64) {
    problems.push('the signed server/index.js bakes in a different key');
  }
  if (bundleServerJs !== undefined && bakedPublicKey(bundleServerJs) !== keyB64) {
    problems.push('the bundle bakes in a different key from the update');
  }
  problems.push(...appliesOnlySignedProblems(serverJs, 'update'));
  if (typeof bundleServerJs === 'string') problems.push(...appliesOnlySignedProblems(bundleServerJs, 'bundle'));
  return problems;
}

// ── Gate 3: no token in config ──────────────────────────────────────────────

/** The text of one PowerShell function, up to the next one. */
function psFunction(script, name) {
  const start = script.indexOf(`function ${name} {`);
  if (start < 0) return null;
  const next = script.indexOf('\nfunction ', start + 1);
  return script.slice(start, next < 0 ? undefined : next);
}

/** Static: install.sh clears the token before its config step, which never names it. */
function installShProblems(script) {
  const at = script.indexOf('write_mcp_entry() {');
  if (at < 0) return ['install.sh has no write_mcp_entry function'];
  const problems = [];
  if (!script.slice(0, at).includes('SESSION_TOKEN=""')) problems.push('install.sh does not clear SESSION_TOKEN before the config step');
  const tail = script.slice(at);
  const m = tail.match(/SESSION_TOKEN|Bearer|Authorization:/);
  if (m) problems.push(`the install.sh config step mentions "${m[0]}"`);
  if (!tail.includes('"--token-store"')) problems.push('the install.sh config entry does not name the token store');
  return problems;
}

/** Static: the install.ps1 function that writes config never sees the token. */
function installPs1Problems(script) {
  const fn = psFunction(script, 'Update-ClaudeConfig');
  if (!fn) return ['install.ps1 has no Update-ClaudeConfig function'];
  const problems = [];
  const m = fn.match(/SessionToken|Bearer|Authorization/i);
  if (m) problems.push(`install.ps1 Update-ClaudeConfig mentions "${m[0]}"`);
  if (!fn.includes("'--token-store'")) problems.push('the install.ps1 config entry does not name the token store');
  return problems;
}

/** The extension's manifest starts the bridge with nothing in its arguments or environment. */
function manifestProblems(manifest) {
  const cfg = manifest && manifest.server && manifest.server.mcp_config;
  if (!cfg) return ['the manifest has no server.mcp_config'];
  const problems = [];
  if (JSON.stringify(cfg.env) !== '{}') problems.push(`the extension starts with env ${JSON.stringify(cfg.env)}, not an empty block`);
  if (JSON.stringify(cfg.args) !== JSON.stringify(['${__dirname}/server/index.js'])) {
    problems.push(`the extension starts with args ${JSON.stringify(cfg.args)}`);
  }
  const m = JSON.stringify(manifest).match(/token|bearer|authorization/i);
  if (m) problems.push(`the manifest mentions "${m[0]}"`);
  return problems;
}

function sessionToken(kind) {
  return `${kind}-session-token-${randomBytes(8).toString('hex')}`;
}

/** A config an earlier installer wrote: the entry carries a bearer header and an env block. */
function seedConfig(dir, name, origin, oldToken) {
  const configPath = path.join(dir, 'claude_desktop_config.json');
  fs.writeFileSync(configPath, JSON.stringify({
    theme: 'dark',
    mcpServers: {
      other: { command: 'other-server', args: ['--flag'] },
      [name]: {
        command: 'npx',
        args: ['-y', 'supergateway', '--streamableHttp', `${origin}/api/mcp`, '--header', `Authorization: Bearer ${oldToken}`],
        env: { M365_MCP_ATTACH_ROOTS: ATTACH_ROOTS },
      },
    },
  }, null, 2));
  return configPath;
}

/** Problems with the config a writer produced from seedConfig, and with its backup. */
function writtenConfigProblems(configPath, { name, origin, shim, nodeBin, oldToken, newToken, backupMode }) {
  const problems = [];
  const written = fs.readFileSync(configPath, 'utf8');
  const backupPath = `${configPath}.backup`;
  const backup = fs.existsSync(backupPath) ? fs.readFileSync(backupPath, 'utf8') : null;
  if (backup === null) problems.push('no backup was written');
  for (const [label, text] of [['the config', written], ['the backup', backup]]) {
    if (text === null) continue;
    if (text.includes(oldToken)) problems.push(`${label} still carries the earlier installer's token`);
    if (text.includes(newToken)) problems.push(`${label} carries the session token`);
    const m = text.match(/Bearer|Authorization/i);
    if (m) problems.push(`${label} mentions "${m[0]}"`);
  }

  let config;
  try {
    config = JSON.parse(written);
  } catch {
    return [...problems, 'the config is not JSON'];
  }
  const entry = config.mcpServers && config.mcpServers[name];
  if (!entry) return [...problems, `the config has no mcpServers.${name} entry`];
  if (entry.command !== nodeBin) problems.push(`the entry runs ${entry.command}, not node`);
  const want = [shim, '--streamableHttp', `${origin}/api/mcp`, '--token-store', name];
  if (JSON.stringify(entry.args) !== JSON.stringify(want)) {
    problems.push(`the entry args are ${JSON.stringify(entry.args)}, want ${JSON.stringify(want)}`);
  }
  if (JSON.stringify(entry.env) !== JSON.stringify({ M365_MCP_ATTACH_ROOTS: ATTACH_ROOTS })) {
    problems.push('the env block of the earlier entry was not carried over');
  }
  if (JSON.stringify(config.mcpServers.other) !== JSON.stringify({ command: 'other-server', args: ['--flag'] }) || config.theme !== 'dark') {
    problems.push('other config content was not preserved');
  }
  if (backup !== null) {
    let args;
    try {
      args = JSON.parse(backup).mcpServers[name].args;
    } catch {
      args = null;
    }
    if (!Array.isArray(args)) problems.push('the backup lost the entry');
    else if (args.includes('--header')) problems.push("the backup keeps the earlier installer's bearer header");
    if (backupMode !== undefined) {
      const mode = fs.statSync(backupPath).mode & 0o777;
      if (mode !== backupMode) problems.push(`the backup is mode ${mode.toString(8)}, not ${backupMode.toString(8)}`);
    }
  }
  return problems;
}

/**
 * Dynamic: extract write_mcp_entry from the served install.sh and run it with
 * the session token in its environment, against a config an earlier installer
 * wrote. Needs bash and python3, as the installer itself does.
 */
function runShConfigWriter(script, { origin }) {
  const name = (script.match(/^MCP_NAME="([^"]+)"$/m) || [])[1];
  if (!name) return ['install.sh does not set MCP_NAME'];
  const start = script.indexOf('write_mcp_entry() {');
  const END = '\nPYEOF\n}';
  const end = start < 0 ? -1 : script.indexOf(END, start);
  if (end < 0) return ['install.sh write_mcp_entry could not be extracted'];

  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'installer-gate-'));
  try {
    const oldToken = sessionToken('old');
    const newToken = sessionToken('new');
    const configPath = seedConfig(dir, name, origin, oldToken);
    const shim = path.join(dir, `${name}-shim.js`);
    const nodeBin = process.execPath;
    const fnFile = path.join(dir, 'write_mcp_entry.sh');
    fs.writeFileSync(fnFile, script.slice(start, end + END.length));
    const run = spawnSync('bash', ['-c', 'source "$1" && write_mcp_entry "$2" 1 1', 'gate', fnFile, configPath], {
      env: { ...process.env, MCP_NAME: name, MCP_URL: origin, NODE_BIN: nodeBin, SHIM_PATH: shim, SESSION_TOKEN: newToken },
      encoding: 'utf8',
    });
    if (run.error) return [`install.sh write_mcp_entry could not run: ${run.error.message}`];
    if (run.status !== 0) return [`install.sh write_mcp_entry exited ${run.status}: ${(run.stderr || '').trim()}`];
    return writtenConfigProblems(configPath, { name, origin, shim, nodeBin, oldToken, newToken, backupMode: 0o600 });
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
}

function pwshAvailable() {
  const run = spawnSync('pwsh', ['-NoProfile', '-NonInteractive', '-Command', '$PSVersionTable.PSVersion.Major'], { encoding: 'utf8' });
  return !run.error && run.status === 0;
}

/**
 * Dynamic: the same for install.ps1's Update-ClaudeConfig, under pwsh, with
 * the session the script holds in scope the way the installer does.
 */
function runPs1ConfigWriter(script, { origin }) {
  const name = (script.match(/^\$MCP_NAME = "([^"]+)"$/m) || [])[1];
  const remove = psFunction(script, 'Remove-TokenArgs');
  const update = psFunction(script, 'Update-ClaudeConfig');
  if (!name || !remove || !update) return ['install.ps1 config functions could not be extracted'];

  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'installer-gate-'));
  try {
    const oldToken = sessionToken('old');
    const newToken = sessionToken('new');
    const configPath = seedConfig(dir, name, origin, oldToken);
    const shim = path.join(dir, `${name}-shim.js`);
    const nodeBin = process.execPath;
    const harness = [
      "$ErrorActionPreference = 'Stop'",
      `$MCP_URL = '${origin}'`,
      `$MCP_NAME = '${name}'`,
      `$NODE_BIN = '${nodeBin}'`,
      `$SHIM_PATH = '${shim}'`,
      `$session = @{ SessionToken = '${newToken}' }`,
      remove,
      update,
      `if (-not (Update-ClaudeConfig -ConfigPath '${configPath}' -IncludeStdioType $true)) { throw 'config not found' }`,
    ].join('\n');
    const ps1 = path.join(dir, 'gate.ps1');
    fs.writeFileSync(ps1, harness);
    const run = spawnSync('pwsh', ['-NoProfile', '-NonInteractive', '-File', ps1], { encoding: 'utf8' });
    if (run.error) return [`install.ps1 Update-ClaudeConfig could not run: ${run.error.message}`];
    if (run.status !== 0) return [`install.ps1 Update-ClaudeConfig exited ${run.status}: ${(run.stderr || run.stdout || '').trim()}`];
    return writtenConfigProblems(configPath, { name, origin, shim, nodeBin, oldToken, newToken });
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
}

// ── Against a running instance ──────────────────────────────────────────────

/** GET `route` from `base` with every host header forged. */
function get(base, route) {
  return new Promise((resolve, reject) => {
    const url = new URL(route, base);
    const req = (url.protocol === 'https:' ? https : http).request(url, {
      method: 'GET',
      headers: {
        Host: FORGED_HOST,
        'X-Forwarded-Host': FORGED_HOST,
        'X-Forwarded-Proto': 'http',
        Forwarded: `host=${FORGED_HOST};proto=http`,
      },
      timeout: 20000,
    }, (res) => {
      const chunks = [];
      res.on('data', (c) => chunks.push(c));
      res.on('end', () => resolve({ status: res.statusCode, body: Buffer.concat(chunks) }));
    });
    req.on('error', reject);
    req.on('timeout', () => req.destroy(new Error(`timeout fetching ${route}`)));
    req.end();
  });
}

async function main(argv) {
  const [base, expected] = argv;
  let origin;
  let baseUrl;
  try {
    origin = new URL(expected).origin;
    baseUrl = new URL(base).toString();
  } catch {
    console.error('usage: installer-release-gates.js <base-url> <expected-origin>');
    return 2;
  }

  let failed = 0;
  const report = (label, problems) => {
    if (problems.length === 0) {
      console.log(`ok   ${label}`);
      return;
    }
    failed++;
    for (const p of problems) console.error(`FAIL ${label}: ${p}`);
  };

  const fetched = new Map();
  for (const route of ROUTES) {
    const res = await get(baseUrl, route);
    if (res.status !== 200) report(route, [`answered ${res.status}`]);
    else fetched.set(route, res.body);
  }
  const text = (route) => (route === '/install.mcpb'
    ? zipEntries(fetched.get(route)).map((e) => e.content).join('\n')
    : fetched.get(route).toString('utf8'));

  for (const route of fetched.keys()) {
    const problems = originProblems(text(route), origin);
    if (ROUTES_NAMING_ORIGIN.includes(route) && !text(route).includes(origin)) problems.push(`does not name ${origin}`);
    report(`origin: ${route}`, problems);
  }

  if (['/api/extension-update', '/api/extension-version', '/install.mcpb'].every((r) => fetched.has(r))) {
    const update = JSON.parse(fetched.get('/api/extension-update').toString('utf8'));
    const { version } = JSON.parse(fetched.get('/api/extension-version').toString('utf8'));
    const bundle = zipEntries(fetched.get('/install.mcpb'));
    const entry = (name) => (bundle.find((e) => e.name === name) || {}).content;
    report('signing: update payload and served extension', signingProblems({ update, version, bundleServerJs: entry('server/index.js') }));
    report('config: extension manifest', manifestProblems(JSON.parse(entry('manifest.json') || 'null')));
  }

  if (fetched.has('/install.sh')) {
    const sh = text('/install.sh');
    report('config: install.sh', installShProblems(sh));
    report('config: install.sh config writer', runShConfigWriter(sh, { origin }));
  }
  if (fetched.has('/install.ps1')) {
    const ps1 = text('/install.ps1');
    report('config: install.ps1', installPs1Problems(ps1));
    if (pwshAvailable()) report('config: install.ps1 config writer', runPs1ConfigWriter(ps1, { origin }));
    else if (process.env.CI) report('config: install.ps1 config writer', ['pwsh is not installed, so the gate cannot run']);
    else console.log('skip config: install.ps1 config writer (pwsh is not installed)');
  }

  return failed === 0 ? 0 : 1;
}

if (require.main === module) {
  main(process.argv.slice(2)).then(
    (code) => process.exit(code),
    (err) => {
      console.error(err);
      process.exit(2);
    },
  );
}

module.exports = {
  FORGED_HOST,
  DOC_LINKS,
  ROUTES,
  ROUTES_NAMING_ORIGIN,
  zipEntries,
  urlsIn,
  originProblems,
  bakedPublicKey,
  signingProblems,
  installShProblems,
  installPs1Problems,
  manifestProblems,
  runShConfigWriter,
  runPs1ConfigWriter,
  pwshAvailable,
};
