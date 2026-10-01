import { app, HttpRequest, HttpResponseInit, InvocationContext } from '@azure/functions';
import * as fs from 'fs';
import * as path from 'path';
import * as zlib from 'zlib';
import { withSecurity, installLandingHeaders } from '../../services/securityHeaders.js';

/**
 * Dynamic .mcpb extension bundle generator with auto-update support.
 *
 * Serves a Claude Desktop extension bundle (.mcpb) with the server URL,
 * instance name, and keychain service baked in from runtime env vars.
 * Each deployment auto-serves its own bundle — no separate distribution step.
 *
 * The generated extension checks /api/extension-version on every startup
 * and self-updates if the server has a newer version.
 *
 * Routes:
 *   GET /install                — HTML landing page (browser) with download link
 *   GET /install.mcpb           — dynamic .mcpb bundle download
 *   GET /api/extension-version  — returns { version } for update checks
 *   GET /api/extension-update   — returns all extension files as JSON for self-update
 *   GET /install.ps1            — rendered PowerShell installer (Windows)
 *   GET /install.sh             — rendered shell installer (macOS)
 *   GET /install/m365-mcp-shim.js — the stdio shim both installers configure
 */

// Bumped 2.8.0 -> 2.9.0 (AC): field-installed 2.8.0 bundles predate the
// PKCE hardening and re-auth against the current server with the old 32-char
// nonce, which the server now 400s — so an expired token silently kills the
// extension. Their auto-update only fires when the server advertises a STRICTLY
// newer version, so the version had to move for the corrected server/index.js
// (PKCE authenticate + real-auth verifyToken) to reach them.
const EXTENSION_VERSION = '2.9.0';

// ── Helpers ──────────────────────────────────────────────────────────────────

function getPublicOrigin(request: HttpRequest): string {
  const proto =
    request.headers.get('x-forwarded-proto')?.split(',')[0].trim() ?? 'https';
  const host =
    request.headers.get('x-forwarded-host')?.split(',')[0].trim() ??
    request.headers.get('host') ??
    'localhost';
  return `${proto}://${host}`;
}

function getMcpSlug(host: string): string {
  const instanceName = process.env.MCP_INSTANCE_NAME ?? '';
  if (instanceName) {
    return instanceName
      .toLowerCase()
      .replace(/m365 mcp$/i, 'm365')
      .replace(/[^a-z0-9]+/g, '-')
      .replace(/^-+|-+$/g, '');
  }
  const firstSegment = host.split('.')[0].replace(/^mcp-?/i, '').replace(/[^a-z0-9]+/gi, '-');
  return `${firstSegment || 'mcp'}-m365`;
}

function getHost(request: HttpRequest): string {
  return (
    request.headers.get('x-forwarded-host')?.split(',')[0].trim() ??
    request.headers.get('host') ??
    'localhost'
  );
}

// ── ZIP builder (minimal, no dependencies) ───────────────────────────────────
// Builds a valid ZIP archive from an array of {name, data} entries.
// Uses DEFLATE compression via Node's built-in zlib.

interface ZipEntry {
  name: string;
  data: Buffer;
}

function buildZip(entries: ZipEntry[]): Buffer {
  const centralDir: Buffer[] = [];
  const fileChunks: Buffer[] = [];
  let offset = 0;

  for (const entry of entries) {
    const nameBuffer = Buffer.from(entry.name, 'utf-8');
    const compressed = zlib.deflateRawSync(entry.data);
    const crc = crc32(entry.data);

    // Local file header (30 bytes + name + compressed data)
    const local = Buffer.alloc(30);
    local.writeUInt32LE(0x04034b50, 0);   // signature
    local.writeUInt16LE(20, 4);            // version needed (2.0)
    local.writeUInt16LE(0, 6);             // flags
    local.writeUInt16LE(8, 8);             // compression: deflate
    local.writeUInt16LE(0, 10);            // mod time
    local.writeUInt16LE(0, 12);            // mod date
    local.writeUInt32LE(crc, 14);          // crc-32
    local.writeUInt32LE(compressed.length, 18);  // compressed size
    local.writeUInt32LE(entry.data.length, 22);  // uncompressed size
    local.writeUInt16LE(nameBuffer.length, 26);  // file name length
    local.writeUInt16LE(0, 28);            // extra field length

    fileChunks.push(local, nameBuffer, compressed);

    // Central directory entry (46 bytes + name)
    const central = Buffer.alloc(46);
    central.writeUInt32LE(0x02014b50, 0);  // signature
    central.writeUInt16LE(20, 4);           // version made by
    central.writeUInt16LE(20, 6);           // version needed
    central.writeUInt16LE(0, 8);            // flags
    central.writeUInt16LE(8, 10);           // compression: deflate
    central.writeUInt16LE(0, 12);           // mod time
    central.writeUInt16LE(0, 14);           // mod date
    central.writeUInt32LE(crc, 16);         // crc-32
    central.writeUInt32LE(compressed.length, 20);  // compressed size
    central.writeUInt32LE(entry.data.length, 24);  // uncompressed size
    central.writeUInt16LE(nameBuffer.length, 28);  // file name length
    central.writeUInt16LE(0, 30);           // extra field length
    central.writeUInt16LE(0, 32);           // file comment length
    central.writeUInt16LE(0, 34);           // disk number start
    central.writeUInt16LE(0, 36);           // internal file attributes
    central.writeUInt32LE(0, 38);           // external file attributes
    central.writeUInt32LE(offset, 42);      // relative offset of local header

    centralDir.push(central, nameBuffer);
    offset += 30 + nameBuffer.length + compressed.length;
  }

  const centralDirBuffer = Buffer.concat(centralDir);
  const centralDirOffset = offset;

  // End of central directory record (22 bytes)
  const eocd = Buffer.alloc(22);
  eocd.writeUInt32LE(0x06054b50, 0);           // signature
  eocd.writeUInt16LE(0, 4);                     // disk number
  eocd.writeUInt16LE(0, 6);                     // disk with central dir
  eocd.writeUInt16LE(entries.length, 8);         // entries on this disk
  eocd.writeUInt16LE(entries.length, 10);        // total entries
  eocd.writeUInt32LE(centralDirBuffer.length, 12); // central dir size
  eocd.writeUInt32LE(centralDirOffset, 16);      // central dir offset
  eocd.writeUInt16LE(0, 20);                     // comment length

  return Buffer.concat([...fileChunks, centralDirBuffer, eocd]);
}

// CRC-32 (ISO 3309)
const crcTable = (() => {
  const table = new Uint32Array(256);
  for (let i = 0; i < 256; i++) {
    let c = i;
    for (let j = 0; j < 8; j++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    table[i] = c;
  }
  return table;
})();

function crc32(buf: Buffer): number {
  let crc = 0xffffffff;
  for (let i = 0; i < buf.length; i++) crc = crcTable[(crc ^ buf[i]) & 0xff] ^ (crc >>> 8);
  return (crc ^ 0xffffffff) >>> 0;
}

// ── Extension bundle templates ───────────────────────────────────────────────

function renderManifest(slug: string, displayName: string): string {
  return JSON.stringify({
    manifest_version: '0.2',
    name: slug,
    version: EXTENSION_VERSION,
    display_name: displayName,
    description: 'Access Microsoft 365 email, calendar, SharePoint, OneDrive, contacts, and Teams.',
    author: { name: 'Standard Gauge, LLC', email: 'support@standardgauge.ai' },
    server: {
      type: 'node',
      entry_point: 'server/index.js',
      mcp_config: {
        command: 'node',
        args: ['${__dirname}/server/index.js'],
        env: {},
      },
    },
    license: 'AGPL-3.0-only',
  }, null, 2) + '\n';
}

function renderPackageJson(slug: string): string {
  return JSON.stringify({
    name: `${slug}-extension`,
    version: EXTENSION_VERSION,
    main: 'index.js',
    type: 'commonjs',
  }, null, 2) + '\n';
}

function renderServerJs(mcpUrl: string, mcpName: string, keychainService: string): string {
  return `#!/usr/bin/env node
/**
 * M365 MCP Extension — Claude Desktop entry point.
 *
 * Bridges stdio (Claude Desktop) <-> HTTP (MCP server) with:
 *   - OS-native encrypted token storage (macOS Keychain / Windows DPAPI /
 *     Linux libsecret), with a 0600-file fallback
 *   - Nonce-based OAuth flow for first-time sign-in
 *   - Auto re-auth on session expiry
 *   - No secrets in process argv
 *
 * Auto-generated by ${mcpUrl}/install.mcpb
 */

const { execFileSync } = require('child_process');
const fs = require('fs');
const path = require('path');
const os = require('os');
const https = require('https');
const http = require('http');
const crypto = require('crypto');
const readline = require('readline');

const MCP_URL = '${mcpUrl}';
const MCP_NAME = '${mcpName}';
const KEYCHAIN_SERVICE = '${keychainService}';
const KEYCHAIN_ACCOUNT = 'session-token';
const EXTENSION_VERSION = '${EXTENSION_VERSION}';
const POLL_INTERVAL = 2000;
const POLL_MAX = 150; // 5 minutes

// -- Auto-update --
// On startup, checks the server for a newer extension version.
// If strictly newer (semver), downloads updated files and overwrites in place.
// Next Claude Desktop restart picks up the new code automatically.
// Bounded by a 5-second timeout so a stalled server never blocks startup.

function withTimeout(promise, ms) {
  return Promise.race([
    promise,
    new Promise((_, reject) => setTimeout(() => reject(new Error('timeout')), ms)),
  ]);
}

function isNewerVersion(remote, local) {
  const parse = (v) => (v || '').split('.').map(Number);
  const r = parse(remote);
  const l = parse(local);
  for (let i = 0; i < Math.max(r.length, l.length); i++) {
    const rv = r[i] || 0;
    const lv = l[i] || 0;
    if (rv > lv) return true;
    if (rv < lv) return false;
  }
  return false;
}

async function autoUpdate() {
  try {
    const res = await withTimeout(httpRequest(MCP_URL + '/api/extension-version', 'GET'), 5000);
    if (res.status !== 200) return;
    const data = JSON.parse(res.body);
    if (!data.version || !isNewerVersion(data.version, EXTENSION_VERSION)) return;

    process.stderr.write('[' + MCP_NAME + '] Updating extension ' + EXTENSION_VERSION + ' -> ' + data.version + '...\\n');
    const updateRes = await withTimeout(httpRequest(MCP_URL + '/api/extension-update', 'GET'), 10000);
    if (updateRes.status !== 200) {
      process.stderr.write('[' + MCP_NAME + '] Update download failed (HTTP ' + updateRes.status + '), continuing with current version.\\n');
      return;
    }
    const update = JSON.parse(updateRes.body);
    if (!update.files) return;

    // Write updated files atomically (temp + rename)
    const extDir = path.resolve(__dirname, '..');
    for (const [filePath, content] of Object.entries(update.files)) {
      const fullPath = path.join(extDir, filePath);
      const dir = path.dirname(fullPath);
      if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
      const tmpPath = fullPath + '.tmp';
      fs.writeFileSync(tmpPath, content, 'utf8');
      fs.renameSync(tmpPath, fullPath);
    }
    process.stderr.write('[' + MCP_NAME + '] Updated to ' + data.version + '. Changes take effect on next restart.\\n');
  } catch {
    // Update check failed silently — not fatal, continue with current version
  }
}

// -- Token storage --
// The session token is encrypted at rest using the OS-native credential store,
// so it never sits in a readable plaintext file on disk:
//   macOS:   Keychain via the 'security' CLI (token briefly in argv — accepted,
//            see).
//   Windows: DPAPI via PowerShell's built-in ConvertTo-SecureString /
//            ConvertFrom-SecureString. The ciphertext is bound to the current
//            Windows login and stored in ~/.<name>-token; the plaintext token
//            is passed to PowerShell over stdin, never on argv.
//   Linux:   libsecret via the 'secret-tool' CLI (GNOME Keyring / KWallet);
//            the secret is passed over stdin.
// On any platform where the native store is unavailable (no secret service on a
// headless Linux box, PowerShell missing, etc.) we fall back to a 0600 file —
// the prior behavior, and still on par with how AWS CLI / gh / kubectl persist
// credentials. See.

function getTokenFile() {
  return path.join(os.homedir(), \`.\${MCP_NAME}-token\`);
}

// Run a PowerShell one-liner, feeding 'input' on stdin and returning trimmed
// stdout. Tries Windows PowerShell first, then PowerShell Core (pwsh).
function runPowerShell(script, input) {
  let lastErr;
  for (const exe of ['powershell.exe', 'pwsh']) {
    try {
      return execFileSync(exe, ['-NoProfile', '-NonInteractive', '-Command', script], {
        input: input == null ? '' : input,
        encoding: 'utf8',
        stdio: ['pipe', 'pipe', 'pipe'],
      }).trim();
    } catch (err) {
      lastErr = err;
    }
  }
  throw lastErr || new Error('PowerShell not available');
}

// Windows DPAPI: encrypt plaintext (read from stdin) into a user-bound blob.
function dpapiEncrypt(token) {
  return runPowerShell(
    '$t = [Console]::In.ReadToEnd(); ' +
    'ConvertTo-SecureString -String $t -AsPlainText -Force | ConvertFrom-SecureString',
    token
  );
}

// Windows DPAPI: decrypt a blob (read from stdin) back to plaintext.
function dpapiDecrypt(blob) {
  return runPowerShell(
    '$e = [Console]::In.ReadToEnd().Trim(); ' +
    '$s = ConvertTo-SecureString -String $e; ' +
    "[System.Net.NetworkCredential]::new('', $s).Password",
    blob
  );
}

// Linux libsecret helpers. The secret is read from stdin by 'secret-tool store',
// so it never appears in argv or the process list.
function secretToolStore(token) {
  execFileSync('secret-tool', [
    'store', '--label=' + MCP_NAME + ' session token',
    'service', KEYCHAIN_SERVICE, 'account', KEYCHAIN_ACCOUNT,
  ], { input: token, stdio: ['pipe', 'pipe', 'pipe'] });
}

function secretToolLookup() {
  return execFileSync('secret-tool', [
    'lookup', 'service', KEYCHAIN_SERVICE, 'account', KEYCHAIN_ACCOUNT,
  ], { encoding: 'utf8', stdio: ['pipe', 'pipe', 'pipe'] }).trim();
}

function secretToolClear() {
  execFileSync('secret-tool', [
    'clear', 'service', KEYCHAIN_SERVICE, 'account', KEYCHAIN_ACCOUNT,
  ], { stdio: 'pipe' });
}

function loadToken() {
  // macOS: Keychain
  if (process.platform === 'darwin') {
    try {
      const token = execFileSync('security', [
        'find-generic-password', '-s', KEYCHAIN_SERVICE, '-a', KEYCHAIN_ACCOUNT, '-w'
      ], { encoding: 'utf8', stdio: ['pipe', 'pipe', 'pipe'] }).trim();
      if (token) return token;
    } catch { /* keychain miss — fall through */ }
  }

  // Linux: libsecret via secret-tool
  if (process.platform === 'linux') {
    try {
      const token = secretToolLookup();
      if (token) return token;
    } catch { /* no secret service or miss — fall through to file */ }
  }

  // All platforms: file fallback. On Windows the stored value is a DPAPI blob.
  try {
    const tokenFile = getTokenFile();
    if (fs.existsSync(tokenFile)) {
      const data = JSON.parse(fs.readFileSync(tokenFile, 'utf8'));
      if (data.enc === 'dpapi' && data.token) {
        try {
          const token = dpapiDecrypt(data.token);
          if (token) return token;
        } catch { return null; } // wrong user / corrupt blob — force re-auth
      }
      if (data.token) return data.token;
    }
  } catch { /* ignore */ }
  return null;
}

function saveToken(token) {
  // macOS: Keychain
  if (process.platform === 'darwin') {
    try {
      try { execFileSync('security', ['delete-generic-password', '-s', KEYCHAIN_SERVICE, '-a', KEYCHAIN_ACCOUNT], { stdio: 'pipe' }); } catch { /* ok */ }
      execFileSync('security', [
        'add-generic-password', '-s', KEYCHAIN_SERVICE, '-a', KEYCHAIN_ACCOUNT, '-w', token, '-U'
      ], { stdio: 'pipe' });
      return;
    } catch { /* fall through to file */ }
  }

  // Linux: libsecret via secret-tool
  if (process.platform === 'linux') {
    try {
      secretToolStore(token);
      // Stored in the keyring — drop any stale plaintext fallback file.
      try { fs.unlinkSync(getTokenFile()); } catch { /* ok */ }
      return;
    } catch { /* no secret service — fall through to file */ }
  }

  // Windows: DPAPI-encrypted blob bound to the current Windows login.
  if (process.platform === 'win32') {
    try {
      const blob = dpapiEncrypt(token);
      if (blob) {
        fs.writeFileSync(getTokenFile(), JSON.stringify({ enc: 'dpapi', token: blob }), { mode: 0o600 });
        return;
      }
    } catch { /* PowerShell/DPAPI unavailable — fall through to plaintext */ }
  }

  // Last resort: plaintext file (mode 0600)
  fs.writeFileSync(getTokenFile(), JSON.stringify({ token }), { mode: 0o600 });
}

function deleteToken() {
  if (process.platform === 'darwin') {
    try {
      execFileSync('security', ['delete-generic-password', '-s', KEYCHAIN_SERVICE, '-a', KEYCHAIN_ACCOUNT], { stdio: 'pipe' });
    } catch { /* ok */ }
  }
  if (process.platform === 'linux') {
    try { secretToolClear(); } catch { /* ok */ }
  }
  try { fs.unlinkSync(getTokenFile()); } catch { /* ok */ }
}

// -- HTTP helpers --

function httpRequest(url, method, body, headers) {
  headers = headers || {};
  const parsed = new URL(url);
  const mod = parsed.protocol === 'https:' ? https : http;
  const bodyStr = body ? (typeof body === 'string' ? body : JSON.stringify(body)) : null;
  return new Promise((resolve, reject) => {
    const reqHeaders = Object.assign({ 'Content-Type': 'application/json' }, headers);
    if (bodyStr) reqHeaders['Content-Length'] = Buffer.byteLength(bodyStr);
    const req = mod.request({
      hostname: parsed.hostname,
      port: parsed.port || (parsed.protocol === 'https:' ? 443 : 80),
      path: parsed.pathname + parsed.search,
      method: method,
      headers: reqHeaders,
    }, (res) => {
      let data = '';
      res.on('data', (chunk) => data += chunk);
      res.on('end', () => resolve({ status: res.statusCode, body: data }));
    });
    req.on('error', reject);
    if (bodyStr) req.write(bodyStr);
    req.end();
  });
}

// -- OAuth nonce flow (PKCE-style binding) --
// The browser URL carries only SHA256(verifier). The raw verifier stays in
// memory and is sent only to install-poll. An attacker who intercepts the
// login URL cannot poll — SHA256 is preimage-resistant.

async function authenticate() {
  const verifier = crypto.randomBytes(16).toString('hex');
  const challenge = crypto.createHash('sha256').update(verifier).digest('hex');
  const loginUrl = MCP_URL + '/api/auth/login?install_nonce=' + challenge;

  process.stderr.write('\\n[' + MCP_NAME + '] Opening browser to sign in with Microsoft 365...\\n');
  process.stderr.write('[' + MCP_NAME + '] If the browser does not open, visit: ' + loginUrl + '\\n\\n');

  try {
    if (process.platform === 'darwin') {
      execFileSync('open', [loginUrl], { stdio: 'ignore' });
    } else if (process.platform === 'win32') {
      execFileSync('cmd.exe', ['/c', 'start', '', loginUrl], { stdio: 'ignore' });
    } else {
      execFileSync('xdg-open', [loginUrl], { stdio: 'ignore' });
    }
  } catch { /* user will open manually */ }

  const pollUrl = MCP_URL + '/api/auth/install-poll?nonce_verifier=' + verifier;
  for (let i = 0; i < POLL_MAX; i++) {
    await new Promise(r => setTimeout(r, POLL_INTERVAL));
    try {
      const res = await httpRequest(pollUrl, 'GET');
      if (res.status === 200) {
        const data = JSON.parse(res.body);
        process.stderr.write('[' + MCP_NAME + '] Signed in as ' + data.displayName + ' <' + data.email + '>\\n');
        return data.sessionToken;
      }
      if (res.status === 410) {
        process.stderr.write('[' + MCP_NAME + '] Sign-in expired. Restart to try again.\\n');
        process.exit(1);
      }
    } catch { /* network error, retry */ }
  }
  process.stderr.write('[' + MCP_NAME + '] Timed out waiting for sign-in.\\n');
  process.exit(1);
}

// -- Verify token --
// /api/mcp always answers HTTP 200: an expired or missing session is reported
// INSIDE the JSON-RPC body as an isError result whose text asks the user to
// re-authenticate, never as an HTTP 401. And tools/list answers 200 even
// unauthenticated, so the old 'tools/list + status === 200' check treated a
// dead token as live — the extension then limped on, every real call failing.
// Make an AUTHENTICATED tools/call and inspect the body. Key on the re-auth
// signal specifically (not any isError) so a service-disabled or transient
// tool error on an otherwise-valid token never forces a pointless re-auth loop.

async function verifyToken(token) {
  try {
    const res = await httpRequest(
      MCP_URL + '/api/mcp', 'POST',
      JSON.stringify({ jsonrpc: '2.0', id: 1, method: 'tools/call', params: { name: 'list_folders_mail', arguments: {} } }),
      { 'Authorization': 'Bearer ' + token }
    );
    if (res.status !== 200) return false;
    let data;
    try { data = JSON.parse(res.body); } catch { return false; }
    const result = data && data.result;
    const text = result && Array.isArray(result.content)
      ? result.content.map((c) => (c && c.text) || '').join(' ')
      : '';
    if (result && result.isError && /re-authenticate|session expired/i.test(text)) {
      return false;
    }
    return true;
  } catch {
    return false;
  }
}

// -- MCP stdio <-> HTTP bridge --

async function bridgeMcp(token) {
  const rl = readline.createInterface({ input: process.stdin, terminal: false });

  for await (const line of rl) {
    if (!line.trim()) continue;
    let msg;
    try {
      msg = JSON.parse(line);
    } catch (err) {
      process.stdout.write(JSON.stringify({
        jsonrpc: '2.0',
        id: null,
        error: { code: -32700, message: 'Parse error: ' + err.message }
      }) + '\\n');
      continue;
    }
    // JSON-RPC notifications carry no id and must not receive a response.
    const isNotification = msg.id === undefined || msg.id === null;
    try {
      const res = await httpRequest(
        MCP_URL + '/api/mcp', 'POST',
        JSON.stringify(msg),
        { 'Authorization': 'Bearer ' + token }
      );

      let body = res.body;
      if (res.status === 401) {
        process.stderr.write('[' + MCP_NAME + '] Session expired. Re-authenticating...\\n');
        deleteToken();
        token = await authenticate();
        saveToken(token);
        const retry = await httpRequest(
          MCP_URL + '/api/mcp', 'POST',
          JSON.stringify(msg),
          { 'Authorization': 'Bearer ' + token }
        );
        body = retry.body;
      }

      if (isNotification) continue;
      if (body && body.trim()) process.stdout.write(body + '\\n');
    } catch (err) {
      if (isNotification) continue;
      process.stdout.write(JSON.stringify({
        jsonrpc: '2.0',
        id: msg.id ?? null,
        error: { code: -32603, message: 'Bridge error: ' + err.message }
      }) + '\\n');
    }
  }
}

// -- Main --

async function main() {
  await autoUpdate();

  let token = loadToken();

  if (token) {
    const valid = await verifyToken(token);
    if (!valid) {
      process.stderr.write('[' + MCP_NAME + '] Session expired. Re-authenticating...\\n');
      deleteToken();
      token = null;
    }
  }

  if (!token) {
    token = await authenticate();
    saveToken(token);
  }

  process.stderr.write('[' + MCP_NAME + '] Connected. MCP bridge running.\\n');
  await bridgeMcp(token);
}

main().catch((err) => {
  process.stderr.write('[' + MCP_NAME + '] Fatal error: ' + err.message + '\\n');
  process.exit(1);
});
`;
}

// ── Bundle generator ─────────────────────────────────────────────────────────

function generateBundle(mcpUrl: string, slug: string, displayName: string): Buffer {
  const keychainService = `ai.standardgauge.${slug}`;

  const entries: ZipEntry[] = [
    { name: 'manifest.json', data: Buffer.from(renderManifest(slug, displayName), 'utf-8') },
    { name: 'package.json', data: Buffer.from(renderPackageJson(slug), 'utf-8') },
    { name: 'server/index.js', data: Buffer.from(renderServerJs(mcpUrl, slug, keychainService), 'utf-8') },
  ];

  return buildZip(entries);
}

// ── Landing page ─────────────────────────────────────────────────────────────

function renderLandingPage(origin: string, slug: string): string {
  const instanceName = process.env.MCP_INSTANCE_NAME ?? 'M365 MCP';
  const escaped = instanceName.replace(/[<>&]/g, (c) =>
    ({ '<': '&lt;', '>': '&gt;', '&': '&amp;' }[c] ?? c)
  );
  return `<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<title>${escaped} — Install</title>
<meta name="viewport" content="width=device-width,initial-scale=1">
<style>
  body { font: 16px/1.5 -apple-system, BlinkMacSystemFont, "Segoe UI", system-ui, sans-serif;
         max-width: 720px; margin: 3em auto; padding: 0 1.5em; color: #1a1a1a; }
  h1 { font-size: 1.6em; margin-bottom: 0.3em; }
  h2 { font-size: 1.1em; margin-top: 2em; margin-bottom: 0.5em; color: #444; }
  p { margin: 0.7em 0; }
  .download-btn {
    display: inline-block; margin: 1.2em 0; padding: 0.7em 1.8em;
    background: #0078d4; color: #fff; border-radius: 6px;
    text-decoration: none; font-size: 1em; font-weight: 600;
  }
  .download-btn:hover { background: #005ea2; }
  .filename { font-family: "SF Mono", Menlo, Monaco, Consolas, monospace;
              font-size: 0.92em; color: #444; }
  ol { padding-left: 1.5em; }
  ol li { margin: 0.5em 0; }
  .meta { color: #666; font-size: 0.9em; margin-top: 3em;
          padding-top: 1em; border-top: 1px solid #e0e0e2; }
</style>
</head>
<body>
  <h1>${escaped}</h1>
  <p>This server provides Microsoft 365 access (mail, OneDrive, SharePoint, Calendar, Contacts, OneNote) to Claude Desktop and Claude Code via the Model Context Protocol.</p>

  <h2>Install for Claude Desktop</h2>
  <a class="download-btn" href="${origin}/install.mcpb" download="${slug}.mcpb">Download Extension</a>
  <p class="filename">${slug}.mcpb</p>

  <h2>Setup</h2>
  <ol>
    <li>Download the extension file above.</li>
    <li>Open Claude Desktop and go to <strong>Settings</strong> (gear icon).</li>
    <li>Navigate to <strong>Desktop App</strong>, then <strong>Extensions</strong>.</li>
    <li>Click "Install from file" and select the downloaded <span class="filename">.mcpb</span> file.</li>
    <li>On first use, a browser window will open for Microsoft 365 sign-in. Sign in and the extension will connect automatically.</li>
  </ol>

  <h2>What it does</h2>
  <p>The extension runs a lightweight bridge between Claude Desktop and this server. Your session token is stored locally — in macOS Keychain on Mac, or in a private file in your home directory on Windows and Linux.</p>
  <p>If your session expires, the extension will automatically open a browser window to re-authenticate.</p>

  <p class="meta">Per-user authentication. No shared API keys. Server-side sessions stored encrypted at rest. To re-authenticate, restart the extension or reinstall.</p>
</body>
</html>`;
}

// ── Route handlers ───────────────────────────────────────────────────────────

async function installLandingHandler(
  request: HttpRequest,
  _context: InvocationContext,
): Promise<HttpResponseInit> {
  const origin = getPublicOrigin(request);
  const host = getHost(request);
  const slug = getMcpSlug(host);

  return {
    status: 200,
    headers: {
      'Content-Type': 'text/html; charset=utf-8',
      'Cache-Control': 'public, max-age=300',
      ...installLandingHeaders,
    },
    body: renderLandingPage(origin, slug),
  };
}

async function installBundleHandler(
  request: HttpRequest,
  context: InvocationContext,
): Promise<HttpResponseInit> {
  try {
    const origin = getPublicOrigin(request);
    const host = getHost(request);
    const slug = getMcpSlug(host);
    const displayName = process.env.MCP_INSTANCE_NAME ?? `${slug} MCP`;
    const bundle = generateBundle(origin, slug, displayName);

    return {
      status: 200,
      headers: {
        'Content-Type': 'application/octet-stream',
        'Content-Disposition': `attachment; filename="${slug}.mcpb"`,
        'Content-Length': bundle.length.toString(),
        'Cache-Control': 'public, max-age=300',
      },
      body: bundle,
    };
  } catch (err) {
    context.error('install.mcpb endpoint error:', err);
    return {
      status: 500,
      jsonBody: { error: 'Failed to generate extension bundle', detail: String(err) },
    };
  }
}

// ── Extension version + update endpoints ─────────────────────────────────────

async function extensionVersionHandler(
  _request: HttpRequest,
  _context: InvocationContext,
): Promise<HttpResponseInit> {
  return {
    status: 200,
    headers: { 'Content-Type': 'application/json', 'Cache-Control': 'public, max-age=60' },
    jsonBody: { version: EXTENSION_VERSION },
  };
}

async function extensionUpdateHandler(
  request: HttpRequest,
  context: InvocationContext,
): Promise<HttpResponseInit> {
  try {
    const origin = getPublicOrigin(request);
    const host = getHost(request);
    const slug = getMcpSlug(host);
    const displayName = process.env.MCP_INSTANCE_NAME ?? `${slug} MCP`;
    const keychainService = `ai.standardgauge.${slug}`;

    return {
      status: 200,
      headers: { 'Content-Type': 'application/json', 'Cache-Control': 'public, max-age=60' },
      jsonBody: {
        version: EXTENSION_VERSION,
        files: {
          'manifest.json': renderManifest(slug, displayName),
          'package.json': renderPackageJson(slug),
          'server/index.js': renderServerJs(origin, slug, keychainService),
        },
      },
    };
  } catch (err) {
    context.error('extension-update endpoint error:', err);
    return { status: 500, jsonBody: { error: 'Failed to generate update payload' } };
  }
}

// ── Scripted installers + stdio shim ────────────────────────────────────────

/**
 * Read a file shipped under src/install/.
 *
 * Resolution order:
 *   1. dist/install/ — the production path after `npm run build` (postbuild copies
 *      src/install → dist/install, which is what the Docker runtime image contains).
 *   2. src/install/ — the source tree fallback for running tests before a build.
 */
function readInstallFile(name: string): string {
  const cwd = process.cwd();
  const candidates = [
    path.join(cwd, 'dist', 'install', name),
    path.join(cwd, 'src', 'install', name),
  ];
  for (const p of candidates) {
    try {
      return fs.readFileSync(p, 'utf-8');
    } catch { /* try next candidate */ }
  }
  throw new Error(`${name} not found; searched: ${candidates.join(', ')}`);
}

function renderInstallTemplate(name: string, mcpUrl: string, mcpName: string): string {
  const instanceName = process.env.MCP_INSTANCE_NAME ?? 'M365 MCP';
  return readInstallFile(name)
    .replace(/\{\{MCP_URL\}\}/g, mcpUrl)
    .replace(/\{\{MCP_NAME\}\}/g, mcpName)
    .replace(/\{\{INSTANCE_NAME\}\}/g, instanceName);
}

function renderPs1Script(mcpUrl: string, mcpName: string): string {
  return renderInstallTemplate('install-mcp.ps1.template', mcpUrl, mcpName);
}

// The local stdio process the installers configure in place of
// `npx -y supergateway`. Served verbatim: the server URL and token
// reach it as arguments in the MCP client config, so one file serves every
// tenant and a re-install picks up whatever canonical last shipped.
const SHIM_FILE = 'm365-mcp-shim.js';

async function installShimHandler(
  _request: HttpRequest,
  context: InvocationContext,
): Promise<HttpResponseInit> {
  try {
    return {
      status: 200,
      headers: {
        'Content-Type': 'application/javascript; charset=utf-8',
        'Content-Disposition': `inline; filename="${SHIM_FILE}"`,
        'Cache-Control': 'public, max-age=300',
      },
      body: readInstallFile(SHIM_FILE),
    };
  } catch (err) {
    context.error('install shim endpoint error:', err);
    return { status: 500, jsonBody: { error: 'Failed to load shim' } };
  }
}

async function installShHandler(
  request: HttpRequest,
  context: InvocationContext,
): Promise<HttpResponseInit> {
  try {
    const origin = getPublicOrigin(request);
    const slug = getMcpSlug(getHost(request));
    return {
      status: 200,
      headers: {
        'Content-Type': 'text/plain; charset=utf-8',
        'Content-Disposition': 'inline; filename="install.sh"',
        'Cache-Control': 'public, max-age=300',
      },
      body: renderInstallTemplate('install-mcp.sh.template', origin, slug),
    };
  } catch (err) {
    context.error('install.sh endpoint error:', err);
    return { status: 500, jsonBody: { error: 'Failed to render installer script' } };
  }
}

async function installPs1Handler(
  request: HttpRequest,
  context: InvocationContext,
): Promise<HttpResponseInit> {
  try {
    const origin = getPublicOrigin(request);
    const host = getHost(request);
    const slug = getMcpSlug(host);
    const script = renderPs1Script(origin, slug);

    return {
      status: 200,
      headers: {
        'Content-Type': 'text/plain; charset=utf-8',
        'Content-Disposition': 'inline; filename="install.ps1"',
        'Cache-Control': 'public, max-age=300',
      },
      body: script,
    };
  } catch (err) {
    context.error('install.ps1 endpoint error:', err);
    return {
      status: 500,
      jsonBody: { error: 'Failed to render installer script', detail: String(err) },
    };
  }
}

// ── Function registrations ───────────────────────────────────────────────────

app.http('installLanding', {
  methods: ['GET'],
  authLevel: 'anonymous',
  route: 'install',
  handler: withSecurity(installLandingHandler),
});

app.http('installBundle', {
  methods: ['GET'],
  authLevel: 'anonymous',
  route: 'install.mcpb',
  handler: withSecurity(installBundleHandler),
});

app.http('extensionVersion', {
  methods: ['GET'],
  authLevel: 'anonymous',
  route: 'api/extension-version',
  handler: withSecurity(extensionVersionHandler),
});

app.http('extensionUpdate', {
  methods: ['GET'],
  authLevel: 'anonymous',
  route: 'api/extension-update',
  handler: withSecurity(extensionUpdateHandler),
});

app.http('installPs1', {
  methods: ['GET'],
  authLevel: 'anonymous',
  route: 'install.ps1',
  handler: withSecurity(installPs1Handler),
});

app.http('installSh', {
  methods: ['GET'],
  authLevel: 'anonymous',
  route: 'install.sh',
  handler: withSecurity(installShHandler),
});

app.http('installShim', {
  methods: ['GET'],
  authLevel: 'anonymous',
  route: `install/${SHIM_FILE}`,
  handler: withSecurity(installShimHandler),
});
