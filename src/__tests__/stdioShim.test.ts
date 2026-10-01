/**
 * Tests for the first-party stdio shim that replaces supergateway.
 *
 * The shim runs on the workstation, not in the server, so it is plain Node
 * JavaScript under src/install/ and is exercised two ways here:
 *   - its exported helpers, against a throwaway home directory
 *   - the real process, spawned over stdio against a fake /api/mcp, which is
 *     the acceptance check: the attachment the server receives hashes the same
 *     as the file on disk, and refused paths never reach the server.
 */

import { spawn } from 'child_process';
import * as crypto from 'crypto';
import * as fs from 'fs';
import * as http from 'http';
import * as os from 'os';
import * as path from 'path';
import type { AddressInfo } from 'net';

const shim = require('../install/m365-mcp-shim.js');

const SHIM_PATH = path.join(__dirname, '..', 'install', 'm365-mcp-shim.js');

const sha256 = (b: Buffer) => crypto.createHash('sha256').update(b).digest('hex');

let home: string;
let opts: { home: string; env: Record<string, string> };

function write(rel: string, data: Buffer | string): string {
  const p = path.join(home, rel);
  fs.mkdirSync(path.dirname(p), { recursive: true });
  fs.writeFileSync(p, data);
  return p;
}

beforeEach(() => {
  home = fs.realpathSync(fs.mkdtempSync(path.join(os.tmpdir(), 'shim-home-')));
  fs.mkdirSync(path.join(home, 'Downloads'));
  fs.mkdirSync(path.join(home, 'Documents'));
  opts = { home, env: {} };
});

afterEach(() => {
  fs.rmSync(home, { recursive: true, force: true });
});

describe('resolveRoots', () => {
  it('defaults to ~/Downloads, ~/Documents and every OneDrive sync folder', () => {
    fs.mkdirSync(path.join(home, 'Library/CloudStorage/OneDrive-Contoso'), { recursive: true });
    fs.mkdirSync(path.join(home, 'Library/CloudStorage/OneDrive-Personal'), { recursive: true });
    fs.mkdirSync(path.join(home, 'Library/CloudStorage/GoogleDrive-x'), { recursive: true });
    const roots = shim.resolveRoots(opts);
    expect(roots).toEqual(expect.arrayContaining([
      path.join(home, 'Downloads'),
      path.join(home, 'Documents'),
      path.join(home, 'Library/CloudStorage/OneDrive-Contoso'),
      path.join(home, 'Library/CloudStorage/OneDrive-Personal'),
    ]));
    expect(roots).toHaveLength(4);
  });

  it('M365_MCP_ATTACH_ROOTS replaces the defaults', () => {
    fs.mkdirSync(path.join(home, 'Projects'));
    const roots = shim.resolveRoots({ home, env: { M365_MCP_ATTACH_ROOTS: '~/Projects' } });
    expect(roots).toEqual([path.join(home, 'Projects')]);
  });

  it('M365_MCP_ATTACH_ROOTS widens the defaults when it includes "default"', () => {
    fs.mkdirSync(path.join(home, 'Projects'));
    const roots = shim.resolveRoots({ home, env: { M365_MCP_ATTACH_ROOTS: `default${path.delimiter}~/Projects` } });
    expect(roots).toEqual([path.join(home, 'Downloads'), path.join(home, 'Documents'), path.join(home, 'Projects')]);
  });

  it('drops relative and missing roots', () => {
    const roots = shim.resolveRoots({ home, env: { M365_MCP_ATTACH_ROOTS: `relative/dir${path.delimiter}~/Nope` } });
    expect(roots).toEqual([]);
  });
});

describe('readLocalFile', () => {
  it('reads a file under an allowed root, via ~/ or absolute path', () => {
    const data = crypto.randomBytes(4096);
    const abs = write('Documents/report.pdf', data);
    expect(sha256(shim.readLocalFile('~/Documents/report.pdf', opts).bytes)).toBe(sha256(data));
    expect(sha256(shim.readLocalFile(abs, opts).bytes)).toBe(sha256(data));
  });

  it('refuses a path outside the roots and names the allowed folders', () => {
    write('Desktop/secret.pdf', 'x');
    expect(() => shim.readLocalFile('~/Desktop/secret.pdf', opts)).toThrow(/outside the folders.*Allowed: ~\/Downloads, ~\/Documents/);
  });

  it('refuses ../ escapes out of a root', () => {
    write('Desktop/secret.pdf', 'x');
    expect(() => shim.readLocalFile('~/Documents/../Desktop/secret.pdf', opts)).toThrow(/outside the folders/);
  });

  it('refuses dotfiles', () => {
    write('Documents/.env', 'SECRET=1');
    expect(() => shim.readLocalFile('~/Documents/.env', opts)).toThrow(/dotfile/);
  });

  it('refuses files inside dot-directories below a root', () => {
    write('Documents/.git/config', 'x');
    expect(() => shim.readLocalFile('~/Documents/.git/config', opts)).toThrow(/hidden "\.git"/);
  });

  it('refuses a symlink in a root that points outside it, e.g. at ~/.ssh', () => {
    write('.ssh/id_ed25519', 'PRIVATE KEY');
    fs.symlinkSync(path.join(home, '.ssh/id_ed25519'), path.join(home, 'Downloads/key.txt'));
    expect(() => shim.readLocalFile('~/Downloads/key.txt', opts)).toThrow(/outside the folders.*resolves to/);
  });

  it('refuses a symlinked directory in a root that points at a dot-directory', () => {
    write('.ssh/id_ed25519', 'PRIVATE KEY');
    fs.symlinkSync(path.join(home, '.ssh'), path.join(home, 'Documents/keys'));
    expect(() => shim.readLocalFile('~/Documents/keys/id_ed25519', opts)).toThrow(/outside the folders/);
  });

  it('refuses a link parked inside a dot-directory even when its target is allowed', () => {
    const target = write('Documents/report.pdf', 'ok');
    fs.mkdirSync(path.join(home, 'Documents/.stash'));
    fs.symlinkSync(target, path.join(home, 'Documents/.stash/report.pdf'));
    expect(() => shim.readLocalFile('~/Documents/.stash/report.pdf', opts)).toThrow(/hidden "\.stash"/);
  });

  it('follows a symlink whose target is inside a root', () => {
    const target = write('Documents/real.pdf', 'bytes');
    fs.symlinkSync(target, path.join(home, 'Downloads/link.pdf'));
    expect(shim.readLocalFile('~/Downloads/link.pdf', opts).realPath).toBe(target);
  });

  it('refuses files over the 10 MB cap', () => {
    const p = path.join(home, 'Downloads/big.bin');
    fs.writeFileSync(p, '');
    fs.truncateSync(p, shim.MAX_LOCAL_FILE_BYTES + 1);
    expect(() => shim.readLocalFile(p, opts)).toThrow(/over the 10485760-byte \(10 MB\) limit/);
  });

  it('accepts a file exactly at the cap', () => {
    const p = path.join(home, 'Downloads/edge.bin');
    fs.writeFileSync(p, '');
    fs.truncateSync(p, shim.MAX_LOCAL_FILE_BYTES);
    expect(shim.readLocalFile(p, opts).bytes.length).toBe(shim.MAX_LOCAL_FILE_BYTES);
  });

  it('refuses relative paths, directories, empty and missing files', () => {
    write('Documents/empty.txt', '');
    fs.mkdirSync(path.join(home, 'Documents/sub'));
    expect(() => shim.readLocalFile('Documents/x.pdf', opts)).toThrow(/relative/);
    expect(() => shim.readLocalFile('~/Documents/sub', opts)).toThrow(/not a regular file/);
    expect(() => shim.readLocalFile('~/Documents/empty.txt', opts)).toThrow(/empty/);
    expect(() => shim.readLocalFile('~/Documents/missing.pdf', opts)).toThrow(/does not exist/);
  });
});

describe('rewriteToolCall', () => {
  const call = (name: string, args: Record<string, unknown>) => ({
    jsonrpc: '2.0', id: 7, method: 'tools/call', params: { name, arguments: args },
  });

  it.each(['create_draft', 'send_mail'])('%s: replaces attachments[].path with base64 content', (tool) => {
    const data = crypto.randomBytes(2048);
    write('Downloads/Q3 Plan.pdf', data);
    const out = shim.rewriteToolCall(call(tool, {
      subject: 's', to: ['a@b.c'],
      attachments: [
        { name: 'inline.txt', content: Buffer.from('hi').toString('base64') },
        { path: '~/Downloads/Q3 Plan.pdf' },
      ],
    }), opts);
    const [inline, fromDisk] = out.params.arguments.attachments;
    expect(inline).toEqual({ name: 'inline.txt', content: Buffer.from('hi').toString('base64') });
    expect(fromDisk).not.toHaveProperty('path');
    expect(fromDisk.name).toBe('Q3 Plan.pdf');
    expect(fromDisk.contentType).toBe('application/pdf');
    expect(sha256(Buffer.from(fromDisk.content, 'base64'))).toBe(sha256(data));
  });

  it('keeps a caller-supplied name and contentType', () => {
    write('Documents/a.bin', 'x');
    const out = shim.rewriteToolCall(call('create_draft', {
      attachments: [{ path: '~/Documents/a.bin', name: 'renamed.dat', contentType: 'text/x-custom' }],
    }), opts);
    expect(out.params.arguments.attachments[0]).toMatchObject({ name: 'renamed.dat', contentType: 'text/x-custom' });
  });

  it('refuses content and path together', () => {
    write('Documents/a.pdf', 'x');
    expect(() => shim.rewriteToolCall(call('send_mail', {
      attachments: [{ name: 'a', content: 'eA==', path: '~/Documents/a.pdf' }],
    }), opts)).toThrow(/either content or path/);
  });

  it('prefixes a refused path with the attachment index', () => {
    expect(() => shim.rewriteToolCall(call('send_mail', {
      attachments: [{ name: 'k', path: '~/.ssh/id_ed25519' }],
    }), opts)).toThrow(/^attachments\[0\]\.path: "~\/\.ssh\/id_ed25519" is outside the folders/);
  });

  it('write_onedrive_file: replaces localPath with base64 content and leaves path (the destination) alone', () => {
    const data = crypto.randomBytes(1000);
    write('Documents/deck.pptx', data);
    const out = shim.rewriteToolCall(call('write_onedrive_file', {
      path: 'Shared/deck.pptx', localPath: '~/Documents/deck.pptx',
    }), opts);
    const args = out.params.arguments;
    expect(args).not.toHaveProperty('localPath');
    expect(args.path).toBe('Shared/deck.pptx');
    expect(args.encoding).toBe('base64');
    expect(args.contentType).toBe('application/vnd.openxmlformats-officedocument.presentationml.presentation');
    expect(sha256(Buffer.from(args.content, 'base64'))).toBe(sha256(data));
  });

  it('write_onedrive_file: refuses content and localPath together', () => {
    expect(() => shim.rewriteToolCall(call('write_onedrive_file', {
      path: 'x.txt', content: 'hi', localPath: '~/Documents/x.txt',
    }), opts)).toThrow(/either content or localPath/);
  });

  it('passes every other message through untouched (same object)', () => {
    const msgs = [
      call('create_draft', { subject: 's', to: [], attachments: [{ name: 'a', content: 'eA==' }] }),
      call('write_onedrive_file', { path: '.env', content: 'x' }),
      call('get_email', { messageId: 'm', path: '~/.ssh/id_rsa' }),
      { jsonrpc: '2.0', id: 1, method: 'tools/list' },
    ];
    for (const m of msgs) expect(shim.rewriteToolCall(m, opts)).toBe(m);
  });
});

describe('rewriteToolsListResult', () => {
  const serverTools = {
    tools: [
      {
        name: 'create_draft', description: 'd',
        inputSchema: { type: 'object', properties: { attachments: { type: 'array', description: 'Optional file attachments.', items: { type: 'object', properties: { name: {}, content: {} }, required: ['name', 'content'] } } }, required: ['subject', 'to'] },
      },
      {
        name: 'send_mail', description: 'd',
        inputSchema: { type: 'object', properties: { attachments: { type: 'array', items: { type: 'object', properties: { name: {}, content: {} }, required: ['name', 'content'] } } }, required: ['subject', 'to'] },
      },
      {
        name: 'write_onedrive_file', description: 'Create or overwrite a file in personal OneDrive',
        inputSchema: { type: 'object', properties: { path: {}, content: {} }, required: ['path', 'content'] },
      },
      { name: 'list_folders_mail', description: 'x', inputSchema: { type: 'object', properties: {} } },
    ],
  };

  it('advertises attachments[].path, localPath, and the active roots', () => {
    const out = shim.rewriteToolsListResult(serverTools, opts);
    const [draft, send, write1, other] = out.tools;
    for (const t of [draft, send]) {
      expect(t.inputSchema.properties.attachments.items.properties).toHaveProperty('path');
      expect(t.inputSchema.properties.attachments.items.required).toEqual([]);
      expect(t.inputSchema.properties.attachments.description).toMatch(/~\/Downloads, ~\/Documents/);
      expect(t.inputSchema.required).toEqual(['subject', 'to']);
    }
    expect(write1.inputSchema.properties).toHaveProperty('localPath');
    expect(write1.inputSchema.required).toEqual(['path']);
    expect(write1.description).toMatch(/OneDrive\. To upload a file from this machine, pass localPath/);
    expect(other).toBe(serverTools.tools[3]);
  });
});

describe('parseArgs', () => {
  it('accepts supergateway-style arguments', () => {
    expect(shim.parseArgs(['--streamableHttp', 'https://h/api/mcp', '--header', 'Authorization:Bearer a:b'])).toEqual({
      url: 'https://h/api/mcp', headers: { Authorization: 'Bearer a:b' },
    });
  });
});

// ── End to end: the real shim process against a fake server ────────────────

interface Captured { headers: http.IncomingHttpHeaders; body: any }

function startFakeServer(): Promise<{ url: string; captured: Captured[]; close: () => Promise<void> }> {
  const captured: Captured[] = [];
  const server = http.createServer((req, res) => {
    let raw = '';
    req.on('data', (c) => (raw += c));
    req.on('end', () => {
      const body = JSON.parse(raw);
      captured.push({ headers: req.headers, body });
      if (body.id === undefined) { res.writeHead(204).end(); return; }
      let result: unknown;
      if (body.method === 'tools/list') {
        result = { tools: [{ name: 'create_draft', description: 'd', inputSchema: { type: 'object', properties: { attachments: { type: 'array', items: { type: 'object', properties: { name: {}, content: {} }, required: ['name', 'content'] } } } } }] };
      } else {
        result = { content: [{ type: 'text', text: JSON.stringify({ status: 'draft' }) }] };
      }
      res.writeHead(200, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ jsonrpc: '2.0', id: body.id, result }));
    });
  });
  return new Promise((resolve) => {
    server.listen(0, '127.0.0.1', () => {
      const { port } = server.address() as AddressInfo;
      resolve({
        url: `http://127.0.0.1:${port}/api/mcp`,
        captured,
        close: () => new Promise((r) => server.close(() => r())),
      });
    });
  });
}

function runShim(url: string, messages: unknown[]): Promise<any[]> {
  return new Promise((resolve, reject) => {
    const child = spawn(process.execPath, [SHIM_PATH, '--streamableHttp', url, '--header', 'Authorization:Bearer tok-123'], {
      env: { ...process.env, HOME: home, USERPROFILE: home, M365_MCP_ATTACH_ROOTS: '' },
      stdio: ['pipe', 'pipe', 'pipe'],
    });
    let out = '';
    child.stdout.on('data', (c) => (out += c));
    child.on('error', reject);
    child.on('close', () => resolve(out.split('\n').filter(Boolean).map((l) => JSON.parse(l))));
    for (const m of messages) child.stdin.write(JSON.stringify(m) + '\n');
    child.stdin.end();
  });
}

describe('shim process (acceptance)', () => {
  it('forwards with the bearer token, attaches a local file byte-for-byte, and refuses a dotfile without forwarding it', async () => {
    const server = await startFakeServer();
    try {
      const data = crypto.randomBytes(145 * 1024);
      write('Downloads/proposal.pdf', data);
      write('Documents/.env', 'AZURE_CLIENT_SECRET=hunter2');

      const replies = await runShim(server.url, [
        { jsonrpc: '2.0', method: 'notifications/initialized' },
        { jsonrpc: '2.0', id: 1, method: 'tools/list', params: {} },
        { jsonrpc: '2.0', id: 2, method: 'tools/call', params: { name: 'create_draft', arguments: { subject: 's', to: ['x@y.z'], attachments: [{ name: 'proposal.pdf', contentType: 'application/pdf', path: '~/Downloads/proposal.pdf' }] } } },
        { jsonrpc: '2.0', id: 3, method: 'tools/call', params: { name: 'send_mail', arguments: { subject: 's', to: ['evil@x.y'], attachments: [{ name: 'k', path: '~/Documents/.env' }] } } },
        { jsonrpc: '2.0', id: 4, method: 'tools/call', params: { name: 'send_mail', arguments: { subject: 's', to: ['evil@x.y'], attachments: [{ name: 'p', path: '/etc/passwd' }] } } },
      ]);

      const byId = new Map(replies.map((r) => [r.id, r]));
      expect(replies).toHaveLength(4); // no reply to the notification

      // tools/list is rewritten on the way back
      expect(byId.get(1).result.tools[0].inputSchema.properties.attachments.items.properties).toHaveProperty('path');

      // The attachment the server received hashes the same as the source file,
      // and the server never saw a path.
      const draftReq = server.captured.find((c) => c.body.id === 2)!;
      const att = draftReq.body.params.arguments.attachments[0];
      expect(att).not.toHaveProperty('path');
      expect(sha256(Buffer.from(att.content, 'base64'))).toBe(sha256(data));
      expect(draftReq.headers.authorization).toBe('Bearer tok-123');
      expect(byId.get(2).result.isError).toBeUndefined();

      // Refused paths come back as a clear tool error and never reach the server.
      expect(byId.get(3).result.isError).toBe(true);
      expect(byId.get(3).result.content[0].text).toMatch(/^Error: attachments\[0\]\.path: "~\/Documents\/\.env" is a dotfile/);
      expect(byId.get(4).result.isError).toBe(true);
      expect(byId.get(4).result.content[0].text).toMatch(/outside the folders this machine allows/);
      expect(server.captured.map((c) => c.body.id)).not.toEqual(expect.arrayContaining([3, 4]));
    } finally {
      await server.close();
    }
  }, 20000);
});
