#!/usr/bin/env node
/**
 * M365 MCP stdio shim.
 *
 * The local process an MCP client (Claude Code, Claude Desktop) launches over
 * stdio. It forwards every JSON-RPC message to the server's /api/mcp endpoint
 * and writes the reply back, the job `npx -y supergateway` used to do. It is
 * served by the server at /install/m365-mcp-shim.js and written to disk by the
 * installers, so it is versioned with canonical and nothing is fetched from npm
 * at launch.
 *
 * It does one thing supergateway did not: it lets a tool call name a LOCAL file
 * instead of carrying its bytes.
 *
 *   create_draft / send_mail   attachments[i].path  -> attachments[i].content
 *   write_onedrive_file        localPath            -> content + encoding: base64
 *
 * The shim reads the file and substitutes base64 before forwarding, so the
 * bytes never pass through model context and the server never sees a path.
 * tools/list is rewritten so clients see the extra fields.
 *
 * Local reads are the exfiltration surface (a prompt-injected agent asked to
 * mail ~/.ssh/id_ed25519 somewhere), so every path is checked here:
 *   - absolute or ~/ only; relative paths are refused
 *   - symlinks are resolved first, then the real path must sit under an
 *     allowed root (also resolved)
 *   - no dotfile or dot-directory anywhere below the root
 *   - regular files only, at most 10 MB
 * Default roots: ~/Downloads, ~/Documents, ~/Library/CloudStorage/OneDrive-*
 * (on Windows, ~/OneDrive* instead of the CloudStorage glob).
 * M365_MCP_ATTACH_ROOTS (path-delimiter separated, `:` on macOS/Linux, `;` on
 * Windows) replaces them; include the entry `default` to keep the defaults and
 * add to them.
 *
 * Usage (arguments match supergateway's, so configs change only the command):
 *   node m365-mcp-shim.js --streamableHttp https://host/api/mcp \
 *       --header "Authorization:Bearer <token>"
 *
 * Node 18+ (global fetch). No dependencies.
 */

'use strict';

const fs = require('fs');
const os = require('os');
const path = require('path');
const readline = require('readline');

const SHIM_VERSION = '1.0.0';
const MAX_LOCAL_FILE_BYTES = 10 * 1024 * 1024;
const ATTACHMENT_TOOLS = new Set(['create_draft', 'send_mail']);
const LOCAL_PATH_TOOL = 'write_onedrive_file';

// Extension -> MIME for the common office and image types. Anything else goes
// up as application/octet-stream, which Outlook and OneDrive both handle.
const CONTENT_TYPES = {
  '.pdf': 'application/pdf',
  '.doc': 'application/msword',
  '.docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  '.xls': 'application/vnd.ms-excel',
  '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  '.ppt': 'application/vnd.ms-powerpoint',
  '.pptx': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
  '.csv': 'text/csv',
  '.txt': 'text/plain',
  '.md': 'text/markdown',
  '.html': 'text/html',
  '.json': 'application/json',
  '.xml': 'application/xml',
  '.zip': 'application/zip',
  '.png': 'image/png',
  '.jpg': 'image/jpeg',
  '.jpeg': 'image/jpeg',
  '.gif': 'image/gif',
  '.webp': 'image/webp',
  '.svg': 'image/svg+xml',
  '.ics': 'text/calendar',
  '.eml': 'message/rfc822',
};

function guessContentType(filePath) {
  return CONTENT_TYPES[path.extname(filePath).toLowerCase()] || 'application/octet-stream';
}

/** Error whose message is safe to hand back to the model verbatim. */
class LocalPathError extends Error {}

// ── Allowed roots ────────────────────────────────────────────────────────────

function expandHome(p, home) {
  if (p === '~') return home;
  if (p.startsWith('~/') || p.startsWith('~\\')) return path.join(home, p.slice(2));
  return p;
}

/**
 * Expand one root entry to concrete directories. A `*` is honoured in the last
 * segment only (enough for OneDrive-*); missing directories are dropped.
 */
function expandRootEntry(entry, home) {
  const expanded = expandHome(entry.trim(), home);
  if (!path.isAbsolute(expanded)) return [];
  const base = path.basename(expanded);
  if (!base.includes('*')) return [expanded];
  const parent = path.dirname(expanded);
  const pattern = new RegExp(
    '^' + base.split('*').map((s) => s.replace(/[.+?^${}()|[\]\\]/g, '\\$&')).join('.*') + '$',
  );
  let names;
  try {
    names = fs.readdirSync(parent);
  } catch {
    return [];
  }
  return names.filter((n) => pattern.test(n)).map((n) => path.join(parent, n));
}

// The local OneDrive sync folders are ~/Library/CloudStorage/OneDrive-<org> on
// macOS and ~/OneDrive - <org> (or plain ~/OneDrive) on Windows.
function defaultRootEntries(platform) {
  const oneDrive = (platform || process.platform) === 'win32' ? '~/OneDrive*' : '~/Library/CloudStorage/OneDrive-*';
  return ['~/Downloads', '~/Documents', oneDrive];
}

/**
 * The allowed roots as real (symlink-resolved) directory paths. Entries that
 * don't exist or aren't directories are skipped. Recomputed per call so a
 * OneDrive folder that appears after launch is picked up.
 */
function rootDirs(opts) {
  const home = (opts && opts.home) || os.homedir();
  const env = (opts && opts.env) || process.env;
  const raw = env.M365_MCP_ATTACH_ROOTS;
  let entries;
  if (raw && raw.trim()) {
    entries = [];
    for (const e of raw.split(path.delimiter)) {
      if (!e.trim()) continue;
      if (e.trim() === 'default') entries.push(...defaultRootEntries());
      else entries.push(e);
    }
  } else {
    entries = defaultRootEntries();
  }
  return entries.flatMap((entry) => expandRootEntry(entry, home)).map((d) => path.normalize(d));
}

function resolveRoots(opts) {
  const roots = [];
  for (const dir of rootDirs(opts)) {
    try {
      const real = fs.realpathSync(dir);
      if (fs.statSync(real).isDirectory() && !roots.includes(real)) roots.push(real);
    } catch {
      /* missing root: skip */
    }
  }
  return roots;
}

/** First path segment below a root that starts with a dot, if any. */
function hiddenSegment(root, target) {
  return path.relative(root, target).split(path.sep).find((seg) => seg.startsWith('.'));
}

/** Roots for display in tool descriptions and errors, with $HOME shown as ~. */
function describeRoots(roots, home) {
  if (roots.length === 0) return '(none: no allowed folder exists on this machine)';
  const h = home || os.homedir();
  return roots
    .map((r) => (r === h || r.startsWith(h + path.sep) ? '~' + r.slice(h.length) : r))
    .join(', ');
}

// ── Local file read ──────────────────────────────────────────────────────────

function isInside(root, target) {
  const rel = path.relative(root, target);
  return rel !== '' && !rel.startsWith('..') && !path.isAbsolute(rel);
}

/**
 * Resolve and read a caller-supplied local path under the allowlist. Returns
 * { realPath, bytes }. Throws LocalPathError with a message meant for the model.
 */
function readLocalFile(requested, opts) {
  const home = (opts && opts.home) || os.homedir();
  const maxBytes = (opts && opts.maxBytes) || MAX_LOCAL_FILE_BYTES;
  if (typeof requested !== 'string' || requested.trim() === '') {
    throw new LocalPathError('path must be a non-empty string');
  }
  const expanded = expandHome(requested.trim(), home);
  if (!path.isAbsolute(expanded)) {
    throw new LocalPathError(`"${requested}" is relative; give an absolute path or one starting with ~/`);
  }
  const normalized = path.normalize(expanded);
  if (path.basename(normalized).startsWith('.')) {
    throw new LocalPathError(`"${requested}" is a dotfile; hidden files cannot be attached`);
  }
  // Check the path as written before touching the filesystem, so nothing
  // outside the roots can be probed for existence, and so a link parked inside
  // a dot-directory is refused even when its target is not. The resolved path
  // is checked again below.
  const roots = resolveRoots(opts);
  const lexicalRoots = [...rootDirs(opts), ...roots];
  let realPath;
  const outside = () => {
    const via = realPath && realPath !== normalized ? ` (it resolves to ${realPath})` : '';
    return new LocalPathError(
      `"${requested}" is outside the folders this machine allows attaching from${via}. ` +
        `Allowed: ${describeRoots(roots, home)}. Move or copy the file into one of them, ` +
        'or set M365_MCP_ATTACH_ROOTS in the MCP server config.',
    );
  };
  if (!lexicalRoots.some((dir) => isInside(dir, normalized))) throw outside();
  for (const dir of lexicalRoots) {
    const seg = isInside(dir, normalized) && hiddenSegment(dir, normalized);
    if (seg) throw new LocalPathError(`"${requested}" is inside hidden "${seg}"; dotfiles and dot-directories cannot be attached`);
  }

  try {
    realPath = fs.realpathSync(normalized);
  } catch (err) {
    const code = err && err.code;
    if (code === 'ENOENT' || code === 'ENOTDIR') throw new LocalPathError(`"${requested}" does not exist`);
    if (code === 'EACCES' || code === 'EPERM') throw new LocalPathError(`"${requested}" is not readable`);
    throw new LocalPathError(`"${requested}" could not be resolved`);
  }

  const root = roots.find((r) => isInside(r, realPath));
  if (!root) throw outside();
  const hidden = hiddenSegment(root, realPath);
  if (hidden) {
    throw new LocalPathError(`"${requested}" is inside hidden "${hidden}"; dotfiles and dot-directories cannot be attached`);
  }

  // Open the resolved path without following a final symlink (where the OS
  // supports it), then check and read through the same descriptor, so the file
  // that passed the checks is the file that gets read.
  const noFollow = fs.constants.O_NOFOLLOW || 0;
  let fd;
  try {
    fd = fs.openSync(realPath, fs.constants.O_RDONLY | noFollow);
  } catch {
    throw new LocalPathError(`"${requested}" could not be opened`);
  }
  try {
    const st = fs.fstatSync(fd);
    if (!st.isFile()) throw new LocalPathError(`"${requested}" is not a regular file`);
    if (st.size === 0) throw new LocalPathError(`"${requested}" is empty`);
    if (st.size > maxBytes) {
      throw new LocalPathError(
        `"${requested}" is ${st.size} bytes, over the ${maxBytes}-byte (${Math.round(maxBytes / 1048576)} MB) ` +
          'limit for local files; share it from OneDrive or SharePoint instead',
      );
    }
    const bytes = Buffer.alloc(st.size);
    let off = 0;
    while (off < st.size) {
      const n = fs.readSync(fd, bytes, off, st.size - off, off);
      if (n === 0) break;
      off += n;
    }
    if (off !== st.size) throw new LocalPathError(`"${requested}" changed while being read; try again`);
    return { realPath, bytes };
  } finally {
    fs.closeSync(fd);
  }
}

// ── tools/call rewrite ───────────────────────────────────────────────────────

function isPlainObject(v) {
  return typeof v === 'object' && v !== null && !Array.isArray(v);
}

/**
 * Return a copy of a tools/call message with local paths replaced by base64
 * content, or the original message when there is nothing to replace. Throws LocalPathError when a path is refused.
 */
function rewriteToolCall(msg, opts) {
  if (!msg || msg.method !== 'tools/call' || !isPlainObject(msg.params)) return msg;
  const name = msg.params.name;
  const args = msg.params.arguments;
  if (!isPlainObject(args)) return msg;

  if (ATTACHMENT_TOOLS.has(name) && Array.isArray(args.attachments)) {
    if (!args.attachments.some((a) => isPlainObject(a) && a.path !== undefined)) return msg;
    const attachments = args.attachments.map((a, i) => {
      if (!isPlainObject(a) || a.path === undefined) return a;
      if (a.content !== undefined) {
        throw new LocalPathError(`attachments[${i}]: give either content or path, not both`);
      }
      let file;
      try {
        file = readLocalFile(a.path, opts);
      } catch (err) {
        if (err instanceof LocalPathError) throw new LocalPathError(`attachments[${i}].path: ${err.message}`);
        throw err;
      }
      const { path: _omit, ...rest } = a;
      return {
        ...rest,
        name: typeof a.name === 'string' && a.name.trim() ? a.name : path.basename(file.realPath),
        contentType: typeof a.contentType === 'string' && a.contentType ? a.contentType : guessContentType(file.realPath),
        content: file.bytes.toString('base64'),
      };
    });
    return { ...msg, params: { ...msg.params, arguments: { ...args, attachments } } };
  }

  if (name === LOCAL_PATH_TOOL && args.localPath !== undefined) {
    if (args.content !== undefined) {
      throw new LocalPathError('give either content or localPath, not both');
    }
    let file;
    try {
      file = readLocalFile(args.localPath, opts);
    } catch (err) {
      if (err instanceof LocalPathError) throw new LocalPathError(`localPath: ${err.message}`);
      throw err;
    }
    const { localPath: _omit, ...rest } = args;
    return {
      ...msg,
      params: {
        ...msg.params,
        arguments: {
          ...rest,
          content: file.bytes.toString('base64'),
          encoding: 'base64',
          contentType: typeof args.contentType === 'string' && args.contentType ? args.contentType : guessContentType(file.realPath),
        },
      },
    };
  }

  return msg;
}

// ── tools/list rewrite ───────────────────────────────────────────────────────

function withLocalPathSchema(tool, rootsText) {
  const schema = tool && tool.inputSchema;
  if (!isPlainObject(schema) || !isPlainObject(schema.properties)) return tool;

  if (ATTACHMENT_TOOLS.has(tool.name) && isPlainObject(schema.properties.attachments)) {
    const att = schema.properties.attachments;
    const items = isPlainObject(att.items) ? att.items : { type: 'object', properties: {} };
    return {
      ...tool,
      inputSchema: {
        ...schema,
        properties: {
          ...schema.properties,
          attachments: {
            ...att,
            description:
              (att.description ? att.description + ' ' : '') +
              'Each item takes EITHER content (base64) OR path, the absolute path of a file on this machine; ' +
              'prefer path for any file already on disk, so the bytes never pass through the conversation. ' +
              `Files up to 10 MB, under: ${rootsText}. Hidden files are refused.`,
            items: {
              ...items,
              properties: {
                ...(items.properties || {}),
                path: {
                  type: 'string',
                  description: 'Absolute (or ~/) path of a local file to attach, instead of content. name defaults to the file name.',
                },
              },
              // name defaults to the file name when path is used, so nothing is strictly required.
              required: [],
            },
          },
        },
      },
    };
  }

  if (tool.name === LOCAL_PATH_TOOL) {
    return {
      ...tool,
      description:
        (tool.description || '').replace(/\.?\s*$/, '') +
        '. To upload a file from this machine, pass localPath instead of content; the file is read locally and sent as binary.',
      inputSchema: {
        ...schema,
        properties: {
          ...schema.properties,
          localPath: {
            type: 'string',
            description:
              `Absolute (or ~/) path of a local file to upload, instead of content. Up to 10 MB, under: ${rootsText}. ` +
              'Hidden files are refused. Not the OneDrive destination; that is path.',
          },
        },
        required: (Array.isArray(schema.required) ? schema.required : []).filter((r) => r !== 'content'),
      },
    };
  }

  return tool;
}

/** Rewrite a tools/list result so the local-path fields are advertised. */
function rewriteToolsListResult(result, opts) {
  if (!isPlainObject(result) || !Array.isArray(result.tools)) return result;
  const home = (opts && opts.home) || os.homedir();
  const rootsText = describeRoots(resolveRoots(opts), home);
  return { ...result, tools: result.tools.map((t) => withLocalPathSchema(t, rootsText)) };
}

// ── Transport ────────────────────────────────────────────────────────────────

function parseArgs(argv) {
  const out = { url: null, headers: {} };
  for (let i = 0; i < argv.length; i++) {
    const a = argv[i];
    if ((a === '--streamableHttp' || a === '--url') && i + 1 < argv.length) {
      out.url = argv[++i];
    } else if (a === '--header' && i + 1 < argv.length) {
      const h = argv[++i];
      const idx = h.indexOf(':');
      if (idx > 0) out.headers[h.slice(0, idx).trim()] = h.slice(idx + 1).trim();
    } else if (a === '--version') {
      out.version = true;
    }
  }
  return out;
}

/** Pull JSON-RPC messages out of a text/event-stream body. */
function parseSse(text) {
  const messages = [];
  for (const block of text.split(/\r?\n\r?\n/)) {
    const data = block
      .split(/\r?\n/)
      .filter((l) => l.startsWith('data:'))
      .map((l) => l.slice(5).replace(/^ /, ''))
      .join('\n');
    if (!data) continue;
    try {
      messages.push(JSON.parse(data));
    } catch {
      /* ignore non-JSON event */
    }
  }
  return messages;
}

function toolError(id, text) {
  return { jsonrpc: '2.0', id, result: { content: [{ type: 'text', text: `Error: ${text}` }], isError: true } };
}

function rpcError(id, code, message) {
  return { jsonrpc: '2.0', id: id === undefined ? null : id, error: { code, message } };
}

/**
 * Build a forwarder bound to one server. Returns handle(msg) -> array of
 * responses to write (empty for notifications).
 */
function createForwarder(config) {
  const fetchImpl = config.fetch || globalThis.fetch;
  const opts = config.opts || {};
  let sessionId = null;

  async function post(msg) {
    const headers = {
      'Content-Type': 'application/json',
      Accept: 'application/json, text/event-stream',
      ...config.headers,
    };
    if (sessionId) headers['Mcp-Session-Id'] = sessionId;
    const res = await fetchImpl(config.url, { method: 'POST', headers, body: JSON.stringify(msg) });
    const sid = res.headers.get('mcp-session-id');
    if (sid) sessionId = sid;
    const text = await res.text();
    if (res.status === 202 || res.status === 204 || !text.trim()) return { status: res.status, messages: [] };
    const type = res.headers.get('content-type') || '';
    let messages;
    if (type.includes('text/event-stream')) {
      messages = parseSse(text);
    } else {
      try {
        const parsed = JSON.parse(text);
        messages = Array.isArray(parsed) ? parsed : [parsed];
      } catch {
        messages = null;
      }
    }
    return { status: res.status, messages, text };
  }

  async function handle(msg) {
    const isNotification = msg.id === undefined || msg.id === null;
    let outgoing;
    try {
      outgoing = rewriteToolCall(msg, opts);
    } catch (err) {
      if (isNotification) return [];
      if (err instanceof LocalPathError) return [toolError(msg.id, err.message)];
      return [rpcError(msg.id, -32603, 'Shim error: ' + (err && err.message ? err.message : String(err)))];
    }
    try {
      const { status, messages } = await post(outgoing);
      if (isNotification) return [];
      if (messages === null || status < 200 || status >= 300) {
        return [rpcError(msg.id, -32603, `Server returned HTTP ${status}`)];
      }
      if (msg.method === 'tools/list') {
        return messages.map((m) => (m && m.id === msg.id && m.result ? { ...m, result: rewriteToolsListResult(m.result, opts) } : m));
      }
      return messages;
    } catch (err) {
      if (isNotification) return [];
      return [rpcError(msg.id, -32603, 'Shim could not reach the server: ' + (err && err.message ? err.message : String(err)))];
    }
  }

  return { handle };
}

function main() {
  const args = parseArgs(process.argv.slice(2));
  if (args.version) {
    process.stdout.write(SHIM_VERSION + '\n');
    return;
  }
  if (!args.url) {
    process.stderr.write('m365-mcp-shim: --streamableHttp <server>/api/mcp is required\n');
    process.exit(2);
  }
  if (typeof globalThis.fetch !== 'function') {
    process.stderr.write('m365-mcp-shim: Node.js 18 or newer is required\n');
    process.exit(2);
  }
  const forwarder = createForwarder({ url: args.url, headers: args.headers });
  const write = (m) => process.stdout.write(JSON.stringify(m) + '\n');
  const pending = new Set();

  process.stderr.write(`[m365-mcp-shim ${SHIM_VERSION}] forwarding to ${args.url}\n`);
  const rl = readline.createInterface({ input: process.stdin, terminal: false });
  rl.on('line', (line) => {
    if (!line.trim()) return;
    let msg;
    try {
      msg = JSON.parse(line);
    } catch (err) {
      write(rpcError(null, -32700, 'Parse error: ' + err.message));
      return;
    }
    // Messages are handled concurrently so one slow tool call (a large upload)
    // does not hold up the rest; each reply carries its own id.
    const batch = Array.isArray(msg) ? msg : [msg];
    const p = Promise.all(batch.map((m) => forwarder.handle(m)))
      .then((results) => {
        const out = results.flat();
        if (Array.isArray(msg)) {
          if (out.length) write(out);
        } else {
          out.forEach(write);
        }
      })
      .catch((err) => process.stderr.write('[m365-mcp-shim] ' + (err && err.stack ? err.stack : String(err)) + '\n'))
      .finally(() => pending.delete(p));
    pending.add(p);
  });
  // On stdin EOF, finish in-flight calls, then exit once stdout has flushed
  // (pipe writes are asynchronous on macOS and Windows).
  rl.on('close', () => {
    Promise.allSettled([...pending]).then(() => process.stdout.write('', () => process.exit(0)));
  });
}

module.exports = {
  SHIM_VERSION,
  MAX_LOCAL_FILE_BYTES,
  LocalPathError,
  resolveRoots,
  readLocalFile,
  rewriteToolCall,
  rewriteToolsListResult,
  createForwarder,
  parseArgs,
  parseSse,
};

if (require.main === module) main();
