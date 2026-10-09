/**
 * Desktop extension self-update: verify a signed payload, then write it inside
 * the extension directory and nowhere else.
 *
 * This file is inlined verbatim into the generated server/index.js (see
 * renderServerJs in src/functions/install/installEndpoint.ts), wrapped in a
 * function so its requires do not collide with the entry point's. It lives
 * here as a real module so the tests exercise the code that ships rather than
 * a copy of it.
 *
 * The instance signs every update with an Ed25519 key and bakes the matching
 * public key into the code it serves. An update is applied only when:
 *   - the signature over the exact payload bytes verifies against that baked-in
 *     key, so a response that did not come from the key holder is refused;
 *   - the version inside the signed payload is strictly newer than the running
 *     one, so an old signed payload cannot be replayed to roll a client back;
 *   - every file path in it resolves inside the extension directory, checked
 *     for the whole set before anything is written.
 */

const crypto = require('crypto');
const fs = require('fs');
const path = require('path');

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

/**
 * Check the update response's `signed` block and return the payload it
 * carries. `signed.payload` is a JSON string and the signature covers its
 * UTF-8 bytes as sent, so nothing has to be re-serialised to verify it.
 * Throws on anything short of a valid, newer, signed payload.
 */
function verifySignedUpdate(signed, publicKeyB64, localVersion) {
  if (!signed || typeof signed.payload !== 'string' || typeof signed.signature !== 'string') {
    throw new Error('update is not signed');
  }
  const key = crypto.createPublicKey({
    key: Buffer.from(publicKeyB64, 'base64'),
    format: 'der',
    type: 'spki',
  });
  const ok = crypto.verify(
    null,
    Buffer.from(signed.payload, 'utf8'),
    key,
    Buffer.from(signed.signature, 'base64'),
  );
  if (!ok) throw new Error('update signature does not verify');

  const payload = JSON.parse(signed.payload);
  if (!payload || typeof payload.version !== 'string' || !isNewerVersion(payload.version, localVersion)) {
    throw new Error('signed update is not newer than ' + localVersion);
  }
  if (!payload.files || typeof payload.files !== 'object') {
    throw new Error('signed update carries no files');
  }
  return payload;
}

/**
 * Resolve one payload path against the extension directory, refusing anything
 * that would land outside it: absolute paths, drive or UNC prefixes, `..`
 * segments, and the directory itself.
 */
function resolveUpdatePath(extDir, filePath) {
  if (typeof filePath !== 'string' || filePath === '' || filePath.includes('\0')) {
    throw new Error('invalid update path');
  }
  if (path.isAbsolute(filePath) || path.win32.isAbsolute(filePath) || /^[a-zA-Z]:/.test(filePath)) {
    throw new Error('update path is absolute: ' + filePath);
  }
  const root = path.resolve(extDir);
  const full = path.resolve(root, filePath);
  const rel = path.relative(root, full);
  if (rel === '' || rel === '..' || rel.startsWith('..' + path.sep) || path.isAbsolute(rel)) {
    throw new Error('update path escapes the extension directory: ' + filePath);
  }
  return full;
}

/** True when `child` is `root` or sits beneath it, after resolving symlinks. */
function isWithin(root, child) {
  const rel = path.relative(root, child);
  return rel === '' || (!rel.startsWith('..' + path.sep) && rel !== '..' && !path.isAbsolute(rel));
}

/**
 * Write every file in `files` under `extDir`. All paths are validated before
 * the first write, so a payload with one bad path changes nothing. Each target
 * directory is re-checked after symlink resolution, so a symlinked directory
 * inside the extension cannot redirect a write outside it. Each file is written
 * to a temp name and renamed into place.
 */
function applyUpdate(extDir, files) {
  const root = path.resolve(extDir);
  const writes = Object.entries(files).map(([filePath, content]) => {
    if (typeof content !== 'string') throw new Error('update content is not text: ' + filePath);
    return { fullPath: resolveUpdatePath(root, filePath), content };
  });

  const realRoot = fs.realpathSync(root);
  for (const { fullPath, content } of writes) {
    const dir = path.dirname(fullPath);
    // Check the deepest directory that already exists before creating anything
    // below it, so mkdir cannot follow a link out either.
    let existing = dir;
    while (!fs.existsSync(existing)) existing = path.dirname(existing);
    if (!isWithin(realRoot, fs.realpathSync(existing))) {
      throw new Error('update path escapes the extension directory via a link: ' + fullPath);
    }
    if (existing !== dir) fs.mkdirSync(dir, { recursive: true });
    const tmpPath = fullPath + '.tmp';
    fs.writeFileSync(tmpPath, content, 'utf8');
    fs.renameSync(tmpPath, fullPath);
  }
}

module.exports = { isNewerVersion, verifySignedUpdate, resolveUpdatePath, applyUpdate };
