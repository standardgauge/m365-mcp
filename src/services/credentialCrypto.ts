import { createHmac, randomBytes, createCipheriv, createDecipheriv } from 'crypto';

/**
 * Credential crypto helpers for at-rest protection of secrets in Azure Table
 * Storage. Two distinct primitives:
 *
 * 1. **HMAC-SHA256 for session tokens** — opaque random values that the server
 *    only ever needs to *compare* against. We hash with a server-side secret
 *    and store the hash. Stolen storage data is useless without the secret.
 *
 * 2. **AES-256-GCM envelope encryption for access tokens and MSAL cache** —
 *    values the server needs to *use* (i.e. send to Microsoft Graph), so
 *    one-way hashing isn't an option. We encrypt with a server-side data
 *    encryption key (DEK), storing the ciphertext, IV, and auth tag.
 *
 * Both keys (`MCP_SESSION_HMAC_KEY` and `MCP_DATA_ENCRYPTION_KEY`) live in
 * Azure Key Vault and are bound to the Container App as secretRef env vars.
 * Container App managed identity needs `Key Vault Secrets User` on the vault.
 *
 * Loss of either key permanently invalidates the corresponding stored data:
 *  - Lose HMAC key → all sessions become unauthenticatable, every user re-OAuths
 *  - Lose DEK    → all stored access tokens and MSAL refresh tokens are lost
 *                  (next request triggers MSAL re-auth via the user's browser)
 *
 * Both keys MUST be 32 random bytes encoded as hex (64 hex chars). Generate via:
 *   openssl rand -hex 32
 */

// Keys are read from env vars on each call (rather than cached at module
// load) so that tests can set them in beforeAll without import-order races.
// The cost is one Buffer.from per crypto operation, which is negligible.

function getHmacKey(): Buffer {
  const hex = process.env.MCP_SESSION_HMAC_KEY ?? '';
  if (!hex) {
    throw new Error(
      'MCP_SESSION_HMAC_KEY env var is not set. Server cannot persist sessions safely. ' +
      'Bind a 64-hex-char key from Key Vault before starting the container.'
    );
  }
  const bytes = Buffer.from(hex, 'hex');
  if (bytes.length !== 32) {
    throw new Error(
      `MCP_SESSION_HMAC_KEY must decode to exactly 32 bytes (64 hex chars), got ${bytes.length}`
    );
  }
  return bytes;
}

function getDek(): Buffer {
  const hex = process.env.MCP_DATA_ENCRYPTION_KEY ?? '';
  if (!hex) {
    throw new Error(
      'MCP_DATA_ENCRYPTION_KEY env var is not set. Server cannot persist credentials safely. ' +
      'Bind a 64-hex-char key from Key Vault before starting the container.'
    );
  }
  const bytes = Buffer.from(hex, 'hex');
  if (bytes.length !== 32) {
    throw new Error(
      `MCP_DATA_ENCRYPTION_KEY must decode to exactly 32 bytes (64 hex chars), got ${bytes.length}`
    );
  }
  return bytes;
}

export function isCryptoConfigured(): boolean {
  try {
    getHmacKey();
    getDek();
    return true;
  } catch {
    return false;
  }
}

export function requireCryptoConfigured(): void {
  getHmacKey();
  getDek();
}

// ── HMAC for session token lookup ──────────────────────────────────────────────

/**
 * Compute the HMAC-SHA256 hash of a session token. The hash is the row-level
 * lookup key in Azure Table Storage. Storing only the hash means that an
 * attacker who exfiltrates the table cannot use the values to authenticate.
 *
 * Returns a 64-char hex string suitable for use as a Table Storage row key.
 */
export function hashSessionToken(token: string): string {
  return createHmac('sha256', getHmacKey()).update(token, 'utf8').digest('hex');
}

// ── AES-256-GCM envelope encryption for sensitive blobs ───────────────────────

export interface EnvelopeCiphertext {
  /** AES-GCM ciphertext (base64) */
  ciphertext: string;
  /** 96-bit IV (base64) — unique per encryption call */
  iv: string;
  /** 128-bit GCM auth tag (base64) */
  authTag: string;
}

/**
 * Encrypt an arbitrary plaintext string under the data encryption key.
 * Returns the three pieces of state needed to decrypt: ciphertext, IV, tag.
 *
 * IV is freshly random per call (12 bytes per NIST SP 800-38D recommendation
 * for AES-GCM). Auth tag is 16 bytes (default).
 *
 * Output is base64-encoded for clean storage in Azure Table Storage string
 * properties. Three columns per encrypted field is the cost — alternatives
 * (single concatenated string, JSON envelope) trade clarity for compactness
 * and aren't worth it at our scale.
 */
export function encryptWithDek(plaintext: string): EnvelopeCiphertext {
  const dek = getDek();
  const iv = randomBytes(12);
  const cipher = createCipheriv('aes-256-gcm', dek, iv);
  const enc = Buffer.concat([cipher.update(plaintext, 'utf8'), cipher.final()]);
  const authTag = cipher.getAuthTag();
  return {
    ciphertext: enc.toString('base64'),
    iv: iv.toString('base64'),
    authTag: authTag.toString('base64'),
  };
}

/**
 * Decrypt an EnvelopeCiphertext back to its original plaintext. Throws if
 * the auth tag doesn't validate, which guards against tampering with the
 * stored ciphertext (e.g. swapping in a value from a different user).
 */
export function decryptWithDek(envelope: EnvelopeCiphertext): string {
  const dek = getDek();
  const iv = Buffer.from(envelope.iv, 'base64');
  const ciphertext = Buffer.from(envelope.ciphertext, 'base64');
  const authTag = Buffer.from(envelope.authTag, 'base64');

  // Reject a truncated auth tag before it reaches setAuthTag. Node accepts GCM
  // tags as short as 4 bytes, and a shorter tag weakens forgery resistance
  // (a truncation attack needs storage-write access, so this is defensive —
  // / F10). We only ever emit 16-byte tags in encryptWithDek.
  if (authTag.length !== 16) {
    throw new Error(
      `Invalid GCM auth tag length: expected 16 bytes, got ${authTag.length}`
    );
  }

  // authTagLength pins the expected tag size so the GCM verification cannot be
  // silently downgraded to a shorter tag.
  const decipher = createDecipheriv('aes-256-gcm', dek, iv, { authTagLength: 16 });
  decipher.setAuthTag(authTag);
  const dec = Buffer.concat([decipher.update(ciphertext), decipher.final()]);
  return dec.toString('utf8');
}
