import { createHmac, randomBytes, createCipheriv, createDecipheriv, hkdfSync } from 'crypto';

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
 *    encryption key (DEK), storing the ciphertext, IV, and auth tag. Each
 *    envelope is bound to the row it is stored in via GCM additional
 *    authenticated data (see envelopeAad), so it cannot be moved to another
 *    user's row and still decrypt.
 *
 * Both keys (`MCP_SESSION_HMAC_KEY` and `MCP_DATA_ENCRYPTION_KEY`) reach the
 * process as env vars bound to Container App secrets (secretRef). The infra/
 * templates create those secrets as Key Vault references resolved by the app's
 * runtime identity (infra/key-vault.bicep), so the values live in the vault and
 * not in the Container App. An instance configured by hand may still hold them
 * as plain secrets; see "Application keys" in docs/operations-runbook.md for
 * moving them, rotation and revocation. Both are
 * validated before the Functions host starts (src/startup/checkKeys.ts), so a
 * missing or malformed key stops the container instead of the first sign-in.
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

export const CRYPTO_KEY_ENV_VARS = ['MCP_SESSION_HMAC_KEY', 'MCP_DATA_ENCRYPTION_KEY'] as const;
export type CryptoKeyEnvVar = (typeof CRYPTO_KEY_ENV_VARS)[number];

const KEY_PURPOSE: Record<CryptoKeyEnvVar, string> = {
  MCP_SESSION_HMAC_KEY: 'Server cannot persist sessions safely.',
  MCP_DATA_ENCRYPTION_KEY: 'Server cannot persist credentials safely.',
};

/**
 * Decode one key, or say why it is unusable. Surrounding whitespace is
 * ignored, since a value pasted into Key Vault often carries a trailing
 * newline. Anything else must be exactly 64 hex characters.
 *
 * Buffer.from(hex, 'hex') is not a validator on its own: it stops at the first
 * non-hex character and drops a trailing odd nibble, so a 65-character value or
 * a valid key with junk appended decodes to 32 bytes without complaint. The
 * regex is what rejects those.
 */
function parseKey(name: CryptoKeyEnvVar, raw: string | undefined): { key: Buffer } | { error: string } {
  const hex = (raw ?? '').trim();
  if (!hex) {
    return {
      error:
        `${name} env var is not set. ${KEY_PURPOSE[name]} ` +
        'Bind a 64-hex-char key from Key Vault before starting the container.',
    };
  }
  if (!/^[0-9a-fA-F]+$/.test(hex)) {
    return { error: `${name} must be hex-encoded (64 hex chars); it contains non-hex characters` };
  }
  if (hex.length !== 64) {
    return {
      error: `${name} must decode to exactly 32 bytes (64 hex chars), got ${hex.length} hex chars`,
    };
  }
  return { key: Buffer.from(hex, 'hex') };
}

function getKey(name: CryptoKeyEnvVar): Buffer {
  const parsed = parseKey(name, process.env[name]);
  if ('error' in parsed) throw new Error(parsed.error);
  return parsed.key;
}

function getHmacKey(): Buffer {
  return getKey('MCP_SESSION_HMAC_KEY');
}

function getDek(): Buffer {
  return getKey('MCP_DATA_ENCRYPTION_KEY');
}

/**
 * Every problem with the two keys in `env`, one message per key. Empty means
 * both are usable. Reports both keys at once so an operator fixing a fresh
 * deployment is not sent round the loop twice. Never includes key material.
 */
export function cryptoKeyProblems(env: NodeJS.ProcessEnv = process.env): string[] {
  const problems: string[] = [];
  for (const name of CRYPTO_KEY_ENV_VARS) {
    const parsed = parseKey(name, env[name]);
    if ('error' in parsed) problems.push(parsed.error);
  }
  return problems;
}

export function isCryptoConfigured(env: NodeJS.ProcessEnv = process.env): boolean {
  return cryptoKeyProblems(env).length === 0;
}

/** Throws one error naming every unusable key. Called at host startup. */
export function requireCryptoConfigured(env: NodeJS.ProcessEnv = process.env): void {
  const problems = cryptoKeyProblems(env);
  if (problems.length) throw new Error(problems.join('\n'));
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

/**
 * HMAC-SHA256 under the session key, domain-separated by `purpose` so a MAC
 * minted for one use can never stand in for another, or for a session-token
 * hash. Session tokens are hex, so they can never contain the newline that
 * separates the purpose from the data.
 */
export function macWithSessionKey(purpose: string, data: string): string {
  return createHmac('sha256', getHmacKey()).update(`${purpose}\n${data}`, 'utf8').digest('hex');
}

/**
 * A 32-byte key for another purpose, derived from the HMAC key with
 * HKDF-SHA256 and `purpose` as the info string. Distinct purposes give
 * independent keys, and none of them reveals the HMAC key, so a new purpose
 * needs no new secret on any deployment. The cost is shared rotation: rotating
 * MCP_SESSION_HMAC_KEY rotates every derived key with it.
 */
export function deriveFromHmacKey(purpose: string): Buffer {
  if (!purpose) throw new Error('deriveFromHmacKey: purpose must be non-empty');
  return Buffer.from(hkdfSync('sha256', getHmacKey(), Buffer.alloc(0), purpose, 32));
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
 * Additional authenticated data (AAD) for an envelope stored at a given Table
 * Storage location. Every envelope is under the one DEK, so without AAD the
 * GCM tag only proves "this was encrypted by us" — not "this was encrypted for
 * this row". Binding the table, partition, row and column into the tag means
 * an envelope copied from user A's row into user B's row fails to decrypt
 * instead of handing B's session A's Graph token.
 *
 * The table and column are included alongside the row key so an envelope
 * cannot be moved between tables or columns either.
 */
export function envelopeAad(
  table: string,
  partitionKey: string,
  rowKey: string,
  column: string,
): string {
  for (const [name, value] of Object.entries({ table, partitionKey, rowKey, column })) {
    if (!value) throw new Error(`envelopeAad: ${name} must be non-empty`);
  }
  return `m365-mcp/v1/${table}/${partitionKey}/${rowKey}/${column}`;
}

/**
 * Extend a location AAD with plaintext columns that sit beside the envelope on
 * the same row. Those columns are then integrity-protected by the envelope's
 * GCM tag: editing any of them makes the envelope fail to decrypt, so the row
 * fails authentication as a whole.
 *
 * The values are JSON-encoded in the order given, so no value can smuggle a
 * separator and make two different column sets produce the same AAD.
 */
export function boundColumnsAad(
  locationAad: string,
  columns: ReadonlyArray<readonly [name: string, value: string]>,
): string {
  if (!columns.length) throw new Error('boundColumnsAad: at least one column is required');
  return `${locationAad}#bound=${JSON.stringify(columns)}`;
}

function aadBytes(aad: string): Buffer {
  // Empty AAD is cryptographically identical to no AAD in GCM, so an empty
  // value would silently produce an unbound envelope.
  if (!aad) throw new Error('AES-GCM additional authenticated data must be non-empty');
  return Buffer.from(aad, 'utf8');
}

/**
 * Encrypt an arbitrary plaintext string under the data encryption key, bound
 * to `aad` (see envelopeAad). Returns the three pieces of state needed to
 * decrypt: ciphertext, IV, tag. The AAD is not stored; the reader recomputes
 * it from where it found the envelope.
 *
 * IV is freshly random per call (12 bytes per NIST SP 800-38D recommendation
 * for AES-GCM). Auth tag is 16 bytes (default).
 *
 * Output is base64-encoded for clean storage in Azure Table Storage string
 * properties. Three columns per encrypted field is the cost — alternatives
 * (single concatenated string, JSON envelope) trade clarity for compactness
 * and aren't worth it at our scale.
 */
export function encryptWithDek(plaintext: string, aad: string): EnvelopeCiphertext {
  const dek = getDek();
  const iv = randomBytes(12);
  const cipher = createCipheriv('aes-256-gcm', dek, iv);
  cipher.setAAD(aadBytes(aad));
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
 * the auth tag doesn't validate against the ciphertext, IV and `aad`. That
 * detects tampering with the stored bytes, and — because `aad` names the row
 * the envelope was read from — an envelope moved in from a different row.
 *
 * The tag alone does NOT detect a swap: every envelope is under the same DEK,
 * so an unmodified envelope from another user's row carries a perfectly valid
 * tag. Only the AAD binding catches that, which is why `aad` is required.
 */
export function decryptWithDek(envelope: EnvelopeCiphertext, aad: string): string {
  return decryptEnvelope(envelope, aadBytes(aad));
}

/**
 * Whether envelopes written before AAD binding may still be read.
 * On by default so a deploy does not log every user out; set
 * MCP_ENVELOPE_REQUIRE_AAD=true once existing rows have been rebound (they are
 * rewritten with AAD on first read, see tableStorage.ts). While legacy reads
 * are allowed, an attacker with storage write can still move an *unbound*
 * envelope between rows, so the window should be short.
 */
export function isLegacyEnvelopeReadAllowed(): boolean {
  return (process.env.MCP_ENVELOPE_REQUIRE_AAD ?? '').toLowerCase() !== 'true';
}

/**
 * Whether a session access-token envelope bound only to its row (before the
 * row's identity columns were added to the AAD) may still be read.
 * On by default so a deploy does not log every user out; set
 * MCP_SESSION_REQUIRE_IDENTITY_BINDING=true once existing rows have been
 * rebound (they are rewritten on first read, see tableStorage.ts). While these
 * reads are allowed, a storage writer can still edit the identity columns of a
 * row that has not been rebound, or of a row whose older envelope they kept.
 *
 * Requiring identity binding also refuses unbound envelopes, whatever
 * MCP_ENVELOPE_REQUIRE_AAD says: an unbound envelope binds nothing either.
 */
export function isLegacyIdentityBindingAllowed(): boolean {
  return (process.env.MCP_SESSION_REQUIRE_IDENTITY_BINDING ?? '').toLowerCase() !== 'true';
}

export interface DecryptResult {
  plaintext: string;
  /** True when the envelope predates AAD binding and should be rewritten. */
  legacy: boolean;
}

/**
 * Migration-aware decrypt. Tries the AAD-bound form first; if that fails and
 * legacy reads are allowed, retries as an unbound pre-binding envelope. A
 * bound envelope never decrypts without its AAD, so the fallback cannot be
 * used to read a bound envelope out of the wrong row.
 */
export function decryptWithDekMigrating(
  envelope: EnvelopeCiphertext,
  aad: string,
): DecryptResult {
  try {
    return { plaintext: decryptWithDek(envelope, aad), legacy: false };
  } catch (err) {
    if (!isLegacyEnvelopeReadAllowed()) throw err;
    try {
      return { plaintext: decryptEnvelope(envelope, null), legacy: true };
    } catch {
      throw err;
    }
  }
}

function decryptEnvelope(envelope: EnvelopeCiphertext, aad: Buffer | null): string {
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
  if (aad) decipher.setAAD(aad);
  decipher.setAuthTag(authTag);
  const dec = Buffer.concat([decipher.update(ciphertext), decipher.final()]);
  return dec.toString('utf8');
}
