import { createPrivateKey, createPublicKey, sign, KeyObject } from 'crypto';
import { deriveFromHmacKey } from './credentialCrypto.js';

/**
 * Ed25519 signing for desktop extension self-updates.
 *
 * The instance serves extension code to every user's machine, and the
 * extension overwrites itself with whatever /api/extension-update returns. The
 * served code carries this instance's public key, and from 2.11.0 on the
 * extension applies only payloads that verify against it (see
 * src/install/extension-update.js). A response that did not come from a holder
 * of the private key — a proxy, a cache, anything that can shape the response
 * without reading the instance's secrets — cannot push code.
 *
 * The private key is derived from MCP_SESSION_HMAC_KEY (HKDF, purpose below)
 * rather than read from a new secret, so every deployment signs as soon as it
 * runs this release with nothing to configure. It stays stable across replicas
 * and restarts for the same reason. Rotating the HMAC key rotates this key, and
 * installed extensions then refuse updates until reinstalled from /install;
 * docs/operations-runbook.md → "Application keys" says so.
 */

const SIGNING_KEY_PURPOSE = 'm365-mcp/v1/extension-update-signing';

// PKCS#8 DER header for a raw 32-byte Ed25519 seed (RFC 8410).
const ED25519_PKCS8_PREFIX = Buffer.from('302e020100300506032b657004220420', 'hex');

function signingKey(): KeyObject {
  const seed = deriveFromHmacKey(SIGNING_KEY_PURPOSE);
  return createPrivateKey({
    key: Buffer.concat([ED25519_PKCS8_PREFIX, seed]),
    format: 'der',
    type: 'pkcs8',
  });
}

/** The public key the served extension verifies against, as base64 SPKI DER. */
export function extensionPublicKey(): string {
  return createPublicKey(signingKey()).export({ format: 'der', type: 'spki' }).toString('base64');
}

export interface SignedPayload {
  /** JSON text of the payload; the signature covers its UTF-8 bytes. */
  payload: string;
  /** Ed25519 signature, base64. */
  signature: string;
}

/** Serialise `payload` and sign the exact bytes the client will verify. */
export function signExtensionPayload(payload: unknown): SignedPayload {
  const text = JSON.stringify(payload);
  const signature = sign(null, Buffer.from(text, 'utf8'), signingKey()).toString('base64');
  return { payload: text, signature };
}
