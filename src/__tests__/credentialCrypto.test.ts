/**
 * Tests for credentialCrypto.ts — the HMAC + AES-GCM helpers introduced in
 * to harden credential storage in Azure Table Storage.
 *
 * The crypto module reads its keys from MCP_SESSION_HMAC_KEY and
 * MCP_DATA_ENCRYPTION_KEY env vars at module load time, so we set them
 * before requiring the module under test.
 */

import { randomBytes } from 'crypto';
import {
  hashSessionToken,
  encryptWithDek,
  decryptWithDek,
  decryptWithDekMigrating,
  envelopeAad,
  isCryptoConfigured,
} from '../services/credentialCrypto.js';
import { createCipheriv } from 'crypto';

const TEST_HMAC_KEY = randomBytes(32).toString('hex');
const TEST_DEK = randomBytes(32).toString('hex');

beforeAll(() => {
  process.env.MCP_SESSION_HMAC_KEY = TEST_HMAC_KEY;
  process.env.MCP_DATA_ENCRYPTION_KEY = TEST_DEK;
});

// Local alias for test ergonomics
const crypto = { hashSessionToken, encryptWithDek, decryptWithDek, isCryptoConfigured };

const AAD = envelopeAad('mcpSessions', 'session', 'row-a', 'accessToken');

/** Build an envelope the way encryptWithDek did before row binding: no AAD. */
function encryptUnbound(plaintext: string) {
  const iv = randomBytes(12);
  const cipher = createCipheriv('aes-256-gcm', Buffer.from(TEST_DEK, 'hex'), iv);
  const enc = Buffer.concat([cipher.update(plaintext, 'utf8'), cipher.final()]);
  return {
    ciphertext: enc.toString('base64'),
    iv: iv.toString('base64'),
    authTag: cipher.getAuthTag().toString('base64'),
  };
}

describe('credentialCrypto', () => {
  describe('hashSessionToken', () => {
    test('produces a 64-char hex string (256 bits)', () => {
      const hash = crypto.hashSessionToken('some-token');
      expect(hash).toMatch(/^[a-f0-9]{64}$/);
    });

    test('is deterministic for the same input', () => {
      const a = crypto.hashSessionToken('the-same-token');
      const b = crypto.hashSessionToken('the-same-token');
      expect(a).toBe(b);
    });

    test('produces different hashes for different inputs', () => {
      const a = crypto.hashSessionToken('token-one');
      const b = crypto.hashSessionToken('token-two');
      expect(a).not.toBe(b);
    });

    test('is sensitive to single-character changes', () => {
      const a = crypto.hashSessionToken('abc123def');
      const b = crypto.hashSessionToken('abc123deg'); // last char changed
      expect(a).not.toBe(b);
    });
  });

  describe('encryptWithDek / decryptWithDek round-trip', () => {
    test('round-trips a short ASCII string', () => {
      const plaintext = 'hello world';
      const envelope = crypto.encryptWithDek(plaintext, AAD);
      const recovered = crypto.decryptWithDek(envelope, AAD);
      expect(recovered).toBe(plaintext);
    });

    test('round-trips a long random JSON-like blob', () => {
      const plaintext = JSON.stringify({
        userId: 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee',
        accessToken: randomBytes(800).toString('base64'),
        scopes: ['Mail.Read', 'Files.Read.All', 'Sites.ReadWrite.All'],
        expiresAt: 1772169238000,
      });
      const envelope = crypto.encryptWithDek(plaintext, AAD);
      const recovered = crypto.decryptWithDek(envelope, AAD);
      expect(recovered).toBe(plaintext);
    });

    test('round-trips multibyte unicode', () => {
      const plaintext = 'café — résumé — 日本語 — 🔐';
      const envelope = crypto.encryptWithDek(plaintext, AAD);
      const recovered = crypto.decryptWithDek(envelope, AAD);
      expect(recovered).toBe(plaintext);
    });

    test('produces a different envelope every call (random IV)', () => {
      const plaintext = 'same input each time';
      const e1 = crypto.encryptWithDek(plaintext, AAD);
      const e2 = crypto.encryptWithDek(plaintext, AAD);
      expect(e1.iv).not.toBe(e2.iv);
      expect(e1.ciphertext).not.toBe(e2.ciphertext);
      // Both still decrypt to the same thing
      expect(crypto.decryptWithDek(e1, AAD)).toBe(plaintext);
      expect(crypto.decryptWithDek(e2, AAD)).toBe(plaintext);
    });

    test('returns base64-encoded fields suitable for Azure Table Storage', () => {
      const envelope = crypto.encryptWithDek('storage-friendly', AAD);
      // Base64 alphabet only
      expect(envelope.ciphertext).toMatch(/^[A-Za-z0-9+/]+=*$/);
      expect(envelope.iv).toMatch(/^[A-Za-z0-9+/]+=*$/);
      expect(envelope.authTag).toMatch(/^[A-Za-z0-9+/]+=*$/);
    });
  });

  describe('decryptWithDek tamper detection', () => {
    test('throws if the ciphertext is altered', () => {
      const envelope = crypto.encryptWithDek('plaintext-A', AAD);
      // Flip a byte in the ciphertext
      const tamperedBytes = Buffer.from(envelope.ciphertext, 'base64');
      tamperedBytes[0] ^= 0xff;
      const tampered = {
        ...envelope,
        ciphertext: tamperedBytes.toString('base64'),
      };
      expect(() => crypto.decryptWithDek(tampered, AAD)).toThrow();
    });

    test('throws if the auth tag is altered', () => {
      const envelope = crypto.encryptWithDek('plaintext-B', AAD);
      const tamperedTag = Buffer.from(envelope.authTag, 'base64');
      tamperedTag[0] ^= 0xff;
      const tampered = {
        ...envelope,
        authTag: tamperedTag.toString('base64'),
      };
      expect(() => crypto.decryptWithDek(tampered, AAD)).toThrow();
    });

    test('throws if the IV is altered', () => {
      const envelope = crypto.encryptWithDek('plaintext-C', AAD);
      const tamperedIv = Buffer.from(envelope.iv, 'base64');
      tamperedIv[0] ^= 0xff;
      const tampered = {
        ...envelope,
        iv: tamperedIv.toString('base64'),
      };
      expect(() => crypto.decryptWithDek(tampered, AAD)).toThrow();
    });

    // / F10: reject a truncated GCM auth tag before setAuthTag.
    test('throws if the auth tag is truncated below 16 bytes', () => {
      const envelope = crypto.encryptWithDek('plaintext-D', AAD);
      const fullTag = Buffer.from(envelope.authTag, 'base64');
      expect(fullTag.length).toBe(16);
      const truncated = {
        ...envelope,
        authTag: fullTag.subarray(0, 8).toString('base64'), // 8-byte tag
      };
      expect(() => crypto.decryptWithDek(truncated, AAD)).toThrow(/auth tag length/i);
    });
  });

  // Every envelope is under one DEK, so only AAD stops an unmodified
  // envelope from another row decrypting cleanly.
  describe('AAD row binding', () => {
    const rowB = envelopeAad('mcpSessions', 'session', 'row-b', 'accessToken');

    test('an envelope moved to another row fails to decrypt', () => {
      const envelope = crypto.encryptWithDek('user-A-graph-token', AAD);
      expect(() => crypto.decryptWithDek(envelope, rowB)).toThrow();
    });

    test('an envelope moved to another column or table fails to decrypt', () => {
      const envelope = crypto.encryptWithDek('x', AAD);
      expect(() => crypto.decryptWithDek(
        envelope, envelopeAad('mcpSessions', 'session', 'row-a', 'other'),
      )).toThrow();
      expect(() => crypto.decryptWithDek(
        envelope, envelopeAad('mcpMsalCache', 'session', 'row-a', 'accessToken'),
      )).toThrow();
    });

    test('rejects empty AAD, which GCM treats as no AAD', () => {
      expect(() => crypto.encryptWithDek('x', '')).toThrow(/non-empty/);
      const envelope = crypto.encryptWithDek('x', AAD);
      expect(() => crypto.decryptWithDek(envelope, '')).toThrow(/non-empty/);
      expect(() => envelopeAad('mcpSessions', 'session', '', 'accessToken')).toThrow(/rowKey/);
    });

    test('an unbound legacy envelope does not decrypt through the strict path', () => {
      const unbound = encryptUnbound('legacy');
      expect(() => crypto.decryptWithDek(unbound, AAD)).toThrow();
    });
  });

  describe('decryptWithDekMigrating', () => {
    const rowB = envelopeAad('mcpSessions', 'session', 'row-b', 'accessToken');

    afterEach(() => {
      delete process.env.MCP_ENVELOPE_REQUIRE_AAD;
    });

    test('reads a bound envelope as non-legacy', () => {
      const envelope = crypto.encryptWithDek('bound', AAD);
      expect(decryptWithDekMigrating(envelope, AAD)).toEqual({ plaintext: 'bound', legacy: false });
    });

    test('reads a pre-binding unbound envelope and flags it for rebinding', () => {
      const unbound = encryptUnbound('legacy-token');
      expect(decryptWithDekMigrating(unbound, AAD)).toEqual({ plaintext: 'legacy-token', legacy: true });
    });

    test('never falls back for a bound envelope read from the wrong row', () => {
      const envelope = crypto.encryptWithDek('user-A', AAD);
      expect(() => decryptWithDekMigrating(envelope, rowB)).toThrow();
    });

    test('refuses unbound envelopes when MCP_ENVELOPE_REQUIRE_AAD=true', () => {
      process.env.MCP_ENVELOPE_REQUIRE_AAD = 'true';
      const unbound = encryptUnbound('legacy-token');
      expect(() => decryptWithDekMigrating(unbound, AAD)).toThrow();
    });

    test('still rejects a tampered unbound envelope', () => {
      const unbound = encryptUnbound('legacy-token');
      const tag = Buffer.from(unbound.authTag, 'base64');
      tag[0] ^= 0xff;
      expect(() => decryptWithDekMigrating({ ...unbound, authTag: tag.toString('base64') }, AAD)).toThrow();
    });
  });

  describe('isCryptoConfigured', () => {
    test('returns true when both keys are set', () => {
      expect(crypto.isCryptoConfigured()).toBe(true);
    });
  });
});
