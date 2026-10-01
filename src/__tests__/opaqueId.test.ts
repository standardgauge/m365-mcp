/**
 * Unit tests for opaqueId validation helpers.
 *
 * assertOpaqueId() must reject any string that could manipulate a Microsoft
 * Graph API URL path, while accepting the full range of characters that appear
 * in legitimate opaque Graph resource IDs (base64url strings, GUIDs, UPNs,
 * SharePoint composite IDs).
 */

import { assertOpaqueId, assertOpaqueIds } from '../services/opaqueId.js';

// ── assertOpaqueId ────────────────────────────────────────────────────────────

describe('assertOpaqueId — valid IDs pass', () => {
  const valid = [
    // Graph message / folder / item IDs (base64url-encoded)
    'AAMkAGE1Mj6J1aD5YcNhp7A',
    'AQMkADYAAAIBDQAAAA==',
    // GUIDs
    '12345678-1234-1234-1234-123456789abc',
    'b1638132-5b9c-4d48-ac40-18ad9e0b534e',
    // SharePoint composite IDs  (hostname,guid,guid — no slashes)
    'contoso.sharepoint.com,62ec43af-c1e7-46e2-986b-30c50e27c89b,d3ce3a3a-bca3-45b2-a3b0-38e4ff948b79',
    // SharePoint b!-encoded site IDs
    'b!tW8dHucDikyOVRAkXfKHx1Qsv_Qr8kxPp6VY2BpMsCk',
    // User Principal Names (mailboxId)
    'user@example.com',
    'john.doe+alias@contoso.onmicrosoft.com',
    // The 'me' self-reference token
    'me',
    // Drive IDs with leading-zero patterns
    '01BWKFZZAH3GYXNQMQIQ',
    // Opaque IDs with exclamation marks
    'cHJveHkh',
    // Base64 padding characters
    'abc=',
    'abc==',
  ];

  test.each(valid)('accepts %s', (id) => {
    expect(() => assertOpaqueId(id, 'testParam')).not.toThrow();
  });
});

describe('assertOpaqueId — injection vectors are rejected', () => {
  const invalid: Array<[string, string]> = [
    // Path traversal via forward slash
    ['allowed-site/../../../users/victim', 'path traversal with /'],
    ['site-id/drives/drive-id', 'embedded path segment'],
    ['me/messages/AAMkAGE1', 'mailboxId + path override'],
    // OData query injection via ?
    ['site-id?$select=secret', 'OData select injection'],
    ['site-id?$top=1000', 'OData top injection'],
    // Fragment injection
    ['site-id#evil', 'fragment injection'],
    // Percent-encoding bypass
    ['site%2Fid', 'URL-encoded slash'],
    ['id%3Finjection', 'URL-encoded question mark'],
    ['id%23hash', 'URL-encoded hash'],
    ['id%00null', 'URL-encoded null byte'],
    // Backslash
    ['site\\id', 'backslash'],
    // Control characters
    ['id\rInjected-Header: value', 'carriage return (header injection)'],
    ['id\nInjected', 'newline (header injection)'],
    // Null byte
    ['id\x00null', 'null byte'],
  ];

  test.each(invalid)('rejects %s (%s)', (id) => {
    expect(() => assertOpaqueId(id, 'testParam')).toThrow(
      'testParam contains characters that are not permitted in a resource identifier',
    );
  });
});

describe('assertOpaqueId — non-string / empty values', () => {
  it('throws for empty string', () => {
    expect(() => assertOpaqueId('', 'paramName')).toThrow('paramName must be a non-empty string');
  });

  it('includes the paramName in the error message', () => {
    expect(() => assertOpaqueId('bad/value', 'siteId')).toThrow(
      'siteId contains characters',
    );
  });
});

// ── assertOpaqueIds ───────────────────────────────────────────────────────────

describe('assertOpaqueIds — batch validation', () => {
  const idParams: ReadonlySet<string> = new Set([
    'siteId', 'driveId', 'messageId', 'mailboxId',
  ]);

  it('passes when all ID params are clean', () => {
    const args = {
      siteId: 'site-abc',
      driveId: 'drive-xyz',
      messageId: 'AAMkAGE1Mj6J1aD5',
      mailboxId: 'me',
      q: 'search query with /slash and ?question — not validated as an ID',
    };
    expect(() => assertOpaqueIds(args, idParams)).not.toThrow();
  });

  it('throws on the first bad ID param found', () => {
    const args = {
      siteId: 'site-abc',
      driveId: 'drive-xyz/../traversal',
    };
    expect(() => assertOpaqueIds(args, idParams)).toThrow('driveId contains characters');
  });

  it('ignores params not in the idParams set', () => {
    const args = {
      q: 'evil?injection&payload=/bad',  // 'q' is not in idParams
      siteId: 'site-abc',
    };
    expect(() => assertOpaqueIds(args, idParams)).not.toThrow();
  });

  it('skips null and undefined values', () => {
    const args = { siteId: null, driveId: undefined };
    expect(() => assertOpaqueIds(args as unknown as Record<string, unknown>, idParams)).not.toThrow();
  });

  it('skips empty-string values', () => {
    const args = { siteId: '' };
    expect(() => assertOpaqueIds(args, idParams)).not.toThrow();
  });

  it('skips non-string values (numbers, booleans)', () => {
    const args = { siteId: 42, driveId: true };
    expect(() => assertOpaqueIds(args as unknown as Record<string, unknown>, idParams)).not.toThrow();
  });
});
