/**
 * Unit tests for encodeGraphId.
 *
 * Graph event/message IDs are base64 strings that routinely contain `+` and
 * `=` (and, for standard base64, `/`). The Graph SDK forwards path segments
 * verbatim, so a raw `+` is decoded to a space by Graph and the ID no longer
 * resolves — "The Id is invalid". encodeGraphId percent-encodes the value so it
 * round-trips, and neutralizes path/OData-injection payloads as a side effect.
 */

import { encodeGraphId, ValidationError } from '../services/opaqueId.js';

describe('encodeGraphId', () => {
  it('percent-encodes a base64 ID containing + and =', () => {
    const id = 'AAMkAGI2gz6IS4+aCkhUgDjL9rVEUR==';
    const enc = encodeGraphId(id, 'eventId');
    expect(enc).toBe('AAMkAGI2gz6IS4%2BaCkhUgDjL9rVEUR%3D%3D');
    // The raw characters that Graph mis-decodes must be gone.
    expect(enc).not.toContain('+');
    expect(enc.includes('=')).toBe(false);
    // Decoding returns the original ID Graph stored.
    expect(decodeURIComponent(enc)).toBe(id);
  });

  it('percent-encodes standard-base64 / so it cannot split the path', () => {
    const enc = encodeGraphId('abc/def+ghi', 'eventId');
    expect(enc).toBe('abc%2Fdef%2Bghi');
    expect(enc).not.toContain('/');
  });

  it('leaves base64url-safe characters (- _) untouched', () => {
    const id = 'AAMk-Guid_007';
    expect(encodeGraphId(id, 'eventId')).toBe(id);
  });

  it('neutralizes a path-traversal payload by encoding the slashes', () => {
    const enc = encodeGraphId('legit/../../../users/victim', 'eventId');
    expect(enc).not.toContain('/');
    expect(decodeURIComponent(enc)).toBe('legit/../../../users/victim');
  });

  it('throws a ValidationError on an empty string', () => {
    expect(() => encodeGraphId('', 'eventId')).toThrow(ValidationError);
  });

  it('throws a ValidationError on a non-string value', () => {
    expect(() => encodeGraphId(undefined as unknown as string, 'eventId')).toThrow(ValidationError);
  });
});
