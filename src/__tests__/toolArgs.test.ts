/**
 * Unit tests for the tool-argument validator.
 */
import { findUnsupportedArgs, unsupportedArgsMessage } from '../services/toolArgs.js';

const SCHEMA = { properties: { q: {}, maxResults: {}, folderId: {} } };

describe('findUnsupportedArgs', () => {
  it('returns [] when every arg is declared', () => {
    expect(findUnsupportedArgs({ q: 'x', maxResults: 10 }, SCHEMA)).toEqual([]);
  });

  it('names an undeclared arg', () => {
    expect(findUnsupportedArgs({ q: 'x', startDate: '2026-01-01' }, SCHEMA)).toEqual(['startDate']);
  });

  it('names every undeclared arg, preserving order', () => {
    expect(findUnsupportedArgs({ participants: [], before: '', q: 'x' }, SCHEMA)).toEqual(['participants', 'before']);
  });

  it('treats a no-property schema as accepting nothing', () => {
    expect(findUnsupportedArgs({ anything: 1 }, { properties: {} })).toEqual(['anything']);
    expect(findUnsupportedArgs({ anything: 1 }, undefined)).toEqual(['anything']);
  });

  it('tolerates non-object args without throwing', () => {
    expect(findUnsupportedArgs(undefined, SCHEMA)).toEqual([]);
    expect(findUnsupportedArgs(null, SCHEMA)).toEqual([]);
    expect(findUnsupportedArgs([1, 2], SCHEMA)).toEqual([]);
  });

  it('accepts an empty arg object', () => {
    expect(findUnsupportedArgs({}, SCHEMA)).toEqual([]);
  });
});

describe('findUnsupportedArgs — nested object/array schemas', () => {
  // Mirrors the real shapes: a declared object with sub-properties (contact
  // addresses), and a declared array whose items are structured objects
  // (attachments, contacts batch).
  const ADDRESS = {
    type: 'object',
    properties: { street: {}, city: {}, state: {}, postalCode: {}, countryOrRegion: {} },
  };
  const CONTACT_ITEM = {
    type: 'object',
    properties: { givenName: {}, surname: {}, homeAddress: ADDRESS },
  };
  const SCHEMA = {
    properties: {
      homeAddress: ADDRESS,
      attachments: { type: 'array', items: { type: 'object', properties: { name: {}, contentType: {}, content: {} } } },
      contacts: { type: 'array', items: CONTACT_ITEM },
      // A string[] must NOT be recursed into (no item properties to check).
      categories: { type: 'array', items: { type: 'string' } },
    },
  };

  it('accepts recognized sub-fields of a declared object', () => {
    expect(findUnsupportedArgs({ homeAddress: { street: '1', city: 'SF' } }, SCHEMA)).toEqual([]);
  });

  it('names an undeclared sub-field of a declared object by path', () => {
    expect(findUnsupportedArgs({ homeAddress: { street: '1', building: 'HQ' } }, SCHEMA)).toEqual([
      'homeAddress.building',
    ]);
  });

  it('names an undeclared field inside an array item by indexed path', () => {
    expect(
      findUnsupportedArgs({ attachments: [{ name: 'a', content: 'x' }, { name: 'b', content: 'y', retentionLabel: 'z' }] }, SCHEMA),
    ).toEqual(['attachments[1].retentionLabel']);
  });

  it('names undeclared fields inside batch contact items, including nested addresses', () => {
    expect(
      findUnsupportedArgs(
        { contacts: [{ givenName: 'A', assistantName: 'B' }, { givenName: 'C', homeAddress: { street: '1', building: 'HQ' } }] },
        SCHEMA,
      ),
    ).toEqual(['contacts[0].assistantName', 'contacts[1].homeAddress.building']);
  });

  it('does not recurse into a declared string[] (no item properties to check)', () => {
    expect(findUnsupportedArgs({ categories: ['a', 'b'] }, SCHEMA)).toEqual([]);
  });

  it('tolerates a scalar where an object/array is declared', () => {
    // Wrong-type values are the handlers' concern; this check must not throw.
    expect(findUnsupportedArgs({ homeAddress: 'nope', attachments: 'nope' }, SCHEMA)).toEqual([]);
  });
});

describe('unsupportedArgsMessage', () => {
  it('names the offending params and lists what is accepted', () => {
    const msg = unsupportedArgsMessage('search_mail', ['participants'], SCHEMA);
    expect(msg).toContain('search_mail');
    expect(msg).toContain('participants');
    expect(msg).toContain('q, maxResults, folderId');
    expect(msg).toContain('rejects parameters it does not implement');
  });

  it('says so when the tool takes no parameters', () => {
    expect(unsupportedArgsMessage('list_sites', ['x'], { properties: {} })).toContain('takes no parameters');
  });
});
