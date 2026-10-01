/**
 * Unit tests for the two-tier deny list logic.
 *
 * Azure Table Storage is mocked so these tests run without any cloud
 * infrastructure.  The focus is on the path-matching rules:
 *   - Exact path match
 *   - Prefix (ancestor folder) match blocks all descendants
 *   - Tier-1 (global admin) blocks override per-user lists — a user cannot
 *     "unblock" a path that the admin has blocked
 *   - Tier-2 (per-user) blocks are additive
 */

import { jest } from '@jest/globals';

// ── Shared mock state ──────────────────────────────────────────────────────────

const globalEntries: Array<{ partitionKey: string; rowKey: string; path: string; addedBy: string; addedAt: string; description: string }> = [];
const userEntries: Array<{ partitionKey: string; rowKey: string; path: string; addedBy: string; addedAt: string; description: string }> = [];

// Mock @azure/data-tables before importing denyList so the module sees the mock
jest.mock('@azure/data-tables', () => {
  const makeAsyncIterable = (rows: unknown[]) => ({
    [Symbol.asyncIterator]: async function* () {
      for (const row of rows) yield row;
    },
  });

  return {
    TableClient: {
      fromConnectionString: jest.fn((_conn: string, tableName: string) => ({
        createTable: jest.fn<() => Promise<void>>().mockResolvedValue(undefined),
        upsertEntity: jest.fn<() => Promise<void>>().mockResolvedValue(undefined),
        deleteEntity: jest.fn<() => Promise<void>>().mockResolvedValue(undefined),
        listEntities: jest.fn(({ queryOptions }: { queryOptions: { filter: string } }) => {
          // Route to the right in-memory list based on table name captured in closure
          const source = tableName === 'GlobalDenyList' ? globalEntries : userEntries;
          // Filter by PartitionKey — mirror odata`PartitionKey eq ${type}` behaviour
          const pkMatch = queryOptions.filter.match(/PartitionKey eq '([^']+)'/);
          const pk = pkMatch ? pkMatch[1] : '';
          return makeAsyncIterable(source.filter((e) => e.partitionKey === pk));
        }),
      })),
    },
    odata: (strings: TemplateStringsArray, ...values: unknown[]) =>
      strings.reduce((acc, str, i) => `${acc}${str}${i < values.length ? `'${values[i]}'` : ''}`, ''),
  };
});

// Set required env var before importing the module
process.env.AZURE_STORAGE_CONNECTION_STRING = 'UseDevelopmentStorage=true';

import { isPathDenied, filterDeniedPaths, canonicalizePath, getDefaultDenyPaths } from '../services/denyList.js';

// ── Helpers ────────────────────────────────────────────────────────────────────

function resetEntries() {
  globalEntries.length = 0;
  userEntries.length = 0;
}

function pushGlobal(type: string, path: string) {
  const entry = { partitionKey: `${TENANT}:${type}`, rowKey: Buffer.from(path).toString('base64'), path, addedBy: 'admin', addedAt: new Date().toISOString(), description: '' };
  globalEntries.push(entry);
}

function pushUser(userId: string, type: string, path: string) {
  const entry = { partitionKey: `${userId}:${type}`, rowKey: Buffer.from(path).toString('base64'), path, addedBy: userId, addedAt: new Date().toISOString(), description: '' };
  userEntries.push(entry);
}

// ── Tests ──────────────────────────────────────────────────────────────────────

const TENANT = 'test-tenant-id';
const USER = 'user-abc';
const TYPE = 'sharepoint' as const;

beforeEach(resetEntries);

describe('isPathDenied — exact match', () => {
  it('returns false when both deny lists are empty', async () => {
    expect(await isPathDenied(TENANT, USER, TYPE, '/sites/HR/documents')).toBe(false);
  });

  it('returns true when path exactly matches a global deny entry', async () => {
    pushGlobal(TYPE, '/sites/HR/documents/Payroll');
    expect(await isPathDenied(TENANT, USER, TYPE, '/sites/HR/documents/Payroll')).toBe(true);
  });

  it('returns true when path exactly matches a user deny entry', async () => {
    pushUser(USER, TYPE, '/sites/Finance/Private');
    expect(await isPathDenied(TENANT, USER, TYPE, '/sites/Finance/Private')).toBe(true);
  });
});

describe('isPathDenied — prefix (ancestor) blocking', () => {
  it('blocks a child path when parent is in the global deny list', async () => {
    pushGlobal(TYPE, '/sites/HR/documents/Payroll');
    expect(await isPathDenied(TENANT, USER, TYPE, '/sites/HR/documents/Payroll/Q1')).toBe(true);
    expect(await isPathDenied(TENANT, USER, TYPE, '/sites/HR/documents/Payroll/Q1/salaries.xlsx')).toBe(true);
  });

  it('does NOT block a sibling path', async () => {
    pushGlobal(TYPE, '/sites/HR/documents/Payroll');
    expect(await isPathDenied(TENANT, USER, TYPE, '/sites/HR/documents/Benefits')).toBe(false);
  });

  it('does NOT block a path that merely starts with the same characters (no slash boundary)', async () => {
    pushGlobal(TYPE, '/sites/HR/documents/Pay');
    // /sites/HR/documents/Payroll shares the prefix string but NOT the folder
    expect(await isPathDenied(TENANT, USER, TYPE, '/sites/HR/documents/Payroll')).toBe(false);
  });

  it('blocks children when user deny entry uses a trailing slash', async () => {
    pushUser(USER, TYPE, '/sites/Finance/Private/');
    expect(await isPathDenied(TENANT, USER, TYPE, '/sites/Finance/Private/Budget.xlsx')).toBe(true);
  });
});

describe('Tier-1 blocks override tier-2 — user cannot unblock admin-denied paths', () => {
  it('path stays blocked even when it is absent from user deny list', async () => {
    // Admin blocked the folder; user has no personal deny list entry
    pushGlobal(TYPE, '/sites/Legal/Contracts');
    expect(await isPathDenied(TENANT, USER, TYPE, '/sites/Legal/Contracts')).toBe(true);
  });

  it('path is blocked by global list regardless of what the user list contains', async () => {
    // Simulate an attempt to "unblock" by having a different (unrelated) entry
    pushGlobal(TYPE, '/sites/Legal/Contracts');
    pushUser(USER, TYPE, '/sites/Finance/Private'); // different path — unrelated
    // Global block still applies
    expect(await isPathDenied(TENANT, USER, TYPE, '/sites/Legal/Contracts')).toBe(true);
    // User-blocked path is also blocked
    expect(await isPathDenied(TENANT, USER, TYPE, '/sites/Finance/Private')).toBe(true);
  });

  it('allows an unrelated path that is in neither list', async () => {
    pushGlobal(TYPE, '/sites/Legal/Contracts');
    pushUser(USER, TYPE, '/sites/Finance/Private');
    expect(await isPathDenied(TENANT, USER, TYPE, '/sites/Engineering/Docs')).toBe(false);
  });
});

describe('filterDeniedPaths', () => {
  it('removes items whose path matches the global deny list', async () => {
    pushGlobal(TYPE, '/sites/HR/Payroll');
    const items = [
      { id: '1', path: '/sites/HR/Payroll', webUrl: '' },
      { id: '2', path: '/sites/HR/Benefits', webUrl: '' },
    ];
    const allowed = await filterDeniedPaths(TENANT, USER, TYPE, items);
    expect(allowed).toHaveLength(1);
    expect(allowed[0].id).toBe('2');
  });

  it('removes items whose webUrl is a denied path when path is absent', async () => {
    pushUser(USER, TYPE, '/sites/Finance/Private');
    const items = [
      { id: '3', webUrl: '/sites/Finance/Private' },
      { id: '4', webUrl: '/sites/Finance/Public' },
    ];
    const allowed = await filterDeniedPaths(TENANT, USER, TYPE, items);
    expect(allowed).toHaveLength(1);
    expect(allowed[0].id).toBe('4');
  });

  it('returns all items when deny lists are empty', async () => {
    const items = [{ id: '1', path: '/sites/A' }, { id: '2', path: '/sites/B' }];
    const allowed = await filterDeniedPaths(TENANT, USER, TYPE, items);
    expect(allowed).toHaveLength(2);
  });

  it('removes items whose display name is denied even when the ID is not', async () => {
    pushGlobal(TYPE, 'HR Sensitive');
    const items = [
      { id: 'cal-guid-1', name: 'HR Sensitive' },
      { id: 'cal-guid-2', name: 'My Calendar' },
    ];
    const allowed = await filterDeniedPaths(TENANT, USER, TYPE, items);
    expect(allowed.map((c) => c.id)).toEqual(['cal-guid-2']);
  });
});

describe('canonicalizePath', () => {
  it('strips Graph drive prefix from parentReference.path', () => {
    expect(canonicalizePath('/drives/b!abc123/root:/Documents/SubFolder')).toBe('/Documents/SubFolder');
  });

  it('handles root-only Graph path (/root: with no trailing slash or segment)', () => {
    expect(canonicalizePath('/drives/b!abc123/root:')).toBe('/');
  });

  it('adds a leading slash to bare drive-relative paths', () => {
    expect(canonicalizePath('Documents/notes.txt')).toBe('/Documents/notes.txt');
  });

  it('leaves slash-prefixed drive-relative paths unchanged', () => {
    expect(canonicalizePath('/Documents/notes.txt')).toBe('/Documents/notes.txt');
  });

  it('leaves mail folder display names unchanged', () => {
    expect(canonicalizePath('Inbox')).toBe('/Inbox');
    expect(canonicalizePath('Sent Items')).toBe('/Sent Items');
  });

  // / F14: canonically-equivalent Unicode must normalize to one form.
  it('normalizes decomposed Unicode to NFC (combining accent === precomposed)', () => {
    const precomposed = '/R\u00e9serv\u00e9'; // e-acute as single U+00E9
    const decomposed = '/Re\u0301serve\u0301'; // e + combining acute U+0301
    expect(precomposed).not.toBe(decomposed); // distinct code-point sequences
    expect(canonicalizePath(decomposed)).toBe(canonicalizePath(precomposed));
    expect(canonicalizePath(decomposed)).toBe(precomposed); // both land on NFC
  });
});

describe('path-format mismatch — Graph vs drive-relative ( regression)', () => {
  it('blocks write_sharepoint_file when deny entry was stored in Graph parentReference.path format', async () => {
    // Admin sees Graph-format path from list_folders / read_file and stores it directly
    pushGlobal(TYPE, '/drives/b!siteCollectionId/root:/Documents/Restricted');
    // write_sharepoint_file checks args.path (drive-relative)
    expect(await isPathDenied(TENANT, USER, TYPE, 'Documents/Restricted/secret.txt')).toBe(true);
    expect(await isPathDenied(TENANT, USER, TYPE, '/Documents/Restricted/secret.txt')).toBe(true);
  });

  it('blocks write_onedrive_file when deny entry was stored in Graph parentReference.path format', async () => {
    pushGlobal('onedrive' as const, '/drives/b!driveId/root:/Confidential');
    expect(await isPathDenied(TENANT, USER, 'onedrive', 'Confidential/report.docx')).toBe(true);
    expect(await isPathDenied(TENANT, USER, 'onedrive', '/Confidential/report.docx')).toBe(true);
  });

  it('does NOT block a write to a sibling folder when only a Graph-format deny entry exists', async () => {
    pushGlobal(TYPE, '/drives/b!siteCollectionId/root:/Documents/Restricted');
    expect(await isPathDenied(TENANT, USER, TYPE, 'Documents/Public/report.txt')).toBe(false);
  });

  it('blocks read_file when deny entry was stored in drive-relative format', async () => {
    // Deny list entry stored as a plain drive-relative path
    pushGlobal(TYPE, '/Documents/Restricted');
    // read_file / delete_sharepoint_file construct filePath from Graph parentReference.path
    const graphPath = '/drives/b!siteCollectionId/root:/Documents/Restricted/secret.txt';
    expect(await isPathDenied(TENANT, USER, TYPE, graphPath)).toBe(true);
  });

  it('filterDeniedPaths removes items with Graph-format paths when entry is drive-relative', async () => {
    pushGlobal(TYPE, '/Documents/HR');
    const items = [
      { id: '1', path: '/drives/b!abc/root:/Documents/HR/salaries.xlsx' },
      { id: '2', path: '/drives/b!abc/root:/Documents/Finance/budget.xlsx' },
    ];
    const allowed = await filterDeniedPaths(TENANT, USER, TYPE, items);
    expect(allowed).toHaveLength(1);
    expect(allowed[0].id).toBe('2');
  });
});

describe('mail folder deny list', () => {
  const MAIL_TYPE = 'mail' as const;

  it('blocks a mail folder by exact display name', async () => {
    pushGlobal(MAIL_TYPE, 'Inbox');
    expect(await isPathDenied(TENANT, USER, MAIL_TYPE, 'Inbox')).toBe(true);
    expect(await isPathDenied(TENANT, USER, MAIL_TYPE, 'Sent Items')).toBe(false);
  });

  it('user-level mail block is independent of sharepoint block', async () => {
    pushGlobal(TYPE, '/sites/HR/Payroll');
    pushUser(USER, MAIL_TYPE, 'Drafts');
    expect(await isPathDenied(TENANT, USER, MAIL_TYPE, 'Drafts')).toBe(true);
    expect(await isPathDenied(TENANT, USER, TYPE, 'Drafts')).toBe(false); // not a SP path
  });
});

// ── Config-driven default deny list ──────────────────────────────
describe('default deny list (env-configured)', () => {
  const MAIL_TYPE = 'mail' as const;

  afterEach(() => {
    delete process.env.DEFAULT_MAIL_DENY_FOLDERS;
    delete process.env.DEFAULT_SHAREPOINT_DENY_PATHS;
  });

  describe('getDefaultDenyPaths', () => {
    it('returns [] when the env var is unset', () => {
      expect(getDefaultDenyPaths(MAIL_TYPE)).toEqual([]);
    });

    it('returns [] for a type with no configured env var', () => {
      process.env.DEFAULT_MAIL_DENY_FOLDERS = 'Finance';
      expect(getDefaultDenyPaths('onedrive')).toEqual([]);
    });

    it('parses, trims, and drops blank entries', () => {
      process.env.DEFAULT_MAIL_DENY_FOLDERS = ' Finance , HR ,, Legal , ';
      expect(getDefaultDenyPaths(MAIL_TYPE)).toEqual(['Finance', 'HR', 'Legal']);
    });
  });

  it('enforces default mail folders even when both tables are empty', async () => {
    process.env.DEFAULT_MAIL_DENY_FOLDERS = 'Finance,HR,Legal,IR,Management';
    for (const folder of ['Finance', 'HR', 'Legal', 'IR', 'Management']) {
      expect(await isPathDenied(TENANT, USER, MAIL_TYPE, folder)).toBe(true);
    }
    expect(await isPathDenied(TENANT, USER, MAIL_TYPE, 'Inbox')).toBe(false);
  });

  it('matches default mail folders case-insensitively', async () => {
    process.env.DEFAULT_MAIL_DENY_FOLDERS = 'Finance,HR';
    expect(await isPathDenied(TENANT, USER, MAIL_TYPE, 'finance')).toBe(true);
    expect(await isPathDenied(TENANT, USER, MAIL_TYPE, 'hr')).toBe(true);
  });

  it('blocks subfolders of a default-denied folder (prefix match)', async () => {
    process.env.DEFAULT_MAIL_DENY_FOLDERS = 'Finance';
    expect(await isPathDenied(TENANT, USER, MAIL_TYPE, 'Finance/2026')).toBe(true);
  });

  it('filterDeniedPaths strips default-denied folders from a listing', async () => {
    process.env.DEFAULT_MAIL_DENY_FOLDERS = 'Finance,HR,Legal,IR,Management';
    const folders = [
      { id: 'f1', path: 'Inbox' },
      { id: 'f2', path: 'Finance' },
      { id: 'f3', path: 'Sent Items' },
      { id: 'f4', path: 'HR' },
    ];
    const allowed = await filterDeniedPaths(TENANT, USER, MAIL_TYPE, folders);
    expect(allowed.map((f) => f.path)).toEqual(['Inbox', 'Sent Items']);
  });

  it('defaults stack on top of admin table entries', async () => {
    process.env.DEFAULT_MAIL_DENY_FOLDERS = 'Finance';
    pushGlobal(MAIL_TYPE, 'Board Reports');
    expect(await isPathDenied(TENANT, USER, MAIL_TYPE, 'Finance')).toBe(true); // default
    expect(await isPathDenied(TENANT, USER, MAIL_TYPE, 'Board Reports')).toBe(true); // table
    expect(await isPathDenied(TENANT, USER, MAIL_TYPE, 'Inbox')).toBe(false);
  });

  it('does not leak mail defaults into other deny-list types', async () => {
    process.env.DEFAULT_MAIL_DENY_FOLDERS = 'Finance';
    // 'Finance' as a SharePoint path is not denied — mail defaults are type-scoped
    expect(await isPathDenied(TENANT, USER, TYPE, 'Finance')).toBe(false);
  });
});
