/**
 * Unit tests for the SharePoint search access-control helpers.
 *
 * These cover the pure building blocks that both search surfaces (the MCP
 * `search_sharepoint` tool and the HTTP `/api/sharepoint/search` route) share:
 * the over-fetch window sizing, site-relative path resolution from a `webUrl`,
 * and the allow-list post-filter. The end-to-end wiring of both handlers is
 * covered in sharepointSearchAccessControl.test.ts.
 */

import {
  searchFetchSize,
  SEARCH_MAX_FETCH,
  siteRelativePathFromWebUrl,
  filterHitsToAllowedSites,
  type AllowedSite,
} from '../services/sharepointSearch.js';

const ROOT: AllowedSite = { id: 'contoso.sharepoint.com,11111111-1111-1111-1111-111111111111,aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa', name: 'Root' };
const ITHUB: AllowedSite = { id: 'contoso.sharepoint.com,22222222-2222-2222-2222-222222222222,bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb', name: 'IT Hub' };

describe('searchFetchSize ( over-fetch)', () => {
  it('requests exactly maxResults when no filter is active (nothing to drop)', () => {
    expect(searchFetchSize(25, false)).toBe(25);
    expect(searchFetchSize(50, false)).toBe(50);
  });

  it('over-fetches to a floor when a filter is active so post-filtering still fills the page', () => {
    // maxResults below the floor is raised so allowed hits are not crowded out.
    expect(searchFetchSize(25, true)).toBe(50);
    expect(searchFetchSize(10, true)).toBe(50);
  });

  it('over-fetches proportionally but never above the per-request cap', () => {
    expect(searchFetchSize(80, true)).toBe(80);
    expect(searchFetchSize(1000, true)).toBe(SEARCH_MAX_FETCH);
  });
});

describe('siteRelativePathFromWebUrl', () => {
  it('strips a /sites/<name> prefix and preserves the document library segment', () => {
    expect(
      siteRelativePathFromWebUrl('https://contoso.sharepoint.com/sites/AricPublic/Shared%20Documents/Finance/deal.xlsx'),
    ).toBe('/Shared Documents/Finance/deal.xlsx');
  });

  it('strips a /teams/<name> prefix', () => {
    expect(
      siteRelativePathFromWebUrl('https://contoso.sharepoint.com/teams/Eng/Shared%20Documents/spec.md'),
    ).toBe('/Shared Documents/spec.md');
  });

  it('strips a /personal/<name> prefix for OneDrive hits', () => {
    expect(
      siteRelativePathFromWebUrl('https://contoso-my.sharepoint.com/personal/nate_contoso_com/Documents/private.docx'),
    ).toBe('/Documents/private.docx');
  });

  it('leaves a root-site path (no site-collection prefix) intact', () => {
    expect(
      siteRelativePathFromWebUrl('https://contoso.sharepoint.com/Shared%20Documents/IR/q3.pptx'),
    ).toBe('/Shared Documents/IR/q3.pptx');
  });

  it('URL-decodes so %20 matches a plain-text deny entry', () => {
    expect(siteRelativePathFromWebUrl('https://c.sharepoint.com/sites/x/A%20B/c')).toBe('/A B/c');
  });

  it('returns null for an absent or unparseable webUrl (fail closed)', () => {
    expect(siteRelativePathFromWebUrl(undefined)).toBeNull();
    expect(siteRelativePathFromWebUrl('')).toBeNull();
    expect(siteRelativePathFromWebUrl('not a url')).toBeNull();
  });
});

describe('filterHitsToAllowedSites', () => {
  const hit = (siteId: string | null | undefined, name: string) => ({ siteId, name });

  it('allows everything when the allow-list is empty', () => {
    const items = [hit('any', 'a'), hit(undefined, 'b')];
    expect(filterHitsToAllowedSites(items, [])).toEqual(items);
  });

  it('keeps hits whose siteId exactly matches an allowed id', () => {
    const items = [hit(ROOT.id, 'root'), hit(ITHUB.id, 'ithub')];
    expect(filterHitsToAllowedSites(items, [ROOT]).map((i) => i.name)).toEqual(['root']);
  });

  it('drops hits from a site outside the allow-list (the AricPublic / OneDrive case)', () => {
    const foreign = 'contoso.sharepoint.com,99999999-9999-9999-9999-999999999999,cccccccc-cccc-cccc-cccc-cccccccccccc';
    const items = [hit(ROOT.id, 'allowed'), hit(foreign, 'AricPublic'), hit(null, 'onedrive')];
    expect(filterHitsToAllowedSites(items, [ROOT, ITHUB]).map((i) => i.name)).toEqual(['allowed']);
  });

  it('matches a full composite id that differs only in host casing/shape but shares both GUIDs', () => {
    // Same site-collection GUID AND same web GUID as ITHUB, different host case.
    const sameWeb = 'CONTOSO.sharepoint.com,22222222-2222-2222-2222-222222222222,bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb';
    const items = [hit(sameWeb, 'ithub-drive')];
    expect(filterHitsToAllowedSites(items, [ITHUB]).map((i) => i.name)).toEqual(['ithub-drive']);
  });

  it('drops a sibling subsite sharing the collection GUID but with a different web GUID (codex CR,)', () => {
    // Same site collection as ITHUB (22222222…), but a DIFFERENT web GUID —
    // a distinct, non-allow-listed subsite. Matching on the collection GUID
    // alone would leak its metadata; web identity must differ, so it is dropped.
    const siblingSubsite = 'contoso.sharepoint.com,22222222-2222-2222-2222-222222222222,eeeeeeee-eeee-eeee-eeee-eeeeeeeeeeee';
    const items = [hit(ITHUB.id, 'ithub'), hit(siblingSubsite, 'ithub-secret-subsite')];
    expect(filterHitsToAllowedSites(items, [ITHUB]).map((i) => i.name)).toEqual(['ithub']);
  });

  it('drops a bare collection GUID with no web GUID (cannot pin a web — fail closed)', () => {
    const bareCollectionGuid = '22222222-2222-2222-2222-222222222222';
    const items = [hit(bareCollectionGuid, 'ithub-drive')];
    expect(filterHitsToAllowedSites(items, [ITHUB])).toEqual([]);
  });

  it('drops a missing siteId when the allow-list is active (fail closed)', () => {
    expect(filterHitsToAllowedSites([hit(undefined, 'x'), hit(null, 'y')], [ROOT])).toEqual([]);
  });

  it('drops a hostname-only siteId with no GUID (fail closed)', () => {
    expect(filterHitsToAllowedSites([hit('contoso.sharepoint.com', 'x')], [ROOT])).toEqual([]);
  });
});
