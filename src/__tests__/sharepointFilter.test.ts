import {
  isSystemSite,
  siteDisambiguator,
  filterAndDisambiguateSites,
} from '../services/sharepointFilter.js';

describe('sharepointFilter', () => {
  describe('isSystemSite', () => {
    test.each([
      ['https://examplecom.sharepoint.com/sites/contentTypeHub', true],
      ['https://examplecom.sharepoint.com/sites/ContentTypeHub', true], // case-insensitive
      ['https://examplecom.sharepoint.com/sites/appcatalog', true],
      ['https://examplecom.sharepoint.com/sites/CompliancePolicyCenter', true],
      ['https://examplecom.sharepoint.com/personal/nate_example_com', true],
      ['https://examplecom.sharepoint.com/personal/someone_example_com', true],
      ['https://examplecom.sharepoint.com/sites/SFPs', false],
      ['https://examplecom.sharepoint.com/sites/Marketing', false],
      ['https://examplecom.sharepoint.com', false],
      ['https://examplecom.sharepoint.com/teams/eng', false],
    ])('%s -> isSystemSite=%s', (url, expected) => {
      expect(isSystemSite({ webUrl: url })).toBe(expected);
    });

    test('returns false on malformed URLs (does not throw)', () => {
      expect(isSystemSite({ webUrl: 'not a url' })).toBe(false);
      expect(isSystemSite({ webUrl: '' })).toBe(false);
    });
  });

  describe('siteDisambiguator', () => {
    test('returns the trimmed pathname', () => {
      expect(siteDisambiguator('https://x.sharepoint.com/sites/marketing')).toBe('sites/marketing');
      expect(siteDisambiguator('https://x.sharepoint.com/teams/eng/sub')).toBe('teams/eng/sub');
    });

    test('returns "(root)" for the root site', () => {
      expect(siteDisambiguator('https://x.sharepoint.com')).toBe('(root)');
      expect(siteDisambiguator('https://x.sharepoint.com/')).toBe('(root)');
    });

    test('returns the original URL on parse failure', () => {
      expect(siteDisambiguator('not a url')).toBe('not a url');
    });
  });

  describe('filterAndDisambiguateSites', () => {
    test('drops system sites', () => {
      const input = [
        { displayName: 'SFPs', webUrl: 'https://x.sharepoint.com/sites/contentTypeHub' },
        { displayName: 'SFPs', webUrl: 'https://x.sharepoint.com/sites/SFPs' },
        { displayName: 'My OneDrive', webUrl: 'https://x.sharepoint.com/personal/me_x_com' },
        { displayName: 'Marketing', webUrl: 'https://x.sharepoint.com/sites/Marketing' },
      ];
      const out = filterAndDisambiguateSites(input);
      expect(out).toHaveLength(2);
      expect(out.map((s) => s.displayName)).toEqual(['SFPs', 'Marketing']);
      expect(out[0].webUrl).toBe('https://x.sharepoint.com/sites/SFPs');
    });

    test('disambiguates remaining displayName collisions', () => {
      const input = [
        { displayName: 'Marketing', webUrl: 'https://x.sharepoint.com/sites/marketing' },
        { displayName: 'Marketing', webUrl: 'https://x.sharepoint.com/sites/marketing-archive' },
        { displayName: 'Sales', webUrl: 'https://x.sharepoint.com/sites/sales' },
      ];
      const out = filterAndDisambiguateSites(input);
      expect(out).toHaveLength(3);
      const names = out.map((s) => s.displayName).sort();
      expect(names).toEqual([
        'Marketing (sites/marketing)',
        'Marketing (sites/marketing-archive)',
        'Sales',
      ]);
    });

    test('does not mutate the input array', () => {
      const input = [
        { displayName: 'A', webUrl: 'https://x.sharepoint.com/sites/a' },
        { displayName: 'A', webUrl: 'https://x.sharepoint.com/sites/a2' },
      ];
      const snapshot = JSON.parse(JSON.stringify(input));
      filterAndDisambiguateSites(input);
      expect(input).toEqual(snapshot);
    });

    test('preserves extra fields on the site object', () => {
      const input = [
        {
          displayName: 'SFPs',
          webUrl: 'https://x.sharepoint.com/sites/sfps',
          id: 'site-id-1',
          createdDateTime: '2025-01-01',
        },
      ];
      const out = filterAndDisambiguateSites(input);
      expect(out[0]).toMatchObject({
        id: 'site-id-1',
        createdDateTime: '2025-01-01',
      });
    });

    test('handles empty input', () => {
      expect(filterAndDisambiguateSites([])).toEqual([]);
    });
  });
});
