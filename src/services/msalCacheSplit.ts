/**
 * Splits a serialized MSAL token cache into one cache per account.
 *
 * The server used to keep every user's MSAL cache in a single Table row. That
 * row is now split so each account has its own (tableStorage.ts, "MSAL cache
 * persistence"). This does the splitting: every Account, IdToken, AccessToken
 * and RefreshToken entry carries the `home_account_id` it belongs to and goes
 * to that account's cache. AppMetadata entries describe the client, not a
 * user, so each account's cache gets a copy.
 *
 * Entries with no `home_account_id`, and sections this does not recognise, are
 * dropped: nothing in them can be attributed to an account, and MSAL rebuilds
 * anything it needs on the next sign-in.
 */

const ACCOUNT_SECTIONS = ['Account', 'IdToken', 'AccessToken', 'RefreshToken'] as const;

type Section = Record<string, { home_account_id?: unknown }>;

/**
 * Returns one serialized cache per home account id. Throws if the input is not
 * a JSON object.
 */
export function splitMsalCacheByAccount(serialized: string): Map<string, string> {
  const cache = JSON.parse(serialized) as Record<string, unknown>;
  if (!cache || typeof cache !== 'object' || Array.isArray(cache)) {
    throw new Error('MSAL cache is not a JSON object');
  }

  const appMetadata = isSection(cache.AppMetadata) ? cache.AppMetadata : {};
  const byAccount = new Map<string, Record<string, Section>>();

  for (const section of ACCOUNT_SECTIONS) {
    const entries = cache[section];
    if (!isSection(entries)) continue;
    for (const [key, entry] of Object.entries(entries)) {
      const homeAccountId = entry?.home_account_id;
      if (typeof homeAccountId !== 'string' || !homeAccountId) continue;
      let partition = byAccount.get(homeAccountId);
      if (!partition) {
        partition = {
          Account: {},
          IdToken: {},
          AccessToken: {},
          RefreshToken: {},
          AppMetadata: { ...appMetadata },
        };
        byAccount.set(homeAccountId, partition);
      }
      partition[section][key] = entry;
    }
  }

  const out = new Map<string, string>();
  for (const [homeAccountId, partition] of byAccount) {
    out.set(homeAccountId, JSON.stringify(partition));
  }
  return out;
}

function isSection(value: unknown): value is Section {
  return !!value && typeof value === 'object' && !Array.isArray(value);
}
