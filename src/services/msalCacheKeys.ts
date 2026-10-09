/**
 * Re-keys credentials in a serialized MSAL token cache to the key format the
 * installed @azure/msal-node writes.
 *
 * MSAL finds a credential by matching the fields stored in it, but it writes a
 * renewed credential under a key it computes. msal-node 2.x built that key from
 * eight segments (…-realm-target-requestedClaimsHash-tokenType); msal-node 7
 * builds it from seven (…-realm-target-scheme). msal-node never
 * migrates the old keys (msal-common's updateCredentialCacheKey has no caller
 * in the node package), so a cache written before the upgrade ends up holding
 * two refresh tokens for the same account after its first refresh: the
 * original under the old key and the rotated one under the new key. Lookup
 * returns the first match, which is the original, so every later refresh keeps
 * presenting the token issued at sign-in until it expires, and the user is sent
 * back to sign in although a newer refresh token is sitting in the cache.
 *
 * Moving each credential to the key the current library would compute makes
 * the renewed token overwrite the old one, as it does for a cache the current
 * library wrote. Where an entry already sits at the current key, that entry was
 * written by the current library and is kept; the legacy duplicate is dropped.
 *
 * This mirrors generateCredentialKey in @azure/msal-common 16. The test in
 * msalAuthFlows.test.ts that runs a cache written by the installed msal-node
 * through this function and expects it back unchanged is what catches a future
 * change to that format: if it fails after an MSAL upgrade, update this to
 * match rather than deleting the test.
 */

type SerializedCredential = {
  home_account_id?: string;
  environment?: string;
  credential_type?: string;
  client_id?: string;
  family_id?: string;
  realm?: string;
  target?: string;
  token_type?: string;
  additionalCacheKeyComponents?: Record<string, string>;
};

const CREDENTIAL_SECTIONS = new Set(['IdToken', 'AccessToken', 'RefreshToken']);

export function currentCredentialKey(credential: SerializedCredential): string {
  const familyId =
    (credential.credential_type === 'RefreshToken' && credential.family_id) || credential.client_id;
  const scheme =
    credential.token_type && credential.token_type.toLowerCase() !== 'bearer'
      ? credential.token_type.toLowerCase()
      : '';
  return [
    credential.home_account_id,
    credential.environment,
    credential.credential_type,
    familyId,
    credential.realm || '',
    credential.target || '',
    scheme,
  ]
    .join('-')
    .toLowerCase();
}

/**
 * Returns the serialized cache with every credential under its current key, or
 * the input unchanged when nothing needed moving (or it is not a JSON object).
 */
export function migrateLegacyCredentialKeys(serialized: string): string {
  let cache: Record<string, unknown>;
  try {
    cache = JSON.parse(serialized) as Record<string, unknown>;
  } catch {
    return serialized;
  }
  if (!cache || typeof cache !== 'object') return serialized;

  let changed = false;
  for (const [section, value] of Object.entries(cache)) {
    if (!CREDENTIAL_SECTIONS.has(section) || !value || typeof value !== 'object') continue;
    const entries = value as Record<string, SerializedCredential>;

    for (const [key, credential] of Object.entries(entries)) {
      if (!credential || typeof credential !== 'object') continue;
      // Keys carrying an extra hash segment are a feature this server does not
      // use; leave them to MSAL rather than reproduce the hash.
      if (
        credential.additionalCacheKeyComponents &&
        Object.keys(credential.additionalCacheKeyComponents).length > 0
      ) {
        continue;
      }
      const target = currentCredentialKey(credential);
      if (target === key) continue;

      if (!Object.hasOwn(entries, target)) {
        // eslint-disable-next-line security/detect-object-injection -- keys come from the cache blob MSAL wrote, not from a request
        entries[target] = credential;
      }
      // eslint-disable-next-line security/detect-object-injection -- as above
      delete entries[key];
      changed = true;
    }
  }

  return changed ? JSON.stringify(cache) : serialized;
}
