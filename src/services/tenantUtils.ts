/**
 * Extract tenant ID from MSAL homeAccountId.
 * Format: "{oid}.{tid}" — the part after the first dot is the tenant ID.
 */
export function extractTenantId(homeAccountId: string): string {
  const dotIndex = homeAccountId.indexOf('.');
  if (dotIndex === -1) throw new Error('Invalid homeAccountId format');
  return homeAccountId.substring(dotIndex + 1);
}
