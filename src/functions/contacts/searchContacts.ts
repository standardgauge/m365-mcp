import { app, HttpRequest, HttpResponseInit, InvocationContext } from '@azure/functions';
import { getValidAccessTokenForSession, getTenantId } from '../../services/tokenCache.js';
import { createGraphClient } from '../../services/graphClient.js';
import { isPathDenied } from '../../services/denyList.js';
import { withPolicyEnforcement } from '../../services/policyEnforcement.js';
import { normalizeContactFolderId } from '../../services/containerResolver.js';
import { assertOpaqueId } from '../../services/opaqueId.js';
import type { AuthResult } from '../../services/authMiddleware.js';
import { withSecurity } from '../../services/securityHeaders.js';
import { CONTACT_SELECT_FIELDS, contactsApiPath, shapeContact } from '../../services/contactFields.js';

/** Escape single quotes for safe OData $filter interpolation. */
function sanitizeODataValue(val: string): string {
  return val.replace(/'/g, "''");
}

async function searchContactsHandler(
  request: HttpRequest,
  context: InvocationContext,
  auth: AuthResult,
): Promise<HttpResponseInit> {
  const userId = auth.userId;

  const q = request.query.get('q');
  const folderId = request.query.get('folderId');
  const maxResults = parseInt(request.query.get('maxResults') ?? '25', 10);

  if (folderId) assertOpaqueId(folderId, 'folderId');

  const accessToken = await getValidAccessTokenForSession(auth.session);
  const graph = createGraphClient(accessToken);

  // Route through the shared helper so the synthetic 'contacts-root' id
  // (returned by list_contact_folders for the default folder) maps back to
  // /me/contacts instead of /me/contactFolders/contacts-root/contacts.
  const apiPath = contactsApiPath(folderId ?? undefined);

  let query = graph
    .api(apiPath)
    .select(CONTACT_SELECT_FIELDS)
    .top(Math.min(maxResults, 50));

  if (q) {
    const safeQ = sanitizeODataValue(q);
    query = query.filter(`startswith(displayName,'${safeQ}') or startswith(givenName,'${safeQ}') or startswith(surname,'${safeQ}')`);
  }

  const result = await query.get();

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const allContacts = (result.value ?? []).map((c: any) => shapeContact(c));

  // Post-filter: when no folderId was specified, check each contact's parentFolderId.
  // Normalize raw Graph parentFolderIds to the synthetic 'contacts-root' key so
  // deny-list comparisons match what the admin UI stores.
  let contacts = allContacts;
  if (!folderId) {
    const tenantId = await getTenantId(userId);
    const filtered = [];
    for (const c of allContacts) {
      const normalizedFolder = await normalizeContactFolderId(graph, userId, c.parentFolderId ?? '');
      if (await isPathDenied(tenantId, userId, 'contacts', normalizedFolder)) continue;
      filtered.push(c);
    }
    contacts = filtered;
  }

  // Strip internal parentFolderId from response
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const cleaned = contacts.map((c: any) => {
    const { parentFolderId: _pf, ...rest } = c;
    return rest;
  });

  return { status: 200, jsonBody: { contacts: cleaned, count: cleaned.length } };
}

app.http('searchContacts', {
  methods: ['GET'],
  authLevel: 'anonymous',
  route: 'api/contacts/search',
  // Deny-list filtering is handled inside the handler — it checks each
  // contact's parentFolderId when no explicit folderId is provided.
  handler: withSecurity(withPolicyEnforcement('contacts', searchContactsHandler, {
    getDenyListPaths: (req) => {
      const folderId = req.query.get('folderId');
      return folderId ? [folderId] : [];
    },
  })),
});
