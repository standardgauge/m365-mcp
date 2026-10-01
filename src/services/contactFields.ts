/**
 * Shared contact field mapping for the contacts surface.
 *
 * `create_contact`, `update_contact`, and `create_contacts_batch` all translate
 * the same tool argument set into a Microsoft Graph `contact` resource body, and
 * `search_contacts` shapes the same field set back out. Keeping the mapping in one
 * place is what makes the round-trip (write a field, read it back through
 * search_contacts) stay in sync across every entry point.
 *
 * Graph contact resource reference:
 * https://learn.microsoft.com/graph/api/resources/contact
 */

/** Microsoft Graph `physicalAddress` — every sub-field is optional. */
export interface PhysicalAddress {
  street?: string;
  city?: string;
  state?: string;
  postalCode?: string;
  countryOrRegion?: string;
}

/** Tool-facing contact fields accepted by create/update/batch. */
export interface ContactInput {
  givenName?: string;
  surname?: string;
  middleName?: string;
  nickName?: string;
  emailAddresses?: string[];
  businessPhones?: string[];
  homePhones?: string[];
  mobilePhone?: string;
  companyName?: string;
  jobTitle?: string;
  personalNotes?: string;
  categories?: string[];
  birthday?: string;
  spouseName?: string;
  homeAddress?: PhysicalAddress;
  businessAddress?: PhysicalAddress;
  otherAddress?: PhysicalAddress;
}

/**
 * Graph `$select` list for reading a contact back with the full surface. Kept in
 * lock-step with {@link shapeContact} and {@link buildContactBody} so a field a
 * caller can write is a field search_contacts returns.
 */
export const CONTACT_SELECT_FIELDS =
  'id,displayName,givenName,surname,middleName,nickName,emailAddresses,' +
  'businessPhones,homePhones,mobilePhone,companyName,jobTitle,personalNotes,' +
  'categories,birthday,spouseName,homeAddress,businessAddress,otherAddress,parentFolderId';

const ADDRESS_SUBFIELDS: (keyof PhysicalAddress)[] = [
  'street',
  'city',
  'state',
  'postalCode',
  'countryOrRegion',
];

/**
 * JSON-Schema fragment describing a `physicalAddress` argument, used verbatim in
 * the MCP inputSchema for the three postal-address properties.
 */
export const ADDRESS_SCHEMA = {
  type: 'object',
  description: 'Postal address (street, city, state, postalCode, countryOrRegion — all optional)',
  properties: {
    street: { type: 'string', description: 'Street address' },
    city: { type: 'string', description: 'City' },
    state: { type: 'string', description: 'State or province' },
    postalCode: { type: 'string', description: 'ZIP / postal code' },
    countryOrRegion: { type: 'string', description: 'Country or region' },
  },
} as const;

/** Copy only the recognised sub-fields out of a caller-supplied address object. */
function pickAddress(a: unknown): PhysicalAddress | undefined {
  if (!a || typeof a !== 'object') return undefined;
  const src = a as Record<string, unknown>;
  const out: PhysicalAddress = {};
  for (const k of ADDRESS_SUBFIELDS) {
    if (src[k] !== undefined) out[k] = src[k] as string;
  }
  return Object.keys(out).length > 0 ? out : undefined;
}

/**
 * Build a Graph `contact` request body from tool arguments.
 *
 * Field-presence semantics are `!== undefined`, so the same builder serves
 * create (an absent field is simply not sent) and PATCH update (an absent field
 * leaves the stored value unchanged). `emailAddresses` is a plain string array on
 * the tool surface and is expanded to Graph's `{ address, name }` shape here.
 */
// eslint-disable-next-line @typescript-eslint/no-explicit-any
export function buildContactBody(args: Record<string, any>): Record<string, unknown> {
  const contact: Record<string, unknown> = {};

  const stringFields = [
    'givenName',
    'surname',
    'middleName',
    'nickName',
    'mobilePhone',
    'companyName',
    'jobTitle',
    'personalNotes',
    'birthday',
    'spouseName',
  ];
  for (const f of stringFields) {
    if (args[f] !== undefined) contact[f] = args[f];
  }

  const arrayFields = ['businessPhones', 'homePhones', 'categories'];
  for (const f of arrayFields) {
    if (args[f] !== undefined) contact[f] = args[f];
  }

  if (args.emailAddresses !== undefined) {
    contact.emailAddresses = (args.emailAddresses as string[]).map((addr) => ({
      address: addr,
      name: addr,
    }));
  }

  const home = pickAddress(args.homeAddress);
  if (home) contact.homeAddress = home;
  const business = pickAddress(args.businessAddress);
  if (business) contact.businessAddress = business;
  const other = pickAddress(args.otherAddress);
  if (other) contact.otherAddress = other;

  return contact;
}

/**
 * Shape a raw Graph contact into the search_contacts response object. Retains the
 * original compact keys (`emails`, `phones`, `company`) for backward
 * compatibility and adds the full field surface. `parentFolderId` is carried for
 * internal deny-list filtering and is stripped before the response is returned.
 */
// eslint-disable-next-line @typescript-eslint/no-explicit-any
export function shapeContact(c: Record<string, any>): Record<string, unknown> {
  return {
    id: c.id,
    displayName: c.displayName,
    givenName: c.givenName ?? null,
    surname: c.surname ?? null,
    middleName: c.middleName ?? null,
    nickName: c.nickName ?? null,
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    emails: (c.emailAddresses ?? []).map((e: any) => e.address),
    phones: [...(c.businessPhones ?? []), ...(c.homePhones ?? []), c.mobilePhone].filter(Boolean),
    mobilePhone: c.mobilePhone ?? null,
    businessPhones: c.businessPhones ?? [],
    homePhones: c.homePhones ?? [],
    company: c.companyName ?? null,
    jobTitle: c.jobTitle ?? null,
    personalNotes: c.personalNotes ?? null,
    categories: c.categories ?? [],
    birthday: c.birthday ?? null,
    spouseName: c.spouseName ?? null,
    homeAddress: c.homeAddress ?? null,
    businessAddress: c.businessAddress ?? null,
    otherAddress: c.otherAddress ?? null,
    parentFolderId: c.parentFolderId,
  };
}

/**
 * Resolve the Graph collection path for reading/writing contacts in a folder.
 * The synthetic `contacts-root` key (used by list_contact_folders to represent
 * the default folder) maps back to `/me/contacts`, so a folderId returned by
 * list_contact_folders is always usable as a create/batch target.
 */
export function contactsApiPath(folderId?: string): string {
  if (!folderId || folderId === 'contacts-root') return '/me/contacts';
  return `/me/contactFolders/${folderId}/contacts`;
}

/** Per-entry outcome of a bulk contact create. */
export interface ContactBatchResult {
  index: number;
  status: 'created' | 'failed';
  id?: string;
  error?: string;
}

/** Graph `$batch` caps a single request at 20 sub-operations. */
export const CONTACT_BATCH_LIMIT = 20;

/**
 * Create contacts via Graph `$batch`, chunked at the 20-op-per-request limit
 *. Entries missing `givenName` are rejected client-side — parity with
 * create_contact — without consuming a Graph request. Result order matches input
 * order; Graph may return sub-responses out of order, so they are matched by the
 * per-chunk id assigned here. A whole-chunk failure (network/token) marks every
 * entry in that chunk failed and continues with the next chunk.
 */
export async function runContactBatch(
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  graph: { api: (path: string) => { post: (body: unknown) => Promise<any> } },
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  contacts: Record<string, any>[],
  url: string,
): Promise<ContactBatchResult[]> {
  const results: ContactBatchResult[] = new Array(contacts.length);

  const sendable: number[] = [];
  contacts.forEach((c, index) => {
    if (!c || typeof c !== 'object' || c.givenName === undefined || c.givenName === null || c.givenName === '') {
      results[index] = { index, status: 'failed', error: 'givenName is required' };
    } else {
      sendable.push(index);
    }
  });

  for (let i = 0; i < sendable.length; i += CONTACT_BATCH_LIMIT) {
    const chunk = sendable.slice(i, i + CONTACT_BATCH_LIMIT);
    const requests = chunk.map((originalIndex, j) => ({
      id: String(j),
      method: 'POST',
      url,
      headers: { 'Content-Type': 'application/json' },
      body: buildContactBody(contacts[originalIndex]),
    }));
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    let responses: any[] = [];
    try {
      const batchResult = await graph.api('/$batch').post({ requests });
      responses = batchResult?.responses ?? [];
    } catch (err) {
      const msg = err instanceof Error ? err.message : 'batch request failed';
      for (const originalIndex of chunk) results[originalIndex] = { index: originalIndex, status: 'failed', error: msg };
      continue;
    }
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const byId = new Map<string, any>();
    for (const r of responses) byId.set(String(r.id), r);
    chunk.forEach((originalIndex, j) => {
      const r = byId.get(String(j));
      if (r && r.status >= 200 && r.status < 300) {
        results[originalIndex] = { index: originalIndex, status: 'created', id: r.body?.id };
      } else {
        const error = r?.body?.error?.message ?? `Graph returned status ${r?.status ?? 'unknown'}`;
        results[originalIndex] = { index: originalIndex, status: 'failed', error };
      }
    });
  }

  return results;
}
