# Scale review: failure points in very large organizations

Reviewed at commit `ef421e5` (main, 2026-10-10). Every instance deployed so
far has served a tenant of fifteen users or fewer. This review reads the code
for what breaks first when one instance serves thousands to tens of thousands of
users: mailboxes with hundreds of thousands of messages, tenants with thousands
of sites, org-wide Teams, and administrators who belong to hundreds of groups.

It is analysis only. Nothing here changes behavior. Each finding carries the
evidence (file and line at the reviewed commit), how likely it is to fire at the
target scale, what it takes down when it does, and the change that fixes it.
The [proposals](#proposals) at the end are the work queue the findings imply.
Two things are measured rather than estimated: the number of Table Storage
operations each policy read issues, and the heap cost of the per-account MSAL
applications. Everything else is derived from the code and the published
service limits in [Sources](#sources).

## Summary

Priorities:

- **P1**: expected to fail outright, or to take a replica or the whole instance
  down, at ten thousand users. Fix before the first large deployment.
- **P2**: degrades a tool class or needs an operator to notice and intervene.
  Fix before relying on the instance at that scale.
- **P3**: latent or narrow. Schedule.

Likelihood is judged at ten thousand users with ordinary agent usage (the
[load model](#load-model) below). Blast radius is what fails when it fires.

| ID | Finding | Likelihood at 10k users | Blast radius | Priority |
|---|---|---|---|---|
| [F1](#f1-every-tool-call-fans-out-to-table-storage) | Every tool call fans out to Table Storage: 11 to 400 operations per call, uncached, on a handful of partitions, failing closed | Certain (latency); failure likely | Instance: every tool refused while probes stay green | P1 |
| [F2](#f2-file-and-attachment-reads-are-unbounded-in-memory) | File and attachment reads buffer the whole object, then copy it three to five times | High | Replica: out of memory kills every in-flight request | P1 |
| [F3](#f3-the-runtime-envelope-stops-at-three-small-replicas) | Three replicas of 0.5 vCPU, a 240 second ingress timeout, no per-request deadline | Certain | Instance: queueing collapse | P1 |
| [F4](#f4-site-listing-reads-one-page-and-filters-afterwards) | Site listing reads one page of a tenant-wide search and filters afterwards | Certain above 100 sites | SharePoint discovery, admin site picker | P1 |
| [F5](#f5-the-global-administrator-check-reads-one-page) | The Global Administrator check reads one page of memberships | High | Admin console: every admin function refused | P1 |
| [F6](#f6-session-rows-are-scanned-decrypted-and-never-purged) | Session rows share one partition, are scanned and decrypted whole, and are never purged | High | Admin pages, cold start, logout, memory | P1 |
| [F7](#f7-the-outbound-policy-reads-storage-through-the-wrong-client) | The outbound policy builds its own storage client from a connection string the templates do not set | Certain on template deployments | Calendar and Teams writes refused outright | P1 (not scale) |
| [F8](#f8-nothing-reports-degradation) | The health check is a constant and no metric reports throttling, latency or cache growth | Certain | Operability: every finding above is invisible | P1 |
| [F9](#f9-the-rate-limiter-keys-on-the-egress-address) | The MCP rate limiter keys on the client address, which a large organization shares | High behind a proxy or VPN | Everyone behind that address | P2 |
| [F10](#f10-graph-retries-have-no-deadline-and-429s-are-opaque) | Graph retries can sleep nine minutes, carry no jitter, and surface 429 as plain text | Medium | One user, amplified across the mailbox | P2 |
| [F11](#f11-offset-paging-re-walks-from-page-one) | Offset paging re-walks a library from page one and cannot reach past 10,000 items | High for large libraries | SharePoint listing, SharePoint resource units | P2 |
| [F12](#f12-sharepoint-search-filters-a-tenant-wide-window) | SharePoint search filters a tenant-wide window of 200 hits against the allow-list | High with an allow-list | Search returns nothing | P2 |
| [F13](#f13-the-internal-teams-policy-enumerates-members) | The internal-only Teams policy enumerates members and refuses large channels | High for org-wide teams | Teams sends under that policy | P2 |
| [F14](#f14-the-audit-table-grows-faster-than-it-is-purged) | The audit table grows faster than its purge can delete, and admin queries scan it | Certain over months | Admin audit page, storage cost | P2 |
| [F15](#f15-the-tenant-user-picker-enumerates-the-directory) | The tenant user picker enumerates the whole directory into memory and one response | Medium; certain above 50k users | Admin user management | P2 |
| [F16](#f16-in-memory-caches-grow-for-the-life-of-the-process) | Session, resolver and device-code caches grow for the life of the process | Medium | Replica memory over weeks | P2 |
| [F17](#f17-list-tools-that-stop-at-one-page) | Eleven list tools stop at one page with a flag and no way to continue | High | Completeness of listings | P3 |
| [F18](#f18-one-mailbox-four-concurrent-requests) | Exchange allows four concurrent requests per mailbox; an agent fanning out exceeds it | Medium | One user | P3 |
| [F19](#f19-outbound-calls-have-no-timeouts-and-readiness-ignores-storage-init) | Outbound calls have no timeouts; readiness ignores storage initialization | Medium | Replica during an outage | P3 |
| [F20](#f20-device-code-flows-poll-entra-for-fifteen-minutes-each) | Each device-code sign-in polls Entra for fifteen minutes and leaves tokens in memory | Low | Replica memory | P3 |
| [F21](#f21-per-account-msal-applications-and-refresh-races) | Per-account MSAL applications and cross-replica refresh races | Low | None found | P3 |
| [F22](#f22-text-filters-over-very-large-folders) | Text search inside a very large folder uses an unindexed filter that times out | Medium | One call | P3 |

The order of work is at the end: [Proposals](#proposals).

## The deployment under load

### Request path and per-call cost

Every tool call is one HTTP `POST /api/mcp` carrying one JSON-RPC message
(`src/functions/mcp/mcpEndpoint.ts:2939-2998`). There is no streaming and no
server-side MCP session: the request is authenticated, the dispatcher reads the
tenant's policy, the tool handler runs, an audit row is written, and the
response is returned. Nothing is cached between calls except the session
(`src/services/tokenCache.ts:147-149`), a few Graph ID lookups
(`src/services/containerResolver.ts:18-23`) and the tenant's verified domains
(`src/services/outboundPolicy.ts:214-246`).

The policy reads are the cost that scales. Each read first calls
`ensureTable()`, which issues a `createTable` request every time (the SDK
swallows the 409; `@azure/data-tables` 13.3.2, `handleTableAlreadyExists`), then
a point read or a partition query. Measured by stubbing the Table SDK and
counting calls:

| Primitive | Table Storage operations | Where |
|---|---|---|
| `isPathDenied` (one subject) | 4 (2 `createTable`, 2 partition queries) | `src/services/denyList.ts:280-294` via `loadDeniedPaths` `212-226` |
| `isPathDenied` (delegated mailbox: caller and owner) | 6 | same, two per-user lists |
| `filterDeniedPaths` (any list length) | 4 | `src/services/denyList.ts:371-392` |
| `filterDeniedSearchHits` | 4 | `src/services/denyList.ts:342-365` |
| `getEnabledServices` | 2 | `src/services/serviceSettings.ts:26-37` |
| `getReadOnlyServices` | 2 | `src/services/serviceSettings.ts:58-69` |
| `getAllowedSites` | 2 | `src/services/serviceSettings.ts:84-95` |
| `getUserServiceOverrides` | 2 | `src/services/userServiceOverrides.ts:42-53` |
| `isMailIndexingDisabled` | 2 | `src/services/userMailConfig.ts:109-117` |
| `getUserEmailSettings` | 5 | `src/services/userEmailSettings.ts:207` |

The dispatcher reads `getEnabledServices` and `getUserServiceOverrides` on every
call (`mcpEndpoint.ts:2719`, `2734`), `isMailIndexingDisabled` on every mail
call (`2749`), `getReadOnlyServices` on every write (`2768`) and
`getAllowedSites` on every site-scoped call (`2787`). `tools/list` reads all
four (`2673-2676`). Handlers then add their own deny-list checks, and two of
them check per item: `stripDeniedFolders` calls `isPathDenied` for every
message in a mailbox-wide result (`mcpEndpoint.ts:211-229`, called at `961` and
`1001`), and `search_contacts` calls it for every contact (`2248-2251`).

Worked tallies, Table Storage operations per call:

| Tool call | Dispatcher | Handler | Audit | Total | Graph calls |
|---|---|---|---|---|---|
| `read_message` | 6 | 4 | 1 | 11 | 2 (one folder-name lookup, cached per process) |
| `list_folders` (SharePoint) | 6 | 4 | 1 | 11 | 1 to 50 pages |
| `send_mail` with one inline attachment | 8 | 9 | 1 | 18 | 1 to 3 |
| `search_mail`, mailbox-wide, 25 results | 6 | 25 × 4 | 1 | 107 | 1 to 11, plus up to 25 folder lookups the first time |
| `list_messages`, 100 results | 6 | 100 × 4 | 1 | 407 | 1 |
| `search_contacts`, 50 results, no folder | 4 | 50 × 4 | 1 | 205 | 1 to 2 |

The per-message and per-contact checks run sequentially (`for ... await`), so
at 5 to 15 ms per Table Storage operation a `list_messages` of 100 rows spends
two to six seconds in storage before anything else. The same calls would cost
one read each with the deny list loaded once per call.

### Capacity envelope

- **Compute.** Each replica is 0.5 vCPU and 1 GiB (`infra/main.bicep:341-344`,
  `infra/container-app.bicep:270-273`). The Azure Functions host runs one Node
  worker, so one event loop per replica does all the AES-GCM, HMAC, base64 and
  JSON work, and shares the 1 GiB with the host process (`Dockerfile:26`).
- **Scale.** One to three replicas, scaled out when a replica averages more than
  20 concurrent requests (`infra/main.bicep:376-389`). Beyond three there is no
  more scaling; each replica then absorbs whatever arrives.
- **Time.** Container Apps ingress gives a request 240 seconds by default. The
  Functions host has no `functionTimeout` set (`host.json`), so a handler keeps
  running after the ingress has answered 504. Nothing in `src/` passes an
  `AbortSignal` or a timeout to a Graph or Table call, and Node's fetch has no
  overall deadline of its own.
- **Storage.** Azure Table Storage allows about 20,000 operations per second per
  account and 2,000 per second per partition. This instance concentrates its
  traffic on a few partitions: `serviceSettings/<tenant>`,
  `GlobalDenyList/<tenant>:<type>`, `UserServiceOverrides/<tenant>`,
  `mcpSessions/session`, `mcpMsalCache/account` and `auditLog/<tenant>`.
- **Graph.** The limits that matter here, from [Sources](#sources): Outlook
  allows 10,000 requests per 10 minutes per app per mailbox and 4 concurrent
  requests per app per mailbox; SharePoint and OneDrive allow 1,250 resource
  units per minute per app per tenant for a tenant under 1,000 licenses, rising
  to 6,250 for 50,000 and above, and a multi-item read costs 2 units; delegated
  search is throttled at 10 requests per second per user; directory reads are
  metered at 8,000 resource units per 10 seconds per app for a large tenant; the
  global ceiling is 130,000 requests per 10 seconds per app.

### Load model

The tiers below are the assumptions the findings are judged against. They are
deliberately ordinary: a fraction of users active in the peak hour, each with an
agent making a few calls a minute. A tenant that rolls agents out to everyone,
or runs scheduled agents, lands higher.

| Tier | Users | Active in peak hour | Calls per active user per minute | Tool calls per second | Table operations per second at ~50 per call | Requests in flight at ~1 s each |
|---|---|---|---|---|---|---|
| A | 1,000 | 10% | 2 | 3 | 170 | 3 |
| B | 10,000 | 10% | 2 | 33 | 1,700 | 33 |
| C | 50,000 | 10% | 2 | 170 | 8,500 | 170 |
| C heavy | 50,000 | 20% | 5 | 830 | 42,000 | 830 |

Against the envelope: tier B already needs two or three replicas for the
concurrency alone and puts the shared deny-list partition near its 2,000
operations per second; tier C exceeds the three-replica ceiling several times
over and the heavy variant exceeds the storage account's total throughput.

## Findings

### P1

### F1 Every tool call fans out to Table Storage

**What the code does.** Every policy read goes to storage, every time, with an
extra `createTable` round trip in front of it:
`src/services/serviceSettings.ts:16-24` and `26-37`,
`src/services/userServiceOverrides.ts:28-36` and `42-53`,
`src/services/userMailConfig.ts:36-44` and `50-61`,
`src/services/userEmailSettings.ts:92-96`,
`src/services/denyList.ts:24-32`, `77-92` and `131-146`. `isPathDenied` reloads
the default, global and per-user lists on each call
(`src/services/denyList.ts:212-226`, `280-294`), and the two per-item loops
call it once per row (`src/functions/mcp/mcpEndpoint.ts:211-229`,
`2248-2251`). Every one of these paths fails closed: a storage error becomes
"denied" (`denyList.ts:289-293`) or an empty list (`362-364`, `388-391`), and
`isMailIndexingDisabled` becomes "mail disabled" (`userMailConfig.ts:113-116`).

**Why it fails at scale.** Three effects compound.

1. *Latency on every call.* The [tallies](#request-path-and-per-call-cost)
   above are sequential round trips. A 100-row listing carries two to six
   seconds of storage latency that has nothing to do with Graph.
2. *Hot partitions.* The global deny list for a tenant is one partition per
   type (`denyList.ts:53`, `83`) and the tenant settings are one partition
   (`serviceSettings.ts:43`). At tier B the deny-list partition takes several
   hundred reads per second from the per-message loops alone; at tier C it
   passes the 2,000 per second partition ceiling and the account approaches
   its 20,000. Table Storage answers with 503 `ServerBusy` and `Retry-After`.
3. *Fail-closed cascade.* When storage throttles, the Table SDK retries each
   operation with exponential backoff (three attempts by default), so a call
   that issues 100 operations stalls for the ingress timeout, and the ones that
   complete return "denied". Every user on the instance sees "Access restricted
   by deny list" or empty listings. `/health` is still 200 (see
   [F8](#f8-nothing-reports-degradation)), so nothing restarts and nothing
   alerts.

**Likelihood and blast radius.** Latency: certain today. Throttling: likely at
tier B under per-message loops, certain at tier C. Blast radius: the whole
instance, every tool.

**Fix.** An in-process policy cache with a short time to live, and one load
per call:

- Cache `getEnabledServices`, `getReadOnlyServices`, `getAllowedSites`, the
  tenant-wide email and outbound policies, the global deny lists, and the
  per-user entries (overrides, mail config, email settings, deny lists) in
  memory, keyed by tenant or tenant plus user, with a 30 to 60 second TTL and a
  bounded size (per-user entries evict by last use). Serve a stale entry while
  a refresh is in flight so a storage blip does not fail closed inside the TTL.
- Load the deny list once at the top of a handler and match every item in
  memory; `matchesDenyList` already works on an in-memory array
  (`denyList.ts:265-273`). `stripDeniedFolders` and `search_contacts` then cost
  one load instead of one per row.
- Create tables once (at startup, or on the first write, remembered per
  process) and drop `ensureTable()` from every read path.
- Keep the fail-closed semantics, but only after the cache has nothing to
  serve, and count those events (see F8).

Size: medium. This single change removes well over 90% of the instance's
storage traffic. The staleness window is a security decision, recorded as
[D2](#decisions).

### F2 File and attachment reads are unbounded in memory

**What the code does.** `read_file` and `read_onedrive_file` stream the whole
object into an array of chunks, concatenate it, and base64-encode it into the
JSON response, with no size check before or during the download
(`src/functions/mcp/mcpEndpoint.ts:565-571` and `1452-1458`; the REST routes do
the same at `src/functions/sharepoint/readFile.ts:73-93` and
`src/functions/onedrive/readOneDriveFile.ts:53-73`). The response is then
pretty-printed with `JSON.stringify(result, null, 2)` (`mcpEndpoint.ts:2883`),
a second full copy. Attachments decode their base64 several times over:
`decodedByteLength` materializes the bytes just to measure them
(`src/services/mailAttachments.ts:81-83`, called at `165` and again in
`requiresUploadSession` at `194-196`), `partitionAttachments` runs twice in
`sendMailWithAttachments` (`294-296`), and `uploadAttachmentViaSession` decodes
once more (`220`). A drive-item attachment is buffered, base64-encoded
(`src/services/driveAttachments.ts:130-151`) and decoded again for upload. An
attached file comes back from Graph inline as `contentBytes` and is decoded and
re-encoded (`src/services/mailAttachmentContent.ts:89-94`). The attachment cap
is 150 MB (`mailAttachments.ts:72`).

**Why it fails at scale.** Peak memory per request is three to five times the
object size. A 300 MB file read, or one 150 MB attachment, needs most of the
replica's 1 GiB, which the Functions host already shares. Two such requests at
once, or one on a replica that is holding a few thousand cached sessions
([F16](#f16-in-memory-caches-grow-for-the-life-of-the-process)), exhaust it.
Node dies, every in-flight request on that replica fails, the client retries
the same read, and the next replica dies. Large files are routine in a large
organization: recordings, decks, exports, archives, and an agent asked to
"read this" will try. Even below the kill threshold, the base64 and JSON work
runs on the single event loop for seconds and starves every other request on
the replica, including the liveness probe (`infra/probes.json:19-24`, 5 second
timeout, three failures), which restarts the replica anyway.

**Likelihood and blast radius.** High. Blast radius: the replica and every user
on it, with a rolling pattern across replicas.

**Fix.**

- Check `size` from the item metadata before downloading (both read tools
  already fetch it: `mcpEndpoint.ts:560`, `1447`) and refuse above a
  configurable cap, returning the size and `webUrl` instead. A default around
  25 MB for the read tools and 10 MB for inline attachments matches what a
  model can use and what the shim already enforces for local files.
- Keep a running total while streaming and abort past the cap, as
  `driveAttachments.ts:133-141` already does.
- For drive-item attachments, stream the bytes straight into the upload
  session in 320 KiB multiples instead of base64-encoding and decoding them.
- Compute an attachment's decoded length once from the base64 length.
- Drop the pretty-printing for large results, or stream the JSON.

Size: medium. The cap defaults are part of [D1](#decisions).

### F3 The runtime envelope stops at three small replicas

**What the code does.** `maxReplicas: 3`, 0.5 vCPU and 1 GiB each, scale rule
at 20 concurrent requests (`infra/main.bicep:341-344`, `376-389`;
`infra/container-app.bicep:270-273`, `299-312`). The runbook describes the
allocation as adequate for about twenty concurrent users
(`docs/operations-runbook.md`, Right-sizing). No handler has a deadline, and
the ingress has a 240 second one.

**Why it fails at scale.** At tier B the instance already runs near its
scale-out ceiling (33 requests in flight against a trigger of 20 per replica
and a cap of three). At tier C the ceiling is passed several times over. Past
it, requests queue inside each replica's event loop; latency climbs until the
ingress returns 504, the handler keeps running anyway, and the client retries,
adding load to an instance that is already over capacity. Because the CPU-bound
work per call (decrypting sessions, HMAC, JSON) is on one loop at half a core,
the collapse arrives earlier than the raw request count suggests.

**Likelihood and blast radius.** Certain at tier C, likely at tier B during
events (an all-hands, a rollout). Blast radius: the instance.

**Fix.**

- Parameterize the templates and raise the defaults for a large tier: 1 vCPU
  and 2 GiB per replica, `maxReplicas` 10 or more, and a scale trigger of 10 to
  20 concurrent requests. Document per-tier sizing in the runbook.
- Give every tool call a deadline below the ingress timeout (for example 200
  seconds overall, 60 seconds per Graph call, 5 seconds per Table operation)
  through `AbortSignal`, and return a retryable JSON-RPC error when it is
  spent, so work the client has stopped waiting for is not finished anyway.
- Document the ingress `requestTimeoutSeconds` setting for operators who want
  longer calls, and keep the server's deadline under it.

Size: small for the configuration, medium for the deadlines. The target tier
is [D1](#decisions).

### F4 Site listing reads one page and filters afterwards

**What the code does.** `list_sites` calls `/sites?search=*` with `$top=100`,
takes that single page, drops system sites, and then removes anything not on
the allow-list (`src/functions/mcp/mcpEndpoint.ts:474-492`). The REST route
behind the admin site picker does the same without even `$top`, and without a
truncation flag (`src/functions/sharepoint/listSites.ts:31-34`, `58-83`).

**Why it fails at scale.** A large tenant has thousands of site collections,
and `search=*` returns them in no useful order. The allow-listed sites may not
be in the first hundred, in which case the tool returns an empty list with
`truncated: true` and no way to page, and an agent concludes the user has no
sites. The admin picker has the same hole, which is why "Add by URL" exists
(`listSites.ts:51-55`); at scale it is the only path that works.

**Likelihood and blast radius.** Certain above a hundred sites. Blast radius:
SharePoint discovery for every user, and the allow-list administration.

**Fix.**

- When an allow-list exists, enumerate it directly: `/sites/{id}` per entry
  (`$select=id,displayName,webUrl`, cached for minutes), which is also the
  cheaper call.
- Without one, give `list_sites` a `q` argument that maps to
  `/sites?search={q}` and a continuation token (see
  [F11](#f11-offset-paging-re-walks-from-page-one)), and page the REST route
  for the picker.

Size: small to medium.

### F5 The Global Administrator check reads one page

**What the code does.** `checkGlobalAdmin` calls `/me/transitiveMemberOf` with
`$select=id,roleTemplateId` and looks for the Global Administrator template ID
in the response (`src/services/authMiddleware.ts:272-286`). It reads one page,
never follows `@odata.nextLink`, and fails closed. It is called on every admin
API request (`authorizeAdmin` `307-311`, `requireGlobalAdmin` `318-326`), on
`/api/auth/me` (`src/functions/auth/me.ts:61`), on the admin bypass
(`src/services/policyEnforcement.ts:227-228`) and from the SharePoint routes
(`src/functions/sharepoint/listSites.ts:59`, `resolveSite.ts:28`). The result is
not cached, so each of those is a directory read.

**Why it fails at scale.** `transitiveMemberOf` returns groups and directory
roles in one collection, 100 objects per page by default, in no guaranteed
order. In a large organization an administrator is transitively a member of
hundreds of groups, so the role object is often not on the first page and the
check answers "not an administrator". Every admin function then returns 403.
The repeated directory reads also count against the tenant's identity
resource units, though at admin volumes that is minor.

**Likelihood and blast radius.** High. Blast radius: the admin console in
exactly the organizations this review is about.

**Fix.** Ask for roles only and page:
`/me/transitiveMemberOf/microsoft.graph.directoryRole?$select=roleTemplateId&$top=999`
(or `/me/memberOf/microsoft.graph.directoryRole`), following `nextLink` until
exhausted or matched, and cache the answer per session for a few minutes.
Size: small.

### F6 Session rows are scanned decrypted and never purged

**What the code does.** Every session row lives in one partition,
`PartitionKey = 'session'` (`src/services/tableStorage.ts:69-70`). Lookups by
token are point reads (`424-441`), but everything keyed by user is a scan of
the whole partition with an AES-GCM decrypt per row: `loadSession(userId)`
(`400-417`), `removeSession(userId)` (`545-557`, run on every logout,
`src/functions/auth/logout.ts:48`), and `listAllSessions()` (`559-569`), which
backs the admin sessions page (`src/functions/admin/getSessions.ts:22`) and the
tenant user picker (`src/functions/admin/getTenantUsers.ts:93-96`) and pins
every row it reads into the replica's memory cache
(`src/services/tokenCache.ts:466-492`). `migrateLegacySessionRows` scans the
partition on every process start, before the first request is served
(`tableStorage.ts:34-49`, `460-530`). Nothing deletes a row once its session
passes the 30 day absolute lifetime: `isAbsoluteLifetimeExceeded` refuses it
(`tokenCache.ts:111-115`) and leaves it in place. Every install, sign-in and
device login adds a row (`src/functions/auth/callback.ts:161`,
`src/functions/auth/deviceLogin.ts:108`).

**Why it fails at scale.** The table only grows: ten thousand users with a
laptop, a desktop extension and the odd reinstall is 30,000 to 50,000 rows in a
year, each a kilobyte or more with its ciphertext. A scan pages through them a
thousand at a time, decrypts every access-token envelope on the event loop,
and builds a `StoredSession` for each, which is tens to hundreds of megabytes
for one admin page load, on a 1 GiB replica, and then stays resident. A new
replica pays the scan before it answers its first request, while the readiness
probe has already told the ingress it is ready, so the first users routed to
it wait on `initPromise`. A logout scans the table. The `getSession(userId)`
fallback (`tokenCache.ts:256-273`) scans it whenever a user-keyed lookup misses
the replica's index, which `getTenantId(userId)` callers across the REST
routes and `src/services/policyEnforcement.ts:39-124` can trigger.

**Likelihood and blast radius.** High within months at tier B. Blast radius:
the admin sessions and user pages (slow, then out of memory), every cold start,
every logout, and replica memory.

**Fix.**

- A user-keyed index: a second row per session under
  `PartitionKey = user:<userId>` (or a small index table) so `loadSession` and
  `removeSession` are partition queries.
- Column projection for listings: never read the ciphertext columns when
  listing sessions, and never decrypt to list.
- Paginate `/api/manage/sessions`, and stop caching what it lists.
- Purge rows past the absolute lifetime on the same once-a-day pattern the
  audit purge uses, and rows whose refresh has failed for longer than the
  renewal window.
- Gate the legacy migration behind a marker row so it runs once per
  deployment rather than once per process, and remove it once no legacy rows
  remain.

Size: medium.

### F7 The outbound policy reads storage through the wrong client

This is a correctness defect found while reading the storage paths, not a scale
finding. It is here because it blocks the Teams and scheduling surfaces on
exactly the deployment shape the templates produce.

**What the code does.** Every other module reaches Table Storage through
`src/services/storageClient.ts:52-57`, which uses the managed identity when
`AZURE_STORAGE_TABLE_ENDPOINT` is set and the connection string otherwise. The
outbound policy builds its own client from
`AZURE_STORAGE_CONNECTION_STRING`, falling back to the local emulator address
(`src/services/outboundPolicy.ts:117-122`). The templates set the endpoint and
no connection string (`infra/main.bicep:346-373`). On such an instance every
policy read (`141-157`, `191-199`) tries to reach an emulator that is not
there, and `enforceOutboundPolicy` fails closed (`302-319`): `create_event`
with attendees, `update_event` on any notifying field, `move_event`,
`respond_to_event` with a comment, `send_chat_message` and
`send_channel_message` are all refused with "the outbound policy could not be
read", and the admin policy endpoints fail. The tests mock
`TableClient.fromConnectionString` directly
(`src/__tests__/outboundPolicy.test.ts:25-27`), so CI cannot see it.

**Fix.** Use `getTableClient('OutboundPolicy')` from `storageClient.ts`, and
add a test that sets the endpoint variable and asserts no connection-string
client is constructed. Size: small.

### F8 Nothing reports degradation

**What the code does.** `/health` returns a constant
(`src/functions/health.ts:19-28`), and all three probes read it
(`infra/probes.json`). Application Insights dependency and performance
collection are turned off (`src/services/telemetry.ts:17-20`), so no outbound
call to Graph or storage is recorded anywhere. `host.json` enables adaptive
sampling for everything but requests, so console output is sampled once the
instance is busy. No metric exists for Graph 429s, retries or latency, Table
Storage latency or throttling, per-tool latency, cache sizes, or the
fail-closed events in [F1](#f1-every-tool-call-fans-out-to-table-storage).
The only signals are console lines: `[denyList] Storage unavailable`
(`denyList.ts:291`), `rate limit: ... refused` (`rateLimit.ts:131`),
`[tokenCache] Silent token refresh FAILED` (`tokenCache.ts:426-431`),
`MCP tool ... error` (`mcpEndpoint.ts:2888`), `Hit MAX_PAGES`
(`getTenantUsers.ts:60`) and the audit upload failure
(`auditLogAnalytics.ts:151-154`). The runbook's monitoring section lists
platform metrics (requests, CPU, memory, replica count, latency), which show
symptoms, not causes.

**Why it matters at scale.** Every failure above is silent from the outside.
A replica with a dead storage credential, a stalled `initPromise`, a saturated
event loop or a throttled Graph tenant is "healthy" to the probes and to an
external monitor. An operator learns about it from users.

**Fix.**

- Keep `/health` shallow for the probes. Add `/health/deep` (or `?deep=1`)
  for external monitors, returning cached results of cheap dependency checks:
  a Table Storage point read with a two second budget, a managed identity
  token, and the MSAL configuration, plus the process's heap size and event
  loop delay.
- Emit custom metrics (Application Insights `trackMetric`, or OpenTelemetry):
  `graph.requests` by host and status, `graph.retries`,
  `graph.retry_after_seconds`, `graph.latency`, `table.ops` by table and
  status, `table.latency`, `tool.duration` by tool, `tool.outbound_calls`,
  `policy.fail_closed`, `sessions.cached`, `rate_limit.refusals`.
- Log one structured line per tool call: tool, duration, Graph calls, Table
  operations, retries, outcome.
- Exclude the warning and error categories from sampling, or raise the cap.
- Document alert rules and the Kusto queries for them in the runbook: 429
  rate, p95 tool duration, replica count at maximum, memory above 80%, any
  `policy.fail_closed`, audit upload failures.

Size: medium.

### P2

### F9 The rate limiter keys on the egress address

**What the code does.** `/api/mcp` is limited to 1,200 requests per minute per
client address per replica, in memory (`src/services/rateLimit.ts:5-16`,
`32-38`, `120-143`; applied at `mcpEndpoint.ts:3004`). The address is the one
the ingress appended to `X-Forwarded-For` (`82-85`). Over the limit, the
server answers 429 with `Retry-After`; the shim turns any non-2xx into a
JSON-RPC error that says only `Server returned HTTP 429`
(`src/install/m365-mcp-shim.js:652-655`), without the header and without
waiting.

**Why it fails at scale.** A large organization reaches the internet through a
small set of addresses: a proxy, a VPN concentrator, a secure web gateway,
virtual desktop hosts. Thousands of users share one bucket. The budget is 20
calls per second per address per replica, about 60 per second across three
replicas if the ingress spreads them evenly. Tier B behind one proxy uses
about half of it in the peak hour; tier C, or any burst, exceeds it, and
everyone behind that address is refused for the rest of the minute. Because
the shim strips `Retry-After`, agents retry immediately and keep the bucket
full.

**Likelihood and blast radius.** High behind shared egress. Blast radius:
every user behind that address.

**Fix.** Limit authenticated `/api/mcp` calls per session (hash of the token,
after authentication) with a per-user budget, keep the address limiter for the
unauthenticated routes and as a high coarse ceiling, return
`retryAfterSeconds` in the JSON-RPC error `data`, and have the shim wait and
retry once. The keying and the default are [D5](#decisions). Size: small.

### F10 Graph retries have no deadline and 429s are opaque

**What the code does.** `createGraphClient` uses the SDK's default middleware
(`src/services/graphClient.ts:375-384`), whose `RetryHandler` retries 429, 503
and 504 three times, sleeping `Retry-After` when present and otherwise 3
seconds with exponential backoff, capped at 180 seconds per attempt
(`@microsoft/microsoft-graph-client` 3.0.7, `RetryHandler.getDelay`). It adds
no jitter when `Retry-After` is present, and an unparseable or past HTTP-date
value yields a `NaN` or negative delay, which is an immediate retry. When the
retries are exhausted the `GraphError` is caught and returned as
`Error: <message>` text (`mcpEndpoint.ts:2886-2922`); the status code and the
`Retry-After` value are dropped. Several tools make four to six sequential
Graph calls per invocation (`createReplyDraft` `379-457`; `move_event` copies
attachments one at a time, `1941-1963`).

**Why it fails at scale.** Throttling is normal at scale, and the first time it
happens three things go wrong at once. A single tool call can sleep for up to
nine minutes inside the handler while the ingress has long since returned 504,
so the client retries a call that is still running. Every throttled request
for a mailbox or tenant wakes at the same instant, so the retry wave is as
large as the original. And the agent cannot tell a throttle from a bug: it sees
"Error: Too Many Requests" with no wait hint and tries again.

**Likelihood and blast radius.** Medium at tier B, high at tier C. Blast
radius: a user or a mailbox, amplified by the retries.

**Fix.**

- Supply `RetryHandlerOptions` with a `shouldRetry` that enforces a total
  budget per tool call (for example, stop retrying once 60 seconds have been
  spent, and never sleep more than 30 seconds in one go), sanitize
  `Retry-After` to a bounded non-negative number, and add jitter.
- Map `GraphError` 429, 503 and 504 to a JSON-RPC error with
  `data: { retryable: true, retryAfterSeconds, graphStatus }`, and teach the
  shim to surface it unchanged so a client can wait.
- Gate concurrency per user: a semaphore of four per `homeAccountId` per
  replica, so an agent's fan-out queues locally instead of being throttled by
  Exchange (see [F18](#f18-one-mailbox-four-concurrent-requests)).

Size: medium.

### F11 Offset paging re-walks from page one

**What the code does.** `list_folders` and `list_sharepoint_list_items` take a
numeric `offset` and reach it by walking Graph's `nextLink` pages from the
start on every call, 200 items per page, up to 50 pages
(`src/services/graphPaging.ts:34`, `45-73`; page size at
`src/functions/mcp/mcpEndpoint.ts:234` and
`src/functions/sharepoint/listListItems.ts:11`). When the page cap stops the
walk before `offset`, the result is zero items with `hasMore: true` and
`nextOffset` equal to the offset that was asked for (`69-72`).

**Why it fails at scale.** Reaching item 8,000 costs 41 Graph requests, and
paging through a 10,000 item library in 200 item steps costs about 1,275
requests in total, each a multi-item read at 2 SharePoint resource units. A
tenant under 1,000 licenses has 1,250 units per minute per app; one user
browsing one large library can spend it. Items past 10,000 are unreachable,
and a client that pages until `hasMore` is false never stops, because the
cursor it is handed back is the one it sent. Lists and libraries in the tens
of thousands are ordinary in a large tenant.

**Likelihood and blast radius.** High wherever libraries are large. Blast
radius: SharePoint listing for that user, and the tenant's SharePoint resource
budget for everyone.

**Fix.** Replace the numeric offset with an opaque continuation token that
carries Graph's `nextLink` (HMAC-signed with the session key, or stored
server-side keyed by session and path, so the caller still never supplies a
URL), report `pageCapReached` honestly, and offer the same continuation
argument to the first-page tools in
[F17](#f17-list-tools-that-stop-at-one-page). The contract change is
[D4](#decisions). Size: medium.

### F12 SharePoint search filters a tenant-wide window

**What the code does.** `search_sharepoint` asks Microsoft Search for up to 200
tenant-wide hits and then removes those outside the allow-list and the deny
list (`src/functions/mcp/mcpEndpoint.ts:590-621`;
`src/services/sharepointSearch.ts:77-105`, `191-199`). The window is widened
when a filter is active (`102-105`), but it is still a window over the whole
tenant.

**Why it fails at scale.** With thousands of sites and an allow-list of a few,
the 200 most relevant hits for a common query can all sit outside the list.
The tool returns nothing, with `truncated: true`, and no way to look further.
The more sites a tenant has, the more often this happens. A separate ticket
already tracks search returning nothing on some tenants; scoping at the source
addresses both.

**Likelihood and blast radius.** High with an allow-list in a large tenant.
Blast radius: SharePoint search.

**Fix.** Scope the query at the source: build a KQL `path:` restriction from
the allow-listed sites' URLs (store `webUrl` on the allow-list entry), or, for
a short list, run `/sites/{id}/drive/root/search(q=...)` per site. Keep the
post-filters as defense in depth. Size: medium.

### F13 The internal Teams policy enumerates members

**What the code does.** Under the `internal` outbound mode every Teams send
lists the chat's or channel's members, up to 20 pages, and refuses the send if
the list is longer (`src/services/outboundPolicy.ts:409-450`; called per send
at `mcpEndpoint.ts:2567-2570` and `2602-2609`). Nothing is cached between
sends.

**Why it fails at scale.** Org-wide and department teams have thousands of
members. Graph pages member collections at a few hundred per page, so a large
channel exceeds the budget and every send to it is refused with "member list
too long to check", and every attempt first spends 20 Graph calls finding that
out. Smaller channels cost a page or more of member reads per message.

**Likelihood and blast radius.** High for organizations that turn the policy
on. Blast radius: Teams sends to large channels, and Teams request volume.

**Fix.** Ask for `$top=999` and `$select=email,tenantId` where the endpoint
supports it, cache the membership verdict per chat or channel for a few
minutes, and consider deciding by team composition instead of enumeration: a
channel of a team with no guest members (`/teams/{id}/members` filtered to the
guest role with `$count`) and standard membership type is internal. That is a
semantic choice, [D6](#decisions). Size: small to medium.

### F14 The audit table grows faster than it is purged

**What the code does.** Every tool call writes one audit row, fire-and-forget,
with no bound on the number of pending writes (`src/services/auditLog.ts:93-127`).
Rows for a tenant share one partition (`105-106`). The retention purge runs at
most once a day per process, deletes at most 50,000 rows per run, and finds
them with a `Timestamp lt` filter, which Table Storage can only satisfy by
scanning the table (`139-145`, `201-233`). The admin query scans the tenant
partition newest-first and caps the scan only when an operation filter is set
(`273-361`, `286-289`); a `userEmail` filter for a quiet user walks the whole
partition (`src/functions/admin/getAuditLog.ts:27-36`).

**Why it fails at scale.** Tier B writes a few hundred thousand rows a day.
The purge removes 50,000 per day per process and gets slower as the table
grows, so the table grows without bound: tens of millions of rows within a
year, in one partition. The admin audit page then scans millions of rows for
a filtered query and times out at the ingress. Under storage throttling the
pending audit writes accumulate in memory, each holding its row.

**Likelihood and blast radius.** Certain over months at tier B. Blast radius:
the admin audit page, storage cost, and replica memory during throttling.

**Fix.** Bucket partitions by day (`<tenant>|<yyyymmdd>`) so the purge deletes
whole partitions by range and date-bounded queries stay within a few
partitions; or purge by the reverse RowKey range per partition. Cap the rows
scanned for every filter and return `scanned` and `complete` with the result.
Put a bounded queue with a drop counter in front of the writes. Whether the
storage copy should stay a 365 day record at all, or become a 30 day window
with Log Analytics as the long-term record, is [D3](#decisions). Size: medium.

### F15 The tenant user picker enumerates the directory

**What the code does.** `/api/manage/tenant-users` lists every user in the
tenant, 999 per page, up to 50 pages, caches the list per tenant for five
minutes, and returns the whole filtered list in one response
(`src/functions/admin/getTenantUsers.ts:7-9`, `35-72`, `93-131`), alongside a
full session listing ([F6](#f6-session-rows-are-scanned-decrypted-and-never-purged)).
The page cap is reported only to the console (`60`). The admin page loads it
all (`src/admin/components/UserManagement.tsx:57`).

**Why it fails at scale.** At 10,000 users this is ten directory pages, a few
megabytes of JSON and a table of ten thousand rows in the browser: slow but
working. Above 50,000 users the list is silently cut at the cap and the
administrator cannot see the users that were dropped.

**Likelihood and blast radius.** Medium at tier B, certain at 50,000. Blast
radius: user management.

**Fix.** Search instead of enumerate: `/users?$search="displayName:..."` or a
`startswith` filter with `$top=25` and server-side paging, and surface the page
cap in the response if enumeration is kept for small tenants. Size: small.

### F16 In-memory caches grow for the life of the process

**What the code does.** The session cache holds every session a replica has
seen, with its access token, until logout or revocation
(`src/services/tokenCache.ts:147-149`), and `listActiveSessions` pins every row
it lists (`474-484`). The resolver caches hold every mail folder, calendar,
contact and section ID ever resolved, with no bound and no expiry
(`src/services/containerResolver.ts:18-23`, `155`). The public MSAL
application used for device-code sign-in has no cache plugin, so every account
that signs in that way stays in its in-memory cache, refresh token included
(`src/services/graphClient.ts:227-243`, `src/functions/auth/deviceLogin.ts:42-62`).
The per-account application map is bounded at 1,000 (`graphClient.ts:170`) and
measured at 1.4 KiB per application before its token cache is loaded, so it is
fine.

**Why it fails at scale.** None of these is large per entry; all of them grow
with distinct users and distinct objects, and a replica lives for weeks. Tens
of thousands of cached sessions at a few kilobytes each is a meaningful share
of 1 GiB, on top of [F2](#f2-file-and-attachment-reads-are-unbounded-in-memory).
The resolver caches also never expire, so a mail folder renamed into a denied
name stays reachable on warm replicas, and one renamed out of a denied name
stays blocked, until the process restarts.

**Likelihood and blast radius.** Medium. Blast radius: replica memory over
time; deny-list drift after renames.

**Fix.** Bound the session cache by last use (a few thousand entries, idle
expiry of an hour) and stop filling it from listings; give the resolver caches
a size bound and a 10 to 15 minute TTL; remove device-code accounts from the
public application once the flow completes (the missing cache plugin for that
flow is tracked separately). Size: small.

### P3

### F17 List tools that stop at one page

These tools read one Graph page, flag `truncated` honestly through
`src/services/resultEnvelope.ts`, and offer no way to continue:
`list_folders_mail` (`mcpEndpoint.ts:804`, 100), `list_sharepoint_lists`
(`696`, 100), `list_calendars` (`1555`, 100), `list_events` (`1633-1656`, up to
100), `list_contact_folders` (`2335`, 100), `search_contacts` (`2235`, up to
50), `list_notebooks` (`2422`, 100), `list_sections` (`2448`, 100),
`list_teams` (`2509`, 100), `list_rooms` (`2198`, up to 200) and the
attachment list in `get_attachments` (`1053`). In a large organization users
belong to more than a hundred teams, buildings have more than two hundred
rooms, and busy calendars have more than a hundred events in a week. The
continuation token from [F11](#f11-offset-paging-re-walks-from-page-one) is the
fix for all of them at once.

### F18 One mailbox four concurrent requests

Exchange allows four concurrent requests per app per mailbox. Each handler is
sequential, but agents issue tool calls in parallel, and a user may run more
than one agent. Five parallel reads of one mailbox are throttled as a group,
and the server does nothing to smooth that out. The per-user semaphore in
[F10](#f10-graph-retries-have-no-deadline-and-429s-are-opaque) is the fix.

### F19 Outbound calls have no timeouts and readiness ignores storage init

No Graph or Table call carries a timeout. Node's fetch waits up to 300 seconds
for headers by default, and the Table SDK retries three times with backoff, so
a storage outage turns each of a call's hundred operations into a stall. The
first request to a new replica waits on `initPromise`
(`src/services/tableStorage.ts:23-32`, `46-48`) while `/health` already
answers 200. A replica whose event loop is blocked for 15 seconds is killed by
the liveness probe (`infra/probes.json:19-24`). The deadlines in
[F3](#f3-the-runtime-envelope-stops-at-three-small-replicas) and the deep
health check in [F8](#f8-nothing-reports-degradation) cover this.

### F20 Device-code flows poll Entra for fifteen minutes each

Each `/api/auth/device` call starts a background poll against Entra for up to
fifteen minutes (`src/functions/auth/deviceLogin.ts:51-62`). The route is
limited to 10 per minute per address, not per instance, so a rollout that
sends thousands of users through device code keeps thousands of polls running
on the replicas that started them. With the per-account cache plugin this
flow lacks (tracked separately), the tokens would at least not accumulate in
memory. Note only.

### F21 Per-account MSAL applications and refresh races

Token refresh is cheap at scale. Each account has its own application, kept
in a 1,000 entry LRU (`src/services/graphClient.ts:170-197`), measured at 1.4
KiB each plus its loaded cache. A silent acquisition loads the account's row
from storage twice (once for `getAllAccounts`, once when MSAL overwrites its
in-memory cache before the lookup) and writes it back when it changes, but
only when the access token is within five minutes of expiry
(`src/services/tokenCache.ts:386-404`), so about once an hour per active
session: roughly three per second at tier B, on the `account` partition, well
within its limit. Two replicas refreshing the same account at once leave the
loser with a token that is still valid. The per-account lock (`200-209`)
serializes refreshes within a replica. No action beyond metrics.

### F22 Text filters over very large folders

A folder-scoped text search with no address criteria uses a `contains()` filter
over subject and sender (`src/services/mailSearch.ts:391-416`). Exchange
cannot index that; over a folder of a few hundred thousand messages it is slow
and sometimes times out, which the SDK retries as a 504. The newest-first scan
is bounded at 1,000 messages and ten pages (`100-107`) and reports its
horizon, so it is the safe path. Route text-only folder searches through the
scan when `since` is given, and say so in the tool description. Note only.

## Measuring it

None of this has been measured against a real large tenant, and it should be
before the fixes are declared done. A harness that runs the compiled server
against a fake Graph (an `undici` `MockAgent` answering the paths the tools
use, with configurable latency and 429 bursts) and a real Table Storage
account (or the emulator) can report the numbers that matter:

- Table Storage operations per tool call, per tool, before and after
  [F1](#f1-every-tool-call-fans-out-to-table-storage). The counts in
  [Request path and per-call cost](#request-path-and-per-call-cost) are the
  baseline.
- p50 and p95 tool duration at 30 and 170 calls per second with a realistic
  mix (half reads, a quarter searches, a quarter writes).
- Memory per replica during a 200 MB `read_file` and during ten concurrent
  25 MB attachments.
- Time to first response on a cold replica with 50,000 session rows.
- Behavior under a sustained 429 burst from the fake Graph: retry counts, wall
  time per call, and what the client sees.
- Replica count and 5xx rate at each tier against the templates' defaults.

Acceptance per tier belongs with [D1](#decisions): for tier B, fewer than ten
storage operations per tool call, p95 under three seconds for reads, no
replica above 70% memory with two large reads in flight, and every throttle
visible as a metric within a minute.

## Proposals

The build items are the fixes above, one ticket each, in the order they should
land. The decisions are the choices an operator has to make before some of
them can be finished.

### Build items

| ID | Title | Section | Summary |
|---|---|---|---|
| B1 | Cache policy reads in process and load the deny list once per call | [F1](#f1-every-tool-call-fans-out-to-table-storage) | Short-TTL bounded caches for settings, allow-list, overrides and deny lists; one load per handler; no `createTable` on reads; count fail-closed events. |
| B2 | Cap and stream file and attachment reads | [F2](#f2-file-and-attachment-reads-are-unbounded-in-memory) | Size gate from metadata, configurable caps, running-total abort, no base64 round trips for drive attachments, no pretty-printing of large results. |
| B3 | Raise the runtime envelope and give every call a deadline | [F3](#f3-the-runtime-envelope-stops-at-three-small-replicas) | Template defaults for a large tier, `AbortSignal` budgets per call, per Graph call and per Table operation, documented ingress timeout. |
| B4 | List allow-listed sites directly and page site search | [F4](#f4-site-listing-reads-one-page-and-filters-afterwards) | Enumerate allow-list entries by ID; `q` plus continuation for `list_sites`; page the admin picker route. |
| B5 | Page the Global Administrator check and cache it | [F5](#f5-the-global-administrator-check-reads-one-page) | Directory-role cast with `$top=999` and `nextLink`; per-session cache for a few minutes. |
| B6 | Index, project, paginate and purge the session table | [F6](#f6-session-rows-are-scanned-decrypted-and-never-purged) | User-keyed index rows, no decrypt on listings, paginated sessions API, purge past the absolute lifetime, run the legacy migration once. |
| B7 | Route the outbound policy through the shared storage client | [F7](#f7-the-outbound-policy-reads-storage-through-the-wrong-client) | Use `storageClient.getTableClient`; add a test under the endpoint configuration. |
| B8 | Deep health check, metrics and alert rules | [F8](#f8-nothing-reports-degradation) | `/health/deep` with cached dependency checks, custom metrics for Graph, storage and tools, one structured line per call, unsampled warnings, runbook alerts and queries. |
| B9 | Rate-limit authenticated MCP calls per session and pass `Retry-After` through | [F9](#f9-the-rate-limiter-keys-on-the-egress-address) | Per-session budget after authentication, address limiter kept as a coarse ceiling, `retryAfterSeconds` in error data, shim waits once. |
| B10 | Bound Graph retries, surface throttling, gate per-user concurrency | [F10](#f10-graph-retries-have-no-deadline-and-429s-are-opaque) | Custom retry options with a total budget and jitter, sanitized `Retry-After`, structured retryable errors, a semaphore of four per account. |
| B11 | Opaque continuation tokens for the list tools | [F11](#f11-offset-paging-re-walks-from-page-one) | Signed or server-held `nextLink` cursors replacing numeric offsets; `pageCapReached`; the same argument for the first-page tools in F17. |
| B12 | Scope SharePoint search to the allow-list at the source | [F12](#f12-sharepoint-search-filters-a-tenant-wide-window) | KQL `path:` restrictions from allow-listed site URLs, or per-site drive search for short lists; keep the post-filters. |
| B13 | Make the internal Teams check scale to large channels | [F13](#f13-the-internal-teams-policy-enumerates-members) | `$top=999` and `$select`, cached membership verdicts, or a team-composition decision. |
| B14 | Day-bucketed audit partitions, bounded scans, bounded writes | [F14](#f14-the-audit-table-grows-faster-than-it-is-purged) | Partition by day for range purges and bounded queries, scan caps for every filter, a bounded write queue with a drop counter. |
| B15 | Search the directory instead of enumerating it | [F15](#f15-the-tenant-user-picker-enumerates-the-directory) | `$search` or `startswith` with `$top=25` and paging; report the cap when enumeration is kept. |
| B16 | Bound the session, resolver and device-code caches | [F16](#f16-in-memory-caches-grow-for-the-life-of-the-process) | LRU with idle expiry for sessions, size and TTL for resolvers, no caching from listings, device-code accounts removed after the flow. |
| B17 | Load-test harness against a fake Graph and real storage | [Measuring it](#measuring-it) | Reproducible runs for the tier B and C scenarios, reporting operations per call, latency, memory, retries and 5xx rate. |

B1, B2, B3, B7 and B8 come first: they are the ones that take an instance down
or blind the operator, and B1 changes the numbers every other measurement
depends on. B4 and B5 follow because they break the SharePoint and admin
surfaces outright in a large tenant. B17 should land alongside B1 so the
before-and-after is recorded.

### Decisions

| ID | Question | Section |
|---|---|---|
| D1 | Which tier the templates and defaults target: replica size and count, scale trigger, file and attachment caps, and the acceptance numbers the load tests must meet. | [F3](#f3-the-runtime-envelope-stops-at-three-small-replicas), [F2](#f2-file-and-attachment-reads-are-unbounded-in-memory), [Measuring it](#measuring-it) |
| D2 | How long a policy change (deny list, settings, allow-list, enforced modes) may take to apply across replicas once reads are cached; 30 to 60 seconds is proposed, and it is a security trade. | [F1](#f1-every-tool-call-fans-out-to-table-storage) |
| D3 | Whether the storage copy of the audit log stays a 365 day record (and gets the partitioning to support that) or becomes a 30 day window with Log Analytics as the long-term record the admin screen also reads. | [F14](#f14-the-audit-table-grows-faster-than-it-is-purged) |
| D4 | Whether the list tools move from numeric `offset` to opaque continuation tokens, which changes a tool contract that clients and prompts already use. | [F11](#f11-offset-paging-re-walks-from-page-one) |
| D5 | How authenticated MCP calls are rate-limited: per session or per address, and the default budget for a large organization behind shared egress. | [F9](#f9-the-rate-limiter-keys-on-the-egress-address) |
| D6 | Whether the internal-only Teams policy keeps enumerating members (exact, expensive, refuses large channels) or decides by team composition (cheap, approximate). | [F13](#f13-the-internal-teams-policy-enumerates-members) |

## Sources

- Microsoft Graph service-specific throttling limits:
  <https://learn.microsoft.com/en-us/graph/throttling-limits>
- Avoid getting throttled or blocked in SharePoint Online (resource units,
  per-app per-tenant limits, delegated search limit):
  <https://learn.microsoft.com/en-us/sharepoint/dev/general-development/how-to-avoid-getting-throttled-or-blocked-in-sharepoint-online>
- Azure Container Apps ingress (request timeout):
  <https://learn.microsoft.com/en-us/azure/container-apps/ingress-overview>
- Azure Table Storage scalability targets:
  <https://learn.microsoft.com/en-us/azure/storage/tables/scalability-targets>
- Microsoft Graph JavaScript client retry handler (`RetryHandler`,
  `RetryHandlerOptions`), version 3.0.7, read from `node_modules`.
