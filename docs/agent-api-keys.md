# Agent API keys

A signed-in user can mint an **agent key**: a bearer credential for an app or
agent they run that reaches this server as them, with less than they have.
This page is the design and the operating notes.

## The problem it solves

Before keys, the connector permissioned a user's MCP as one unit. A user could
hide a folder from AI, and an administrator could make a service read-only
for everyone, but nobody could say *the user may see this in Claude, the app
they wrote may not*, or *the user may write, the app may only read*. Both came
up the first time a client wanted to let staff build with AI against
production SharePoint while keeping what those builds could touch narrower
than the staff themselves.

The alternative that gets proposed in that situation is a second SharePoint
with a copy of the data for AI to use. Keys exist so nobody has to do that:
the data stays in one place and the narrowing happens at the credential.

## What a key is

A key is **not a new identity** and never an application permission. It
resolves to a delegated session of the user who minted it, so every Graph
call under it is made as that user and bounded by what they can reach.
Everything in the README's authentication model still holds: no service
account, no app-only Graph scope, no client-supplied user id.

On top of the user's and the tenant's policy, a key carries its own:

| Field | Meaning | Can only |
|---|---|---|
| `services` | The services the key may use | narrow: a subset of what the user can use |
| `readOnlyServices` | Services whose write tools the key may not call | narrow: adds to the tenant's read-only set |
| `allowedSites` | SharePoint sites the key may see | narrow: intersects the tenant's allow-list |
| deny entries | Folders, mail folders, calendars and the like the key may not reach | narrow: added to the owner's and the tenant's deny lists (see **A key's own deny entries**) |
| `expiresInDays` | How long the key lives, 1 to 365 | bound by the tenant's maximum |

A Global Administrator sets a tenant-wide **key policy** that no key can
loosen: keys on or off, services no key may use, services every key is
read-only on, the longest lifetime a key may be given
(`maxKeyLifetimeDays`, default 365), and whether owners may widen a key
after minting it (`ownersMayWiden`, default on; see **Editing a key**). The
policy is read on every request, so tightening it applies to existing keys
at once; a lowered maximum applies to keys minted from then on, since
existing keys keep the lifetime they were given.

The policy in force is the intersection, computed in
`effectiveAgentKeyPolicy` ([`agentKeys.ts`](../src/services/agentKeys.ts)):

```
services   = key.services − tenantKeyPolicy.disabledServices
readOnly   = (key.readOnlyServices ∪ tenantKeyPolicy.readOnlyServices) ∩ services
sites      = tenant allow-list empty  → key.allowedSites
             key.allowedSites empty   → tenant allow-list
             both set                 → intersection by site id
             intersection empty       → no site at all (never the tenant list)
```

Everything else the user is subject to still applies under the key: tenant
enabled services, the user's own opt-outs, mail-indexing disable, both deny
list tiers, draft mode and its enforcement, outbound policy.

## How it is bound to the user

Minting a key creates two rows.

1. A **client session** in `mcpSessions`, the way the install flow mints one
   for an installed client: same user, same `homeAccountId`, a fresh random
   session token nobody is given, `kind: client`, device label
   `agent-key:<label>#<keyId>`. Unlike an installer's session it is anchored
   at the key's creation, not at the owner's sign-in: its absolute cap is the
   key's own lifetime, so a one-year key lives a year. See **Lifetime** below
   for why that is safe and what it costs.
2. A **key row** in `agentKeys`, PartitionKey `key`, RowKey = HMAC(key)[:32],
   holding the full HMAC, the key's public id, owner, label, policy, expiry,
   and the RowKey of that session. The key value itself is stored nowhere and
   shown once.

On a request, `authenticateRequest` sees the `m365k_` prefix, resolves the key
row by one point read, compares the HMAC in constant time, refuses a revoked,
expired or tenant-disabled key, loads the bound session by RowKey, and runs
the same absolute-cap, renewal and sliding-window logic as a session token.
The result is a *copy* of the session carrying `agentKey` (id, label,
effective policy). The cached session and anything written back through
`storeSession` never carry it: `saveSession` writes named columns only.

Because the session is an ordinary client session with the key's name on it,
every audit row written under the key already names it in `deviceLabel`, on
the Table Storage row and in the Log Analytics stream, with no change to the
audit schema. The admin sessions list shows it the same way.

## Lifetime

Every key has an expiry, chosen by the owner at creation, from one day to
the tenant's maximum (365 by default). The key's session carries that as its
absolute cap in place of the 30 days a sign-in gets
(`authenticateSession(session, expiresAt − createdAt)`), so the two never
disagree and nothing else about session handling changes.

This is a deliberate loosening of the 30-day rule, for a credential that is
meant to run unattended. What makes it acceptable:

- The owner picks the lifetime and the tenant caps it. An administrator who
  wants keys no longer than a sign-in sets `maxKeyLifetimeDays` to 30.
- Revocation is immediate (30 s across replicas) and available to both the
  owner and the administrator, and the owner is told when expiry is near
  (below).
- The underlying Graph access still comes from the owner's refresh token.
  Entra ends that on sign-in revocation, password reset or its own inactivity
  rule (90 days unused by default), and the key's session ends with it at
  its next renewal. A key used at least every few days keeps the refresh
  token warm; a key unused for three months may find the owner has to sign
  in again, which is the right outcome.
- A session token still gets 30 days. Only a key, minted on purpose through
  the browser with a chosen expiry, can live longer.

### The reminder

With `reminder` set (the form's default) the server puts a 30-minute event
on the owner's own calendar before the key expires: seven days before for
keys of two weeks or more, two days before for keys of four days or more,
none for shorter keys. The event is at 09:00 in the mailbox's own time zone,
shown as free, with an Outlook reminder at its start, and its body names the
key, its id, what it may use, the exact expiry, and where to mint a
replacement. Nobody is invited, so nothing is sent and the outbound policy is
not involved. It is written with the owner's own delegated token
(`Calendars.ReadWrite`, already consented for the calendar tools) by
[`agentKeyReminder.ts`](../src/services/agentKeyReminder.ts).

The time of day comes from the mailbox's own zone (`/me/mailboxSettings`,
which needs `MailboxSettings.Read`). When that cannot be read, because the
scope is not consented or the setting is empty, the reminder is still
written, at 16:00 UTC, a working hour from the US West Coast through the
UK, with a line in the body saying why and asking for the permission, and
the creation is audited as `policy.agent_key.reminder` allowed with the
reason. The user-facing calendar tools keep their loud failure on a missing
zone on purpose; a best-effort reminder is better late in the day than
missing.

The reminder is otherwise best effort. If Graph refuses the write itself the
key is still created, the response says so, the attempt is audited as
`policy.agent_key.reminder` denied, and the owner is told to note the date
themselves. Revoking the key
removes the event, as the owner, through the key's own session, so it works
whether the owner or an administrator revokes.

## Editing a key

A key's label and policy can be changed in place, through
`PATCH /api/manage/agent-keys?keyId=`, the Edit button on the owner's table,
or the Narrow button on the administrator's. The key value and its expiry
never change; the app holding the key sees the new policy on its next call.

Who may do what:

| Actor | On | May |
|---|---|---|
| Owner | their own key | narrow or widen, within what they could grant today (their reach less the tenant key policy) |
| Owner, tenant `ownersMayWiden` off | their own key | narrow only; anything wider is refused with a message to mint a new key |
| Global Administrator | another user's key | narrow only: remove a service, make one read-only, remove a site |

"Narrow" is `isNarrowingOf`: no new service, no read-only flag cleared, and
a site list that is a subset, where an empty list means the tenant's whole
allow-list and so counts as the widest. Every edit is audited as
`policy.agent_key.update` with `reason` `narrowed` or `widened` and the full
policy before and after; refusals are audited too. The key row records
`policyUpdatedAt` and `policyUpdatedBy`, and both tables show "edited" with
the time.

Why editable in both directions by default, when a session token is not.
The two families of credential differ: a token that carries its own grant
(an OAuth access token, a SAS token) is immutable and widening means a new
grant, while a credential that names a principal whose permissions live in a
policy (an IAM access key, a GitHub fine-grained token) is edited in place
and audited. Agent keys are the second kind: the policy sits on the row and
is read on every request. Editing in place keeps the secret where it is, and
making keys cheap to mint and change is what stops people sharing one key
across several apps, which is the outcome that actually weakens the audit
trail. What the restriction protects, a key already deployed somewhere less
trusted quietly gaining power, is covered by the audit row, by the owner
being the one who widens, and by the interface saying in so many words that
a widening reaches wherever the key lives. A tenant that wants the stricter
model sets `ownersMayWiden` off, and administrators are narrow-only in every
case, since a key acts on its owner's data and widening is theirs to decide.

## A key's own deny entries

A key can be kept out of one folder, mail folder or calendar its owner can
still see: the owner's Claude may read the Board mail folder, the reporting
app holding a key may not. These are deny entries of the same kind as the
personal deny list, held on the key.

**Storage.** They are ordinary tier 2 rows in `UserDenyList` under a
synthetic subject, `key:<keyId>` (`agentKeyDenySubject` in
[`denyList.ts`](../src/services/denyList.ts)), so they match exactly as the
owner's entries do: same canonicalisation, same case-insensitive prefix rule,
names as well as paths, the ancestor-segment rule for search hits. The
subject cannot collide with a user's, which is an Entra object id.

**Enforcement.** When `authenticateAgentKey` accepts a key it records the
key id on the request scope (`markRequestAgentKey` in
[`clientAddress.ts`](../src/services/clientAddress.ts), the same
`AsyncLocalStorage` that carries the audit address), and `loadDeniedPaths`
adds `key:<keyId>` to the subjects of every check made while that request is
handled. Every caller of `isPathDenied`, `checkDenyList`,
`filterDeniedPaths` and `filterDeniedSearchHits`, on REST and on MCP, inherits
the key's entries without being changed or knowing keys exist. Delegated
access still adds the data owner's subject alongside, so a key reading a
shared mailbox is held to its own entries, its owner's and the mailbox
owner's.

The scope is set by `withSecurity`, which wraps every route. A key presented
somewhere without a request scope would have its own entries silently
skipped, so it is refused there instead ("no request scope to carry the
key's deny entries").

**Who may change them.** Through `/api/manage/deny-list/user` with `keyId`
in place of `targetUserId`, under the same rules as **Editing a key**:

| Actor | May |
|---|---|
| Owner | add and remove |
| Owner, tenant `ownersMayWiden` off | add only; removing an entry widens the key |
| Global Administrator, another user's key | read and add; not remove |

A revoked key's entries cannot be changed, and revocation deletes them
(best effort; a revoked key's rows are inert either way). The key subject is
refused as a `targetUserId`, so the narrow-only rules cannot be stepped round
through the personal-list path. Changes are audited as
`policy.deny_list.agent_key.add` and `.remove`, with the resource
`key:<keyId>/<type>:<path>` and the entry before and after.

**UI.** **My Settings → Agent Keys → Folders** on a key opens its list and
the same folder browsers the personal deny list uses, limited to the key's
services and, for SharePoint, to the key's sites. Block there writes to the
key, not to the owner.

## What ends a key

| Event | Effect |
|---|---|
| Owner or admin revokes it | Row marked, calendar reminder removed, session row deleted, its own deny entries removed. Other replicas drop their cached session within `SESSION_REVALIDATE_MS` (30 s), the same bound as logout. |
| Its own expiry (`expiresInDays`, 1 to the tenant maximum, at most 365) | Refused at the next request. The key's session is capped by the same instant. |
| The owner signs out, or their sessions are purged | `deleteAllUserSessions` removes the key's session too. Refused with "the session the key is bound to no longer exists". |
| Entra revokes the owner's refresh token, or it lapses from inactivity | The silent renewal fails at the next renewal window and the key's session ends. |
| Admin turns keys off for the tenant | Every key refused; existing rows kept so turning them back on restores them. |

Revocation and refusals are audited as `auth.agent_key` denied, with the key
id as the resource.

## What a key can never do

Enforced in code and covered by `agentKeyEnforcement.test.ts` and
`authAgentKey.test.ts`:

- Reach `/api/manage/*`, including the key routes. Those take only a browser
  console session (threat model 7.3), so a key cannot mint keys, list them,
  or read policy.
- Change email output mode, on either surface. `POST /api/mail/settings`
  refuses a key and audits it, and the MCP tool `set_email_output_mode` is
  hidden from a key's `tools/list` and refused at `tools/call` even when the
  key holds mail (`KEY_REFUSED_TOOLS` in `mcpEndpoint.ts`). Loosening draft
  mode is the one thing an injected prompt would want a key for, and a key
  has no business tightening its owner's setting either.
- Sign the owner out (`POST /api/auth/logout` refuses a key).
- Refresh (`/api/auth/refresh` takes session tokens only; a key is unknown to
  it).
- See a tool it may not call: `tools/list` is filtered by the key's policy and
  `tools/call` refuses as the backstop, both logged as denied.
- Write through REST on a service it is read-only on. The policy wrapper treats
  every method other than GET, HEAD and OPTIONS as mutating whether or not the
  route sets `mutating`, so a write route that forgets the flag still fails
  closed. The same gate now covers the tenant's own read-only setting on REST.
- Use the SharePoint admin views. `GET /api/sharepoint/sites?admin=true` gives
  a key the same narrowed list as without the flag, and
  `GET /api/sharepoint/resolve-site` refuses a key, even when the owner is a
  Global Administrator.

## Routes

| Route | Who | Does |
|---|---|---|
| `GET /api/manage/agent-keys` | any console user | the caller's keys; `?all=true` or `?userId=` needs Global Administrator |
| `POST /api/manage/agent-keys` | any console user | mint: `{ label, services, expiresInDays, readOnlyServices?, allowedSites?, reminder? }` → `201 { key, keyId, reminder, … }`, the key once |
| `PATCH /api/manage/agent-keys?keyId=` | owner (either way, or narrow-only per tenant policy); Global Administrator (narrow only) | edit: `{ label?, services?, readOnlyServices?, allowedSites? }` → `200 { widened, … }` |
| `DELETE /api/manage/agent-keys?keyId=` | owner, or Global Administrator | revoke |
| `GET /api/manage/deny-list/user?type=&keyId=` | owner, or Global Administrator | the key's own deny entries of one type |
| `POST`, `DELETE /api/manage/deny-list/user` with `{ type, path, keyId }` | owner (remove only if owners may widen); Global Administrator (add only) | add or remove one of the key's own deny entries |
| `GET /api/manage/agent-key-policy` | any console user | the tenant's key policy, so the form can show what is grantable |
| `POST /api/manage/agent-key-policy` | Global Administrator | `{ allowKeys?, disabledServices?, readOnlyServices?, maxKeyLifetimeDays?, ownersMayWiden? }`, merged over the stored policy |

Every change is audited with before and after: `policy.agent_key.create`,
`policy.agent_key.update`, `policy.agent_key.revoke`,
`policy.agent_key_policy.set`, `policy.deny_list.agent_key.add` and
`.remove`, and the admin refusals under the same names.

The UI is **My Settings → Agent Keys** for owners and **Admin → Agent Keys**
for the tenant policy and the list of every key.

## Using a key

Send it where a session token goes:

```
Authorization: Bearer m365k_…
```

against `/api/mcp` or the REST routes, or hand it to the stdio shim with
`--store-token` as you would a session token. Keep it in a secret store. A
key in a repository or a config file is a credential leak like any other, and
the remedy is revocation.

## First-party use: the post-deploy smoke

The first consumer of keys is this repository's own deploy workflow. After a
deploy, its **Functional smoke** job reads a key from the fork's
`SMOKE_AGENT_KEY` secret and drives the live instance with it
(`scripts/post-deploy-smoke.js`; checks and failure meanings in the
[operations runbook](operations-runbook.md#functional-smoke)). A key is the
right credential for it because it is everything an unattended probe needs and
nothing more: it acts as a real person, so every read is a real delegated Graph
call that a missing consent or a broken client fails exactly as a user's would;
it can be narrowed to read-only, so the probe can prove the write gates refuse
without being able to write; and it reaches neither the admin API nor its
owner's settings, which the probe also proves.

How the smoke key is minted, and why each choice:

- **By a Global Administrator, as themselves, under My Settings → Agent Keys.**
  The key can only reach what its owner can, so the owner's reach is the
  probe's reach. An administrator's mailbox, calendar and drive exist on every
  tenant; a pilot user's may not.
- **Every service the tenant enables, every one of them read-only.** The probe
  reads each service it can see and expects every write refused. A service left
  off the key goes untested; a service left writable fails the run at "no write
  tool listed" and the write probes are skipped. A service the tenant has not
  enabled is fine to leave off: the probe calls it to prove it is refused.
- **No site restriction**, unless the tenant's own allow-list is the point:
  `list_sites` is a read, and an empty key list means the tenant's.
- **365 days, with the reminder on.** The reminder puts the replacement date on
  the owner's calendar a week ahead. Rotating is: mint the new key, `gh secret
  set SMOKE_AGENT_KEY`, dispatch a deploy to prove it, revoke the old one.
- **Label it for what it is** (`deploy-smoke`), so its audit rows read as the
  pipeline's and not as an app the administrator forgot about.

The run leaves a short burst of audit rows under that label on every deploy:
the reads `allowed`, and the write, service, settings and logout refusals
`denied`. That burst is expected, and its absence after a deploy is worth
noticing.

## Handling keys

For owners minting keys and the operators who answer for them.

### One key per app

Mint a key for each app or agent, and never hand one key to two of them. The
audit trail names the key, not the program holding it, so two apps on one key
read as one in every audit row, and revoking it to stop one breaks the other.
Separate keys also let each app be narrowed to what it needs, and keep the
`deploy-smoke` key's expected burst of refusals out of any other key's history.
Keys are cheap to mint and to edit for exactly this reason (see **Editing a
key**). Label each one for the app that holds it.

### How long a key should live

The tenant maximum is a ceiling, not a recommendation. Pick the shortest
lifetime that does not make rotation a chore, and leave the reminder on.

| Use | Longest lifetime | Why |
|---|---|---|
| A one-off test, a demo, a prompt being iterated on | 1 to 7 days | Nobody needs it next week, and an expired key needs no clean-up |
| An app in development on someone's own machine | 30 days | The same cap as a sign-in. Conditional Access is not evaluated when a key is used (see **Operating notes**), so a copy on a laptop is the case it does not cover |
| An unattended job on a server, key in a secret store (a nightly report) | 90 days | Rotated quarterly on the reminder. Long enough that rotation is a routine, short enough that a copy that leaked unnoticed dies on its own |
| The post-deploy smoke key, read-only, held only as a pipeline secret | 365 days | Read-only on every service, so a leaked copy can read but not write. See **First-party use** above |
| A key that can write mail or files, held anywhere but a secret store | 7 days | Write reach outside a secret store is the combination to keep short |

To make these limits rather than advice, set the tenant's
`maxKeyLifetimeDays` (**Admin → Agent Keys**) to the longest row you accept,
for example 90. The maximum applies to every key, the smoke key included: below
365, the smoke key is minted for the maximum and rotated on its reminder.

### How fast a revocation takes effect

The key row is read from storage on every request, without a cache, and
revoking marks that row before anything else. A request that arrives after the
revoke returns is refused on every replica, and logged as `auth.agent_key`
denied with reason `key revoked`. A request already in flight finishes.

Plan around **30 seconds** all the same. That is `SESSION_REVALIDATE_MS`, the
bound on the key's cached session, which revocation also deletes. It is the
documented guarantee if the first check ever regresses, and the same bound as
logout. If storage is unreachable, a key request is refused with 503 rather
than served from cache, so an outage does not keep a revoked key alive.

### If a key may have been disclosed

A key in a repository, a log, a chat or a ticket, or on a machine that was
lost, is disclosed. Revoke it first and investigate afterwards: a replacement
takes a minute, and the investigation does not need the key to stay live.

1. **Revoke** it under **My Settings → Agent Keys** (the owner) or **Admin →
   Agent Keys** (a Global Administrator).
2. **Review the audit log** for everything done under the key's device label,
   `agent-key:<label>#<keyId>`, over the key's whole life.
3. **Rotate**: mint a replacement for the app with the same or a narrower
   policy, put it in the app's secret store, and confirm the app's next calls
   appear under the new key id.

The queries, what to look for, and when one key becomes a wider incident are in
[incident response](incident-response.md#a-disclosed-agent-key).

### Customer-pilot smoke

Before a customer pilot, show on the customer's own tenant that each key
control holds on a live instance, not only in the unit suite. Four of the six
checks are the post-deploy **Functional smoke**, which runs on every deploy once
go-live item 10.5 is done; the other two are one-time manual checks with a
throwaway key.

| Check | How it is shown |
|---|---|
| Mint | Go-live [item 10.5](customer-go-live.md#10-first-user-first-call-sign-off): the administrator mints the smoke key through the browser |
| Allow/deny matrix | The [functional smoke](operations-runbook.md#functional-smoke): one real read per held service, and every service outside the key refused |
| Read-only | The functional smoke: no write tool listed, a write on every held service refused, and `POST /api/mail/settings`, the admin API and logout refused |
| Audit | The functional smoke's burst of `allowed` and `denied` rows under `agent-key:deploy-smoke#<keyId>`; confirm it in `M365McpAudit_CL` after the first run |
| Site restriction | Manual, below. The smoke key has no site restriction by design |
| Revoke | Manual, below. The smoke key has to stay live |

For the two manual checks, mint a key labelled `pilot-smoke`, SharePoint only,
read-only, restricted to one site, one day, reminder off. Then, with the key in
`KEY` and the id of a site outside the key in `OTHER_SITE` (the administrator's own
`list_sites` lists them):

```bash
mcp() {
  curl -sS https://your-mcp-host.example.com/api/mcp \
    -H "Authorization: Bearer $KEY" -H 'Content-Type: application/json' \
    -d "{\"jsonrpc\":\"2.0\",\"id\":1,\"method\":\"tools/call\",\"params\":$1}"
}
mcp '{"name":"list_sites","arguments":{}}'                       # only the one site
mcp "{\"name\":\"list_sharepoint_lists\",\"arguments\":{\"siteId\":\"$OTHER_SITE\"}}"   # "not in the allowed sites list"
```

Revoke the key in the browser, then repeat the first call at once. It answers
with a `Session expired` tool error instead of the site, and the audit log shows the site refusal (`list_sharepoint_lists`
denied, `Site not in allowedSites`) and `auth.agent_key` denied with `key
revoked`, both under `agent-key:pilot-smoke#<keyId>`. Record the four results
with the go-live record.

## Operating notes

- A key costs one extra Table Storage point read per request over a session
  token (the key row). Sessions are cached as before. `lastUsedAt` is written
  at most every five minutes per key.
- Keys are for an app or agent the user runs **as themselves**: a nightly
  report, a build they are testing, a second assistant with a tighter brief.
  A process that should not run as any person needs a service identity, which
  this server deliberately does not offer; see the threat model's
  assumptions.
- **Conditional Access is evaluated at the owner's sign-in, not at key use.**
  A key copied to a non-compliant machine works until it is revoked or
  expires. Tenants that rely on CA for "corporate machines only" should treat
  keys as server-side credentials and keep lifetimes short. In the threat
  model's gap register; tracked separately.
- Nothing yet alerts on a burst of refusals or a destructive call under a
  key; the audit rows are there to be consumed. Tracked separately.

## Tests

- `agentKeys.test.ts`: key shape, requested-policy validation, the intersection, site narrowing.
- `authAgentKey.test.ts`: the authentication path, every refusal, the cap, touch rate, the refresh path.
- `manageAgentKeys.test.ts`: both routes, console-only, minting binds a client session anchored at the key, the tenant ceiling, reminder success and failure, editing by owner and by administrator under both tenant settings, admin boundaries, audit rows.
- `agentKeyEnforcement.test.ts`: REST checks, the SharePoint REST write and admin routes end to end through the wrapper, `tools/list` and `tools/call` under a key, and the routes a key must never reach.
- `agentKeyReminder.test.ts`: lead time by lifetime, the event's shape, failure handling, deletion.
- `agentKeyDenyList.test.ts`: a key's own deny entries applied by `isPathDenied`, `filterDeniedPaths` and `filterDeniedSearchHits` under the key and not under the owner's session, another key, or no request scope; delegated subjects alongside; clearing on revoke.
- `agentKeyDenyListRoute.test.ts`: the `keyId` path of the deny-list route, owner and administrator rules under both tenant settings, revoked and unknown keys, and the key subject refused as a `targetUserId`.
