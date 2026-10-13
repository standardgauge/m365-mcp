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

A Global Administrator sets a tenant-wide **key policy** that no key can
loosen: keys on or off, services no key may use, services every key is
read-only on. Both the key's policy and the tenant's are read on every
request, so tightening either applies to existing keys at once.

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

1. A **client session** in `mcpSessions`, exactly as the install flow mints
   one for an installed client: same user, same `homeAccountId`, a fresh
   random session token nobody is given, `kind: client`, and the browser
   session's immutable absolute-lifetime anchor, so the key can never outlive
   the sign-in that created it (30 days). Its device label is
   `agent-key:<label>#<keyId>`.
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

## What ends a key

| Event | Effect |
|---|---|
| Owner or admin revokes it | Row marked, session row deleted. Other replicas drop their cached session within `SESSION_REVALIDATE_MS` (30 s), the same bound as logout. |
| Its own expiry (`expiresInDays`, at most 365) | Refused at the next request. |
| The owner signs out, or their sessions are purged | `deleteAllUserSessions` removes the key's session too. Refused with "the session the key is bound to no longer exists". |
| 30 days from the owner's sign-in | The inherited absolute cap. The owner signs in again and mints a new key. |
| Entra revokes the owner's refresh token | The silent renewal fails at the next renewal window and the key's session ends. |
| Admin turns keys off for the tenant | Every key refused; existing rows kept so turning them back on restores them. |

Revocation and refusals are audited as `auth.agent_key` denied, with the key
id as the resource.

## What a key can never do

Enforced in code and covered by `agentKeyEnforcement.test.ts` and
`authAgentKey.test.ts`:

- Reach `/api/manage/*`, including the key routes. Those take only a browser
  console session (threat model 7.3), so a key cannot mint keys, list them,
  or read policy.
- Change email output mode (`POST /api/mail/settings` refuses a key and audits
  it). Loosening draft mode is the one thing an injected prompt would want a
  key for.
- Sign the owner out (`POST /api/auth/logout` refuses a key).
- Refresh (`/api/auth/refresh` takes session tokens only; a key is unknown to
  it).
- See a tool it may not call: `tools/list` is filtered by the key's policy and
  `tools/call` refuses as the backstop, both logged as denied.

## Routes

| Route | Who | Does |
|---|---|---|
| `GET /api/manage/agent-keys` | any console user | the caller's keys; `?all=true` or `?userId=` needs Global Administrator |
| `POST /api/manage/agent-keys` | any console user | mint: `{ label, services, readOnlyServices?, allowedSites?, expiresInDays? }` → `201 { key, keyId, … }`, the key once |
| `DELETE /api/manage/agent-keys?keyId=` | owner, or Global Administrator | revoke |
| `GET /api/manage/agent-key-policy` | any console user | the tenant's key policy, so the form can show what is grantable |
| `POST /api/manage/agent-key-policy` | Global Administrator | `{ allowKeys?, disabledServices?, readOnlyServices? }`, merged over the stored policy |

Every change is audited with before and after: `policy.agent_key.create`,
`policy.agent_key.revoke`, `policy.agent_key_policy.set`, and the admin
refusals under the same names.

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
- A key cannot yet carry its own folder, mail-folder or calendar deny
  entries; it inherits the owner's. Also in the gap register; tracked
  separately.
- Nothing yet alerts on a burst of refusals or a destructive call under a
  key; the audit rows are there to be consumed. Tracked separately.

## Tests

- `agentKeys.test.ts`: key shape, requested-policy validation, the intersection, site narrowing.
- `authAgentKey.test.ts`: the authentication path, every refusal, the cap, touch rate, the refresh path.
- `manageAgentKeys.test.ts`: both routes, console-only, minting binds a client session with the inherited anchor, admin boundaries, audit rows.
- `agentKeyEnforcement.test.ts`: REST checks, `tools/list` and `tools/call` under a key, and the routes a key must never reach.
