# M365 MCP Server — Operations Runbook

**Audience:** Operators managing a live M365 MCP Container App. The worked
example throughout is the `your-mcp-host.example.com` instance; every command takes
the same shape for any tenant with the resource names swapped. \
**Last updated:** October 2026


Standing a tenant up from nothing starts with [`entra-setup.md`](entra-setup.md).

* * *

## Service overview

Property| Value \
---|--- \
Live URL| https://your-mcp-host.example.com \
MCP endpoint| https://your-mcp-host.example.com/api/mcp \
Admin UI| https://your-mcp-host.example.com/ (site root — see Reserved paths) \
Container App| m365-mcp (rg-m365-mcp, Azure East US) \
ACR| somamcpacr.azurecr.io \
Storage Account| stm365mcp \
Entra app| your M365 MCP app registration (app ID from the Entra portal) \
Repo (deploy fork)| your-org/your-deploy-fork \
CI/CD| GitHub Actions — the deploy workflow (push to main) \

* * *

## Reserved paths

The Azure Functions host claims some URL paths for itself before any of our
functions see the request. Do not point user-facing config at them.

Path| Owner| Behaviour \
---|---|--- \
`/admin/*`| Functions host admin API| 401 with `WWW-Authenticate: Bearer`, or an empty 404. Master-key protected. \
`/runtime/*`| Functions host| Reserved by the runtime. \

`serveAdmin` is registered as a catch-all (`{*restOfPath}`) and serves the admin
SPA from the **site root**, with SPA fallback for any unmatched path. It would
happily serve `/admin` too, but the host intercepts that path first, so the
request never arrives.

Two consequences worth remembering:

  * The admin UI URL is the bare origin, e.g. `https://your-mcp-host.example.com/`. There is no `/admin` route and there cannot be one.
  * `FRONTEND_URL` / the `frontend-url` secret must be the origin. Setting it to `<host>/admin` ends every login on an empty 404 — which is exactly what shipped on all three tenants until. The Bicep templates still write that value on a fresh provision; set the secret to the root right after deploying them.

To tell an app-served 404 from a host-served one: every response from our code
carries the security headers (`content-security-policy`, `strict-transport-security`).
A bare 404 with `content-length: 0` and none of those headers came from the host.

* * *

## Monitoring

### Azure Monitor — Container App metrics

Portal path: **rg-m365-mcp → m365-mcp → Metrics**

  * **Requests** — request volume and HTTP status code distribution. Spike in 5xx indicates an application error.
  * **CPU utilization** and **Memory utilization** — normal idle: <10% CPU, <300 MiB. Sustained >80% CPU warrants investigation or scale-up.
  * **Replica count** — should stay at 1 during normal operation. Replicas will scale up to 3 under load (20 concurrent requests per replica threshold).
  * **Response latency** — P99 above 5 seconds during non-heavy-Graph calls indicates a problem.



### Health probe

The endpoint is `/health`, not `/api/health`. Check the body, not just the exit
code — `curl -f` alone is not a health check here.

```
# Quick health check from any machine
curl -fsS https://your-mcp-host.example.com/health | jq -e '.status == "ok"' >/dev/null \
  && echo "OK" || echo "FAIL"
# Expected: HTTP 200, application/json, {"status":"ok","sha":"<deployed commit>"}
```

Why the body assertion: until this runbook said `curl -f .../api/health`.
That route has never existed. The admin SPA is served by a catch-all that fell
back to `index.html` for any unmatched path, so the probe received 200
`text/html`, `curl -f` succeeded, and the documented check reported OK for
months — it could not have failed even with the API completely dead. Unmatched
`/api/*` now returns a real 404, but the habit is the point: assert on content.

`.sha` is the deployed commit, so the same call answers "is it up" and "is it
running what I think it is".

There is no external probe unless you add one. Run these same conditions from it.

### Session table check
```
# Count active sessions in Azure Table Storage
az storage entity query \
  --account-name stm365mcp \
  --table-name mcpSessions \
  --auth-mode login \
  --query "items | length(@)"
```

### GitHub Actions deploy status
```
gh run list -R your-org/your-deploy-fork --workflow the deploy workflow --limit 5
```

* * *

## Log access

### Live logs (streaming)
```
# Stream logs from the active revision
az containerapp logs show \
  --name m365-mcp \
  --resource-group rg-m365-mcp \
  --follow

# Show the last 100 log lines without streaming
az containerapp logs show \
  --name m365-mcp \
  --resource-group rg-m365-mcp \
  --tail 100
```

### Logs via Azure portal

Portal path: **rg-m365-mcp → m365-mcp → Logs (Log Analytics)**

Useful queries:
```
// Recent errors
ContainerAppConsoleLogs_CL
| where ContainerAppName_s == "m365-mcp"
| where Log_s contains "ERROR" or Log_s contains "error"
| order by TimeGenerated desc
| take 50

// All logs from the last hour
ContainerAppConsoleLogs_CL
| where ContainerAppName_s == "m365-mcp"
| where TimeGenerated > ago(1h)
| order by TimeGenerated desc
```

### Application audit log

Every allowed and denied call on both the REST and MCP surfaces is written to
the `auditLog` table in the tenant's storage account: tenant, user, device
label, operation, resource, result, reason, source, client IP, timestamp. A
Global Administrator reads it in the admin UI (Audit Log) or exports it:

```
# JSON, newest first, filters are optional
curl -H "Cookie: mcp_session=<admin session>" \
  "https://your-mcp-host.example.com/api/manage/audit-log?result=denied&limit=500"
# CSV
curl -H "Cookie: mcp_session=<admin session>" \
  "https://your-mcp-host.example.com/api/manage/audit-log?format=csv" -o audit-log.csv
```

Retention: the table has **no automatic purge** today; it grows until someone
truncates it. Log Analytics (console logs, telemetry) is set to 90 days in the
Bicep; confirm on the workspace, since the setting only applies
on provisioning:

```
az monitor log-analytics workspace show -g rg-m365-mcp -n m365-mcp-logs --query retentionInDays
```

### Revision history
```
az containerapp revision list \
  --name m365-mcp \
  --resource-group rg-m365-mcp \
  --query "[].{name:name,image:properties.template.containers[0].image,active:properties.active,created:properties.createdTime}" \
  --output table
```

* * *

## Troubleshooting

### Issue: MCP calls return 401 / "Session not found"

  1. Check that the user has authenticated: visit `https://your-mcp-host.example.com/api/auth/login` and complete the OAuth flow.
  2. If authentication fails at the Entra redirect: verify the Entra app registration redirect URI matches `https://your-mcp-host.example.com/api/auth/callback` exactly.
  3. If the session exists but calls still fail: the session has passed its 30-day absolute lifetime or 7-day idle window (`src/services/tokenCache.ts`), or the MSAL refresh token was revoked in Entra. Either way the user re-authenticates; the extension does this on its own at next start.



### Issue: SharePoint tools return "No sites enabled"

  1. Navigate to the admin UI: `https://your-mcp-host.example.com/`.
  2. Check **Sites → Allowlist**. Once the list has any entry, a site must be on it to be reachable. An **empty** list allows every site.



### Issue: Container App is down / returning 5xx
```
# 1. Check revision status
az containerapp revision list -n m365-mcp -g rg-m365-mcp \
  --query "[?properties.active].{name:name,health:properties.healthState,replicas:properties.replicas}" \
  --output table

# 2. Check recent logs for the error
az containerapp logs show -n m365-mcp -g rg-m365-mcp --tail 200 | grep -i error

# 3. Force a new revision (restarts all containers)
az containerapp revision restart \
  --name m365-mcp \
  --resource-group rg-m365-mcp \
  --revision $(az containerapp revision list -n m365-mcp -g rg-m365-mcp \
    --query "[?properties.active] | [0].name" -o tsv)
```

### Issue: a new revision never becomes ready, previous revision still serving

Check the new revision's console log for `[startup] FATAL`. The container
refuses to start when `MCP_SESSION_HMAC_KEY` or `MCP_DATA_ENCRYPTION_KEY` is
missing or malformed, or when a Key Vault reference behind one cannot be
resolved. See [Application keys](#application-keys): fix the binding and roll a
new revision. Do not generate a new key to get past it unless you mean to sign
every user out.

### Issue: Azure Table Storage connection errors
```
# Verify the connection string secret is set
az containerapp secret list -n m365-mcp -g rg-m365-mcp \
  --query "[?name==azure-storage-connection-string].name" -o tsv

# Verify the storage account is accessible
az storage account show -n stm365mcp -g rg-m365-mcp --query "statusOfPrimary" -o tsv
# Expected: available
```

### Issue: OAuth redirect URI mismatch (after domain change)
```
# Update the secrets and create a new revision
az containerapp secret set -n m365-mcp -g rg-m365-mcp \
  --secrets \
    oauth-redirect-uri="https://your-mcp-host.example.com/api/auth/callback" \
    frontend-url="https://your-mcp-host.example.com/"

# Force a NEW revision to pick up secret changes.
# `revision restart` is NOT enough: it restarts the replicas but they come back
# with the revision's original secret values. Verified on example-mcp.
az containerapp update -n m365-mcp -g rg-m365-mcp --revision-suffix "$(date +%Y%m%d%H%M)"
```

Always verify the value the container actually resolved, rather than the value
the platform reports for the secret. The two disagree after a plain restart:

```
az containerapp exec -n m365-mcp -g rg-m365-mcp --command "printenv FRONTEND_URL"
```

Note: `--query "[?properties.active][0].name"` silently returns empty. If you
need the active revision name, the working form is
`--query "[?properties.active] | [0].name"`.

* * *

## Scaling

### Current configuration

  * **Min replicas:** 1 (keeps at least one instance warm — avoids cold start for users)
  * **Max replicas:** 3
  * **Scale trigger:** HTTP concurrent requests ≥20 per replica



### Adjusting scale limits
```
# Increase max replicas (e.g., ahead of a large all-hands event)
az containerapp update \
  --name m365-mcp \
  --resource-group rg-m365-mcp \
  --max-replicas 5

# Scale to zero when idle (cost optimization for dev/test environments only)
az containerapp update \
  --name m365-mcp \
  --resource-group rg-m365-mcp \
  --min-replicas 0
```

Setting min-replicas to 0 causes cold starts of 10-30 seconds. Do not set this on the production instance — users will see authentication timeouts on the first request after an idle period.

### Right-sizing vCPU/memory
```
# Current allocation: 0.5 vCPU / 1 GiB — adequate for up to ~20 concurrent users
# To increase (requires a new revision):
az containerapp update \
  --name m365-mcp \
  --resource-group rg-m365-mcp \
  --cpu 1 --memory 2Gi
```

* * *

## Image updates and rollback

### Normal update path (CI/CD)

Push to `main` on `your-org/your-deploy-fork` (which the fork sync does for
you when `AUTO_UPDATE=true`). GitHub Actions builds, tags (`sha-<short>` \+ `latest`),
rolls the Container App, and then runs a post-deploy smoke that waits for
`/health` to report the commit it just built and checks the frontend URL serves
the admin SPA. A green deploy run means the new commit is live.

Before it rolls the image, the workflow converges the Container App's ingress
`targetPort` to the port the image binds (8080 since the non-root image). The deploy only swaps the image and never applies Bicep, so
without this step a port change in the image strands every tenant on its old
revision: that is what happened from 2026-09-26 until the step was added. The step is
a no-op when the port already matches. The Dockerfile, the Bicep `targetPort`
and each workflow's `CONTAINER_PORT` are pinned together by
`src/__tests__/containerPortInvariant.test.ts`, so change all of them or none.

For a **pinned** instance (`AUTO_UPDATE` unset), dispatch the sync and then the
deploy by hand:

```
gh workflow run "Sync from canonical" -R your-org/your-deploy-fork
gh workflow run the deploy workflow -R your-org/your-deploy-fork            # builds the fork head
gh workflow run the deploy workflow -R your-org/your-deploy-fork -f tag=sha-<short>   # or an existing image
```

### Manual deploy
```
cd ~/Git/m365-mcp-m365-mcp
./infra/infra/deploy.sh --tag v1.2.0
```

Break-glass only: a manual image tag does not correspond to a fork commit, which
makes "what is running?" unanswerable from `/health`.

### Rollback
```
# List recent revisions to find the prior good image
az containerapp revision list -n m365-mcp -g rg-m365-mcp \
  --query "[].{name:name,image:properties.template.containers[0].image,active:properties.active}" \
  --output table

# Roll back to a prior image tag
az containerapp update -n m365-mcp -g rg-m365-mcp \
  --image somamcpacr.azurecr.io/m365-mcp:sha-<previous-short-sha>
```

Then confirm with `curl -fsS https://your-mcp-host.example.com/health`. If the
instance is tracking upstream, the next canonical merge will roll it forward
again; pin it (`gh variable delete AUTO_UPDATE`) if you need it to stay put.

* * *

## Secret rotation

Every rotation below ends with a **new revision**, not a restart. A restarted
replica comes back with the revision's original secret values (see the redirect
URI section above).

### AZURE_CLIENT_SECRET (recommended every 12 months)

  1. Entra ID → App registrations → your M365 MCP app registration → Certificates & secrets → New client secret → copy value. Leave the old secret in place.
  2. Update the Container App secret and roll a new revision:
```
az containerapp secret set -n m365-mcp -g rg-m365-mcp \
  --secrets azure-client-secret="<new-value>"

az containerapp update -n m365-mcp -g rg-m365-mcp --revision-suffix "$(date +%Y%m%d%H%M)"
```

  3. Verify sign-in and one tool call, then delete the old secret from the Entra app registration.



### MCP_SESSION_HMAC_KEY and MCP_DATA_ENCRYPTION_KEY

These have their own section, [Application keys](#application-keys), because
rotating either one signs every user out.

### Credential envelope row binding (MCP_ENVELOPE_REQUIRE_AAD)

Stored access tokens and the MSAL cache are AES-256-GCM envelopes under the
one `MCP_DATA_ENCRYPTION_KEY`. Each envelope is bound to its table, partition,
row and column through GCM additional authenticated data, so an envelope
copied from one user's session row into another's fails to decrypt instead of
handing the second session the first user's Graph token.

Envelopes written by releases before that change carry no binding. The server
still reads them, and rewrites each one bound the first time it reads it. To
finish the migration and close the swap window for good:

  1. Deploy the release and leave it running long enough for active sessions to be read (an hour covers every active user, since access tokens refresh hourly).
  2. Open the user list in the admin UI once. Listing reads every session row, which rebinds the dormant ones too.
  3. Turn off legacy reads: `az containerapp update -n m365-mcp -g rg-m365-mcp --set-env-vars MCP_ENVELOPE_REQUIRE_AAD=true`

After step 3 an unbound envelope is treated like a tampered one: that session
fails authentication and the user signs in again. If you would rather not wait,
skip to step 3 directly and accept that every user re-authenticates once, or run
`infra/scripts/purge-credentials.sh` for the same effect.

### AZURE_CREDENTIALS (deploy service principal)

```
SP_APP_ID=$(az ad sp list --display-name gh-actions-m365-mcp --query '[0].appId' -o tsv)
az ad sp credential reset --id "$SP_APP_ID" --sdk-auth | gh secret set AZURE_CREDENTIALS -R your-org/your-deploy-fork
gh workflow run the deploy workflow -R your-org/your-deploy-fork    # prove it
```

Deploys authenticate to ACR with this principal (`az acr login`); there is no
ACR admin user and no `ACR_PASSWORD` secret.

### CANONICAL_SYNC_TOKEN

A fine-grained PAT with Contents: Read on `standardgauge/m365-mcp`, shared across
the forks, expiring 2027-09-04. Mint the replacement, `gh secret set
CANONICAL_SYNC_TOKEN` on every fork, then dispatch `Sync from canonical` on one
fork and confirm a green run. Unnecessary once the upstream is public.

* * *

## Application keys

`MCP_SESSION_HMAC_KEY` and `MCP_DATA_ENCRYPTION_KEY` protect what the server
keeps in Table Storage. Nothing else in the deployment works like them, so read
this section before touching either.

| Key | Protects | Table |
|---|---|---|
| `MCP_SESSION_HMAC_KEY` | Session lookup. Clients hold a random bearer token; storage holds only its HMAC-SHA256, which is also the row key | `mcpSessions` |
| `MCP_DATA_ENCRYPTION_KEY` | Each session's Microsoft Graph access token, and the MSAL token cache (refresh tokens for every signed-in user), as AES-256-GCM envelopes | `mcpSessions`, `mcpMsalCache` |

Each is 32 random bytes written as 64 hex characters (`openssl rand -hex 32`).
Use two different keys, and a fresh pair per deployment and per environment.

### Startup validation

The container checks both keys before the Functions host starts
(`src/startup/checkKeys.ts`, run by the image's `CMD`). A key that is missing,
empty, the wrong length, or not hex stops the container with a line per bad key:

```
[startup] FATAL: MCP_DATA_ENCRYPTION_KEY env var is not set. ...
[startup] Refusing to start. Generate a key with `openssl rand -hex 32`; ...
```

Surrounding whitespace, such as a trailing newline pasted into Key Vault, is
ignored. Nothing else is: a value with junk after a valid key is rejected
rather than quietly truncated. The messages never include key material.

Releases before this check accepted such a value and used its first 64 hex
characters. If an upgrade is refused for a key that has been working, those 64
characters are the key in use: bind exactly them and every session survives.
Binding a new key instead is a rotation.

In single-revision mode, the default here, a revision whose replicas never
become ready does not take traffic, so the previous revision keeps serving and
the deploy workflow's smoke check fails on the stale `sha`. To see why:

```
az containerapp logs show -n m365-mcp -g rg-m365-mcp --type console --tail 50 | grep '\[startup\]'
```

Before this check existed, a missing key surfaced only at the first sign-in,
after the revision had passed `/health`.

### Storage: Key Vault references

Keep both keys in Azure Key Vault and bind the Container App secrets to them as
Key Vault references, so the value lives in one place with access control,
versioning, soft delete and an access log. The Container App's system-assigned
identity (present in both Bicep templates) needs **Key Vault Secrets User** on
the vault.

```
VAULT=kv-m365-mcp            # an RBAC-mode vault with soft delete and purge protection
VAULT_ID=$(az keyvault show -n "$VAULT" --query id -o tsv)
PRINCIPAL=$(az containerapp show -n m365-mcp -g rg-m365-mcp --query identity.principalId -o tsv)

az role assignment create --assignee-object-id "$PRINCIPAL" \
  --assignee-principal-type ServicePrincipal \
  --role "Key Vault Secrets User" --scope "$VAULT_ID"

# Store the current values (moving an existing deployment) or fresh ones (a new
# deployment). Moving the current values keeps everyone signed in; new values
# are a rotation (below).
HMAC_URI=$(az keyvault secret set --vault-name "$VAULT" -n mcp-session-hmac-key \
  --value "$(openssl rand -hex 32)" --query id -o tsv)
DATA_URI=$(az keyvault secret set --vault-name "$VAULT" -n mcp-data-encryption-key \
  --value "$(openssl rand -hex 32)" --query id -o tsv)

az containerapp secret set -n m365-mcp -g rg-m365-mcp --secrets \
  "mcp-session-hmac-key=keyvaultref:${HMAC_URI},identityref:system" \
  "mcp-data-encryption-key=keyvaultref:${DATA_URI},identityref:system"

az containerapp update -n m365-mcp -g rg-m365-mcp --revision-suffix "$(date +%Y%m%d%H%M)"
```

`--query id` returns the **versioned** secret URI, and that is deliberate. With
a versionless reference, writing a new version in the vault rotates the key on
whichever revision happens to start next, which signs everyone out at a time
nobody chose. With a versioned reference the key changes only when you change
the reference.

Two traps:

- **The shipped Bicep writes these secrets as plain values** from its
  `mcpSessionHmacKey` and `mcpDataEncryptionKey` parameters. Re-running it after
  switching to Key Vault references replaces the references with whatever the
  parameters hold. If those differ from the vault, that is an unplanned
  rotation. Pass the current values, or re-apply the references afterwards.
- **`az containerapp secret set name=<value>`** with a literal value does the
  same: it silently turns a reference back into a plain secret.

Confirm what is bound (values are not printed):

```
az containerapp secret list -n m365-mcp -g rg-m365-mcp \
  --query "[?starts_with(name,'mcp-')].{name:name,keyVaultUrl:keyVaultUrl}" -o table
```

### Rotation

There is no dual-key window. The server holds one key of each kind, so data
written under the old key cannot be read under the new one, and **rotating
either key signs every user out**. Rotate on suspected compromise, on staff
changes with vault access, or on a scheduled date with notice to users.

What each rotation does:

| | Sessions | Stored Graph access tokens | MSAL cache (refresh tokens) |
|---|---|---|---|
| HMAC key | Every bearer token stops matching its row. Calls return 401 and the user signs in again; the extension prompts at its next start | Unreadable in practice: still decryptable, but nothing can look the row up | Unaffected |
| Data key | Rows are found, but the access token fails to decrypt, so the session is treated as missing: 401, sign in again | Undecryptable | Undecryptable, treated as empty. Silent refresh fails for everyone until they sign in, and the first new sign-in overwrites the row |
| Both | Union of the two | Undecryptable | Undecryptable |

Old rows are not deleted by a rotation. They are harmless, but they cost
something until removed: a request carrying an old bearer token misses the
direct lookup and falls through to a scan of the whole session table before it
returns 401, and after a data-key rotation the admin UI's session list fails to
read storage (it logs the decrypt error and shows only sessions held in that
replica's memory). So purge after every rotation.

Each replica also holds decrypted sessions in memory. That is why a rotation
must be a new revision: the new replicas start empty, and the old ones are
retired once the new revision is ready.

Procedure, both keys (drop the line for the key you are not rotating):

```
VAULT=kv-m365-mcp
HMAC_URI=$(az keyvault secret set --vault-name "$VAULT" -n mcp-session-hmac-key \
  --value "$(openssl rand -hex 32)" --query id -o tsv)
DATA_URI=$(az keyvault secret set --vault-name "$VAULT" -n mcp-data-encryption-key \
  --value "$(openssl rand -hex 32)" --query id -o tsv)

az containerapp secret set -n m365-mcp -g rg-m365-mcp --secrets \
  "mcp-session-hmac-key=keyvaultref:${HMAC_URI},identityref:system" \
  "mcp-data-encryption-key=keyvaultref:${DATA_URI},identityref:system"

az containerapp update -n m365-mcp -g rg-m365-mcp --revision-suffix "$(date +%Y%m%d%H%M)"

# Once the new revision serves /health and one sign-in works:
infra/scripts/purge-credentials.sh --resource-group rg-m365-mcp --app-name m365-mcp --dry-run
infra/scripts/purge-credentials.sh --resource-group rg-m365-mcp --app-name m365-mcp
```

Until you purge, rollback is possible: point the references back at the
previous secret versions and roll another revision, and the old sessions
work again. After the purge they are gone either way. Once you are satisfied,
disable the previous versions in the vault
(`az keyvault secret set-attributes --id <old-version-uri> --enabled false`).

A deployment still on plain Container App secrets rotates the same way with
`--secrets mcp-session-hmac-key="$(openssl rand -hex 32)"` in place of the
references. Moving to Key Vault at the same time costs nothing extra, since the
users are signing in again regardless.

### Emergency revocation

The keys protect stored data; they are not Entra credentials. What you do
depends on what was exposed.

- **A key, but not storage.** Neither key reveals anything on its own. Rotate
  the exposed key and purge, on your own schedule.
- **Storage and the HMAC key.** Session tokens are 32 random bytes and storage
  holds only their HMAC, so the rows cannot be turned back into bearer tokens.
  Rotate the HMAC key and purge so no stolen row is ever matched again.
- **Storage and the data key.** Treat this as an incident. The attacker can
  decrypt every stored Graph access token (valid for about an hour) and the
  MSAL cache, which holds a refresh token for every signed-in user. Rotating the
  data key stops the *server* from using those tokens; it does nothing to the
  copies already taken. In this order:
  1. Rotate both keys in one revision and purge (procedure above). The
     server's own hold on every token ends.
  2. **Revoke sessions in Entra for every user** of the app (Kill switch,
     step 2). This is what invalidates the stolen refresh tokens.
  3. Rotate `AZURE_CLIENT_SECRET` and delete the old one. Refresh tokens issued
     to this app can only be redeemed with its client credentials, which sit in
     the same Container App as the keys, so assume they went together.
  4. Read the application audit log and Entra sign-in logs for the exposure
     window.
- **The vault itself, or the app's identity.** Both keys, plus
  `AZURE_CLIENT_SECRET` if it is stored there: do all of the above, and review
  the vault's role assignments before writing new versions into it.

**A key that is lost** (secret deleted or disabled in the vault) is not an
outage straight away: running replicas already hold the value. But the next
revision, and any replica that restarts, cannot resolve the reference and will
not start. Recover it first: `az keyvault secret recover --vault-name <vault>
-n <name>` while it is in soft delete, or re-enable the version. If it cannot be
recovered, everything it protected is lost; generate a new key, which is a
rotation with the same effect on users.

* * *

## Users

### Adding a user

Nothing to do server-side. Any enabled, licensed member of the tenant installs
the extension from `https://your-mcp-host.example.com/install` and signs in. To
restrict who can sign in at all, require assignment on the enterprise
application in Entra and assign users or groups. To restrict what a signed-in
user can reach, use User Management in the admin UI: disable services per user,
or disable mail indexing for a user.

### Removing a user

The application has no admin action that revokes one user's session; only the
user's own logout deletes their sessions. Remove access in Entra:

  1. Disable or delete the user, and **Revoke sessions** on the user object. Their refresh token stops working; tool calls fail as soon as the current access token expires (about an hour).
  2. Optionally remove them from the enterprise application's assignment.
  3. Their session row lingers until the 30-day absolute cap and is harmless; `infra/scripts/purge-credentials.sh` removes every session and the MSAL cache if you need it gone now.

* * *

## Kill switch and incident response

Ordered by speed. Each step stands alone; in an incident, do 1 and 3 first.

  1. **Disable the enterprise application** in Entra (Enterprise applications → the app → Properties → "Enabled for users to sign in?" → No). New sign-ins and refresh-token redemptions stop at once. Existing access tokens keep working until they expire, roughly an hour.
  2. **Revoke sessions** in Entra for the affected users, or for all users. Refresh tokens are invalidated; combined with step 1 no new Graph access is possible after current access tokens expire.
  3. **Rotate both application keys** ([Application keys](#application-keys)). Every stored session and the MSAL cache become undecryptable in one revision roll; the server's own hold on tokens is gone regardless of what Entra does.
  4. **Purge the credential tables:** `infra/scripts/purge-credentials.sh --resource-group rg-m365-mcp --app-name m365-mcp` (add `--dry-run` first). Same effect as 3, slower, no key change.
  5. **Delete the client secret** on the app registration. The server can no longer redeem authorization codes or refresh tokens.
  6. **Stop traffic:** `az containerapp ingress disable -n m365-mcp -g rg-m365-mcp`, or scale to zero, or delete the app. Storage and keys survive unless deleted.

Then: export the audit log for the window (`/api/manage/audit-log?startDate=…&format=csv`),
pull Entra sign-in logs for the app registration, and read the container logs
for the same window. Users get back in by re-authenticating once you re-enable
the application.

* * *

## Uninstall and teardown

  1. Pin the instance (`gh variable delete AUTO_UPDATE -R <fork>`) so nothing rolls mid-teardown.
  2. Tell users; the extension will report "Session expired" and can be removed from Claude Desktop.
  3. Export the audit log if the record is needed.
  4. `az group delete -n rg-m365-mcp`: Container App, environment, ACR, storage (every session, cache and audit row) and the Log Analytics workspace go together.
  5. Delete the app registration in Entra, or at least the client secret and the consent grants; every refresh token the instance held becomes unusable.
  6. Delete the fork, or leave it pinned as a record.

The server never stored tenant content, only tokens, policy and audit rows,
all of which lived in the deleted storage account.
