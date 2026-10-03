# M365 MCP Server — Operations Runbook

**Audience:** Operators managing a live M365 MCP Container App. The worked
example throughout is the `your-mcp-host.example.com` instance; every command takes
the same shape for any tenant with the resource names swapped. \
**Last updated:** September 2026


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

**Rotating either key invalidates all existing sessions and the MSAL token
cache.** Every user must re-authenticate afterwards (the extension prompts on
its next start). Rotate on suspected compromise, or on a scheduled date with
notice. This is also the fastest kill switch (below).
```
NEW_HMAC=$(openssl rand -hex 32)
NEW_DATA=$(openssl rand -hex 32)

az containerapp secret set -n m365-mcp -g rg-m365-mcp \
  --secrets \
    mcp-session-hmac-key="$NEW_HMAC" \
    mcp-data-encryption-key="$NEW_DATA"

az containerapp update -n m365-mcp -g rg-m365-mcp --revision-suffix "$(date +%Y%m%d%H%M)"
```

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
  3. **Rotate both application keys** (previous section). Every stored session and the MSAL cache become undecryptable in one revision roll; the server's own hold on tokens is gone regardless of what Entra does.
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
