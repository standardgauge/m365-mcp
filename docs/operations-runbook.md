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

## Admin API sessions

The admin API, `/api/manage/*`, does not accept the session token an MCP client
uses. A request must come from the admin UI in a browser, and carry both:

  * the `mcp_session` cookie, and
  * the `mcp_console` cookie, which only an interactive sign-in at `/api/auth/login` sets. It is HttpOnly, `SameSite=Strict`, scoped to `/api`, and bound by MAC to that exact `mcp_session` value.

It must also pass an Origin check: the `Origin` header, when present, must be
the instance's own origin (or `FRONTEND_URL`'s, for local development), and
`Sec-Fetch-Site` must be `same-origin`. A typed or bookmarked GET (`none`) is
allowed, so an admin can open an `/api/manage/...` URL in the signed-in browser.
A bearer token, `x-session-token`, or a copied `mcp_session` cookie gets a 401.

The console session lasts 30 minutes idle and never more than 8 hours from the
sign-in that started it. The admin UI renews it through `/api/auth/me` every
five minutes while its tab is visible, and sends the user back through sign-in
once it has lapsed; with a live Microsoft session that is a redirect, not a
prompt. Logout is a `POST` to `/api/auth/logout` and also checks Origin. It
ends the console session along with every session row for the user.

Nothing is stored for it: the cookie is checked against `MCP_SESSION_HMAC_KEY`,
so rotating that key ends every console session along with every other session.
Code: `src/services/consoleSession.ts`. Threat model: [section 7](threat-model.md#7-admin-surface).

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

### Host health-check noise

The Container App runs without `AzureWebJobsStorage` by design (see
`infra/container-app.bicep`). The Functions host still registers its own
storage health check and cannot be told not to, so that check reports
`azure.functions.webjobs.storage Unhealthy: Unable to create client for
AzureWebJobsStorage` on every 30-second publish. The service is unaffected:
the probes check the app's own `/health`, not the host's health report.

Left alone, the host's health-check publisher writes that report to the
console on every publish as a `warn:` entry, a header line naming the
category followed by `Process reporting unhealthy: Unhealthy. Health check
entries are {...}`. Any console-log alert keyed on warnings then fires
constantly and real warnings drown. `host.json` sets that category,
`Microsoft.Azure.WebJobs.Script.Diagnostics.HealthChecks.TelemetryHealthCheckPublisher`,
to `Error`, which drops the entry. The cost is that the host's own lifecycle
health reports are not logged either; the probes and `/health` cover the same
ground. The container boot test in CI fails if the entry comes back.

If an instance still shows it, for example after a host update moves the
publisher to another category, exclude it from the query rather than
loosening the alert. The console writes the header and the body as separate
rows, so drop both:

```
ContainerAppConsoleLogs_CL
| where ContainerAppName_s == "m365-mcp"
| where Log_s !has "TelemetryHealthCheckPublisher"
| where Log_s !has "azure.functions.webjobs.storage"
| where Log_s has "warn" or Log_s contains "error"
| order by TimeGenerated desc
| take 50
```

### Application audit log

Every allowed and denied call on both the REST and MCP surfaces, every
authentication event and every admin policy change is written to two places. The **authoritative record** is the `M365McpAudit_CL` table in the
instance's Log Analytics workspace (next section), which is where a security
team or SIEM should read it. A working copy goes to the `auditLog` table in the
tenant's storage account: tenant, user, device label, operation, resource,
result, reason, source, client IP, before, after, timestamp. That copy backs the admin screen;
a Global Administrator reads it in the admin UI (Audit Log), which also exports
the filtered view as CSV. The same endpoint answers a URL typed into the address
bar of the browser that is signed in to the admin UI:

```
# JSON, newest first, filters are optional
https://your-mcp-host.example.com/api/manage/audit-log?result=denied&limit=500
# CSV
https://your-mcp-host.example.com/api/manage/audit-log?format=csv
```

`/api/manage/*` accepts only the admin UI's browser session, never the session
token an MCP client holds, so `curl` with a bearer token or a copied
`mcp_session` cookie gets a 401. See [Admin API sessions](#admin-api-sessions).

Events that are not tool calls:

| Operation | Written when | Actor and target |
|---|---|---|
| `auth.login` | Browser sign-in completes (`resource` `browser` or `install`), or fails: identity-platform error, state mismatch, missing PKCE verifier or nonce, missing code, foreign tenant (`resource` `tenant:<id>`), callback error | The signed-in user; on failure no user, in the instance's own tenant |
| `auth.device_login` | A device-code sign-in completes or fails | As above. The address is the one that started the flow |
| `auth.install_handoff` | The browser enters the installer's code on install-confirm and the session is attached (`resource` `attach`; denied if the callback could not record the pending handoff or the attach failed), the user declines it or it is discarded after five wrong codes (`confirm`, denied), the installer collects it (`poll`), or polls an expired nonce (`poll`, denied) | The signed-in user. Compare the two addresses: `attach` and `confirm` are the browser, `poll` is the machine running the installer |
| `auth.logout` | A logout ended a session, or failed to | The user whose sessions were deleted |
| `auth.refresh`, `auth.session_renew` | `/api/auth/refresh`, or a request on a session idle past its TTL, renews it or fails to | The session's user |
| `policy.deny_list.global.add` / `.remove` | A global deny entry changes | The admin; `resource` is `<type>:<path>` |
| `policy.deny_list.user.add` / `.remove` / `.clear` | A per-user deny list changes | The caller; `resource` is `user:<id>/<type>:<path>` or `user:<id>` |
| `policy.services.set`, `policy.read_only.set`, `policy.allowed_sites.set` | Tenant service settings change | The admin; `resource` `tenant` |
| `policy.user_services.set`, `policy.mail_config.set` | A user's service overrides or mail config change | The caller; `resource` `user:<id>` |
| `set_email_output_policy`, `set_email_output_mode` | Draft-mode policy, or a user's own output mode, changes | The caller; `resource` `tenant`, `user:<id>` or the user id |
| `set_outbound_policy` | The outbound policy for calendar invitations, response comments or Teams sends changes | The admin; `resource` `tenant` or `user:<id>`; `reason` lists the channels set |
| `admin.*.read` | Refusals only: a non-admin asked for the audit log, sessions, tenant users, mail config, draft-mode policy, outbound policy, or another user's deny list or overrides | The caller |

Every policy operation is also written, `denied` with reason `Global
Administrator role required`, when a caller fails the admin check. `before` and
`after` hold the value as JSON, capped at 8,000 characters each.

The client IP is the entry the Container Apps ingress appended to
`X-Forwarded-For`, never one the client sent, with any port removed and IPv6
addresses kept whole. Behind Front Door or an Application Gateway set
`RATE_LIMIT_TRUSTED_PROXY_HOPS=2` (see [Rate limiting](#rate-limiting-and-ingress-restriction)),
or every row records the proxy's address. Rows written before this change may
hold the whole forwarded header (REST) or nothing (MCP).

#### Audit log retention

Rows older than `AUDIT_LOG_RETENTION_DAYS` (default **365**) are deleted
automatically. The purge runs at most once a day per replica, started by the
first audited call after the last run, so it runs whenever the server is in use.
It is not a Functions timer trigger: those need `AzureWebJobsStorage`, which the
Container App runs without. Every run logs one line, including runs that delete
nothing, so you can confirm it is alive:

```
ContainerAppConsoleLogs_CL
| where ContainerAppName_s == "m365-mcp"
| where Log_s contains "Retention purge"
| order by TimeGenerated desc
| take 10
```

No line in the last two days on a server that is taking calls means the purge
has stopped. A `Retention purge failed` line carries the storage error.

To change the window, or set `0` to keep every row:

```
az containerapp update -n m365-mcp -g rg-m365-mcp --set-env-vars AUDIT_LOG_RETENTION_DAYS=730
```

An invalid value (negative, fractional, not a number) logs an error and falls
back to 365. It never turns retention off; only an explicit `0` does.

Before shortening the window, or when a client asks for a record the window
would drop, export first. The export is the archive; the purge does not keep a
copy:

```
# Everything older than the new cutoff, as CSV, from the signed-in browser
https://your-mcp-host.example.com/api/manage/audit-log?endDate=2026-01-01T00:00:00Z&limit=1000&format=csv
```

The endpoint returns at most 1000 rows per call, newest first. For a larger
range, page backwards by setting `endDate` to the oldest `timestamp` in the
previous file until a call returns no rows.

#### Log Analytics retention

Container console logs and telemetry go to the Log Analytics workspace, which
the Bicep provisions with `retentionInDays: 90`. That is the value at
provisioning time; a change in the portal survives until the next Bicep
deployment. Check what each live workspace actually holds:

```
az monitor log-analytics workspace show -g rg-m365-mcp -n m365-mcp-logs --query retentionInDays
```

Change it with
`az monitor log-analytics workspace update -g rg-m365-mcp -n m365-mcp-logs --retention-time <days>`,
and change `retentionInDays` in the Bicep to match so the next deployment does
not put it back. Record the verified value with the rest of the deployment's
own notes, outside this repository.

### Audit trail in Log Analytics

Each audit event is also sent through the Azure Monitor Logs Ingestion API to a
custom table, `M365McpAudit_CL`, in the instance's own Log Analytics workspace
(`m365-mcp-logs`), in the client's subscription. `infra/audit-ingestion.bicep`
defines the table and the data collection rule `m365-mcp-audit-dcr` that feeds
it; `infra/main.bicep` and `infra/container-app.bicep` deploy both and grant the
Container App's system-assigned identity **Monitoring Metrics Publisher on that
rule only**. The identity can append rows to this one stream; it cannot read the
workspace or write anywhere else.

| Column | Type | Meaning |
|---|---|---|
| `TimeGenerated` | datetime | When the server recorded the event (UTC) |
| `EventId` | string | Same value as the `auditLog` table's RowKey, to join the two copies |
| `EntraTenantId` | string | Entra tenant of the caller (`TenantId` is reserved by Log Analytics for the workspace id) |
| `UserId` | string | Entra object id of the caller |
| `UserEmail` | string | Caller's UPN |
| `DeviceLabel` | string | Device label from the session, when present |
| `Operation` | string | e.g. `mail.search_mail`, `sharepoint.read_file`, `auth.login`, `policy.services.set` |
| `TargetResource` | string | Message id, site/path, folder, or the target of an admin change, when the operation has one |
| `Result` | string | `allowed` or `denied` |
| `Reason` | string | Why a call was denied |
| `Source` | string | `mcp` or `http` |
| `ClientIp` | string | Caller IP as the ingress saw it, when known |
| `Before` | string | Policy value before an admin change, as JSON |
| `After` | string | Policy value after it, as JSON |

Retention follows the workspace (90 days in the Bicep). To keep the audit table
longer than the console logs, set it on the table alone:

```
az monitor log-analytics workspace table update -g rg-m365-mcp \
  --workspace-name m365-mcp-logs -n M365McpAudit_CL \
  --retention-time 90 --total-retention-time 730
```

Queries (portal: **rg-m365-mcp → m365-mcp-logs → Logs**):

```
// Denied calls in the last 24 hours, by user and reason
M365McpAudit_CL
| where TimeGenerated > ago(24h) and Result == "denied"
| summarize Denials = count(), Operations = make_set(Operation) by UserEmail, Reason
| order by Denials desc

// Everything one user did in a window
M365McpAudit_CL
| where UserEmail =~ "adele@fabrikam.com"
| where TimeGenerated between (datetime(2026-01-01) .. datetime(2026-01-02))
| project TimeGenerated, Operation, TargetResource, Result, Reason, Source, ClientIp
| order by TimeGenerated asc

// Policy changes and refused admin checks
M365McpAudit_CL
| where Operation startswith "policy." or Operation startswith "admin." or Operation endswith "_email_output_policy"
| project TimeGenerated, UserEmail, Operation, TargetResource, Result, Before, After, ClientIp

// Failed sign-ins by address
M365McpAudit_CL
| where Operation in ("auth.login", "auth.device_login") and Result == "denied"
| summarize Failures = count(), Reasons = make_set(Reason) by ClientIp
| order by Failures desc
```

**The `Before` and `After` columns need the current Bicep.** An instance whose
table and data collection rule predate them drops both columns from the Log
Analytics copy without an error; the storage table and the admin screen still
have them. Re-apply `infra/audit-ingestion.bicep` as described below to add
them.

**Pointing Sentinel or another SIEM at it.** The table lives in the client's
workspace, so the client's tooling owns what happens next:

- *Sentinel on this workspace:* enable Microsoft Sentinel on `m365-mcp-logs` and
  the table is immediately available to hunting queries and analytics rules.
  Use the queries above as the starting point for a scheduled rule.
- *Sentinel on a different workspace:* query across workspaces from the central
  one, `workspace("<m365-mcp-logs resource id>").M365McpAudit_CL`, or export
  the table into it with a data export rule.
- *Any other SIEM (Splunk, Elastic, QRadar…):* export the table to an Event Hub
  and point the SIEM's Event Hub connector at it:

  ```
  az monitor log-analytics workspace data-export create -g rg-m365-mcp \
    --workspace-name m365-mcp-logs -n audit-to-siem \
    --tables M365McpAudit_CL \
    --destination <event hub namespace resource id>
  ```

**When a write fails, the server does nothing about it.** It logs
`[auditLog] Failed to send N audit event(s) to Log Analytics` to the console and
drops that batch; requests are not blocked and the server raises no alert of
its own. The storage-table copy is written independently. Detecting a gap is the
client's security tooling's job. Two signals to alert on:

```
// Upload failures reported by the server
ContainerAppConsoleLogs_CL
| where Log_s has "Failed to send" and Log_s has "Log Analytics"

// The audit stream went quiet while the app was serving traffic
M365McpAudit_CL
| summarize Last = max(TimeGenerated)
| where Last < ago(6h)
```

Events are sent in batches of up to 100, at most one second after the first
event in a batch, so a replica killed mid-batch can lose up to a second of
events from Log Analytics. They still reach the storage table.

**Turning it on for an instance deployed before this existed.** The deploy
workflow only swaps the image; it never applies Bicep. Either re-run the full
`az deployment group create` with `infra/main.bicep`, or add just the audit
pieces:

```
az deployment group create -g rg-m365-mcp \
  --template-file infra/audit-ingestion.bicep \
  --parameters location=<region of m365-mcp-logs> workspaceName=m365-mcp-logs \
               dataCollectionRuleName=m365-mcp-audit-dcr \
  --query properties.outputs
# Note logsIngestionEndpoint and dataCollectionRuleImmutableId from the output.

PRINCIPAL=$(az containerapp show -n m365-mcp -g rg-m365-mcp --query identity.principalId -o tsv)
DCR_ID=$(az resource show -g rg-m365-mcp -n m365-mcp-audit-dcr \
  --resource-type Microsoft.Insights/dataCollectionRules --query id -o tsv)
az role assignment create --assignee-object-id "$PRINCIPAL" \
  --assignee-principal-type ServicePrincipal \
  --role "Monitoring Metrics Publisher" --scope "$DCR_ID"

az containerapp update -n m365-mcp -g rg-m365-mcp --set-env-vars \
  AUDIT_LOGS_INGESTION_ENDPOINT=<logsIngestionEndpoint> \
  AUDIT_DCR_IMMUTABLE_ID=<dataCollectionRuleImmutableId> \
  AUDIT_DCR_STREAM_NAME=Custom-M365McpAudit
```

Without `AUDIT_LOGS_INGESTION_ENDPOINT` and `AUDIT_DCR_IMMUTABLE_ID` the server
writes the storage table only. Expect `403` upload failures in the console for
up to about half an hour after the role assignment while it propagates, and the
first rows to take several minutes to appear in a newly created table. Verify
with one tool call, then `M365McpAudit_CL | take 10`.

**Why not Application Insights.** The server already sends console output to
Application Insights, and `trackEvent` would put rows in the same workspace's
`AppEvents` table. It is not used for the audit record because the App Insights
connection string authenticates nothing (anyone holding it can write
indistinguishable rows), telemetry is subject to sampling and the App Insights
daily cap, and events arrive as an untyped property bag in a table shared with
every other custom event. The Logs Ingestion API path accepts only callers
holding the role on this rule and lands typed columns in a table of its own.

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
  3. If the session exists but calls still fail: the session has passed its 30-day absolute lifetime (`src/services/tokenCache.ts`), or the MSAL refresh token was revoked or expired in Entra. A session left unused for more than 7 days is not ended by that alone: its next request renews it through the refresh token, and fails only if the refresh token no longer works. Either way the user re-authenticates; the extension does this on its own at next start.



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

The server reaches storage with its runtime identity (see
[Storage access](#storage-access)). A `403 AuthorizationPermissionMismatch` means
the identity has no data role on the account; `KeyBasedAuthenticationNotPermitted`
means something is still presenting an account key to an account that refuses
them, usually an instance whose app predates `AZURE_STORAGE_TABLE_ENDPOINT`.

```
# Which path is the app on? An endpoint means Entra auth.
az containerapp show -n m365-mcp -g rg-m365-mcp \
  --query "properties.template.containers[0].env[?starts_with(name,'AZURE_STORAGE')].{name:name,value:value}" -o table

# Does the runtime identity hold its role on the account?
az role assignment list --scope "$(az storage account show -n stm365mcp -g rg-m365-mcp --query id -o tsv)" \
  --query "[?roleDefinitionName=='Storage Table Data Contributor'].principalName" -o tsv

# Is the account up?
az storage account show -n stm365mcp -g rg-m365-mcp --query "statusOfPrimary" -o tsv
# Expected: available
```

A role assignment made a moment ago can take a few minutes to apply.

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

### Rate limiting and ingress restriction

Container Apps ingress has no native rate limit, so the server limits its five
unauthenticated routes itself: `/api/auth/login`, `/api/auth/device`,
`/api/auth/install-poll`, `/api/auth/install-confirm` and `/api/mcp`. Each client address gets a fixed
one-minute window per route, counted in the replica's memory. Over the limit,
the route answers `429` with `Retry-After` and the handler never runs.

| Route | Default per address per minute | Override |
|---|---|---|
| `login` | 30 | `RATE_LIMIT_LOGIN_PER_MINUTE` |
| `device` | 10 | `RATE_LIMIT_DEVICE_PER_MINUTE` |
| `install-poll` | 120 | `RATE_LIMIT_INSTALL_POLL_PER_MINUTE` |
| `install-confirm` | 30 | `RATE_LIMIT_INSTALL_CONFIRM_PER_MINUTE` |
| `mcp` | 1200 | `RATE_LIMIT_MCP_PER_MINUTE` |

The defaults are sized for an office behind one NAT address, not for one user.
`device` is the tightest because every call starts a device-code flow that polls
Entra in the background for up to fifteen minutes. The install scripts poll every
two seconds and back off on a `429`, so a burst of installs from one office slows
down rather than failing.

Things to know:

  * **Per replica.** With three replicas an address can reach up to three times the
    limit. That is deliberate: the limiter makes the routes cost something to
    hammer, it does not meter them.
  * **Client address.** The limiter keys on the entry the ingress appended to
    `X-Forwarded-For` (the rightmost one), never on entries the client sent.
    IPv6 clients are grouped by `/64`. If you put Front Door or an Application
    Gateway in front of the app, set `RATE_LIMIT_TRUSTED_PROXY_HOPS=2`, or every
    user shares the proxy's bucket.
  * **Seeing it fire.** The first refusal for an address in a window logs
    `rate limit: <route> refused <address> over <n>/min` at warning level.
  * **Turning a route off.** Set its variable to `0`.

```
az containerapp update -n m365-mcp -g rg-m365-mcp \
  --set-env-vars RATE_LIMIT_MCP_PER_MINUTE=2400
```

**If your users always connect from known addresses**, restrict ingress as well.
This is the stronger control: requests from anywhere else are refused at the
edge and never reach the app. It only fits a tenant whose users reach the server
from named egress ranges (an office, a VPN, a proxy). Remote users on home
connections will be locked out, and so will Microsoft's sign-in redirect back to
`/api/auth/callback` if the browser is outside the allowed ranges.

```
az containerapp ingress access-restriction set -n m365-mcp -g rg-m365-mcp \
  --rule-name office --action Allow --ip-address 203.0.113.0/24
az containerapp ingress access-restriction list -n m365-mcp -g rg-m365-mcp
az containerapp ingress access-restriction remove -n m365-mcp -g rg-m365-mcp --rule-name office
```

Once any `Allow` rule exists, every address not on the list is denied. The deploy
workflow only swaps the image, so restrictions you set survive deploys.

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
you). With `AUTO_UPDATE=true`, the deploy starts when the fork's CI run on that
commit succeeds, not on the push itself; a red CI run deploys nothing and the
instance keeps serving what it had. If `main` has moved on by the time CI
finishes, the run skips and leaves the newer commit to its own CI run, so a slow
CI run cannot roll an instance backwards. GitHub Actions builds, tags (`sha-<short>` \+ `latest`),
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
deploy by hand. A manual dispatch does not wait for CI, which is what makes it
the rollback path while `main` is red; check the CI run on the commit first when
you are rolling forward:

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
  2. Write it to the vault as a new version, point the reference at it, and roll a new revision (variables as in [Storage: Key Vault references](#storage-key-vault-references)):
```
SECRET_URI=$(az keyvault secret set --vault-name "$VAULT" -n azure-client-secret \
  --value "<new-value>" --query id -o tsv)

az containerapp secret set -n m365-mcp -g rg-m365-mcp \
  --secrets "azure-client-secret=keyvaultref:${SECRET_URI},identityref:${IDENTITY}"

az containerapp update -n m365-mcp -g rg-m365-mcp --revision-suffix "$(date +%Y%m%d%H%M)"
```
     Rotating the client secret does not sign anyone out. The next Bicep deployment writes whatever its `azureClientSecret` parameter holds, so update that too.

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

A session row's access-token envelope also binds that row's `userId`,
`homeAccountId` and `tenantId`. `homeAccountId` picks the MSAL account a silent
refresh draws on, so without this a storage writer holding any session could
point it at another user's account. With it, an edited row fails to decrypt and
the session fails authentication.

Envelopes written by older releases carry no binding, or bind the row but not
its identity columns. The server still reads them, and rewrites each one fully
bound the first time it reads it. To finish the migration and close the window
for good:

  1. Deploy the release and leave it running long enough for active sessions to be read (an hour covers every active user, since access tokens refresh hourly).
  2. Open the user list in the admin UI once. Listing reads every session row, which rebinds the dormant ones too.
  3. Turn off legacy reads: `az containerapp update -n m365-mcp -g rg-m365-mcp --set-env-vars MCP_ENVELOPE_REQUIRE_AAD=true MCP_SESSION_REQUIRE_IDENTITY_BINDING=true`

After step 3 an unbound or row-only envelope is treated like a tampered one:
that session fails authentication and the user signs in again. So does a
session row with no envelope at all. `MCP_SESSION_REQUIRE_IDENTITY_BINDING=true`
on its own already refuses unbound session envelopes; `MCP_ENVELOPE_REQUIRE_AAD`
is still what covers the MSAL cache. If you set `MCP_ENVELOPE_REQUIRE_AAD=true`
under an earlier release, repeat steps 1 and 2 after upgrading before you add
the second variable. If you would rather not wait,
skip to step 3 directly and accept that every user re-authenticates once, or run
`infra/scripts/purge-credentials.sh` for the same effect.

### MSAL token cache: one row per account

`mcpMsalCache` holds one row per signed-in account: partition `account`, row
key the account's MSAL home account id (`<object id>.<tenant id>`, the same value
as `homeAccountId` on that user's `mcpSessions` rows). Each row is that
account's refresh token cache as an AES-256-GCM envelope bound to the row.
Writes are conditional on the row's ETag, so two replicas refreshing the same
user cannot silently overwrite each other.

Releases before this kept every account in one row (partition `cache`, row
`msal-token-cache`). The server splits that row into per-account rows and
deletes it the first time any replica reads the cache after the upgrade; nobody
has to sign in again. During a rolling deploy a replica still on the old release
can write the shared row back. It is split again the next time a replica
starts, and an account that already has its own row keeps it. If the shared row
cannot be decrypted (a rotated `MCP_DATA_ENCRYPTION_KEY`), it is left in place
and logged; delete it, or run `infra/scripts/purge-credentials.sh`.

To sign one user out of Graph access everywhere, delete their row:

```
az storage entity delete --account-name <storage account> --table-name mcpMsalCache \
  --partition-key account --row-key '<object id>.<tenant id>'
```

Every replica sees the row gone on that user's next token refresh (within the
hour, as access tokens expire) and the user signs in again. This removes the
server's copy only; revoke the user's sessions in Entra as well if the refresh
token itself may have leaked.

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

Both keys, and `AZURE_CLIENT_SECRET`, live in Azure Key Vault. The Container
App holds Key Vault references to them, not the values, so the value lives in
one place with access control, versioning, soft delete and an access log, and
`az containerapp secret show` returns a URI. Reading a value takes a data-plane
role on the vault; Contributor on the resource group does not include one.

The Bicep templates set all of this up (`infra/key-vault.bicep`): an RBAC-mode
vault with soft delete and purge protection, the three secrets written from the
template parameters, a user-assigned **runtime identity** (`<app>-runtime`)
holding **Key Vault Secrets User** on the vault, and references resolved through
that identity. It is user-assigned because the app resolves its references
while it is being created, before a system-assigned identity could be granted
anything. The vault name defaults to `kv-` plus a hash of the resource group.

The commands below use:

```
VAULT=$(az keyvault list -g rg-m365-mcp --query "[0].name" -o tsv)   # the deployment's keyVaultName output
IDENTITY=$(az identity show -g rg-m365-mcp -n m365-mcp-runtime --query id -o tsv)
```

An instance set up by hand before the templates did this may bind its
references through the system-assigned identity instead. Use
`IDENTITY=system` there, or move it onto the templates.

**Moving an instance onto Key Vault by hand.** Store the *current* values to
keep everyone signed in; new values are a rotation (below).

```
az role assignment create --assignee-object-id "$(az identity show -g rg-m365-mcp -n m365-mcp-runtime --query principalId -o tsv)" \
  --assignee-principal-type ServicePrincipal \
  --role "Key Vault Secrets User" --scope "$(az keyvault show -n "$VAULT" --query id -o tsv)"

HMAC_URI=$(az keyvault secret set --vault-name "$VAULT" -n mcp-session-hmac-key \
  --value "<current-hmac-key>" --query id -o tsv)
DATA_URI=$(az keyvault secret set --vault-name "$VAULT" -n mcp-data-encryption-key \
  --value "<current-data-key>" --query id -o tsv)
SECRET_URI=$(az keyvault secret set --vault-name "$VAULT" -n azure-client-secret \
  --value "<current-client-secret>" --query id -o tsv)

az containerapp identity assign -n m365-mcp -g rg-m365-mcp --user-assigned "$IDENTITY"
az containerapp secret set -n m365-mcp -g rg-m365-mcp --secrets \
  "mcp-session-hmac-key=keyvaultref:${HMAC_URI},identityref:${IDENTITY}" \
  "mcp-data-encryption-key=keyvaultref:${DATA_URI},identityref:${IDENTITY}" \
  "azure-client-secret=keyvaultref:${SECRET_URI},identityref:${IDENTITY}"

az containerapp update -n m365-mcp -g rg-m365-mcp --revision-suffix "$(date +%Y%m%d%H%M)"
```

`--query id` returns the **versioned** secret URI, and that is deliberate. With
a versionless reference, writing a new version in the vault rotates the key on
whichever revision happens to start next, which signs everyone out at a time
nobody chose. With a versioned reference the key changes only when you change
the reference. The templates bind versioned URIs too.

Two traps:

- **Every Bicep deployment writes its key parameters into the vault** as a new
  version and points the references at it. Deploying with the values already
  in the vault is harmless; deploying with different ones is a rotation. Pass
  the current values (`az keyvault secret show --vault-name "$VAULT" -n <name>
  --query value -o tsv`), never fresh ones, unless you mean to rotate.
- **`az containerapp secret set name=<value>`** with a literal value silently
  turns a reference back into a plain Container App secret.

Confirm what is bound (values are not printed):

```
az containerapp secret list -n m365-mcp -g rg-m365-mcp \
  --query "[?keyVaultUrl!=null].{name:name,keyVaultUrl:keyVaultUrl,identity:identity}" -o table
```

`azure-client-secret`, `mcp-session-hmac-key` and `mcp-data-encryption-key`
should all appear.

Who read them: the vault sends its audit log to the instance's Log Analytics
workspace. Reads by anything other than the runtime identity deserve a look.

```
AzureDiagnostics
| where ResourceProvider == "MICROSOFT.KEYVAULT" and OperationName == "SecretGet"
| project TimeGenerated, identity_claim_oid_g, CallerIPAddress, id_s, ResultSignature
```

### Storage access

The server reaches Table Storage with the same runtime identity, which holds
**Storage Table Data Contributor** on the account and nothing else there. The
app gets `AZURE_STORAGE_TABLE_ENDPOINT` and `AZURE_STORAGE_IDENTITY_CLIENT_ID`
instead of a connection string, and `infra/main.bicep` creates the account
with shared-key access off, so no account key, connection string or key-signed
SAS works against it. `infra/container-app.bicep` grants the role on an account
you bring but cannot change that account's settings; turn shared-key access
off on it yourself (the last step below).

Operators read and purge tables with their own Entra identity
(`--auth-mode login`), which needs the same role on the account:

```
az role assignment create --assignee "<you@example.com>" \
  --role "Storage Table Data Contributor" \
  --scope "$(az storage account show -n stm365mcp -g rg-m365-mcp --query id -o tsv)"
```

**Moving an instance off the connection string.** Order matters: an app still
on the connection string stops working the moment shared-key access goes off.

  1. Deploy a release that reads `AZURE_STORAGE_TABLE_ENDPOINT` (any release carrying `src/services/storageClient.ts`). It keeps using the connection string until the endpoint is set.
  2. Re-run the Bicep, or by hand: assign the runtime identity to the app (`az containerapp identity assign`), grant it Storage Table Data Contributor on the account, and set the two variables:
```
az containerapp update -n m365-mcp -g rg-m365-mcp \
  --set-env-vars "AZURE_STORAGE_TABLE_ENDPOINT=https://stm365mcp.table.core.windows.net" \
  "AZURE_STORAGE_IDENTITY_CLIENT_ID=$(az identity show -g rg-m365-mcp -n m365-mcp-runtime --query clientId -o tsv)" \
  --remove-env-vars AZURE_STORAGE_CONNECTION_STRING \
  --revision-suffix "$(date +%Y%m%d%H%M)"
```
  3. Confirm `/health`, one sign-in and one tool call on the new revision, then drop the old secret: `az containerapp secret remove -n m365-mcp -g rg-m365-mcp --secret-names azure-storage-connection-string`.
  4. Turn shared-key access off and rotate both account keys, since the old connection string sat in the app's secrets and in deployment history:
```
az storage account update -n stm365mcp -g rg-m365-mcp --allow-shared-key-access false
az storage account keys renew -n stm365mcp -g rg-m365-mcp --key primary
az storage account keys renew -n stm365mcp -g rg-m365-mcp --key secondary
```

Sessions survive all four steps: the tables and the application keys are
unchanged.

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

Rotating the HMAC key also rotates the desktop extension's update-signing key,
which is derived from it (`src/services/extensionSigning.ts`). Installed
extensions carry the old public key, so they refuse every later update and keep
running the version they have. Users already have to sign in again; tell them to
reinstall the extension from `/install` at the same time, or they stay on that
version for good.

Each replica also holds decrypted sessions in memory. That is why a rotation
must be a new revision: the new replicas start empty, and the old ones are
retired once the new revision is ready.

Procedure, both keys (drop the line for the key you are not rotating):

```
# VAULT and IDENTITY as in "Storage: Key Vault references" above
HMAC_URI=$(az keyvault secret set --vault-name "$VAULT" -n mcp-session-hmac-key \
  --value "$(openssl rand -hex 32)" --query id -o tsv)
DATA_URI=$(az keyvault secret set --vault-name "$VAULT" -n mcp-data-encryption-key \
  --value "$(openssl rand -hex 32)" --query id -o tsv)

az containerapp secret set -n m365-mcp -g rg-m365-mcp --secrets \
  "mcp-session-hmac-key=keyvaultref:${HMAC_URI},identityref:${IDENTITY}" \
  "mcp-data-encryption-key=keyvaultref:${DATA_URI},identityref:${IDENTITY}"

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
The next Bicep deployment writes its own parameter values back into the vault,
so update the parameters to the new keys before it runs.

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
  2. **Revoke sessions in Entra for every user** of the app ([kill
     switch](incident-response.md#kill-switch-ordered-by-speed), step 2).
     This is what invalidates the stolen refresh tokens.
  3. Rotate `AZURE_CLIENT_SECRET` and delete the old one. Refresh tokens issued
     to this app can only be redeemed with its client credentials, which sit in
     the same vault as the keys, so assume they went together.
  4. Read the application audit log and Entra sign-in logs for the exposure
     window.
- **The vault itself, or the app's runtime identity.** Both keys and
  `AZURE_CLIENT_SECRET`, and with the identity also read and write on every
  table: do all of the above, and review the vault's and the storage account's
  role assignments before writing new versions into the vault.

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

The kill switch, ordered by speed, and what to do around it (detection sources,
triage, evidence to preserve before the kill switch destroys it, who to notify,
and the post-incident rotation list) are in
[`incident-response.md`](incident-response.md). In an incident, start there.

* * *

## Uninstall and teardown

  1. Pin the instance (`gh variable delete AUTO_UPDATE -R <fork>`) so nothing rolls mid-teardown.
  2. Tell users; the extension will report "Session expired" and can be removed from Claude Desktop.
  3. Export the audit log if the record is needed. The authoritative copy is `M365McpAudit_CL` in the Log Analytics workspace, which step 4 deletes; a SIEM that already ingests it keeps its own copy.
  4. `az group delete -n rg-m365-mcp`: Container App, environment, ACR, storage (every session, cache and audit row), the Log Analytics workspace (including the audit table) and the Key Vault go together. The vault has purge protection, so it and its secrets stay recoverable in soft delete for 90 days and its name cannot be reused until then; that is intended, since the data key in it is what decrypts any copy of the tables.
  5. Delete the app registration in Entra, or at least the client secret and the consent grants; every refresh token the instance held becomes unusable.
  6. Delete the fork, or leave it pinned as a record.

The server never stored tenant content, only tokens, policy and audit rows,
all of which lived in the deleted storage account.
