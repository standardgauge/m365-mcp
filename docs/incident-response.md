# Incident response

What the operator of an instance does when something looks wrong: a report
arrives, the audit log shows a user nobody recognises, a secret turns up where
it should not be. It covers where the signals come from, how to contain an
incident with the kill switch, who to tell, what to keep before it is gone, and
what to rotate afterwards.

It is written for the people who run an instance in their own Azure
subscription and Entra tenant. The project maintainers have no access to your
instance and cannot act on it for you; [`SECURITY.md`](../SECURITY.md) covers
what they do with a report about the code.

Commands use the placeholder names from the [operations
runbook](operations-runbook.md): Container App `m365-mcp`, resource group
`rg-m365-mcp`, storage account `stm365mcp`. Swap in your own.

## Before you need it

Write these down now, with your deployment's own notes, not in this repository:

- Who can act in **Entra**: disabling the enterprise application and deleting a
  client secret need Cloud Application Administrator or higher. Revoking
  sessions depends on who the user is: User Administrator is enough for a
  non-admin user, but a user who holds any Entra admin role needs Privileged
  Authentication Administrator or Global Administrator. Operators of this
  instance are usually in the second group, so the operator-compromise case
  below needs one of those two roles on call, not only User Administrator
  ([revoke access](https://learn.microsoft.com/en-us/entra/identity/users/users-revoke-access),
  [privileged roles](https://learn.microsoft.com/en-us/entra/identity/role-based-access-control/privileged-roles-permissions)).
- Who can act in **Azure** (Container App secrets and revisions, the storage
  account, Key Vault). Contributor on the resource group or higher.
- The instance's application (client) ID, its Container App outbound IP
  addresses (`az containerapp show -n m365-mcp -g rg-m365-mcp --query
  properties.outboundIpAddresses`), and where the Key Vault is.
- Who in your organisation decides on notifying users, regulators and
  customers.

If one person holds both the Entra and the Azure role, the kill switch takes
minutes. If they are two people in two teams, find out how to reach both out of
hours before the first incident, not during it.

Watch this repository for security advisories (Watch, Custom, Security alerts)
so a published fix reaches you. A pinned instance (`AUTO_UPDATE` unset) receives
nothing until you sync it.

## Detection sources

| Source | What to look for | Where |
|---|---|---|
| Application audit log | A user, device label or client IP you do not recognise; calls at hours the user does not work; a burst of denials (an agent probing deny lists); sends or writes the user did not ask for | `M365McpAudit_CL` in the instance's Log Analytics workspace, the authoritative record ([runbook](operations-runbook.md#audit-trail-in-log-analytics)); the admin UI's Audit Log and the `auditLog` table hold a working copy ([runbook](operations-runbook.md#application-audit-log)) |
| Entra sign-in logs, non-interactive | Refresh-token redemptions for the app from an address that is **not** one of the Container App's outbound IPs. Redeeming a refresh token for this app needs its client secret, so a redemption from elsewhere means a refresh token and the client secret are both out | Entra, Sign-in logs, Non-interactive user sign-ins, filtered to the application ID |
| Entra sign-in logs, interactive | Sign-ins to the app from unexpected locations or devices; a user signing in who should not have access | Same, Interactive user sign-ins |
| Entra audit logs | A client secret or certificate added to the app registration that you did not add; new owners; new API permissions or consent grants | Entra, Audit logs, filtered to the application |
| Azure Activity Log | Secret reads on the Container App (`listSecrets`), storage key listing (`listKeys`), new revisions or role assignments nobody made | The resource group's Activity log |
| Container logs | `Silent token refresh FAILED` for many users at once; decrypt errors; `[startup] FATAL`; a 401 spike | Log Analytics ([runbook](operations-runbook.md#log-access)) |
| `/health` | A `sha` that is not a commit you deployed | `curl -fsS https://<host>/health` |
| Deploy fork | Pushes to `main` or deploy runs you did not trigger; changed repository secrets or collaborators | The fork's Actions history and audit log |
| Microsoft Graph activity logs, Purview audit | Graph calls under the app's ID, mailbox access, if your tenant's licensing provides them and they are enabled | Entra diagnostic settings; Purview |
| Reports | A user, your security team, or a maintainer advisory | Wherever your organisation takes them |

The audit log and the Entra logs answer different questions. The audit log says
what the server did on a user's behalf. The Entra logs say who obtained tokens
for the app, including anyone who bypassed the server entirely with a stolen
refresh token and the client secret. An incident involving the client secret is
only visible in the second.

## Triage

Decide which case you are in. It sets which kill-switch steps you need.

| Case | Contain with | Then |
|---|---|---|
| **Vulnerability report, no sign of exploitation** | Nothing yet | Read the audit and Entra logs for the indicators the report describes. Apply the fix when it ships, or pin and take the mitigation the advisory gives. |
| **One user's session token stolen** (lost laptop, token in a log or a backup) | Disable the user in Entra and revoke their sessions (step 2), then rotate the HMAC key (step 3; that key alone is enough here) | See the note below on why revoking alone is not enough. |
| **Activity under one user that the user did not do** | As above, and read that user's audit rows and Graph activity for the window | Check the user's sent items and drafts; ask whether their MCP client had a prompt-injected session. |
| **Container App secrets exposed, or storage plus the data key** | Steps 1, 2, 3 and 5, in that order | The full rotation list below. |
| **Code you did not intend is running** (unexpected `sha`, compromised fork or upstream) | Pin (`gh variable delete AUTO_UPDATE -R <fork>`), stop traffic (step 6), roll back to a known-good image ([runbook](operations-runbook.md#rollback)) | That code ran with every secret in the Container App. Treat it as the row above. |
| **Azure or Entra account of an operator compromised** | Your organisation's own account-compromise process first | Then treat the instance's secrets as exposed. |

**A stolen session token outlives an Entra revocation.** Revoking a user's
sessions in Entra invalidates their refresh token, and the stolen session token
stops working within the hour. But every session for a user draws on one shared
MSAL account, and a failed refresh keeps the session row so that it can recover.
When the user signs in again, the stolen session row recovers with them, and
keeps working until its 30-day absolute cap. There is no admin action that
deletes one user's sessions, so the session row has to go another way: keep the
user disabled until the cap passes, or roll the HMAC key, which signs everyone
out.

## Kill switch, ordered by speed

Each step stands alone. With the client secret or the data key exposed, do 1, 2,
3 and 5. When in doubt, do 1 and 3 first: they take minutes and cut both the
Entra side and the server's own hold on tokens.

1. **Disable the enterprise application** in Entra: Enterprise applications,
   the app, Properties, "Enabled for users to sign in?" to No. New sign-ins and
   refresh-token redemptions stop at once. Access tokens already issued keep
   working until they expire, roughly an hour.
2. **Revoke sessions** in Entra for the affected users, or for every user of the
   app. This invalidates their refresh tokens, including copies taken from
   storage. With step 1, nobody obtains new Graph access once current access
   tokens expire.
   If an affected user holds an Entra admin role, this step needs Privileged
   Authentication Administrator or Global Administrator ([Before you need
   it](#before-you-need-it)).
3. **Rotate both application keys** in one new revision ([Application
   keys](operations-runbook.md#rotation)). Every stored session and the MSAL
   cache become undecryptable, every client is signed out, and the new replicas
   start with empty memory. The server's own hold on tokens ends regardless of
   Entra.
4. **Purge the credential tables**: `infra/scripts/purge-credentials.sh
   --resource-group rg-m365-mcp --app-name m365-mcp`, with `--dry-run` first.
   Same effect on storage as step 3, with no key change. Replicas re-check a
   cached session against storage at most every 30 seconds and drop it once its
   row is gone, so purged sessions stop working everywhere within 30 seconds
   ([threat model](threat-model.md) row 3.3). If storage is unreachable, a
   replica keeps serving a cached session for up to 5 minutes. To end that
   window at once, roll a new revision afterwards (`az containerapp update -n
   m365-mcp -g rg-m365-mcp --revision-suffix "$(date +%Y%m%d%H%M)"`).
5. **Delete the client secret** on the app registration. The server, and anyone
   holding a copy of the secret, can no longer redeem codes or refresh tokens.
   The instance cannot sign anyone in until you add a new one.
6. **Stop traffic**: `az containerapp ingress disable -n m365-mcp -g
   rg-m365-mcp`, or scale to zero. Prefer these to deleting the app or the
   resource group, which destroys the evidence below.

Steps 3 and 4 delete or orphan the session rows, and step 1 shuts admins out of
the admin UI within the hour. If the incident is not actively in progress, take
the evidence first. If it is, contain first and take what survives.

## Evidence to preserve

Record every action you take with a UTC timestamp as you go. Keep exports out of
the instance's own storage account and away from anyone whose access is in
question.

| What | Why it is at risk | How |
|---|---|---|
| Audit rows for the window | The `M365McpAudit_CL` table is append-only to the server and survives key rotation and purges, but Log Analytics retention (90 days by default) still ages it out. The storage copy is weaker: retention deletes rows past `AUDIT_LOG_RETENTION_DAYS`, and anyone with the storage key can edit them | Query `M365McpAudit_CL` for the window and export the results ([runbook](operations-runbook.md#audit-trail-in-log-analytics)); this needs no admin session, so it still works after step 1 or 3. For the storage copy, export through the admin UI or `/api/manage/audit-log?startDate=…&endDate=…&format=csv` ([runbook](operations-runbook.md#audit-log-retention)). The API needs a live admin session; once step 1 or 3 has run, read the table directly: `az storage entity query --account-name stm365mcp --table-name auditLog --auth-mode login --filter "Timestamp ge datetime'<start>'"`. Raise the retention window, or set it to `0`, until the incident is closed. |
| Who had sessions, from which device | Steps 3 and 4 delete or orphan the rows | Copy the non-secret columns only: `az storage entity query --account-name stm365mcp --table-name mcpSessions --auth-mode login --select email,userId,deviceLabel,sessionCreatedAt,sessionAbsoluteCreatedAt,Timestamp`. Do not copy the ciphertext columns; a copy of the envelopes is a second thing to protect. |
| Entra sign-in and audit logs | Entra keeps them 30 days with a P1 or P2 licence and 7 without, unless a diagnostic setting already sends them to a workspace | Download for the window, filtered to the application ID, interactive and non-interactive. |
| Container logs | Log Analytics retention, 90 days by default | Export the query results for the window ([runbook](operations-runbook.md#logs-via-azure-portal)). |
| Azure Activity Log | 90 days | Export for the resource group and the Key Vault. |
| What was running | A rollback or redeploy replaces it | The `/health` body, the revision list with images ([runbook](operations-runbook.md#revision-history)), the image digests in the registry, and the fork's commit history and Actions runs. |
| Key Vault access log | Only exists if diagnostic settings were enabled | Export it if it exists; if it does not, note that in the record. |

## Who to notify

- **Your own security function**, at once. They own the incident; this page is
  the instance-specific part of their process.
- **Whoever holds the other half of the kill switch**: the Entra administrator
  or the Azure owner, whichever you are not.
- **Affected users.** After step 3 every user is signed out and the extension
  reports an expired session; tell them before they open tickets. They sign in
  again from `https://<host>/install` once the application is re-enabled. If
  their access was used, ask them to check sent items, drafts, calendar
  responses and shared files for the window.
- **Your privacy or legal function**, if a user's delegated access may have
  been used by someone else. Mail, files and calendar content reachable by that
  user may have been read, which can start notification clocks under your
  obligations. That decision is theirs, and the clock can run from discovery,
  so tell them early rather than once the investigation is complete.
- **The project maintainers**, privately through [`SECURITY.md`](../SECURITY.md),
  if the cause is or might be in this code, its container or its templates.
  Not for incidents whose cause is local, such as a leaked secret. Never in a
  public issue.
- **Microsoft**, through [MSRC](https://msrc.microsoft.com/report), if Entra ID
  or Graph behaved differently from how they document it.

## Post-incident rotation

Rotate everything the exposure could have reached. With Container App secrets
or the runtime exposed, that is the whole list: they sit side by side.

| Credential | How | Notes |
|---|---|---|
| `AZURE_CLIENT_SECRET` | [Runbook](operations-runbook.md#azure_client_secret-recommended-every-12-months) | Delete the old secret, not just add a new one. Then remove any other secret or certificate on the registration you did not create. |
| `MCP_SESSION_HMAC_KEY`, `MCP_DATA_ENCRYPTION_KEY` | [Runbook](operations-runbook.md#rotation) | Both, in one revision, then purge. Disable the previous versions in Key Vault. |
| Storage account keys | `az storage account keys renew -g rg-m365-mcp -n stm365mcp --key secondary`, point `azure-storage-connection-string` at the secondary key and roll a revision, then renew the primary | Both keys: the connection string could have held either. |
| Entra refresh tokens | Revoke sessions for every user of the app (kill switch step 2) | The only thing that invalidates refresh tokens already copied out of storage. |
| `AZURE_CREDENTIALS` (deploy principal) | [Runbook](operations-runbook.md#azure_credentials-deploy-service-principal) | Also review the principal's role assignments. |
| Sync token and any other fork secret | Mint new, `gh secret set` on the fork, revoke the old | Plus the fork's collaborators and deploy keys. |
| Users' client session tokens | Invalidated by the HMAC key rotation | Users reinstall from `/install`. |

Then review what an attacker could have changed rather than read: the app
registration's owners, permissions and consent grants; role assignments on the
resource group and the Key Vault; the admin policy (services, deny lists,
allowed sites, mail settings), which is not audited when changed; and the images
in the registry.

## Recovery

1. Confirm `/health` serves the commit you expect.
2. Re-enable the enterprise application, or re-enable ingress.
3. Sign in once and make one tool call per enabled service.
4. Tell users to reinstall.
5. Restore the audit retention window you raised.

Afterwards, write up the timeline, cause and changes in your own records. If the
cause was in this code, or exposed a gap the [threat model](threat-model.md)
does not list, say so in your report to the maintainers so the threat model can
be updated.
