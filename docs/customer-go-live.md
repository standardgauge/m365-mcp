# Customer go-live checklist

The ordered list of what has to be true before a new tenant's instance takes
its first user. Tick each item in order, run its check, and write the result
in the [go-live record](#go-live-record) at the end. An item without a passed
check is not done, however confident the person who did it feels.

**Why this exists.** The app registration every tenant consents to carries
`Sites.ReadWrite.All`, `Files.ReadWrite.All`, `Mail.ReadWrite` and
`Contacts.ReadWrite`, tenant-wide, for every signed-in user, plus
`Directory.Read.All` for the admin check. That breadth is inherent to what the
server does (see the [threat model](threat-model.md), row 1.3), and it makes
each deployment a high-trust registration in the customer's tenant. A
registration that broad is acceptable only when the deployment controls around
it are actually in place, every time, and this list is how that is shown
rather than assumed.

**Who runs it.** One person can run the whole list if they hold all three
roles below. More often it is two or three people; the record says who did
which item.

| Role | Needed for | Items |
|---|---|---|
| Owner, or Contributor plus User Access Administrator, on the resource group | The Bicep creates role assignments; Contributor alone cannot | 1, 5, 7 |
| Cloud Application Administrator or higher in Entra | Tenant-wide admin consent | 2 |
| Global Administrator in Entra | The admin UI unlocks for this role only | 8, 10 |
| Owner of the DNS zone | The CNAME for the custom hostname | 3 |
| Admin on the deploy fork | Repository variables and secrets | 4 |

**Placeholders.** Commands use the names the templates derive from
`appName=m365-mcp`: Container App `m365-mcp`, resource group `rg-m365-mcp`,
storage account `stm365mcp`, runtime identity `m365-mcp-runtime`, workspace
`m365-mcp-logs`, data collection rule `m365-mcp-audit-dcr`, and the hostname
`your-mcp-host.example.com`. Swap in your own. Nothing about a real deployment
belongs in this repository; the record lives with the deployment's own notes.

**What this list does not do.** It does not explain the controls; the
[README](../README.md), [`entra-setup.md`](entra-setup.md) and the
[operations runbook](operations-runbook.md) do, and each item links to the
section that does. It does not cover day-two operation, rotation or incidents:
[`operations-runbook.md`](operations-runbook.md) and
[`incident-response.md`](incident-response.md). Four improvements that would
change items on this list are tracked separately and listed at the
[end](#tracked-separately), so they are not restated here.

---

## Contents

- [Before you start](#before-you-start)
- [1. Provision with the Bicep: managed identity and Key Vault references](#1-provision-with-the-bicep-managed-identity-and-key-vault-references)
- [2. Entra registration and consent, matched to the services you will enable](#2-entra-registration-and-consent-matched-to-the-services-you-will-enable)
- [3. Hostname, redirect URI and frontend URL](#3-hostname-redirect-uri-and-frontend-url)
- [4. Deploy fork and first image](#4-deploy-fork-and-first-image)
- [5. Storage reached by identity, never by key](#5-storage-reached-by-identity-never-by-key)
- [6. Application keys unique to this tenant and environment](#6-application-keys-unique-to-this-tenant-and-environment)
- [7. Audit trail landing, retention recorded](#7-audit-trail-landing-retention-recorded)
- [8. Policy set before any agent gets access](#8-policy-set-before-any-agent-gets-access)
- [9. Client surface: served origin and update signing](#9-client-surface-served-origin-and-update-signing)
- [10. First user, first call, sign-off](#10-first-user-first-call-sign-off)
- [Tracked separately](#tracked-separately)
- [Go-live record](#go-live-record)

---

## Before you start

Decide these and write them in the record. Several later items depend on them,
and changing them after users have signed in costs a re-sign-in or a consent
round.

- [ ] **Subscription, region and resource group name.** Who holds Owner on the
      resource group.
- [ ] **Hostname** (`your-mcp-host.example.com`) and who owns its DNS zone.
- [ ] **Services to enable.** The default is mail and SharePoint only. Each
      service you add is a consent decision in item 2 and a policy decision in
      item 8.
- [ ] **Read-only services.** Which enabled services stay readable but never
      writable.
- [ ] **Closed or open SharePoint posture.** An empty site allow-list allows
      every site. If the tenant wants closed, you need at least one site ID
      before item 8.
- [ ] **Mail folders and SharePoint paths denied for everyone**, at the
      deployment layer (`DEFAULT_MAIL_DENY_FOLDERS`,
      `DEFAULT_SHAREPOINT_DENY_PATHS`) and in the global deny lists.
- [ ] **Draft mode enforced tenant-wide, or left to each user.** Enforced means
      no agent can send mail directly for anyone.
- [ ] **Outbound policy** for calendar invitations, meeting-response comments
      and Teams messages: allow, internal only, or block. The default is allow.
- [ ] **Audit retention**: how long the audit table and the workspace keep
      rows, and whether a SIEM ingests the table.
- [ ] **Tracking or pinned.** Whether the instance deploys each upstream commit
      as CI passes on it (`AUTO_UPDATE=true`) or only on manual dispatch.
      Pinned is the right answer for a tenant under change control. Leave it
      pinned until item 10 passes either way.
- [ ] **Who is on call** for the kill switch: the Entra administrator and the
      Azure owner, by name ([incident response](incident-response.md#before-you-need-it)).

---

## 1. Provision with the Bicep: managed identity and Key Vault references

The only supported path is `infra/main.bicep` (full stack) or
`infra/container-app.bicep` (bring your own storage account). Both put the
client secret and the two application keys in an RBAC-mode Key Vault with
purge protection, bind them to the Container App as versioned Key Vault
references resolved through a user-assigned runtime identity, and reach Table
Storage with that identity rather than a key. A hand-built app that sets
`az containerapp secret set name=<literal>` has plain Container App secrets and
does not pass this item.

- [ ] **1.0 Do item 2.1 first.** The template takes the client ID, tenant ID
      and client secret as parameters and writes the secret to the vault.
      Create the registration and its secret before this item, so the
      template runs once with real values. Every run writes its parameters
      into the vault as a new version, and a re-run with different values is
      a rotation, so one run is the goal.

- [ ] **1.1 Generate the two application keys inline.** Never from a file,
      never copied from another deployment. Item 6 checks the result. Keep
      this shell open through 1.2; the variables exist nowhere else.

  ```bash
  RG=rg-m365-mcp; LOC=<region>; APP=m365-mcp
  az group create --name "$RG" --location "$LOC"
  HMAC=$(openssl rand -hex 32); DEK=$(openssl rand -hex 32)
  ```

- [ ] **1.2 Deploy the template.**

  ```bash
  az deployment group create --resource-group "$RG" \
    --template-file infra/main.bicep \
    --parameters appName="$APP" location="$LOC" \
      customDomain="your-mcp-host.example.com" \
      azureClientId="<client-id>" azureClientSecret="<client-secret>" azureTenantId="<tenant-id>" \
      mcpSessionHmacKey="$HMAC" mcpDataEncryptionKey="$DEK" \
    --query properties.outputs
  unset HMAC DEK
  ```

  Keep the outputs: `keyVaultName`, `runtimeIdentityId`, `acrLoginServer`,
  `containerAppFqdn`, `logAnalyticsWorkspaceName`,
  `auditDataCollectionRuleName`.

  The registry is empty at this point, so the Container App's first revision
  has no image to pull and the deployment may report the Container App as
  failed. That is expected: the resource exists, item 4 gives it an image,
  and the one resource downstream of the app in the template (the audit
  publisher role assignment on the app's identity) is checked and, if
  missing, added by hand in item 7.2. Do not re-run the template to fix it.

- [ ] **1.3 Check: the three secrets are Key Vault references.** Values are not
      printed.

  ```bash
  az containerapp secret list -n "$APP" -g "$RG" \
    --query "[?keyVaultUrl!=null].{name:name,keyVaultUrl:keyVaultUrl,identity:identity}" -o table
  ```

  Expected: `azure-client-secret`, `mcp-session-hmac-key` and
  `mcp-data-encryption-key`, each with a `keyVaultUrl` that ends in a version
  segment and an `identity` naming `m365-mcp-runtime`. A row missing from the
  list, or a URL with no version, fails the item. See
  [Storage: Key Vault references](operations-runbook.md#storage-key-vault-references).

- [ ] **1.4 Check: the vault is RBAC mode with purge protection, and only the
      runtime identity can read it.**

  ```bash
  VAULT=$(az keyvault list -g "$RG" --query "[0].name" -o tsv)
  az keyvault show -n "$VAULT" --query "{rbac:properties.enableRbacAuthorization,purge:properties.enablePurgeProtection,softDelete:properties.enableSoftDelete}"
  az role assignment list --scope "$(az keyvault show -n "$VAULT" --query id -o tsv)" \
    --query "[].{role:roleDefinitionName,principal:principalName}" -o table
  ```

  Expected: all three `true`; one `Key Vault Secrets User` assignment, for
  `m365-mcp-runtime`. Any human or other principal with a data-plane role on
  the vault goes in the record with a reason, or comes off.

- [ ] **1.5 Record** the resource names, the vault name, the identity resource
      ID and who holds Owner on the resource group.

---

## 2. Entra registration and consent, matched to the services you will enable

The permission table and its reasoning are in the README under
[Entra ID app registration and permissions](../README.md#entra-id-app-registration-and-permissions);
operational detail and the verification commands are in
[`entra-setup.md`](entra-setup.md). This item makes sure the grant matches the
services decided in [Before you start](#before-you-start), and nothing more.

- [ ] **2.1 Create the registration** per the README: single tenant, web
      redirect URI `https://your-mcp-host.example.com/api/auth/callback`, ID
      tokens enabled, logout URL `https://your-mcp-host.example.com/api/auth/logout`,
      one client secret. Put the secret's expiry date in the record and in a
      calendar now; expiry takes the whole instance down at once.

- [ ] **2.2 Add exactly the permissions the enabled services need.** Two
      groups. The first is non-negotiable: every entry of `GRAPH_SCOPES` in
      `src/services/graphClient.ts` is requested at sign-in and must be
      consented or nobody can sign in (the README names the current set and
      a test keeps that sentence in step with the array). The second is per
      service, and a permission for a service you are not enabling stays off
      the registration. The README table is authoritative; this grouping is a
      reading of it.

  | Enabling | Add, beyond the sign-in set |
  |---|---|
  | Sign-in and admin check (always) | `openid`, `profile`, `offline_access`, `User.Read`, `Directory.Read.All`, and the rest of `GRAPH_SCOPES`: `Mail.ReadWrite`, `Mail.ReadBasic`, `Sites.ReadWrite.All`, `Files.ReadWrite.All`, `Contacts.ReadWrite` |
  | Mail | `Mail.Send` (direct sends and `send_draft`; omit it if draft mode is enforced tenant-wide and direct send is never wanted), `User.ReadBasic.All` (resolving other users, delegated mailbox owners) |
  | Calendar | `Calendars.ReadWrite`, `Calendars.Read.Shared`, `Place.Read.All`, `MailboxSettings.Read` |
  | SharePoint, OneDrive, Contacts | Nothing beyond the sign-in set |
  | OneNote | `Notes.ReadWrite.All` |
  | Teams | `Team.ReadBasic.All`, `Channel.ReadBasic.All`, `ChannelMessage.Read.All`, `ChannelMessage.Send`, `ChatMessage.Send`; plus `ChatMember.Read` and `ChannelMember.Read.All` if the outbound policy will set Teams to internal only |

  Do not add anything to `GRAPH_SCOPES` for this tenant. The README explains
  why; the short version is that it breaks sign-in on every other deployment
  of the same image.

- [ ] **2.3 Grant tenant-wide admin consent**, once, for the whole set.

- [ ] **2.4 Check: the grant, not the click.**

  ```bash
  az ad app permission list-grants --id <client-id> --query "[].scope" -o tsv | tr ' ' '\n' | sort
  ```

  Expected: every `GRAPH_SCOPES` entry present, every permission from the
  rows you ticked in 2.2 present, and nothing from the rows you did not
  tick. A scope on the registration but missing here was added without
  consent being re-granted ([verifying consent](entra-setup.md#verifying-consent-actually-landed)).

- [ ] **2.5 Check, once an image is running (after item 4):** the live sign-in
      redirect names this tenant and requests exactly the sign-in set.

  ```bash
  curl -sI https://your-mcp-host.example.com/api/auth/login | grep -i '^location:'
  ```

  Expected: the authority contains your tenant ID and `scope=` lists the
  `GRAPH_SCOPES` entries plus `openid profile offline_access`. Another tenant
  ID means `AZURE_TENANT_ID` on the app is wrong.

- [ ] **2.6 Limit who can sign in.** On the enterprise application, set
      *Assignment required* and assign the pilot users or group. Without it
      every licensed member of the tenant can install and sign in.

- [ ] **2.7 Confirm Conditional Access is not excluding the app.** It needs no
      exclusion ([Conditional Access](entra-setup.md#conditional-access)).

- [ ] **2.8 Record** the client ID, tenant ID, secret expiry, the consented
      scope list from 2.4, and who granted consent.

---

## 3. Hostname, redirect URI and frontend URL

- [ ] **3.1 Publish the CNAME** from `your-mcp-host.example.com` to the
      `containerAppFqdn` output of item 1, then bind it:

  ```bash
  az containerapp hostname add  -n "$APP" -g "$RG" --hostname your-mcp-host.example.com
  az containerapp hostname bind -n "$APP" -g "$RG" --hostname your-mcp-host.example.com \
    --environment "$APP-env" --validation-method CNAME
  ```

- [ ] **3.2 Set the two URL secrets to the custom host, and the frontend URL
      to the site root.** The template writes `frontend-url` as
      `https://<host>/admin`, a path the Azure Functions host reserves, and
      every login ends on an empty 404 until it is the root
      ([reserved paths](operations-runbook.md#reserved-paths)).

  ```bash
  az containerapp secret set -n "$APP" -g "$RG" --secrets \
    oauth-redirect-uri="https://your-mcp-host.example.com/api/auth/callback" \
    frontend-url="https://your-mcp-host.example.com/"
  ```

  A plain restart does not pick up a secret change; item 4's deploy creates
  the new revision that does.

- [ ] **3.3 Check**, after item 4: `curl -sI https://your-mcp-host.example.com/`
      returns HTTP/2 200 with a `strict-transport-security` header, and

  ```bash
  az containerapp exec -n "$APP" -g "$RG" --command "printenv OAUTH_REDIRECT_URI FRONTEND_URL"
  ```

  prints the two values from 3.2, exactly. The value the container resolved
  is the one that matters, not the one the platform reports for the secret.

- [ ] **3.4 Ingress limits set and shown to hold.** Container Apps ingress
      has no rate limit and no body cap, and the app's own limiter is per
      replica, so with three replicas an address gets up to three times each
      route's limit. Set the three limits in
      [production ingress](operations-runbook.md#production-ingress-edge-rate-limit-and-body-size):
      a request body cap equal to `MCP_MAX_BODY_BYTES`, 16 MiB by default
      (`FUNCTIONS_REQUEST_BODY_SIZE_LIMIT`, and at the edge if there is one), an edge rate limit on Front Door Premium or an
      Application Gateway WAF, and ingress restricted to the edge so the
      `*.azurecontainerapps.io` name cannot walk around it, with
      `RATE_LIMIT_TRUSTED_PROXY_HOPS=2`. Behind Front Door the restriction is
      Private Link with the environment's public network access disabled. An
      `Allow` rule for the `AzureFrontDoor.Backend` service tag does not
      count: those ranges admit every Front Door profile in Azure, not only
      this one, and nothing on the Container Apps side can check
      `X-Azure-FDID`. Front Door Standard is therefore not accepted here.
      The edge body cap is a `Content-Length` custom rule using the
      runbook's `RegEx` pattern, not `GreaterThan`: Front Door reads
      `GreaterThan` on a header as the length of the header string, so a
      numeric threshold there never fires, and Application Gateway documents
      no numeric semantics for it either. On an Application Gateway that
      rule goes with WAF request body enforcement off, because the WAF's own
      `maxRequestBodySizeInKb` stops at 2 MB; and its rate-limit counters are
      per gateway instance, so record the gateway's maximum instance count and
      the per-address ceiling it gives (threshold times that count).

  Check, after item 4, with the runbook's three verification commands: an
  over-cap body gets `403` from the edge (`413` only with no edge; a `413`
  behind an edge means the edge rule did not fire), a burst against `/health`
  gets `429`s, and the direct FQDN is refused. Add the runbook's WAF policy
  check for the edge in use (the `Content-Length` rule present with the
  `RegEx` pattern; on an Application Gateway, enforcement off too). Behind an
  Application Gateway, scale the burst to the gateway's instance count.
  Behind Front Door, add the runbook's Private Link
  checks: the origin's private link approved and the environment's public
  network access `Disabled`. An instance with no edge records the body check
  and writes "no edge" against the other two, with who accepted that. That
  instance's only flood control is the per-replica limiter
  ([threat model](threat-model.md#10-availability), row 10.1).

---

## 4. Deploy fork and first image

Deployments run from a fork whose `main` is always equal to this repository's
`main` ([`AGENTS.md`](../AGENTS.md)). The deploy workflow in the fork builds
the image, pushes it to the tenant's registry, rolls the Container App, and
then proves the deploy: `/health` has to report the commit it built and the
frontend URL has to serve the admin UI. Use it rather than the manual
`docker build` path in the README; the proof is the point.

- [ ] **4.1 Fork the repository** into the tenant's organisation and confirm
      the fork's `main` equals upstream `main`:

  ```bash
  git ls-remote https://github.com/<owner>/<fork>.git refs/heads/main
  git ls-remote https://github.com/standardgauge/m365-mcp.git refs/heads/main
  ```

- [ ] **4.2 Set the repository variables** the workflow header in
      [`.github/workflows/deploy.yml`](../.github/workflows/deploy.yml)
      documents: `DEPLOY_ENABLED=true`, `AZURE_RESOURCE_GROUP`,
      `CONTAINER_APP_NAME`, `ACR_NAME`. Leave `AUTO_UPDATE` unset until item
      10 passes, so nothing deploys unattended before the instance is signed
      off.

- [ ] **4.3 Set the deploy credential** as the repository secret
      `AZURE_CREDENTIALS`: a service principal with Contributor on the
      resource group and AcrPush on the registry, piped so the value is never
      displayed.

  ```bash
  SUB=<subscription-id>
  az ad sp create-for-rbac --name "gh-actions-m365-mcp" --role Contributor \
    --scopes "/subscriptions/$SUB/resourceGroups/$RG" --sdk-auth \
    | gh secret set AZURE_CREDENTIALS -R <owner>/<fork>
  SP=$(az ad sp list --display-name "gh-actions-m365-mcp" --query '[0].id' -o tsv)
  az role assignment create --assignee "$SP" --role AcrPush \
    --scope "$(az acr show -n <acr-name> -g "$RG" --query id -o tsv)"
  ```

  Contributor on the resource group is the whole reach of this principal;
  it holds no data-plane role on the vault or the storage account. A
  federated (secretless) form of this credential is tracked separately.

- [ ] **4.4 Dispatch the first deploy by hand and watch it to the end.**

  ```bash
  gh workflow run Deploy -R <owner>/<fork>
  gh run watch -R <owner>/<fork>
  ```

  A green run means the post-deploy smoke passed. A red run at the smoke
  step with a stale `sha` usually means the new revision never became ready;
  read `[startup]` lines in the console log (item 6.3). A red run saying the
  frontend URL's path is reserved means item 3.2 did not land.

- [ ] **4.5 Confirm the first revision is the only revision taking traffic**
      and that the failed image-less one from item 1 is gone:

  ```bash
  az containerapp revision list -n "$APP" -g "$RG" \
    --query "[].{name:name,active:properties.active,health:properties.healthState,image:properties.template.containers[0].image}" -o table
  ```

  Expected: one active, `Healthy` revision on the image the run built.

- [ ] **4.6 Check: what is running is what you built.**

  ```bash
  curl -fsS https://your-mcp-host.example.com/health
  az containerapp show -n "$APP" -g "$RG" --query 'properties.template.containers[0].image' -o tsv
  ```

  Expected: `{"status":"ok","sha":"<commit>"}` as `application/json`, with
  `sha` equal to the fork's `main`, and an image tagged `sha-<short>` of the
  same commit. `/health` is the only endpoint that says what is running
  ([health probe](operations-runbook.md#health-probe)).

- [ ] **4.7 Record** the fork, the deploy principal's name and role scopes,
      the deployed commit and image tag, and the date.

---

## 5. Storage reached by identity, never by key

`infra/main.bicep` creates the account with shared-key access off and gives the
runtime identity Storage Table Data Contributor on it. The server reads
`AZURE_STORAGE_TABLE_ENDPOINT` and `AZURE_STORAGE_IDENTITY_CLIENT_ID`; when the
endpoint is set it wins, and a connection string is ignored
(`src/services/storageClient.ts`). This item proves both halves on the live
instance. Background: [Storage access](operations-runbook.md#storage-access).

- [ ] **5.1 Check: the account refuses keys.**

  ```bash
  az storage account show -n stm365mcp -g "$RG" \
    --query "{sharedKey:allowSharedKeyAccess,oauthDefault:defaultToOAuthAuthentication,https:supportsHttpsTrafficOnly,tls:minimumTlsVersion,publicBlob:allowBlobPublicAccess}"
  ```

  Expected: `sharedKey: false`, `oauthDefault: true`, `https: true`,
  `tls: TLS1_2`, `publicBlob: false`. With `container-app.bicep` on an
  account you brought, the template cannot change the account's settings:
  turn shared-key access off yourself and rotate both account keys, as the
  runbook's last step describes.

- [ ] **5.2 Check: the app is on the endpoint, and no connection string is
      present anywhere on it.**

  ```bash
  az containerapp show -n "$APP" -g "$RG" \
    --query "properties.template.containers[0].env[?starts_with(name,'AZURE_STORAGE')].{name:name,value:value,secretRef:secretRef}" -o table
  az containerapp secret list -n "$APP" -g "$RG" --query "[].name" -o tsv | grep -c connection-string
  ```

  Expected: `AZURE_STORAGE_TABLE_ENDPOINT` set to
  `https://stm365mcp.table.core.windows.net/` and
  `AZURE_STORAGE_IDENTITY_CLIENT_ID` set to the runtime identity's client ID;
  no `AZURE_STORAGE_CONNECTION_STRING` row; the `grep -c` prints `0`. If the
  variable were present, the endpoint would still take precedence, but a
  connection string that exists is a key that exists, and this item is about
  there being none.

- [ ] **5.3 Check: the identity holds exactly its one role on the account.**

  ```bash
  az role assignment list --scope "$(az storage account show -n stm365mcp -g "$RG" --query id -o tsv)" \
    --query "[].{role:roleDefinitionName,principal:principalName}" -o table
  ```

  Expected: `Storage Table Data Contributor` for `m365-mcp-runtime`. Any
  operator who needs to read or purge tables gets the same role under their
  own identity and uses `--auth-mode login`; record each one.

- [ ] **5.4 Check: it works.** The sign-in in item 8.1 writes a session row.
      A `403 AuthorizationPermissionMismatch` or
      `KeyBasedAuthenticationNotPermitted` in the console log means one of
      5.1 to 5.3 is not actually true
      ([troubleshooting](operations-runbook.md#issue-azure-table-storage-connection-errors)).

---

## 6. Application keys unique to this tenant and environment

`MCP_SESSION_HMAC_KEY` hashes every session token at rest and derives the
extension update-signing key; `MCP_DATA_ENCRYPTION_KEY` encrypts every stored
Graph access token and the MSAL cache of refresh tokens. A key shared between
two deployments lets a storage copy of one be read with the other's key, and a
key shared between production and a test environment puts production's refresh
tokens one test-environment compromise away. The threat model marks reuse as
documented but not enforced (row 2.6); this item is the enforcement.
Background: [Application keys](operations-runbook.md#application-keys).

- [ ] **6.1 Confirm the provenance.** Both keys were generated by
      `openssl rand -hex 32` in item 1.1, in the shell that ran the
      deployment, and were not written to a file, a ticket, a password
      manager entry shared with another deployment, or a parameters file
      checked in anywhere. If any of that is not true, rotate now, before
      users exist, following [Rotation](operations-runbook.md#rotation); it
      costs nothing on an instance with no sessions.

- [ ] **6.2 Check: the two keys differ from each other, without printing
      either.**

  ```bash
  for s in mcp-session-hmac-key mcp-data-encryption-key; do
    printf '%s  ' "$s"; az keyvault secret show --vault-name "$VAULT" -n "$s" --query value -o tsv | tr -d '\n' | sha256sum | cut -c1-16
  done
  ```

  Expected: two different prefixes. Keep them in the record: they are what
  lets a later reviewer confirm a key was not reused in another deployment,
  by comparing prefixes rather than values. (This read needs Key Vault
  Secrets User on the vault for the person running it; remove that
  assignment afterwards, and note the read will appear in the vault's audit
  log, which is correct.)

- [ ] **6.3 Check: the container validated both keys at startup.**

  ```bash
  az containerapp logs show -n "$APP" -g "$RG" --type console --tail 200 | grep '\[startup\]'
  ```

  Expected: `[startup] MCP_SESSION_HMAC_KEY and MCP_DATA_ENCRYPTION_KEY validated`
  and no `[startup] FATAL`. The image's `CMD` runs this check before the
  Functions host, so a revision with a bad key never becomes ready.

- [ ] **6.4 Turn on strict envelope binding from day one.** A new instance has
      no rows written by older releases, so there is nothing to migrate and
      no reason to accept unbound envelopes:

  ```bash
  az containerapp update -n "$APP" -g "$RG" --set-env-vars \
    MCP_ENVELOPE_REQUIRE_AAD=true MCP_SESSION_REQUIRE_IDENTITY_BINDING=true \
    --revision-suffix "$(date +%Y%m%d%H%M)"
  ```

  Check with `az containerapp exec ... --command "printenv MCP_ENVELOPE_REQUIRE_AAD MCP_SESSION_REQUIRE_IDENTITY_BINDING"`.
  Background: [credential envelope row binding](operations-runbook.md#credential-envelope-row-binding-mcp_envelope_require_aad).

- [ ] **6.5 Record** the two hash prefixes, the date, and who generated the
      keys. Diarise nothing: these keys do not expire, and rotation is a
      deliberate act that signs every user out.

---

## 7. Audit trail landing, retention recorded

The authoritative audit record is the `M365McpAudit_CL` table in the tenant's
own Log Analytics workspace, fed through a data collection rule the Container
App's system identity may publish to and nothing else; a working copy backs the
admin screen from the `auditLog` storage table. The runbook explains both and
how to point a SIEM at the table ([audit trail in Log
Analytics](operations-runbook.md#audit-trail-in-log-analytics)). A scheduled
retention mechanism for the storage copy and a verified workspace retention on
every live instance are tracked separately; this item confirms the pipeline is
alive on this instance and records what retention it actually has today.

- [ ] **7.1 Check: the app knows where to send rows.**

  ```bash
  az containerapp show -n "$APP" -g "$RG" \
    --query "properties.template.containers[0].env[?starts_with(name,'AUDIT_')].{name:name,value:value}" -o table
  ```

  Expected: `AUDIT_LOGS_INGESTION_ENDPOINT`, `AUDIT_DCR_IMMUTABLE_ID` and
  `AUDIT_DCR_STREAM_NAME=Custom-M365McpAudit`. Without the first two the
  server writes the storage table only and says nothing about it.

- [ ] **7.2 Check: the system identity holds the publisher role on the rule,
      and only there.**

  ```bash
  PRINCIPAL=$(az containerapp show -n "$APP" -g "$RG" --query identity.principalId -o tsv)
  DCR_ID=$(az resource show -g "$RG" -n m365-mcp-audit-dcr --resource-type Microsoft.Insights/dataCollectionRules --query id -o tsv)
  az role assignment list --assignee "$PRINCIPAL" --all --query "[].{role:roleDefinitionName,scope:scope}" -o table
  ```

  Expected: one row, `Monitoring Metrics Publisher`, scoped to `$DCR_ID`. If
  it is missing (the first template run in item 1 could not create it before
  the app had an identity), add it now:

  ```bash
  az role assignment create --assignee-object-id "$PRINCIPAL" \
    --assignee-principal-type ServicePrincipal \
    --role "Monitoring Metrics Publisher" --scope "$DCR_ID"
  ```

  then allow up to half an hour for the role to propagate before 7.3; `403`
  upload failures in the console log during that window are expected
  ([turning it on](operations-runbook.md#audit-trail-in-log-analytics)).

- [ ] **7.3 Check: rows arrive.** After the sign-in and the tool calls in
      items 8 and 10, in the workspace's Logs blade:

  ```
  M365McpAudit_CL | where TimeGenerated > ago(1h) | project TimeGenerated, UserEmail, Operation, Result, Reason, Source, ClientIp | take 20
  ```

  Expected: the `auth.login`, the `policy.*.set` rows from item 8 and the
  tool calls from item 10, with `ClientIp` populated. The first rows in a new
  table can take several minutes. Also confirm the console log has no
  `Failed to send` lines:

  ```bash
  az containerapp logs show -n "$APP" -g "$RG" --type console --tail 500 | grep -c 'Failed to send'
  ```

  Expected: `0`.

- [ ] **7.4 Read and record the retention this instance actually has.**

  ```bash
  az monitor log-analytics workspace show -g "$RG" -n m365-mcp-logs --query retentionInDays
  az monitor log-analytics workspace table show -g "$RG" --workspace-name m365-mcp-logs -n M365McpAudit_CL \
    --query "{interactive:retentionInDays,total:totalRetentionInDays}"
  az containerapp show -n "$APP" -g "$RG" \
    --query "properties.template.containers[0].env[?name=='AUDIT_LOG_RETENTION_DAYS'].value | [0]" -o tsv
  ```

  The template provisions the workspace at 90 days and the table inherits it;
  the storage copy purges rows after `AUDIT_LOG_RETENTION_DAYS` (default 365
  when unset). If the tenant decided on something longer in
  [Before you start](#before-you-start), set it now with the commands under
  [Log Analytics retention](operations-runbook.md#log-analytics-retention)
  and read it back. Whatever the values are, they go in the record. A value
  nobody wrote down is a value nobody will notice drifting.

- [ ] **7.5 Hand the two alert queries to whoever owns the tenant's
      monitoring**: the upload-failure line and the stream-gone-quiet check,
      both under [when a write fails](operations-runbook.md#audit-trail-in-log-analytics).
      The server raises no alert of its own. Record who took them.

- [ ] **7.6 Check: the Container App ingress is the only proxy.** If a Front
      Door or Application Gateway sits in front, set
      `RATE_LIMIT_TRUSTED_PROXY_HOPS=2`, or every audit row records the
      proxy's address ([rate limiting](operations-runbook.md#rate-limiting-and-ingress-restriction)).

---

## 8. Policy set before any agent gets access

Every control below defaults open or permissive on a fresh instance: services
default to mail and SharePoint, the site allow-list is empty (every site), the
global deny lists are empty, draft mode is each user's own choice, and the
outbound policy is allow on all three channels. An agent that connects before
this item is done works against those defaults, and the first thing a
prompt-injected agent would try is a path this item closes. So this item
finishes before anyone is sent to `/install`, and the sign-off in item 10 is
conditional on it.

Everything here is done in the admin UI at the site root by a Global
Administrator, or against `/api/manage/*` from the browser that is signed in
to it. That API takes only the admin UI's browser session; a bearer token gets
a 401 ([admin API sessions](operations-runbook.md#admin-api-sessions)). Every
change writes an audit row with the value before and after, which is how 8.9
checks the whole item.

- [ ] **8.1 Sign in as the Global Administrator** at
      `https://your-mcp-host.example.com/api/auth/login`. You land on the
      admin UI with the admin view visible. If the view is missing,
      `GET /api/auth/me` shows `isGlobalAdmin`; false means the account lacks
      the role or `Directory.Read.All` was not consented (the check fails
      closed).

- [ ] **8.2 Enabled services.** Set the list to the services decided in
      [Before you start](#before-you-start) and nothing else. Admin UI
      (Connected Services), or `POST /api/manage/services` with
      `{"enabledServices":[...]}`. Valid keys:
      `mail`, `sharepoint`, `calendar`, `onedrive`, `onenote`, `contacts`,
      `teams`. Check: `GET /api/manage/services` returns exactly that list.

- [ ] **8.3 Read-only services.** Same endpoint, `readOnlyServices`. A
      read-only service hides its write tools from `tools/list` and refuses
      them on call ([calendar access controls](calendar-access-controls.md#2-read-only-mode--readonlyservices)).
      Check: the `GET` shows the list.

- [ ] **8.4 SharePoint site allow-list.** If the posture is closed, add at
      least one site now (admin UI, Allowed Sites, or `POST /api/manage/allowed-sites`
      with `{"allowedSites":[{"id":"<site-id>","name":"<name>"}]}`). The list
      applies to listing, reading and search alike
      ([search access controls](sharepoint-search-access-controls.md)). Check:
      `GET /api/manage/allowed-sites` is non-empty, or the record says
      "open posture, decided by <who>" explicitly.

- [ ] **8.5 Global deny lists.** Add the tenant's entries for each type:
      `sharepoint` (path prefixes), `mail` (folder display names), `calendar`
      (calendar ID or display name). Admin UI, Global Deny List (the
      per-user Personal Deny List is each user's own and is not part of this
      item), or `POST /api/manage/deny-list/global` with
      `{"type":"...","path":"..."}`.
      Check: `GET /api/manage/deny-list/global?type=<type>` for each type
      matches the decision.

- [ ] **8.6 Deployment-layer defaults.** `main.bicep` sets
      `DEFAULT_MAIL_DENY_FOLDERS=Finance,HR,Legal,IR,Management`;
      `container-app.bicep` sets nothing. These apply on top of the table and
      cannot be cleared from the admin UI. Set them to the tenant's decision:

  ```bash
  az containerapp update -n "$APP" -g "$RG" --set-env-vars \
    "DEFAULT_MAIL_DENY_FOLDERS=<comma-separated>" "DEFAULT_SHAREPOINT_DENY_PATHS=<comma-separated>" \
    --revision-suffix "$(date +%Y%m%d%H%M)"
  ```

  Check: `az containerapp exec ... --command "printenv DEFAULT_MAIL_DENY_FOLDERS DEFAULT_SHAREPOINT_DENY_PATHS"`.

- [ ] **8.7 Mail draft mode.** If the decision was enforced:
      `POST /api/manage/email-output-policy` with
      `{"scope":"tenant","enforceDraft":true}` (admin UI: Email Output
      Policy). While enforced, `send_mail` drafts, `send_draft` refuses, and
      neither the user nor the agent can switch to send
      ([access controls](../README.md#access-controls)). Check:
      `GET /api/manage/email-output-policy` shows the tenant policy. If the
      decision was to leave it to users, write that in the record: the
      default is draft, and only the user can loosen it in the web UI.

- [ ] **8.8 Outbound policy.** Set each channel to the decision:
      `POST /api/manage/outbound-policy` with
      `{"scope":"tenant","calendarInvites":"<mode>","eventResponses":"<mode>","teamsMessages":"<mode>"}`,
      modes `allow`, `internal` or `block` (admin UI: Outbound Policy).
      `internal` for Teams needs the two member-read permissions from item
      2.2 or every Teams send is refused. Check:
      `GET /api/manage/outbound-policy` shows the three modes.

- [ ] **8.9 Check the whole item from the audit log.** In the admin UI's
      Audit Log, or the workspace query from 7.3 filtered to
      `Operation startswith "policy." or Operation == "set_email_output_policy" or Operation == "set_outbound_policy"`:
      one row per change above, each with your email and an `After` value
      that matches the decision. Export that filtered view as CSV
      (`/api/manage/audit-log?format=csv`) and attach it to the record. That
      export is the evidence this item was done, in the tenant's own words.
- [ ] **8.10 Agent key policy.** Decide whether users may mint
      [agent keys](agent-api-keys.md) and what every key is held to:
      `POST /api/manage/agent-key-policy` with `allowKeys`,
      `disabledServices`, `readOnlyServices`, `maxKeyLifetimeDays` and
      `ownersMayWiden`, or **Admin → Agent Keys**. The default allows keys
      with no further restriction beyond the user's own reach, a lifetime of
      up to a year, and owners editing their own keys in either direction; a
      tenant that wants keys no longer than a sign-in sets the maximum to 30,
      and one that wants a wider key to always mean a new key turns
      `ownersMayWiden` off. A tenant that relies on Conditional Access for
      corporate-machine-only access should read the Conditional Access note
      on that page before allowing keys, since a key is evaluated by CA at its
      owner's sign-in and not at each use. Audited as
      `policy.agent_key_policy.set`, so 8.9 picks it up.

---

## 9. Client surface: served origin and update signing

The instance serves the installers, the stdio shim and the desktop extension
to every user's machine, and the extension overwrites itself with whatever
`/api/extension-update` returns. Two properties keep that from being a way to
push code at users: every URL baked into served code comes from
`OAUTH_REDIRECT_URI` and never from a forwarded host header, and every update
payload is signed with a key derived from this instance's HMAC key, which the
served extension verifies ([threat model](threat-model.md), row 4.3). Both
hold in the code, and CI checks them on every pull request against the source
and the built image ([installer release gates](installer-release-gates.md));
this item shows they hold on this deployment, with this configuration.
`node scripts/installer-release-gates.js <instance-url> <instance-url>` runs
9.1 and 9.2 in one go, with forged host headers on every request.

- [ ] **9.1 Check: served code names the configured origin, and only it.**

  ```bash
  for p in install.sh install.ps1 install; do
    printf '%s: ' "$p"
    curl -fsS "https://your-mcp-host.example.com/$p" | grep -oE 'https://[A-Za-z0-9.-]+' | sort -u | tr '\n' ' '; echo
  done
  ```

  Expected: each line lists `https://your-mcp-host.example.com`; the
  installers also carry `https://nodejs.org` and a `learn.microsoft.com`
  page for installing prerequisites, and nothing else. The
  `azurecontainerapps.io` FQDN from item 1 must not appear: if it does,
  `OAUTH_REDIRECT_URI` on the app is not the custom host and item 3.2 did
  not land.

- [ ] **9.2 Check: the update payload is signed, and the bundle carries the
      public key.**

  ```bash
  curl -fsS https://your-mcp-host.example.com/api/extension-update \
    | python3 -c 'import json,sys; d=json.load(sys.stdin); print(d["version"], "signed" if d["signed"]["signature"] else "UNSIGNED")'
  curl -fsS https://your-mcp-host.example.com/api/extension-update \
    | python3 -c 'import json,sys; d=json.load(sys.stdin); print("public key in served code:", "UPDATE_PUBLIC_KEY" in d["files"]["server/index.js"])'
  ```

  Expected: the version and `signed`; then `True`. Record the version. The
  signing key is derived from `MCP_SESSION_HMAC_KEY`, so if that key is ever
  rotated, installed extensions refuse updates until reinstalled from
  `/install`; the rotation procedure says so.

- [ ] **9.3 Check: the installer handoff asks for the code.** Run the
      installer once yourself in item 10.1 and confirm the browser lands on
      the install-confirm page asking for the code the installer printed,
      naming your account and the device label, before any session is
      attached. A sign-in that completes without that page is not this
      release.

- [ ] **9.4 Decide and record the distribution channel.** `curl | bash` from
      the instance's own hostname is the default and is TLS-bound to that
      host (threat model row 4.5). A tenant that distributes the shim through
      its own software channel instead writes that here, with where the
      artefact came from.

---

## 10. First user, first call, sign-off

- [ ] **10.1 Install as a pilot user**, not as the Global Administrator, from
      `https://your-mcp-host.example.com/install` (desktop extension) or the
      `install.sh` / `install.ps1` one-liner from the README. Confirm the
      code prompt from 9.3 and that no token is ever shown or pasted.

- [ ] **10.2 One read call per enabled service**, for example
      `list_folders_mail`, `list_sites`, `list_calendars`. Each succeeds and
      appears in the admin UI's Audit Log as `allowed`, `source: mcp`, with
      the user's email and a device label.

- [ ] **10.3 One refused call per control you set in item 8.** A folder on the
      mail deny list (`denied`, reason naming the deny list); a site off the
      allow-list (`denied`); a write tool on a read-only service (absent from
      `tools/list`, and `denied` if called by name); `send_mail` under
      enforced draft mode (lands in Drafts, and `send_draft` is `denied`); a
      calendar invitation to an external address under `internal` or `block`
      (`denied`, reason naming the outbound policy). Every refusal is a row.
      A control that was set but cannot be shown refusing is not a control
      this list can vouch for.

- [ ] **10.4 Re-run the live checks that waited on a running image**: 2.5,
      3.3, 3.4, 5.4, 7.3.

- [ ] **10.5 Functional smoke key.** As the Global Administrator, mint an
      [agent key](agent-api-keys.md#first-party-use-the-post-deploy-smoke)
      under **My Settings → Agent Keys**: label `deploy-smoke`, every service
      enabled in 8.2, every one of them read-only, no site restriction, 365
      days, reminder on. Store it on the fork without displaying it, and turn
      Issues on so a failing smoke has somewhere to report. This needs the
      key policy from 8.10 to allow keys; a tenant that turns keys off runs
      without the smoke, and the record says so.

  ```bash
  gh secret set SMOKE_AGENT_KEY -R <owner>/<fork>          # paste at the prompt
  gh variable set SMOKE_EXPECT_SERVICES -R <owner>/<fork> --body "mail,sharepoint"   # the 8.2 list
  gh repo edit <owner>/<fork> --enable-issues
  gh workflow run Deploy -R <owner>/<fork> && gh run watch -R <owner>/<fork>
  ```

  Check: the run's **Functional smoke (read-only agent key)** job is green and
  its summary shows a passed read for every enabled service and every refusal
  (write, service outside the key, `POST /api/mail/settings`, the admin API,
  logout). A job that passes with the notice *SMOKE_AGENT_KEY is not set* has
  not run. From here every deploy repeats this item unattended
  ([functional smoke](operations-runbook.md#functional-smoke)); the key's
  replacement date is on the administrator's calendar.

- [ ] **10.6 Sign off.** The record is complete, every check has a result,
      and the people in the roles table have each initialled their items.
      Only now:

  ```bash
  gh variable set AUTO_UPDATE -R <owner>/<fork> --body "true"   # tracking instances only
  ```

  and tell the pilot users where `/install` is. A pinned instance stays
  pinned; its operator dispatches each deploy by hand and re-reads 4.6 after
  every one.

---

## Tracked separately

Four pieces of work would change items on this list and are deliberately not
restated here. Each is referenced by the threat model or the runbook as
"tracked separately"; the project's tracker holds the detail.

- **A scheduled retention mechanism for the `auditLog` storage table, and a
  verified Log Analytics retention on every live instance.** Until it lands,
  item 7.4 records what each instance has by hand.
- **A federated (OIDC) deploy credential in place of the `AZURE_CREDENTIALS`
  service-principal secret.** Item 4.3 describes the credential as it is
  today.
- **Narrowing the SharePoint and OneDrive grant below `ReadWrite` at consent
  time.** Today a read-only tenant enforces that in item 8.3, not in Entra, and
  a reviewer has to trust the connector for it. Item 2.2 says so.
- **A narrower role-read scope in place of `Directory.Read.All` for the admin
  check.** Item 2.2 carries `Directory.Read.All` until that is decided.

When one of these lands, update the item it affects in the same pull request.

---

## Go-live record

Copy this into the deployment's own notes, outside this repository, and fill
it in as you go. A blank cell is an item not done.

| Item | What to record | Value | Who | Date |
|---|---|---|---|---|
| Before you start | Every decision in that list, one line each | | | |
| 1 | Resource group, region, vault name, runtime identity ID, resource group Owner | | | |
| 1.3 | Output of the Key Vault reference check (three names, versioned) | | | |
| 1.4 | Principals with a data-plane role on the vault, with reasons | | | |
| 2 | Client ID, tenant ID, client secret expiry date | | | |
| 2.4 | Consented scope list, verbatim | | | |
| 2.6 | Assignment required: yes; assigned users or group | | | |
| 3 | Hostname; values of `OAUTH_REDIRECT_URI` and `FRONTEND_URL` as resolved | | | |
| 3.4 | Body cap in bytes, in the app and at the edge; edge service and tier, and its two rate-limit thresholds (Application Gateway: its maximum instance count and the resulting per-address ceilings, and the WAF policy check); the edge body-cap rule's `RegEx` pattern; ingress restriction (gateway address, or Front Door Private Link with public access disabled); output of the three checks (the body check's status code says which cap fired) and, behind Front Door, the Private Link checks; or "no edge" and who accepted it | | | |
| 4 | Fork; deploy principal name and role scopes; deployed commit and image tag; the release it belongs to and the CI evidence on that release ([supply chain](supply-chain.md#release-evidence-the-ci-runs-for-the-tagged-commit)) | | | |
| 5 | Storage account settings from 5.1; role assignments from 5.3 | | | |
| 6.2 | The two key hash prefixes | | | |
| 6.4 | Envelope binding variables set: yes | | | |
| 7.4 | Workspace retention; audit table retention; `AUDIT_LOG_RETENTION_DAYS` | | | |
| 7.5 | Who owns the two audit alerts; SIEM, if any | | | |
| 8 | Each policy value as set; the CSV export from 8.9 attached | | | |
| 9 | Served origin check result; extension version; distribution channel | | | |
| 10 | Pilot user; the allowed and denied rows from 10.2 and 10.3; `AUTO_UPDATE` state | | | |
| 10.5 | Smoke key id, owner and expiry date; `SMOKE_EXPECT_SERVICES`; the first green smoke run | | | |
| Sign-off | Names and initials for each role in the roles table | | | |
