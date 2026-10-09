# Threat model

This document says what the M365 MCP server protects, where its trust boundaries
are, what can go wrong at each one, what the code does about it today, and where
it falls short. It is written for two readers: a security reviewer deciding
whether to let an instance into their tenant, and a penetration tester deciding
what to test first.

It describes the code, not any particular deployment. Every instance runs in its
operator's own Azure subscription and Entra tenant, so the operator's own
controls (Conditional Access, Azure RBAC, network policy, Defender) sit on top of
what is described here and are not assumed.

**Baseline.** Reviewed against `main` at commit `e3326b4`. Code references name
files and functions rather than line numbers, so they survive edits.

**Method.** STRIDE (Spoofing, Tampering, Repudiation, Information disclosure,
Denial of service, Elevation of privilege) applied per area. Each threat row
gives the current mitigation and, where one exists, a gap. Gaps carry a stable
identifier (`G1`, `G2`, ...) that the project's tracker references; the
[gap register](#gap-register) lists them in one place.

**Severity.** *High*: exploitable by someone outside the tenant, or yields
another user's Graph access, without needing operator access. *Medium*: needs a
foothold first (a stolen token, storage write, a prompt-injected agent, an
Azure role), or weakens a control a reviewer would rely on. *Low*:
defence in depth, accuracy of a documented claim, or bounded nuisance.

---

## Contents

- [System overview](#system-overview)
- [Assets](#assets)
- [Actors](#actors)
- [Trust boundaries](#trust-boundaries)
- [Assumptions](#assumptions)
- [1. Delegated OAuth and consent](#1-delegated-oauth-and-consent)
- [2. Token cache and credential envelopes](#2-token-cache-and-credential-envelopes)
- [3. MCP session state](#3-mcp-session-state)
- [4. Client install and update](#4-client-install-and-update)
- [5. Tenant isolation](#5-tenant-isolation)
- [6. Table Storage](#6-table-storage)
- [7. Admin surface](#7-admin-surface)
- [8. Audit log](#8-audit-log)
- [9. Prompt and tool abuse, confused-deputy paths](#9-prompt-and-tool-abuse-confused-deputy-paths)
- [10. Availability](#10-availability)
- [11. Supply chain and deployment](#11-supply-chain-and-deployment)
- [Gap register](#gap-register)
- [Penetration test scope](#penetration-test-scope)
- [Keeping this document current](#keeping-this-document-current)

---

## System overview

```
 User device                                   Operator's Azure subscription
 ┌──────────────────────────────┐               ┌──────────────────────────────────────┐
 │ MCP client + model           │               │ Container App (Functions host, Node) │
 │   │ stdio                    │  TB1 HTTPS    │  /api/mcp        JSON-RPC tools      │
 │ shim / desktop extension ────┼──────────────▶│  /api/auth/*     OAuth, install poll │
 │   holds session token        │  bearer token │  /api/<service>  REST tool routes    │
 │   reads local files on ask   │               │  /api/manage/*   admin API           │
 │                              │  TB2 HTTPS    │  /  (catch-all)  admin SPA           │
 │ Browser ─────────────────────┼──────────────▶│  /install*, /api/extension-*         │
 │   mcp_session cookie         │  cookie       │      │ in-memory session cache       │
 └──────────────────────────────┘               │      │                               │
            ▲  TB7 installers, shim and         │  TB5 │ Entra token, runtime identity │
            │  extension code served by the     │      ▼                               │
            │  instance                         │  Table Storage: sessions, MSAL cache,│
            └───────────────────────────────────┤  install nonces, policy, audit       │
                                                │                                      │
   TB6 Azure control plane: RG roles, secrets,  │  Key Vault, by reference: client     │
       deploy principal, fork sync  ───────────▶│  secret, HMAC key, DEK               │
                                                └──────┬──────────────────────┬────────┘
                                                  TB3  │ OAuth (confidential) │ TB4 delegated
                                                       ▼                      ▼ access token
                                                 Entra ID (one tenant)   Microsoft Graph
                                                                          │
                                       TB8 tenant content (mail, files, ◀─┘
                                           events) returned to the model
```

The server is one HTTP service. A user signs in once through Entra with the
authorization-code flow. The server keeps the resulting refresh token in an
encrypted MSAL cache, mints its own random session token, stores only an HMAC of
that token, and hands the token to the user's MCP client through the install
flow. Every later call presents the session token; the server looks up the
session, obtains a fresh delegated access token for that user, applies its own
policy (services, read-only, site allow-list, deny lists, draft enforcement),
calls Graph, and returns the result to the model.

It holds **delegated** permissions only. It never stores mail, file or event
content; it stores credentials, policy and audit rows.

---

## Assets

| ID | Asset | Where it lives | Why it matters |
|---|---|---|---|
| A1 | Refresh tokens for every signed-in user | `mcpMsalCache`, one AES-GCM envelope ([`tableStorage.ts`](../src/services/tableStorage.ts) `saveMsalCache`) | Delegated Graph access to everything each user can reach, for up to 90 days, independent of this server's session lifetime. The highest-value thing the server holds. |
| A2 | Graph access tokens | `mcpSessions` per-session envelope; replica memory | Roughly an hour of the user's delegated access. |
| A3 | Session tokens | Client keychain, DPAPI blob or libsecret (a 0600 file where none is available); HMAC only in `mcpSessions`; for up to 5 minutes during install, an envelope in `mcpInstallNonces` | Bearer credential for the server, valid up to 30 days. Whoever holds it acts as the user through every tool. |
| A4 | Entra client secret | Key Vault; the Container App holds a reference | With a refresh token, redeems new access tokens; with an authorization code, completes sign-in. |
| A5 | `MCP_SESSION_HMAC_KEY`, `MCP_DATA_ENCRYPTION_KEY` | Key Vault; the Container App holds references | Turn a storage copy into usable credentials (DEK), or let an attacker mint a row that matches a token they chose (HMAC key). |
| A6 | Storage data access | The app's runtime identity (Storage Table Data Contributor). Account keys exist but shared-key auth is off | Full read and write on every table, including policy and audit. |
| A7 | Policy state | `GlobalDenyList`, `UserDenyList`, `serviceSettings`, `allowedSites`, `EmailOutputModePolicy`, user overrides | What AI is allowed to touch. Tampering silently widens access. |
| A8 | Audit log | `auditLog` table | The record of what each user's agent did and what was refused. |
| A9 | Tenant content in transit | Graph responses, tool results, model context | Not stored by the server, but passes through it and into a third-party model. |
| A10 | User profile data | `mcpSessions` (display name, email, object ID, tenant ID), audit rows | Personal data under the operator's privacy obligations. |
| A11 | Client-side code | Installers, shim and desktop extension served by the instance | Runs on every user's machine with that user's privileges. |
| A12 | Deployment credentials | Fork repository secrets, deploy principal | Push a new image to the instance. |

---

## Actors

| Actor | Starting position |
|---|---|
| External attacker | Internet access to the instance's public endpoints. No account in the tenant. |
| Malicious content author | Can put text in front of the model: an email to any user, a shared document, a meeting invite, a Teams message. Needs no access to the server at all. |
| Prompt-injected agent | A legitimate user's MCP client, steered by A9 content into calling tools the user did not intend. |
| Tenant insider | A signed-in member of the tenant using their own session for more than they should. |
| Token thief | Holds a copied session token (A3) from a laptop, backup, log, or config file. |
| Storage-only attacker | Read or write on the storage account without the app secrets: a leaked SAS, a backup copy, a storage-scoped role. |
| Azure operator | Contributor or Owner on the resource group. Controls the Container App and what it runs; reading a Key Vault value directly takes a data-plane role on the vault. |
| Upstream contributor | Can land code on this repository's `main`, which deploy forks track. |

---

## Trust boundaries

| ID | Boundary | Crossing it |
|---|---|---|
| TB1 | User device to server | MCP JSON-RPC and install polling over TLS, authenticated by the session token. |
| TB2 | Browser to server | OAuth redirects and the admin SPA, authenticated by the `mcp_session` cookie. |
| TB3 | Server to Entra ID | Confidential-client authorization code and refresh-token redemption. |
| TB4 | Server to Microsoft Graph | Delegated access tokens. Graph enforces the user's own permissions. |
| TB5 | Server to Table Storage | Entra token for the app's runtime identity; the account refuses shared-key auth. Everything at rest is here. |
| TB6 | Azure control plane to runtime | Resource-group roles, Key Vault and storage data roles, the Container App's Key Vault references, the deploy principal. |
| TB7 | Server to user device, as a code source | Installers, the shim and the desktop extension's auto-update are fetched from the instance and executed locally. |
| TB8 | Tenant content to model | Untrusted text from mail, files and events enters model context and can steer tool calls. This is a boundary in data, not in network. |

---

## Assumptions

- Entra ID and Microsoft Graph enforce delegated permissions correctly. A token
  for user U cannot reach what U cannot reach.
- TLS terminates at the Container Apps ingress and the ingress refuses plain HTTP
  (`allowInsecure: false` in [`infra/main.bicep`](../infra/main.bicep)).
- The operator keeps the resource group's privileged roles to people trusted
  with every user's Graph access. Section 6 explains why that is literal.
- The MCP client shows tool calls to the user and lets them refuse. The server
  cannot verify this and does not rely on it for any control in this document.
- The model provider's handling of tool results is governed by the operator's
  agreement with that provider, not by this server.

---

## 1. Delegated OAuth and consent

Sign-in is MSAL's authorization-code flow against a tenant-specific authority
([`graphClient.ts`](../src/services/graphClient.ts) `getMsalApp`,
[`login.ts`](../src/functions/auth/login.ts),
[`callback.ts`](../src/functions/auth/callback.ts)). Consent is granted once by an
administrator for the whole permission set ([`entra-setup.md`](entra-setup.md)).

| # | STRIDE | Threat | Current mitigation | Gap |
|---|---|---|---|---|
| 1.1 | S | Login CSRF: an attacker completes the callback in a victim's browser with the attacker's code, so the victim works inside the attacker's account. | Random `state` in an HttpOnly, Secure, `SameSite=Lax` cookie, compared with the query value; mismatch is a 403. | None found. |
| 1.2 | S, E | Authorization-code interception or injection. | Confidential client: redeeming a code needs the client secret. Redirect URI must match the registration. PKCE with S256 on every sign-in: `login` keeps the verifier in an HttpOnly, Secure, `SameSite=Lax` cookie beside `state`, and the callback refuses to redeem without it, so a code issued to another browser's flow does not redeem here. `login` also sends a random `nonce`, kept the same way, and MSAL rejects an ID token whose `nonce` claim is missing or different. | None found. |
| 1.3 | E | Over-broad consent: the registration grants more than the enabled tools need, so a stolen token reaches more. | All permissions delegated; no application permissions. Setup guide tells operators to drop permissions for services they do not enable. | Sites and Files are granted ReadWrite tenant-wide because Graph has no narrower delegated scope that fits; tracked separately. `Directory.Read.All` is carried only for the admin check; a narrower role-read scope is tracked separately. |
| 1.4 | S | Device-code phishing through the anonymous `GET /api/auth/device` route. | Single-tenant authority; foreign-tenant accounts rejected. The route never returns the session token it creates, so a phisher gains no session. | The route is dead weight: an anonymous endpoint that starts a 15-minute background flow per call and offers a sign-in path Conditional Access may treat differently. Removal is tracked separately. |
| 1.5 | I | Consent prompt hides scope from users. | Admin consent once, so users see no prompt and cannot be tricked into granting more. | Accepted: users are not asked, by design. The operator is the consent authority. |
| 1.6 | D | Client secret expiry takes every user out at once. | Documented rotation in the [runbook](operations-runbook.md). | Expiry alerting is tracked separately. |
| 1.7 | I | Error text leaks configuration. | `login` and the callback return generic errors; the exception text goes to the log only. | None found. |

---

## 2. Token cache and credential envelopes

[`credentialCrypto.ts`](../src/services/credentialCrypto.ts) holds the
primitives: HMAC-SHA256 over session tokens, and AES-256-GCM envelopes for access
tokens and the MSAL cache. Each envelope is bound by GCM additional authenticated
data to its table, partition, row and column, so an envelope copied into another
row fails to decrypt. Envelopes written before that binding are still readable
until the operator sets `MCP_ENVELOPE_REQUIRE_AAD=true`
([runbook](operations-runbook.md#credential-envelope-row-binding-mcp_envelope_require_aad)).

| # | STRIDE | Threat | Current mitigation | Gap |
|---|---|---|---|---|
| 2.1 | I | Storage-only attacker reads refresh and access tokens. | Both are AES-256-GCM ciphertext under a key that is not in storage. Session tokens are stored only as a keyed HMAC. | Holds against a storage-only attacker: the DEK is in Key Vault, and no storage credential sits beside it. Who can still reach the DEK is row 6.2. |
| 2.2 | T, E | Storage writer moves user A's access-token envelope into user B's row, so B's session acts as A. | AAD binds each envelope to its row; a moved envelope fails authentication. 16-byte tag enforced. | Until `MCP_ENVELOPE_REQUIRE_AAD=true` is set, an unbound legacy envelope still moves. Operator action, documented. |
| 2.3 | T, E | Storage writer edits the plaintext columns next to the envelope instead of the envelope. | None. `userId`, `homeAccountId` and `tenantId` are plaintext and outside the AAD. | A writer holding any valid session can repoint that session's `homeAccountId` at another user's MSAL account; the next silent refresh returns the other user's access token into the attacker's session. Editing `userId` changes whose deny lists and settings apply and whose name the audit records. **G3** |
| 2.4 | I, D | One MSAL cache row for all users. | Encrypted and AAD-bound like everything else. | Every user's refresh token is in one blob, so any code path that loads the cache holds all of them. Writes are an unconditional replace, so two replicas refreshing at once can drop a user's newly rotated refresh token. One Table property is capped at 64 KiB, which puts a ceiling on how many users the cache can hold before writes start failing; the failure is logged, not surfaced. **G4** |
| 2.5 | T | Truncated GCM tag accepted. | Tag length pinned to 16 bytes on decrypt. | None. |
| 2.6 | I | Key reuse across deployments. | Generated per deployment; README says never reuse across tenants. | Not enforced. Startup validation of key shape is tracked separately. |
| 2.7 | R, I | Key rotation. | Rotating either key signs everyone out; documented as the fastest kill switch. | No dual-key period, so rotation is always disruptive. Accepted. |

---

## 3. MCP session state

The server keeps no MCP-protocol session state. Each JSON-RPC request carries
the session token, and [`authMiddleware.ts`](../src/services/authMiddleware.ts)
`authenticateRequest` resolves it. A session has a 7-day idle window and a
30-day absolute cap ([`tokenCache.ts`](../src/services/tokenCache.ts)). Each
replica keeps an in-memory cache of sessions it has seen, keyed by token hash.

| # | STRIDE | Threat | Current mitigation | Gap |
|---|---|---|---|---|
| 3.1 | S | Guessing a session token. | 32 random bytes; lookup by HMAC; no client-supplied user ID is trusted. | None. |
| 3.2 | S | Stolen token used indefinitely. | 30-day absolute cap, anchored on an immutable creation time and checked before any refresh. | Sessions with neither timestamp skip the absolute check. The 7-day "idle" window is not a boundary: on expiry the server silently refreshes and carries on, so only the 30-day cap ends a session that holds a live refresh token. The runbook describes it as an idle timeout. **G13** |
| 3.3 | E | Revoked session keeps working. Logout, the credential purge script, or deleting rows by hand. | Logout deletes all of the user's rows and evicts every cached session for the user on the replica that handled it. Other replicas re-check a cached session's row at most every 30 seconds on use and drop it once the row is gone, so a deleted session stops authenticating everywhere within 30 seconds ([`tokenCache.ts`](../src/services/tokenCache.ts) `SESSION_REVALIDATE_MS`). Writes back to an existing session are conditional updates, so a refresh inside that window fails instead of recreating the row, and evicts the session at once. Entra "revoke sessions" kills the refresh token. Key rotation (new revision, new process) clears everything. | While storage is unreachable a replica keeps serving a cached session for up to 5 minutes past its last check, then refuses it (`SESSION_REVALIDATE_MAX_STALE_MS`); accepted, so a storage blip does not sign everyone out. An admin action to revoke a named user's sessions is tracked separately. |
| 3.4 | T | Cross-site request forgery against cookie-authenticated routes. | `mcp_session` is HttpOnly, Secure, `SameSite=Lax`, so it is not sent on cross-site POSTs. No CORS headers are emitted. | Relies on SameSite alone: no Origin check or CSRF token. A same-site origin (a sibling subdomain under the instance's registrable domain) is not cross-site and gets the cookie. Logout is a GET, so a cross-site link logs the user out. Folded into **G9**. |
| 3.5 | I | Session token leaks from the client. | The desktop extension and the script installers both keep it in the OS keychain, DPAPI, or libsecret, falling back to a 0600 file. The installers hand it to the shim on stdin (`--store-token`) and the MCP client config names only the store (`--token-store`), so no config file holds the token. The `.backup` copies the installers write are 0600 and have the earlier installer's bearer header removed. | A config written by an earlier installer keeps its token until the installer is re-run, and that token stays valid until the 30-day cap. The macOS `security` CLI takes the token on argv for the moment it runs, the same as the extension. Accepted. |
| 3.6 | R | Several devices share one identity. | Each device gets its own session row and an optional device label that the audit log records. | None. |
| 3.7 | E | `tools/list` shows tools the user cannot use. | Filtered by enabled services, overrides and read-only. Call-time checks are the backstop. | If policy cannot be read, `tools/list` shows everything. Display only; every call is still checked and those checks fail closed. Accepted. |

---

## 4. Client install and update

Users connect through `install.sh`, `install.ps1` or the desktop extension, all
served by the instance ([`installEndpoint.ts`](../src/functions/install/installEndpoint.ts)).
The installer generates a verifier, opens a sign-in URL carrying its SHA-256, and
polls [`install-poll`](../src/functions/auth/installPoll.ts) with the verifier
until the callback has attached a session to it. The extension also checks the
instance for a newer version of itself at startup and overwrites its own files.

| # | STRIDE | Threat | Current mitigation | Gap |
|---|---|---|---|---|
| 4.1 | S, E | Someone other than the installer that started the sign-in collects the resulting session. | Verifier/challenge split: the URL carries only the hash, so a party that only sees the URL cannot poll. One-time consumption with an ETag race guard. 5-minute window. | The binding between the browser that signs in and the installer that collects the session needs strengthening. Details are withheld from this public document until the fix ships. **G1** |
| 4.2 | I | Session token at rest during handoff. | The `mcpInstallNonces` row holds the token as an AES-GCM envelope bound to its row (the DEK and AAD scheme of section 2, [`tableStorage.ts`](../src/services/tableStorage.ts) `attachSessionToInstallNonce`). A row without a valid envelope, including one written before encryption, is deleted and never returned. The poll response is consumed once, and rows past their 5-minute expiry are deleted by the server whether or not anyone polls (`purgeExpiredInstallNonces`). `install.sh` writes the poll response to a `mktemp` file (0600, unpredictable name) removed on exit, and passes the token to curl and the shim on stdin, not argv. `install.ps1` keeps it in memory. | The purge is not a timer trigger: it starts from session writes, token lookups and install traffic, at most every 10 minutes per replica, so an abandoned row can outlive its expiry on an idle instance. It is encrypted meanwhile. Accepted. |
| 4.3 | T, E | Instance pushes code to every client. | The extension only updates to a strictly newer version and only from its own instance origin. | Updates are unsigned, so whoever controls the instance's responses controls code on every user's machine. File paths from the update payload are joined to the extension directory without a containment check. The origin baked into served code is taken from forwarded host headers. **G10** |
| 4.4 | I | Shim reads local files and sends them to the server. | Only for attachment and upload arguments; limited to `~/Downloads`, `~/Documents` and OneDrive sync folders; refuses dotfiles and dot-directories after resolving symlinks; 10 MB cap. | A steered agent can still attach any file under those folders. See 9.4. |
| 4.5 | T | `curl \| bash` from a spoofed host. | TLS to the instance's own hostname. | Inherent to the pattern. Operators who need more can distribute the shim through their own software channel. Accepted. |

---

## 5. Tenant isolation

Each instance serves one Entra tenant, in that tenant's own Azure subscription,
with its own app registration, keys and storage. There is no shared
infrastructure between deployments.

| # | STRIDE | Threat | Current mitigation | Gap |
|---|---|---|---|---|
| 5.1 | S, E | User from another tenant signs in. | MSAL authority is `login.microsoftonline.com/<AZURE_TENANT_ID>`; the server refuses to build an MSAL client without it. The callback and device flow also compare the account's tenant with `AZURE_TENANT_ID` and refuse a mismatch. | None. The README's multi-tenant section describes a mode the code does not support; foreign-tenant users are refused. |
| 5.2 | E | Guest (B2B) accounts. | The tenant compared is the one in the account's home ID, so a guest whose home tenant differs is refused. | Behaviour to confirm in the penetration test; it is a side effect of 5.1, not a designed control. |
| 5.3 | I | One deployment's credentials work against another. | Keys, client secrets and storage are per deployment. Session tokens are only meaningful against the storage and HMAC key that issued them. | Key reuse across deployments is possible if an operator copies values. Documented, not enforced. |
| 5.4 | E | Upstream change reaches every deployment at once. | `main` is protected, reviewed and CI-gated. | Section 11. |
| 5.5 | I | One user reads another user's data inside the tenant. | Every Graph call uses the caller's own delegated token, so Graph applies that user's permissions. Session lookup is by token, never by a client-supplied user ID. | Server-side policy keyed on `userId` inherits the integrity of the session row (**G3**). |

---

## 6. Table Storage

The server reaches Storage with an Entra token for its user-assigned runtime
identity, which holds Storage Table Data Contributor on the account and Key Vault
Secrets User on the vault ([`storageClient.ts`](../src/services/storageClient.ts),
[`infra/key-vault.bicep`](../infra/key-vault.bicep)). The account created by
[`infra/main.bicep`](../infra/main.bicep) has shared-key access off, so neither
an account key nor a SAS signed with one is accepted. The client secret and both
application keys are Key Vault secrets; the Container App holds versioned
references to them. Tables are created by the Bicep templates or on first use: `mcpSessions`, `mcpMsalCache`, `mcpInstallNonces`, the
policy tables, and `auditLog` ([`infra/main.bicep`](../infra/main.bicep),
[`tableStorage.ts`](../src/services/tableStorage.ts)).

| # | STRIDE | Threat | Current mitigation | Gap |
|---|---|---|---|---|
| 6.1 | I | Storage copy (backup, leaked SAS) yields credentials. | Envelopes and HMAC, section 2, including the install-handoff token in `mcpInstallNonces` (4.2). Shared-key access off, so a leaked account key or key-signed SAS is refused. HTTPS only, TLS 1.2 minimum, no public blob access. | None. |
| 6.2 | E | Azure operator reads everything. | The DEK, HMAC key and client secret are in an RBAC-mode Key Vault with purge protection; the Container App holds references, so reading its secrets yields URIs, not values. Reading a value directly takes a data-plane role on the vault, which Contributor does not include and only Owner or User Access Administrator can grant, and every read is logged to the instance's Log Analytics workspace. Storage takes an Entra token, so there is no storage key in the app or in deployment history. | Residual, not closed: Contributor on the resource group still controls the runtime. It can exec into a replica or roll a revision running its own image and read the resolved keys from the environment, and it can change the vault's and the storage account's settings, including turning shared-key access back on. Those are deliberate, logged control-plane writes rather than a passive read, so the Activity Log is the record; who holds Contributor remains the control that matters. Both the account and the vault are reachable from public networks, because the Consumption profile has no VNet integration; private endpoints are tracked separately. |
| 6.3 | T | Storage writer widens policy (removes deny entries, empties the allow-list, lifts draft enforcement). | None at rest. Policy rows are plain entities. | Covered by 6.2 (who can write) and **G8** (no record when policy changes through the API). |
| 6.4 | T, R | Storage writer edits or deletes audit rows. | None. The audit table is writable by the app's runtime identity and by anyone holding a data role on the account. | **G8**, and audit export to an append-only store is tracked separately. |
| 6.5 | D | Storage unavailable. | Deny-list checks fail closed. Service, read-only, override and draft-policy reads propagate the error, so the call fails. | Audit writes fail open (6.4, section 8). |

---

## 7. Admin surface

The admin SPA is served from the site root and calls `/api/manage/*`. Every
admin route authenticates the session, then calls Graph `/me/transitiveMemberOf`
with the caller's own token and requires the Global Administrator role
(`checkGlobalAdmin`). The check runs on every request and fails closed. The
Functions host's own `/admin/*` API is master-key protected and never reaches
this code ([runbook](operations-runbook.md#reserved-paths)).

| # | STRIDE | Threat | Current mitigation | Gap |
|---|---|---|---|---|
| 7.1 | E | Non-admin calls an admin route. | Per-request Global Administrator check against Graph, fail closed. | Only active role assignments count; that is the intended behaviour. Allowing a lesser role or a group is tracked separately. |
| 7.2 | E | Non-admin edits another user's deny list (IDOR). | User deny-list routes allow a different `targetUserId` only for a Global Administrator. | None. |
| 7.3 | E | The MCP client's token is an admin credential. | None. | `authenticateRequest` accepts the same session token from the bearer header, the alternate header and the cookie, on every route. A Global Administrator who installs the client has an admin-API credential sitting in their MCP config, keychain or desktop extension. Admin sessions should be separate, short-lived and browser-only. **G9** |
| 7.4 | T | XSS in the admin SPA. | React rendering; CSP with `default-src 'none'`, `frame-ancestors 'none'`, `form-action` and `connect-src` restricted; runtime config JSON-escaped for script context. | `script-src` allows `'unsafe-inline'` for the injected config block, so the CSP does not stop an injected inline script. **G9** |
| 7.5 | R | Admin changes are not attributable. | Draft-mode policy changes are audited. | Deny-list, services, read-only, allowed-sites, per-user override and mail-config changes are not. **G8** |
| 7.6 | I | Admin reads tenant directory. | `tenant-users` requires Global Administrator and uses the admin's own token. | None. |

---

## 8. Audit log

[`auditLog.ts`](../src/services/auditLog.ts) `logAccess` writes one row per
allowed or denied tool call on both surfaces, partitioned by tenant, newest
first. A Global Administrator reads or exports it from `/api/manage/audit-log`.

| # | STRIDE | Threat | Current mitigation | Gap |
|---|---|---|---|---|
| 8.1 | R | A tool call happens without a record. | Every MCP `tools/call` path and every REST route through `withPolicyEnforcement` logs. | Writes are fire-and-forget: a failed write is logged to the console and the call proceeds. Retention and a SIEM path are tracked separately. |
| 8.2 | R | A security-relevant event that is not a tool call goes unrecorded. | None. | Sign-in, sign-in failure, logout, session refresh, install handoff, admin policy changes and admin-check failures are not audited. **G8** |
| 8.3 | S, R | Recorded client address is wrong. | REST rows record `x-forwarded-for`. | MCP rows record no address. REST rows record the whole header, whose leftmost entry the client controls. **G8** |
| 8.4 | T | Audit rows edited or deleted. | Only a Global Administrator can read them through the API; nothing in the API deletes them. | Anyone with the storage key can. No hash chain or immutable copy. Section 6.4. |
| 8.5 | I | Audit rows hold tenant content. | Rows carry identifiers and paths, not message or file bodies. | Paths and denial reasons can include folder and file names. Accepted, and stated here so a reviewer does not read "stores no tenant content" as "stores no names". |
| 8.6 | D | Unbounded growth. | Queries capped at 1,000 rows and a 10,000-row scan for substring filters. | No purge; tracked separately. |

---

## 9. Prompt and tool abuse, confused-deputy paths

The model reads untrusted content (TB8) and then chooses tool calls that run
with the user's delegated authority. The server cannot tell an instruction the
user gave from one an email gave. Its controls therefore limit what any call can
do, rather than trying to judge intent.

Controls available to an operator: per-service enablement (mail and SharePoint
only by default), per-user service overrides, per-service read-only mode (write
tools hidden and refused), the SharePoint site allow-list, two-tier deny lists,
default deny folders, enforced draft mode, strict argument validation (unknown
parameters refused, opaque IDs checked before they reach a Graph path), and the
shim's local-file limits.

| # | STRIDE | Threat | Current mitigation | Gap |
|---|---|---|---|---|
| 9.1 | E | Injected prompt sends mail to an outside address. | Draft mode by default; an admin can enforce it so `send_mail` drafts, `send_draft` refuses, and the agent cannot switch modes. | Without enforcement, `set_email_output_mode` is a tool and the agent can switch to send. Documented in the README; enforcement is the answer. |
| 9.2 | I, E | Injected prompt exfiltrates through a channel draft mode does not cover. | Calendar and Teams are off unless enabled; read-only mode can cover calendar. | With calendar enabled and writable, `create_event` and `update_event` with attendees send invitations at once, carrying a body the agent wrote, and `respond_to_event` can carry a comment to the organizer. With Teams enabled, chat and channel sends are immediate. None of these are held by enforced draft mode, and there is no control on external recipients. **G6** |
| 9.3 | T | Injected prompt deletes or overwrites data. | Read-only mode per service; deny lists on paths and folders; Graph's own recycle bins. | No server-side confirmation for destructive calls; the MCP client's approval prompt is the only human step. Accepted, stated so operators choose read-only where it matters. |
| 9.4 | I | Confused deputy on the device: an injected prompt makes the shim read a local file and attach or upload it. | Root folders, dotfile refusal, symlink resolution, size cap (4.4). Enforced draft mode keeps the attachment in Drafts. | Residual: anything under `~/Documents` or `~/Downloads` can be attached to a draft or uploaded to OneDrive. Operators can narrow `M365_MCP_ATTACH_ROOTS`. Accepted. |
| 9.5 | E | Confused deputy across users: a delegate's agent reaches an owner's mailbox or calendar. | Graph enforces the delegation itself. Global (tier 1) deny entries apply to everyone. Every mail tool and route that takes `mailboxId`, and `respond_to_event` on another mailbox's calendar, is checked against the per-user (tier 2) lists of both the caller and the mailbox owner. The owner is looked up in the directory (`/users/{mailboxId}`); if the lookup fails, the call is refused before any message is read (`src/services/mailboxOwner.ts`). | A calendar someone shares into the caller's own calendar list is reached by `calendarId`, not `mailboxId`. It is checked against the caller's list and tier 1 only. The shared copy has its own ID in the recipient's list, and its name there can differ from the owner's, so the owner's entries could not reliably match it anyway. An owner who wants a shared calendar hidden from every agent asks an administrator for a tier 1 entry. The README says so. |
| 9.6 | E | Confused deputy at the server: someone gets the server to act with another user's authority. | Sessions are looked up by token, never by a supplied user ID; the old `x-user-id` trust model is gone. | Through the install handoff (**G1**) or a tampered session row (**G3**). |
| 9.7 | E | Argument injection into Graph paths or OData. | `assertOpaqueIds` on every ID parameter; KQL and filter values escaped; unknown parameters refused. | None found; worth fuzzing (see test scope). |
| 9.8 | I | Model sees content the user hid. | Deny lists filter listings and search hits, including ancestor-folder matches on search results. | Deny matching is by name and path. A renamed folder escapes a name-based entry until the entry is updated. Accepted and documented in the access-control docs. |

---

## 10. Availability

| # | STRIDE | Threat | Current mitigation | Gap |
|---|---|---|---|---|
| 10.1 | D | Flooding public endpoints. | Per-address limit on `login`, `install-poll`, `mcp` and `device`, keyed on the ingress-appended forwarded address (IPv6 by `/64`), answering `429`. Optional ingress IP restriction for tenants with named egress. Container Apps scales to three replicas. | The limit is per replica and in memory, so three replicas allow three times it, and a distributed source with many addresses is not slowed. A volumetric attack needs an edge service (Front Door with WAF rate rules); not in the templates. |
| 10.2 | D | An unknown bearer token costs a table scan. | Fast path is a single row lookup. | On a miss the server scans the whole `mcpSessions` partition looking for legacy rows, for up to three token candidates per request, before rejecting. A JSON-RPC batch has no size limit. **G14** |
| 10.3 | D | MSAL cache write failure signs users out. | Logged. | **G4**. |
| 10.4 | D | Cold start. | Minimum one replica in the templates. | Operator setting. |

---

## 11. Supply chain and deployment

Deploy forks track this repository's `main` exactly and, with `AUTO_UPDATE`
set, deploy each new `main` commit once CI has passed on it. Anything merged here reaches every tracking instance
and, through the extension's auto-update (4.3), every user's desktop.

| # | STRIDE | Threat | Current mitigation | Gap |
|---|---|---|---|---|
| 11.1 | T, E | Malicious or compromised change lands on `main`. | Protected branch: pull request, required CI (type check, lint, tests, identifier scan, dependency audit, secrets scan, CodeQL, non-root container boot test), squash only, no force push. Tenant deploys start only from a successful CI run on the exact commit being deployed (`deploy.yml`, `workflow_run`), so a `main` that goes red after merge ships nowhere. | Actions are pinned by tag rather than SHA; deploy forks authenticate with a long-lived principal secret. All tracked separately. |
| 11.2 | T | Dependency compromise. | `npm audit` gate, Dependabot. | Dependabot cooldown tracked separately. SBOM and image provenance tracked separately. |
| 11.3 | E | Container escape or host-level foothold. | Image runs as a non-root user; CI fails a build that runs anything as root. | None in this model's scope. |
| 11.4 | R | Cannot tell what is running. | `/health` reports the deployed commit. | None. |

---

## Gap register

Gaps found by this review. Each has a tracker item that links back to its row
above. Items the project already tracked before this review are marked "tracked
separately" in the tables and are not repeated here.

| Gap | Severity | Area | Summary |
|---|---|---|---|
| G1 | High | 4 Install | Install sign-in handoff binding. Details withheld until fixed. |
| G3 | Medium | 2 Envelopes | `userId`, `homeAccountId`, `tenantId` on session rows are outside the envelope AAD and unprotected. |
| G4 | Medium | 2 Token cache | One MSAL cache row for every user: shared blast radius, last-writer-wins across replicas, 64 KiB property ceiling. |
| G6 | Medium | 9 Prompt abuse | Enforced draft mode does not cover calendar invitations, event-response comments or Teams sends; no external-recipient control. |
| G8 | Medium | 7, 8 Audit | Auth events and admin policy changes are not audited; MCP rows have no client address; REST rows trust the leftmost forwarded address. |
| G9 | Medium | 7 Admin | The MCP client's session token is also an admin-API credential for Global Administrators; admin CSP allows inline script; CSRF defence is SameSite alone. |
| G10 | Medium | 4 Client update | Extension auto-update is unsigned, writes payload paths without containment, and bakes in an origin from forwarded headers. |
| G13 | Low | 3 Session | The 7-day idle window renews silently instead of ending the session; sessions with no timestamps skip the 30-day cap; runbook overstates the idle timeout. |
| G14 | Low | 10 Availability | An unknown bearer token triggers full-partition scans (up to three per request); JSON-RPC batch size is unbounded. |

When a gap closes, change its row in the relevant table to describe the new
mitigation and delete it from this register in the same pull request.

---

## Penetration test scope

What a test of a non-production instance should cover first, in order. Each
line names the rows it exercises.

1. **Session acquisition without the user's intent.** The install handoff end
   to end, including links the tester did not start from an installer (4.1,
   9.6). Login CSRF and state handling (1.1). Guest and foreign-tenant sign-in
   (5.1, 5.2).
2. **Session lifetime and revocation.** Reuse after logout, after row deletion,
   after Entra revocation, and after key rotation, against more than one replica
   (3.2, 3.3). Idle and absolute limits with a stolen token.
3. **Authorization on every route.** Each `/api/manage/*` route as a non-admin,
   each REST tool route against the deny list, read-only mode and the site
   allow-list, and the same checks through `/api/mcp` (7.1, 7.2, 9.x). The
   session token from an MCP client presented to the admin API (7.3).
4. **Input handling.** Fuzz every opaque-ID parameter, `q` / KQL, `path`,
   `$filter`-bearing arguments and HTML bodies (9.7). Admin SPA XSS (7.4).
5. **Prompt-driven abuse.** Seed mail, documents and invites with instructions
   and measure what leaves the tenant with draft enforcement on, with calendar
   and Teams enabled (9.1, 9.2), and through delegated mailboxes (9.5).
6. **Client code.** Extension update handling against a hostile instance
   response, installer temp files and config permissions (4.2, 4.3), shim path
   containment including symlinks and Windows path forms (4.4).
7. **Storage with write access.** If the engagement includes a storage-scoped
   credential: envelope swap with and without `MCP_ENVELOPE_REQUIRE_AAD`, and
   session-row column edits (2.2, 2.3).
8. **Availability.** Unknown-token request cost, batch size, device-flow
   start-up cost (10.1, 10.2, 1.4).

Out of scope: Microsoft Graph and Entra ID themselves, and any instance the
tester's client does not own (see [SECURITY.md](../SECURITY.md)).

---

## Keeping this document current

Update this document in the same pull request as any change that adds an
endpoint, a table, a stored credential, a tool that sends anything outside the
tenant, or a new way to authenticate, and whenever a gap above closes. Move the
baseline commit forward when the whole document is re-reviewed, not on every
edit.
