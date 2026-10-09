# Entra ID setup, operations, and troubleshooting

The **permission list and the reasoning behind it live in the README**, under
[Entra ID app registration and permissions](../README.md#entra-id-app-registration-and-permissions).
That is the single source of truth. This document covers the operational
material around it: who can grant consent, how Conditional Access interacts
with the flows, verifying consent actually landed, rotating the secret, and
diagnosing the failures that follow from getting consent wrong.

Deliberately no permission table here. Two copies drift, and a drifted
permission table is worse than no permission table, because it is confidently
wrong. This file supersedes `docs/aad-setup.md`, which had drifted: it described
an Azure Functions deployment that no longer exists and named an app
registration from a tenant this project no longer targets.

The step order for a whole tenant, with checkpoints, is
`onboarding-a-tenant.md`.

---

## Who can grant consent, and who can use the admin UI

Every permission is delegated; there are no application permissions. That
changes which roles are needed:

- **Tenant-wide admin consent** can be granted by a Privileged Role
  Administrator, or by a Cloud Application Administrator, Application
  Administrator or AI Administrator. Those three can consent to any delegated
  permission for any API; only Microsoft Graph *application* permissions are
  outside their reach, and this app requests none.
  ([Microsoft: grant tenant-wide admin consent](https://learn.microsoft.com/en-us/entra/identity/enterprise-apps/grant-admin-consent))
- **The admin UI** unlocks for the Global Administrator directory role only.
  The server reads `/me/transitiveMemberOf` and looks for the Global
  Administrator role template (`src/services/authMiddleware.ts:158-172`); it
  fails closed on any error. No other role and no group unlocks it. The
  person who configures deny lists and services needs that role; the person
  who grants consent does not.

`Directory.Read.All` is what lets the role check see `roleTemplateId` on the
returned directory roles; with only `User.Read`, Graph returns those objects
with `id` and type only and the check fails closed for everyone. It also
backs the User Management view's `/users` listing with `accountEnabled` and
`assignedLicenses` (`src/functions/admin/getTenantUsers.ts:35-46`).

## Conditional Access

The OAuth exchange happens in the user's own browser: `/api/auth/login`
redirects to `login.microsoftonline.com/<tenant>/…`, the tenant's Conditional
Access policies run there, and the server only ever receives the resulting
authorization code. So:

- Device-compliance, MFA and location policies apply to this app exactly as
  they apply to any browser sign-in. Do not exclude the app registration from
  Conditional Access; it does not need an exclusion, and the policies are what
  make delegated tokens safe.
- The `.mcpb` extension opens the **system browser** for sign-in, which is why
  it works on tenants with a managed-device policy: the compliant device's SSO
  session satisfies the policy. (April 2026) was the older redirect
  flow failing under such a policy; moving to the extension closed it.
- `GET /api/auth/device` still exists as an anonymous route and starts a
  device-code flow (`src/functions/auth/deviceLogin.ts`). If your policies are
  scoped to browser flows only, extend them to the device-code authentication
  flow, or ask for the route to be removed (it creates sessions that are never
  returned to the caller and has no consumer in the current install paths).
- Refresh tokens are subject to the tenant's token lifetime and revocation
  policy. The application caps a session at 30 days absolute regardless
  (`src/services/tokenCache.ts`). It has no idle timeout of its own: a
  session unused for more than 7 days is renewed through the refresh token on
  its next request, so an Entra sign-in frequency or revocation is what ends
  an idle session sooner.

---

## Verifying consent actually landed

Granting consent in the portal reports success whether or not it did what you
wanted, so check the result rather than the click.

```bash
# List the delegated permissions actually consented for the app.
az ad app permission list-grants --id <client-id> --query "[].scope" -o tsv | tr ' ' '\n' | sort
```

Compare that output against two things:

1. The permission table in the README, for the tools you intend to enable.
2. **`GRAPH_SCOPES` in `src/services/graphClient.ts`.** Every entry in that array
   must appear in the consented set. This is the one that breaks sign-in for
   everybody rather than degrading gracefully, so check it explicitly.

A permission that appears in the registration but not in the grant list has been
added without consent being re-granted. Adding a permission does not consent to
it; the two are separate actions in the portal and it is easy to do the first and
believe you have done both.

A second, live check once an image is running: `curl -sI https://<host>/api/auth/login`
returns a 302 whose `Location` names your tenant ID in the authority and lists
exactly the `GRAPH_SCOPES` entries plus `openid profile offline_access` in
`scope=`. If the tenant ID there is not yours, `AZURE_TENANT_ID` on the
Container App is wrong.

---

## Troubleshooting

### Every tool call fails with `interaction_required`

A scope in `GRAPH_SCOPES` is not consented in this tenant. MSAL requests it on
silent refresh, Entra refuses, and the failure surfaces on every call rather
than only the affected service.

Fix the consent, not the array. Removing the scope from `GRAPH_SCOPES` to make
the error go away will break other tenants running the same image.

### Sign-in works, but one service returns nothing

That service's permissions were never granted. Unlike the case above, a missing
permission for a service not named in `GRAPH_SCOPES` degrades quietly: the token
is issued, the tool is offered, and Graph refuses the individual call.

Check the grant list for the relevant scopes and re-consent.

### Sign-in works, but the admin view never appears

The signed-in user does not hold the Global Administrator role, or
`Directory.Read.All` is not consented so the role check cannot see the role
template and fails closed. `GET /api/auth/me` shows `isGlobalAdmin` either way.

### `AADSTS50011: redirect URI mismatch`

`OAUTH_REDIRECT_URI` does not exactly match a registered redirect URI. The match
is literal, including scheme, host, port, path, and trailing slash.

### `Access denied. Your account belongs to a different tenant.`

The callback compares the token's tenant with `AZURE_TENANT_ID` and refuses a
mismatch (`src/functions/auth/callback.ts:70-78`). Either the user is a guest
from another tenant, which is refused by design, or the Container App's
`azure-tenant-id` secret is wrong.

### Everything breaks at once, months after a working deployment

Check the client secret expiry first. An expired secret takes down the whole
instance simultaneously, which is a distinctive signature: no partial failures,
no service-specific pattern, and nothing changed on your side.

---

## Rotating the client secret

The secret expires. Rotation is routine, but it has an ordering that avoids
downtime:

1. Create the **new** secret in **Certificates & secrets** without deleting the old one. Both are valid at once.
2. Update the Container App secret reference to the new value.
3. Create a new revision so the change is picked up: `az containerapp update -n <app> -g <rg> --revision-suffix "$(date +%Y%m%d%H%M)"`. A plain `revision restart` brings replicas back with the old value (verified on example).
4. Verify sign-in and one tool call.
5. Only then delete the old secret.

Deleting first inverts steps 1 and 5 and produces an outage for the length of the
rollout.

Diarise the next expiry when you create the secret. This failure mode is
avoidable and its blast radius is the entire instance.

---

## Adding a permission later

The order matters, and it is the opposite of what feels natural.

1. Add the permission to the app registration and **grant admin consent**, in every tenant that runs the image.
2. Confirm with `az ad app permission list-grants` that the grant is live.
3. Only then ship code that uses it.

Because a v2.0 access token carries every consented scope for the resource, a
newly consented permission reaches existing sessions on their next silent
refresh. Users do not need to sign in again.

Do not add the scope to `GRAPH_SCOPES` as part of this. See the README for why
that array is deliberately short.

---

## Removing a user's access

There is no admin action in the application that revokes one user's session.
The controls are in Entra:

- **Disable or delete the user**, or **revoke sessions** on the user object.
  Their refresh token stops working; tool calls fail as soon as the current
  access token expires (about an hour), and the server retains the dead session
  row until its 30-day cap so a later re-sign-in can heal it.
- **Require assignment** on the enterprise application and assign users or
  groups, so only assigned users can sign in at all.
- Per-user restrictions short of removal (disable a service, disable mail
  indexing) are in the admin UI's User Management.

Removing every session at once is the kill switch in the
[operations runbook](operations-runbook.md).
