# What an agent can never do

A plain-English summary for compliance, risk and security reviewers. Each
statement describes what the code on `main` enforces. The
[threat model](threat-model.md) carries the detail, the source files and the
tests behind each one.

**This server is one component, not a compliance programme.** Running it does
not make a deployment HIPAA compliant, SEC compliant or compliant with anything
else, and nothing here is certified. Compliance belongs to the deploying
organisation's whole environment, meaning its policies and supervision, its Microsoft
365 and Entra configuration, its archiving, and its agreements, including a
business associate agreement with Microsoft and with the model provider where
HIPAA applies. The mapping below says which requirements a control supports
and marks the gaps as gaps. It is a pointer for your reviewers, not legal
advice.

## The short list

An agent connected through this server can never:

1. **Reach anything its user can't.** Every Microsoft Graph call uses the
   signed-in user's own delegated token. The app registration has no
   application permissions and there is no service account.
2. **Hold more than the person it works for.** An agent key resolves to its
   owner's session and can only narrow it, to fewer services, more of them
   read-only, fewer SharePoint sites, or extra deny entries. An administrator can
   narrow or revoke any key, and switch keys off for the whole tenant.
3. **Send email while draft mode is on, or turn draft mode off.** The agent can
   switch its user to draft mode but never to send mode, and every attempt is
   logged as denied. Only the user, in the web UI, can choose send mode, and an
   administrator can enforce draft mode so the user can't either.
4. **Send an invitation, a meeting-response comment or a Teams message the
   outbound policy forbids.** Each channel is set to allow, internal only or
   block. If the policy or the recipients can't be read, the call is refused.
5. **Open a folder, path, calendar, notebook or plan on a deny list.** That
   holds on the MCP endpoint and on the REST routes, and denied items are
   dropped from search results.
6. **Write to a service set read-only, or reach a SharePoint site outside the
   allow-list** once the allow-list has any entry.
7. **Reach the admin controls.** An agent key can't read or change policy, mint
   or edit keys, change email settings, or sign its owner out.
8. **Reach another tenant.** Each instance serves one Entra tenant and refuses
   accounts whose home tenant is different.

## What gets recorded

Every allowed and denied tool call, on both the MCP endpoint and the REST
routes, writes an audit row with who, which tool, which resource, the result, the
denial reason, the client address, and which device or agent key made the
call. Sign-ins, sign-in failures, logouts and session refreshes are recorded
too, and every administrator policy change is recorded with its value before
and after. Rows go to a Log Analytics table and an Azure Table in your own
subscription. They carry identifiers and paths, not message or file bodies.
Retention in the table is `AUDIT_LOG_RETENTION_DAYS` (365 by default);
retention in Log Analytics is whatever the workspace is set to.

## What it does not do

These are gaps, stated so nobody reads them as covered.

- **Audit writes fail open.** A failed write is logged to the console and the
  call goes ahead. The server raises no alert of its own when the audit stream
  stops, so watching for gaps belongs to your security tooling.
- **The audit table is not tamper-evident.** Anyone holding a data role on the
  storage account can edit or delete rows. There is no hash chain. Treat the
  Log Analytics copy as the record and set its retention and access to match.
- **Conditional Access is evaluated at sign-in**, not each time an agent key is
  used, so a copied key keeps working from a non-compliant device until it is
  revoked or expires.
- **No sensitivity-label awareness and no DLP.** The server applies Graph
  permissions and its own deny lists. An encrypted or labelled file comes back
  as Graph returns it to that user, and nothing inspects content for regulated
  data before it reaches the model.
- **Draft mode binds the agent, not the user's token.** The web UI's
  settings route accepts send mode from anyone holding the user's session
  token, because that is how the user changes it. No tool returns that token,
  so an agent reaches the route only if it can read the token some other way,
  such as off the user's disk. Enforced draft mode closes this, since then the
  route refuses too.
- **No idle timeout.** A session is capped at 30 days from sign-in and ends
  sooner if Entra revokes the refresh token. An agent key lives for the
  lifetime its owner chose, up to the tenant's maximum (365 days by default).
- **No records archive.** The server keeps no copy of mail or messages. Mail
  sent through the MCP tools lands in the user's Sent Items. The REST send
  route accepts `saveToSentItems: false`, so mail sent that way leaves no copy
  there. Tenant journaling or an archiving product is the control for books and
  records, not Sent Items.
- **What the model provider does with content is outside this server.**
  Whatever a tool returns goes to the client's model. Its retention, training
  terms and BAA are between you and the provider.
- **Azure commercial cloud only.** The sign-in authority and Graph endpoint are
  the commercial ones, so GCC High, DoD and Azure China would need code changes.

## Control mapping

"Supports" means the server provides a technical control that a programme can
point to. It never means the requirement is met.

### HIPAA Security Rule, technical safeguards (45 CFR 164.312)

| Requirement | What the server does | Gap, or left to you |
|---|---|---|
| (a)(1) Access control | Delegated tokens only, so Graph applies each user's own permissions. Deny lists, per-service read-only, the SharePoint allow-list and agent keys narrow what an agent reaches below that. | Who has access to what in Microsoft 365 itself is the tenant's configuration. |
| (a)(2)(i) Unique user identification | Every user signs in individually through Entra. Each device gets its own session, and each agent key is named on its audit rows. | None in the server. |
| (a)(2)(ii) Emergency access procedure | Nothing specific. | Gap. Yours to define. |
| (a)(2)(iii) Automatic logoff | 30-day absolute session cap; keys expire on their chosen lifetime; logout and key revocation end a session on every replica within 30 seconds while storage is reachable. | Gap: no idle timeout. |
| (a)(2)(iv) Encryption and decryption | Access tokens and the token cache are AES-256-GCM encrypted at rest under a key held in Key Vault; session tokens are stored only as a keyed hash. The server stores no mail or file content. | Encryption of the content itself is Microsoft 365's. |
| (b) Audit controls | Allowed and denied calls, authentication events and admin policy changes, to Log Analytics and Table Storage in your subscription. | Gap: writes fail open, and the table copy is not tamper-evident. |
| (c)(1) Integrity | Draft mode, read-only services and the outbound policy limit what an agent can change or send. Credential records are authenticated, so one moved or edited in storage fails to decrypt. | Gap: no integrity mechanism for content. Microsoft 365 is the system of record. |
| (d) Person or entity authentication | Entra sign-in, with the tenant's own MFA and Conditional Access. An agent key acts only as its owner. | Gap: Conditional Access is not re-checked when a key is used. |
| (e)(1) Transmission security | HTTPS only at the ingress; the storage account the full-stack template creates requires TLS 1.2 or later; Graph is called over HTTPS. | The leg from the MCP client to the model provider is outside the server. |

### SEC and FINRA themes

| Theme | What the server does | Gap, or left to you |
|---|---|---|
| Access to customer information (Regulation S-P safeguards, 17 CFR 248.30) | The same access controls as above: delegated tokens, deny lists, read-only services, the site allow-list, agent keys no wider than their owner. | Written policies, incident response and vendor oversight are the firm's. |
| Supervision (FINRA Rule 3110) | Draft mode puts a person between the agent and delivery, and can be enforced by an administrator. The outbound policy limits invitations and Teams messages to internal recipients or blocks them. The audit log gives a supervisor a record of what each agent did and was refused. | Gap: the server does no supervisory review. A user in send mode, or a channel the policy allows, delivers without one. |
| Books and records (SEC Rules 17a-4 and 204-2, FINRA Rule 4511) | Mail and Teams messages are sent from the user's own account in your tenant, so the tenant's journaling and retention see them like any other. | Gap: the server keeps no archive, and the REST send route can skip Sent Items. Capture and retention belong to your archiving. |
