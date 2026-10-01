# Security policy

## Reporting a vulnerability

Report privately. Do not open a public issue for a suspected vulnerability.

Use GitHub's [private vulnerability
reporting](https://docs.github.com/en/code-security/security-advisories/guidance-on-reporting-and-writing-information-about-vulnerabilities/privately-reporting-a-security-vulnerability)
on this repository, or email **security@standardgauge.ai**.

Please include what you were doing, what happened, and enough detail to
reproduce it. If you have a proof of concept, say so rather than attaching it
in a first message.

## What to expect

This project is maintained by a small team. We will acknowledge a report within
five working days and tell you whether we consider it in scope and what we
intend to do. We do not operate a bug bounty.

We ask that you give us a reasonable opportunity to fix an issue before
disclosing it publicly, and we will credit you when a fix ships unless you
prefer otherwise.

## Scope

In scope: this source, the container it builds, and the Bicep templates under
`infra/`.

Out of scope: any deployed instance you do not own. Every deployment of this
server runs in its operator's own Azure subscription and Entra tenant, so
testing against someone else's instance is testing against their systems, not
ours. Do not do it.

Also out of scope: Microsoft Graph and Entra ID themselves. Report those to
[Microsoft](https://msrc.microsoft.com/report).

## Design notes a reviewer should have

- The server holds **delegated** access only. Every Graph call carries a token
  for a specific signed-in user; there are no application permissions and no
  client-credentials flow. It cannot reach anything the signed-in user could
  not already reach.
- It stores **no tenant content**. Mail, files and calendar entries are fetched
  per call and returned, never cached or indexed.
- Access tokens and the MSAL cache are encrypted at rest with AES-256-GCM under
  per-deployment keys. The client's session token is stored only as a keyed
  HMAC.
- Deny-list checks **fail closed** when the policy store is unreachable.

Known limitations are in the README and the operations runbook rather than
hidden here. Two worth naming: the container currently runs as root, and there
is no rate limiting on the public authentication endpoints.
