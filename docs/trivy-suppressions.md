# Dependency-vulnerability findings: what gates, what does not, and why

*(Filename is historical — this doc started as Trivy-only and now covers the
`Dependency Audit` job too. Both CI gates are documented here so "why is this CVE
not fixed?" has one answer, not two.)*

The `Container Scan (Trivy)` job in `.github/workflows/ci.yml` builds the runtime
image and scans it in two steps, both with `ignore-unfixed: true`:

1. **App layer — blocking** (`vuln-type: library`). The meaningful gate. Fails the
   build on fixable `CRITICAL`/`HIGH` vulns in packages this repo ships and
   controls (our production npm deps, plus any lang-pkg such as the
   `dotnet-core`-detected `MessagePack.dll`). Skips base-image npm dirs it does
   not control and applies the deliberate suppressions in the repo-root
   `.trivyignore` via the `trivyignores:` input.
2. **Base OS — report-only** (`vuln-type: os`, `exit-code: 0`). Surfaces base-OS
   CVEs for visibility but does not gate. See the section below.

Every `.trivyignore` entry must carry a justification and a re-review date. This
file is the long-form record behind those one-liners.

## Why the base-OS layer is report-only — ``

The runtime image's OS layer is Microsoft's vendor base image
(`mcr.microsoft.com/azure-functions/node:4-node20`, currently `debian 11.11`). We
do not build that image and cannot bump its OS packages independently. `4-node20`
is a rolling tag CI already re-pulls fresh on every `main` push, so the only
remediation available to us for an OS-layer CVE is to wait for Microsoft to
rebuild the host image — or for Dependabot's docker ecosystem watch
(`.github/dependabot.yml`) to bump the tag when a newer (e.g. Debian 12) variant
ships.

Gating on OS-layer findings therefore reddens `main` CI on every new base-OS CVE
wave with no action we can take. The trigger for this change was the July-2026
`linux-libc-dev` batch: 61 HIGH kernel-header CVEs, all `fixed`-status (a Debian
fix exists) but not yet in the base image, landed after the MessagePack suppression and re-reddened the trunk. Each future wave would do the
same. Kernel-header CVEs specifically are non-exploitable in a container — the
package ships only C headers, and a container runs no kernel of its own (it shares
the host's) — so gating on them buys no security while masking real regressions on
a red trunk right as/245 bring auto-deploy live.

**This is not a blanket HIGH ignore.** The app-layer step above still fails the
build on any fixable HIGH/CRITICAL in a package we ship. Only the vendor-owned OS
layer is downgraded to report-only, and it stays visible in the CI log so a
genuinely actionable OS finding (or a base-image bump opportunity) is not hidden.

### Exit criteria / revisit

Restore OS-layer gating if we ever own the base image (e.g. switch to a
self-built runtime), or tighten to gate on specific OS packages we can influence.
Bumping to a Debian 12 base image (option (a) on the ticket) does not by itself
justify re-gating: the newer OS still accrues `linux-libc-dev` kernel-header waves,
so the report-only posture remains the correct long-term stance for a consumed
vendor image.

## `uuid` < 11.1.1 via `@azure/msal-node` — assessed, not acting

`npm audit` reports two linked **moderate** findings:

```
uuid              <11.1.1     Missing buffer bounds check in v3/v5/v6 when buf is provided
@azure/msal-node  <=5.1.4     Depends on vulnerable versions of uuid
  fixAvailable: @azure/msal-node@6.0.0  (isSemVerMajor: true)
```

**Neither gates CI**, and neither should be "fixed" by taking npm's advice.

### Why the suggested fix is wrong here

npm proposes upgrading the direct dependency `@azure/msal-node` from **2.16.3 to
6.0.0** — four major versions, on the library that performs every OAuth token
acquisition for every tenant. The blast radius of getting that wrong is total:
no user on any tenant can authenticate. is a reminder that this code path
fails in subtle, hard-to-detect ways.

Doing that to clear two moderates that gate nothing is a bad trade.

### Reachability (why this is low risk here)

The advisory is specific: a missing bounds check in **`v3`, `v5` and `v6`**, and
only **when the optional `buf` argument is supplied**.

`@azure/msal-node@2.16.3` calls **`v4()` only** — a single call site, for
correlation IDs, with no `buf` argument. `v4` is not in the affected set.

There is exactly one `uuid` copy in the tree (`node_modules/uuid@8.3.2`,
hoisted). The second `@azure/msal-node@5.4.2` present via
`applicationinsights → @azure/identity` does not pull its own.

So the vulnerable code paths are not called, and cannot be reached through this
application's use of the library.

### Why not just override `uuid`

An `overrides` pin to `uuid@^11.1.1` would silence the finding without the major
bump, and was considered. Rejected for now: it forces an untested transitive
substitution (8.x → 11.x, across the CommonJS/ESM export rework) **inside the
auth library**, to remediate a code path that is not reachable. That is real risk
bought with no real gain. Revisit if the finding ever becomes gating or reachable.

### Exit criteria

Any of these should retire this entry:

- The `@azure/msal-node` 2.x line ships a release depending on `uuid` ≥ 11.1.1.
- A deliberate, tested upgrade to `@azure/msal-node` 6.x happens on its own
  merits — with token-flow regression coverage, not as a side effect of `npm
  audit fix --force`.
- msal-node begins calling `v3`/`v5`/`v6`, which would make the advisory
  reachable. Re-check the call sites when the dependency moves.

**Re-review by:** 2026-12-01

## MessagePack (MessagePack-CSharp) 2.5.192 — ``

**Suppressed IDs:** CVE-2026-48109, CVE-2026-48506, CVE-2026-48502,
CVE-2026-48510, CVE-2026-48512, CVE-2026-48515
**Fixed upstream in:** MessagePack 2.5.301 (v2 line) / 3.1.7 (v3 line)
**Re-review by:** 2026-10-01, or immediately when the base image is bumped.

### What broke

Since the last green run (2026-05-28), every push to `main` failed the Trivy job
on `MessagePack 2.5.192`. Because the job is main-push-only (image build + scan
is too slow for every PR), PR CI stayed green while `main` went red — 15+ commits
landed on a red trunk, and a real build regression would have been invisible
behind the standing MessagePack failure. Surfaced by the 2026-07-01 Example/m365-mcp
E2E review (`docs/reviews/2026-07-01-e2e-review.md` in the tracker parent).
`` was closed earlier but addressed the `npm audit` job, not this
container-scan finding — different scanner, different package source.

### Where MessagePack 2.5.192 enters the image

It is a **.NET assembly (`MessagePack.dll`) vendored inside the Azure Functions
host runtime** that ships in our stage-2 base image,
`mcr.microsoft.com/azure-functions/node:4-node20` (see `Dockerfile`). It is:

- **not** an npm dependency of this app — there is no `messagepack` entry in
  `package-lock.json`; and
- **not** part of our application code under `/home/site/wwwroot` — it belongs to
  the Microsoft-maintained Functions host.

### Why we suppress rather than bump

We do not build the host and cannot bump the assembly independently. `4-node20`
is a rolling tag and CI rebuilds the image fresh on every `main` push, yet it
still ships 2.5.192 — so Microsoft has not yet published a `4-node20` host image
carrying MessagePack ≥ 2.5.301. With `ignore-unfixed: true` already set, these
advisories surface only because an upstream fix version exists that the base
image has not adopted. There is no in-repo change that upgrades a
base-image-vendored .NET assembly.

### Reachability (why this is low risk here)

The vulnerable path across this June-2026 advisory batch is LZ4 decompression /
deserialization of **untrusted** MessagePack payloads (CWE-20; CVE-2026-48109 is
a remote DoS via out-of-bounds read, CVSS 8.2). This is a Node.js/TypeScript
Azure Functions app: it neither references nor invokes MessagePack. The host uses
MessagePack only for internal, trusted, in-process communication with its
language workers, not for deserializing attacker-controlled input from the public
HTTP surface. The precondition for exploitation is not met in our deployment.

### Exit criteria

Remove the MessagePack IDs from `.trivyignore` (and this section) once a
`4-node20` base image ships MessagePack ≥ 2.5.301 / 3.1.7. A fresh `main` build
will then pass on its own; if it does not, the assembly is still present and the
suppression is still warranted — re-date and re-review.
