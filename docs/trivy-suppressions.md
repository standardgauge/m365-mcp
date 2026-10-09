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

## `uuid` < 11.1.1 via `@azure/msal-node` — retired

`npm audit --omit=dev` used to report two linked moderates: `uuid` <11.1.1
(GHSA-w5hq-g745-h8pq, a missing bounds check in `v3`/`v5`/`v6` when `buf` is
passed) and `@azure/msal-node` <=5.1.4 for depending on it. The vulnerable calls
were never reachable (msal-node 2.16.3 only called `v4()`), so the entry was held
open until the library could be upgraded deliberately rather than through
`npm audit fix --force`, which an `overrides` pin on `uuid` would have been a
worse version of.

That upgrade has happened: `@azure/msal-node` is on 7.x, which does not depend on
`uuid`, and the production audit is clean. What made it safe to take:

- `src/__tests__/msalAuthFlows.test.ts` runs the real library against an
  in-process stand-in for the Entra endpoints (`src/__tests__/fixtures/fakeEntra.ts`)
  through the authorization code, device code and refresh flows, and through the
  Table Storage cache plugin. Every other test mocks MSAL out, so before this
  nothing would have noticed a change in request shapes or cache handling.
- A cache serialized by 2.16.3 is checked in as a fixture
  (`msal-node-2.16.3-cache.json`) and the test confirms 7.x reads it, finds the
  account, and refreshes with its refresh token, so signed-in users are not sent
  back to sign in by the deploy.
- That test found the one behavioural break: 7.x keys cached credentials
  differently and never migrates 2.x keys, so an upgraded cache kept presenting
  the pre-upgrade refresh token on every refresh while the rotated one sat
  unused beside it. `src/services/msalCacheKeys.ts` re-keys the blob as the
  cache plugin loads it; its header has the detail.

If a later MSAL major changes the credential key format again, the test that
expects a 7.x-written cache to pass through the re-keying unchanged fails, and
`currentCredentialKey` needs updating to match.

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
