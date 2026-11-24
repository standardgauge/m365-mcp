# Releases

A release is a `vX.Y.Z` tag on a commit of `main`, with a section in
[`CHANGELOG.md`](../CHANGELOG.md) saying what changed and what a deployment has
to do before it moves. It exists so an instance can be pinned to a version that
has a name and notes, rather than to whichever commit of `main` it last
deployed.

`main` stays production for every deployment that tracks it
([`AGENTS.md`](../AGENTS.md)). Releases do not change that. They add a second
way to follow the project: move from one named version to the next, on your own
schedule, having read what changed. [Deployment pattern](deployment-pattern.md)
is the design for making that the supported way.

## What the version number promises

Semantic versioning, read from the side of someone running an instance. The
question for each change is what a deployment has to do when it rolls forward.

**Major** (`2.0.0`): a deployment that rolls forward without acting breaks, or
behaves differently in a way its users will notice. Any of:

- A scope added to `GRAPH_SCOPES`. An unconsented scope makes
  `acquireTokenSilent` fail with `interaction_required` for every signed-in
  user, so the consent has to happen first.
- A setting removed or renamed, or a new one that is required.
- A Bicep change the instance has to apply. The deploy workflow only swaps the
  image; it converges the ingress port and the probes, and nothing else.
- A storage change that needs a migration or makes rolling back unsafe.
- A tool, a tool argument or an HTTP route removed or renamed, or a default
  that tightens or loosens what an agent can do.

**Minor** (`1.3.0`): something new that a deployment gets without acting, and
that leaves existing behaviour as it was. New tools, new optional settings,
new optional Bicep parameters, a deprecation announced.

**Patch** (`1.2.1`): fixes, dependency and base image updates, and docs.

Security fixes ship as a patch on the newest release. Only the newest release
is supported; there are no maintenance branches. If a fix has to change
behaviour to close the hole, it is still released, and the Upgrade notes say
so plainly.

There are no pre-release tags yet. `scripts/release-check.sh` refuses anything
but `vX.Y.Z`.

## The changelog

[Keep a Changelog 1.1.0](https://keepachangelog.com/en/1.1.0/), with one group
added in front of the standard six:

```markdown
## [Unreleased]

### Upgrade notes

- Consent `Tasks.ReadWrite` before rolling forward (Planner tools).

### Added

- Planner tools: plans, buckets, tasks.

### Fixed

- `search_mail` honours KQL in the query.
```

- **Upgrade notes** lists every action a deployment takes before or while it
  rolls forward. Every major change has one. A minor change with an optional
  action (a new setting worth turning on) can have one too.
- The other groups are Keep a Changelog's own: Added, Changed, Deprecated,
  Removed, Fixed, Security. No others; the script refuses them.
- A pull request that changes behaviour adds its line under `## [Unreleased]`
  in the same pull request. Docs-only and test-only changes don't need one.
- Releases are `## [X.Y.Z] - YYYY-MM-DD`, newest first, each version once,
  under `## [Unreleased]`.

CI runs `scripts/release-check.sh` on every pull request. It checks the
headings above, and that `package.json` carries the newest released version.

## Cutting a release

Releases are cut by a maintainer, from `main`, in two steps.

**1. A release pull request.** One topic: the version.

```bash
V=1.2.0
npm version "$V" --no-git-tag-version   # package.json and package-lock.json
```

In `CHANGELOG.md`, rename `## [Unreleased]` to `## [1.2.0] - <today>`, put a
new empty `## [Unreleased]` above it, and update the link references at the
foot. Then decide the version from the Upgrade notes rather than the other way
round: an Upgrade note that a deployment cannot skip makes the release major.
`bash scripts/release-check.sh --tag "v$V"` passes locally before the pull
request opens. Merge it once CI is green.

**2. The tag and the release**, on the merge commit, once CI has passed on it:

```bash
SHA=$(gh api repos/standardgauge/m365-mcp/commits/main --jq .sha)
gh release create "v$V" --repo standardgauge/m365-mcp --target "$SHA" \
  --title "v$V" --notes-file <(bash scripts/release-check.sh --notes "$V")
```

Run it with a maintainer's own credentials. A release created with a
workflow's `GITHUB_TOKEN` does not start other workflows, so it would publish
with no SBOM and no provenance.

Publishing starts [`release.yml`](../.github/workflows/release.yml). It
checks the tag against `package.json` and the changelog before it builds
anything, then attaches the SBOMs, the source archive and the checksums, signed
([supply chain](supply-chain.md)). Check the release page carries them.

If the tag check fails, the release is published with no assets. Delete the
release and the tag, fix the tree with another pull request, and cut it again.
Don't move a tag that has been published: somebody may already have pinned it.

### Not automated yet

A workflow that cuts the release itself would need a credential that can start
`release.yml` (a GitHub App or a token stored in this public repository), or
`release.yml` turned into a reusable workflow called by the cutting one. Both
change who signs the provenance, and that is the part a deployer verifies, so
it is left as the manual step above until it is worth that review.

Recommended repository setting, which needs an administrator: a tag ruleset on
`v*` that blocks update and deletion, so a published version cannot be moved.

## Pinning an instance to a release today

A fork running [`deploy.yml`](../.github/workflows/deploy.yml) can deploy a
release now, with no change to the fork:

```bash
FORK=<owner>/<fork>; V=v1.2.0
gh variable delete AUTO_UPDATE -R "$FORK"   # pinned: only a dispatch deploys
git fetch https://github.com/standardgauge/m365-mcp.git "refs/tags/$V:refs/tags/$V"
git push "https://github.com/$FORK.git" "refs/tags/$V"
gh workflow run Deploy -R "$FORK" --ref "$V"
```

A dispatch on a tag builds and deploys the tagged commit, and the post-deploy
smoke checks `/health` reports it. Read the release's Upgrade notes first.

The fork's `main` can keep tracking upstream as before; with `AUTO_UPDATE`
unset, a new commit there changes the source and not the running instance.
