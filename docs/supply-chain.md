# Supply chain: SBOMs and build provenance

Two questions come up whenever someone evaluates an instance of this server:
what is in it, and was it built from the source it claims. This page shows how
to answer both from evidence, for a release of this repository and for the
instance you actually run.

The pieces:

- every release carries an SBOM and a signed provenance attestation
  ([`release.yml`](../.github/workflows/release.yml));
- every release archives the hosted CI runs for its commit with its release
  notes ([release evidence](#release-evidence-the-ci-runs-for-the-tagged-commit));
- every deploy pushes its image with a provenance statement attached
  ([`deploy.yml`](../.github/workflows/deploy.yml), `provenance: mode=max`);
- every running instance reports the commit it was built from on `GET /health`.

## What a release carries

Publishing a release runs `release.yml` on the tagged commit. It builds the
image, catalogues it with [syft](https://github.com/anchore/syft), and attaches
these files to the release:

| File | What it is |
|---|---|
| `m365-mcp-<tag>.spdx.json` | SBOM of the built image, SPDX 2.3 JSON: OS packages from the base image and every npm package in the app |
| `m365-mcp-<tag>.cdx.json` | The same SBOM as CycloneDX JSON |
| `m365-mcp-<tag>-source.tar.gz` | The tagged source, packed with `git archive` |
| `SHA256SUMS` | SHA-256 of the three files above |

Before anything is attached, the workflow checks that every production
dependency in `package.json` appears in both SBOMs at exactly the version
`package-lock.json` resolves, so an SBOM that misses the application's own
packages fails the release rather than shipping.

Each file has a [SLSA build provenance](https://slsa.dev/provenance/v1)
attestation, signed through GitHub's Sigstore instance with a short-lived
certificate bound to this repository and this workflow. The attestation records
the repository, the commit, the tag, the workflow file, and the run.

### Verify a release

With the [GitHub CLI](https://cli.github.com/) 2.49 or later:

```bash
gh release download <tag> --repo <owner>/m365-mcp --dir m365-mcp-<tag>
cd m365-mcp-<tag>
sha256sum --check SHA256SUMS

for f in *.json *.tar.gz SHA256SUMS; do
  gh attestation verify "$f" \
    --repo <owner>/m365-mcp \
    --signer-workflow <owner>/m365-mcp/.github/workflows/release.yml \
    --source-ref refs/tags/<tag>
done
```

`<owner>/m365-mcp` is the canonical repository the release came from.
`--signer-workflow` refuses an attestation made by any other workflow in the
repository, and `--source-ref` refuses one made for any other ref, which
matters because pull requests that touch the build sign throwaway attestations
on their merge ref to prove the signing path works between releases. Add
`--format json` to see the full statement, including `resolvedDependencies`,
which carries the commit SHA.

That commit is what a deploy built from the tag reports on `/health`.

### Release evidence: the CI runs for the tagged commit

The SBOM and the attestation say what a release contains and where it was
built. They do not say the commit passed its gates. That evidence is in two
hosted CI runs, and every release archives both with its release notes, because
Actions logs expire and a link to an expired run proves nothing.

| Run | Event | Gates it carries |
|---|---|---|
| `ci.yml` on the tagged commit | push to `main` | type check, lint, tests, identifier scan, dependency audit, TruffleHog secrets scan, CodeQL, Trivy image scan |
| `ci.yml` on the pull request that merged it | `pull_request` | the same source gates, plus container boot (non-root, healthcheck, key-store and logging assertions) and the [installer release gates](installer-release-gates.md) against the booted image |

Container boot and the installer gates run only on pull requests, and Trivy
only on the push. So no single run covers the release: the push run tested the
tagged commit itself, and the pull-request run tested that pull request's merge
with `main` as it stood then. The two trees are the same when `main` did not
move between that run and the merge. When it did, say so in the notes.

When the release is published, from a checkout of the canonical repository:

```bash
R=<owner>/m365-mcp; TAG=<tag>; SHA=$(git rev-list -n1 "$TAG")
PUSH_RUN=$(gh run list -R "$R" --workflow ci.yml --event push --commit "$SHA" \
  --json databaseId -q '.[0].databaseId')
PR=$(gh api "repos/$R/commits/$SHA/pulls" -q '.[0].number')
PR_HEAD=$(gh pr view "$PR" -R "$R" --json headRefOid -q .headRefOid)
PR_RUN=$(gh run list -R "$R" --workflow ci.yml --event pull_request --commit "$PR_HEAD" \
  --json databaseId -q '.[0].databaseId')

for run in "$PUSH_RUN" "$PR_RUN"; do
  gh run view "$run" -R "$R" --json url,headSha,event,conclusion,jobs \
    -q '"\(.event) \(.headSha) \(.conclusion) \(.url)", (.jobs[] | "  \(.conclusion)\t\(.name)")'
  gh run view "$run" -R "$R" --log > "ci-$run.log"
done | tee ci-evidence.txt

gh release upload "$TAG" ci-evidence.txt "ci-$PUSH_RUN.log" "ci-$PR_RUN.log" -R "$R"
```

Paste `ci-evidence.txt` into the release notes under **CI evidence** as well.
Every job reads `success` except three that are `skipped` by design: `DCO
Sign-off` and `Container Boot (PR)` on the push run, `Container Scan (Trivy)` on
the pull-request run. A release with any other job not green, or with
`Container Boot (PR)` skipped on the pull-request run, is not published.

These files are a record, not an attestation: they are not in `SHA256SUMS` and
nothing signs them. While the runs still exist, the runs are the authority and
the files let anyone check them; after they expire, the files are what is left.

## What a running instance is

An instance runs an image its operator built in their own fork and pushed to
their own registry, so the release attestation does not cover it directly: the
release proves what the source at a tag is and contains, and the three steps
below tie a running instance back to that source.

**1. The commit it reports.**

```bash
curl -s https://<your-host>/health
# {"status":"ok","sha":"<40-character commit>"}
```

The SHA is baked into the image at build time (`GIT_SHA`), and the deploy
workflow refuses to finish until `/health` reports the commit it built.

**2. The image it runs, and the provenance attached to it.** Each deploy pushes
an image index holding the image and a provenance statement. Find the digest
the app runs, then read the statement:

```bash
image=$(az containerapp show -n <app> -g <rg> \
  --query 'properties.template.containers[0].image' -o tsv)
az acr login -n <registry>
docker buildx imagetools inspect "$image" --format '{{ json .Provenance }}' > provenance.json
```

The statement names the fork and commit the image was built from (the BuildKit
`vcs` metadata), the `GIT_SHA` build argument, the deploy workflow run, and the
digest of every base image pulled. The commit should match step 1 and a commit
on the canonical repository's `main`; a deploying fork carries no commits of
its own, so any other commit is a finding.

This statement is **not signed**. It lives in your own registry, under your
own access control, and records what the build said about itself. GitHub only
signs attestations for public repositories and private ones on Enterprise
Cloud, and deploying forks are usually private. Signing it another way, with
keyless Sigstore signing, would write the fork's name and every deploy commit
to Sigstore's public transparency log, which most operators will not want.

**3. What is inside it.** The release SBOM for the tag that commit belongs to
lists the application's npm packages exactly, because they come from the same
lockfile. The OS layer can differ: the runtime base image is pinned to a
versioned host build, but an image built after a later base-image bump carries
that bump's OS packages and .NET runtime. For an exact inventory of the running
image, catalogue it directly:

```bash
syft scan "registry:$image" -o spdx-json > instance.spdx.json
```

## What runs where

| Job | When | Produces |
|---|---|---|
| `release.yml` `build` | release published; pull requests touching the Dockerfile, `package.json`, `package-lock.json` or the workflow | SBOMs, source archive, checksums, attestations (signed only on a public repository), verified in the same job |
| `release.yml` `publish` | release published | the files above, attached to the release |
| `deploy.yml` | every deploy | the image and its provenance statement, in your registry |
| `ci.yml` `container-scan` | every push to `main` | Trivy scan of the image (see [trivy-suppressions.md](trivy-suppressions.md)) |
