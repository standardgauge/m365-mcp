# Supply chain: SBOMs and build provenance

Two questions come up whenever someone evaluates an instance of this server:
what is in it, and was it built from the source it claims. This page shows how
to answer both from evidence, for a release of this repository and for the
instance you actually run.

The pieces:

- every release carries an SBOM and a signed provenance attestation
  ([`release.yml`](../.github/workflows/release.yml));
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
lockfile. The OS layer can differ: the runtime base image tag
(`azure-functions/node:4-node22`) is rolling, so an image built a week after
the release may carry newer OS packages. For an exact inventory of the running
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
