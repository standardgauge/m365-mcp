# Deployment pattern: staying current without a fork that drifts

**Status:** proposal. It compares two supported deployment patterns and
recommends one. Nothing in the code changes until a maintainer picks. The
release process it depends on ([releases](releasing.md)) is in place.

## The problem

The people who deploy this server are often running more than one instance: an
internal IT lead with a production and a test tenant, a managed service
provider or a fractional CTO with one instance per client tenant. Each instance
has to stay current with this repository, because fixes and controls land here,
and each has to stay under its owner's change control, because the server holds
delegated Graph access to mail, files and calendars.

The documented way to do both today is a fork per instance whose `main` is kept
equal to this repository's `main`. The instances this project's maintainers run
are deployed exactly that way. This note looks at what that has cost, and at two
ways to make following upstream cheaper than forking it.

## How deployments track upstream today

- **The fork carries no code of its own.** Its `main` equals upstream `main`
  ([`AGENTS.md`](../AGENTS.md)). Everything that differs per instance lives in
  the fork's repository variables and secrets (the header of
  [`deploy.yml`](../.github/workflows/deploy.yml) lists them), in the Container
  App's own configuration, and in Key Vault.
- **Something has to move the fork.** GitHub does not update a fork by itself,
  so each fork needs a sync: a scheduled job that fast-forwards it to upstream
  `main` and refuses one that has diverged.
- **The fork's CI gates its deploy.** With `AUTO_UPDATE=true`, the fork's
  `Deploy` workflow runs when CI passes on the new tip of the fork's `main`,
  builds the image into the instance's own registry, rolls the Container App,
  and proves it: `/health` has to report the commit, the admin UI has to load,
  and an optional functional smoke makes one real Graph read per service. With
  `AUTO_UPDATE` unset the instance is pinned, and only a manual dispatch
  deploys.

The maintainers' own sync runs outside GitHub, as a scheduled job every 15
minutes. It advances the forks in canary order (the first instance moves on
every upstream merge, the others only once the first one's `/health` serves the
new commit) and raises an alarm when it has not succeeded for two hours.

## What has gone wrong

None of these was a fork carrying its own commits. Every one was the update
path.

1. **Two instances, two behaviours, one tool.** Two forks synced at different
   commits, and the mail search tool built a subject-substring `$filter` on one
   and a KQL `$search` on the other. A date-bounded query returned nothing on
   the older instance, with no error. Nothing named the difference: each
   instance's `/health` reported a commit hash, and there was no version or
   changelog to say what lay between them.
2. **The sync could not carry the files that define the deploy.** The first
   sync ran inside each fork as a GitHub Actions workflow. GitHub refuses a push
   from a workflow's `GITHUB_TOKEN` that touches `.github/workflows/`, and that
   permission cannot be granted to the token. So the sync carried ordinary code
   and could never carry a change to the deploy workflow itself. It went
   unnoticed because its green runs were all no-ops. The sync had to move to a
   job outside GitHub holding a token with workflow-push rights on every fork.
3. **A stopped sync is silent.** The forks froze on one build for five weeks
   because the sync had stopped, and nothing was watching for it. GitHub can
   also disable a scheduled workflow after 60 days without repository
   activity, so an in-fork check can stop without a trace. The fix was an
   alarm outside the forks, and the sync's credential still has an expiry
   that someone has to renew.
4. **Pinned meant "whatever was last dispatched".** An instance under change
   control could hold still, but there was nothing to hold it to except a
   commit hash, and no notes to read before moving it.

So the cost of following upstream today is a sync job, a cross-repository
credential, an alarm, a canary gate, and one fork per instance, all run by the
deployer. The maintainers built that machinery once. An outside deployer with
twenty client tenants would have to build it again, and without it their forks
drift.

## What any supported pattern has to keep

- **The deploy proof.** `/health` reports the deployed commit, the admin UI
  loads, the functional smoke reads Graph. A deploy that is not proven is not
  done.
- **Per-instance change control.** Each instance moves when its owner says so,
  and can be held.
- **Secrets stay in Key Vault and the deployer's own secret store.** None in
  any repository, never in a parameters file.
- **Nothing about any instance in this repository** (the identifier scan).
- **Existing forks keep working** until they choose to move. The fork sync and
  `deploy.yml`'s current triggers stay as they are.

Releases are a prerequisite for both options: each needs a version to pin, and
notes to read before moving it. Those now exist ([releases](releasing.md)).

## Option A: a thin deploy repository that pins a release

The deployer keeps a small repository of their own with no server code in it:

```text
my-m365-mcp-deploy/
  .github/workflows/deploy.yml      # ~20 lines: calls upstream at a tag
  .github/dependabot.yml            # opens the upgrade pull request
  infra/production.parameters.json  # non-secret Bicep parameters
```

The workflow calls this repository's deploy workflow as a reusable workflow,
pinned to a release. The deploy repository belongs to the deployer's own
account or organisation, which is almost never this repository's, so the
caller cannot use `secrets: inherit`: GitHub passes inherited secrets only
between workflows in the
[same organisation or enterprise](https://docs.github.com/en/actions/how-tos/reuse-automations/reuse-workflows#passing-inputs-and-secrets-to-a-reusable-workflow).
Across that boundary a secret reaches the called workflow in one of two ways:
it is a secret the reusable workflow declares by name and the caller maps
explicitly, or it is an environment secret in the caller's repository that the
reusable workflow's own job reads by setting `environment`. The interface below
supports both, and recommends the second.

### The reusable workflow's interface

`deploy.yml` would declare its secrets by name. Each is `required: false`
because an environment can supply it instead of the caller, so the
"Check required configuration" step has to gain a check that fails the run
when `AZURE_CREDENTIALS` arrives from neither.

```yaml
on:
  workflow_call:
    inputs:
      environment:          # a GitHub environment in the CALLER's repository
        type: string
        required: true
      resource-group:
        type: string
        required: true
      container-app:
        type: string
        required: true
      acr-name:
        type: string
        required: true
      smoke-expect-services:
        type: string
        required: false
    secrets:
      AZURE_CREDENTIALS:    # service principal JSON, as documented today
        required: false
      SMOKE_AGENT_KEY:      # read-only agent key; unset skips the smoke job
        required: false

jobs:
  build-push-deploy:
    environment: ${{ inputs.environment }}
    concurrency:
      group: deploy-${{ inputs.environment }}
      cancel-in-progress: false
    # ...every existing step, reading secrets.AZURE_CREDENTIALS and
    # secrets.SMOKE_AGENT_KEY exactly as it does now
```

Three GitHub rules decide how this behaves when a caller in another
organisation runs it:

- **The environment belongs to the caller.** A called workflow runs in the
  caller's context (its `github` context is the caller's), so
  `environment: contoso` names an environment in the deploy repository, with
  that repository's required reviewers, wait timer, deployment-branch policy
  and environment secrets. Nothing about any instance is configured here.
- **Only the called job can name the environment.** A job that calls a
  reusable workflow accepts `uses`, `with`, `secrets`, `strategy`, `needs`,
  `if`, `concurrency`, `permissions` and `name`, and not `environment`. So the
  caller cannot read an environment secret and map it across; the environment
  has to be set inside `deploy.yml`, which is why it is an input.
- **An environment secret wins over a passed one of the same name.** If the
  environment holds `AZURE_CREDENTIALS`, the job uses it and ignores whatever
  the caller mapped. Keep each secret in one place per instance, or the
  reviewed caller file stops being the truth about which credential deployed.

Every job that touches Azure or the instance (the deploy and the functional
smoke) sets the same `environment`. The concurrency group keys on it, so one
tenant's deploy does not queue behind another's, while two deploys to one
tenant still serialise. The interface (input names, secret names, the
environment rule) is public under semver from the first release that ships it.

### Recommended caller: one environment per instance

Each instance gets a GitHub environment in the deploy repository holding its
`AZURE_CREDENTIALS` and `SMOKE_AGENT_KEY` as **environment secrets**, with
required reviewers set to whoever approves changes to that instance and
deployment branches limited to `main`. The caller passes no secrets at all:

```yaml
# my-m365-mcp-deploy/.github/workflows/deploy.yml
on:
  push:
    branches: [main]
  workflow_dispatch:

permissions:
  contents: read

jobs:
  production:
    uses: standardgauge/m365-mcp/.github/workflows/deploy.yml@v1.4.0
    with:
      environment: production
      resource-group: rg-m365-mcp
      container-app: m365-mcp
      acr-name: m365mcpacr
      smoke-expect-services: mail,calendar,sharepoint
    # no secrets: block. AZURE_CREDENTIALS and SMOKE_AGENT_KEY are
    # environment secrets on "production", read by deploy.yml's own jobs.
```

This is the model to use for anything under change control. The credential is
released only to a job that has passed the environment's protection rules: a
named reviewer approves the run, on `main`, before any step can read it. A
repository or organisation secret has no such gate: any workflow run in the
deploy repository can read it, on any branch someone with write access pushes.

A provider with many client tenants keeps one deploy repository, one
environment per tenant and one caller job per tenant, each with its own
credential and reviewers, instead of one fork per tenant:

```yaml
jobs:
  contoso:
    uses: standardgauge/m365-mcp/.github/workflows/deploy.yml@v1.4.0
    with:
      environment: contoso
      resource-group: rg-m365-mcp-contoso
      container-app: m365-mcp
      acr-name: contosom365mcp
  fabrikam:
    needs: contoso          # optional canary order: fabrikam after contoso
    uses: standardgauge/m365-mcp/.github/workflows/deploy.yml@v1.4.0
    with:
      environment: fabrikam
      resource-group: rg-m365-mcp-fabrikam
      container-app: m365-mcp
      acr-name: fabrikamm365mcp
```

Contoso's reviewers approve Contoso's job and nobody else's. A tenant can be
held by leaving its job's approval pending or by pinning its job to an older
tag, and `needs` gives the canary order the fork sync provides today.

### Alternative caller: explicit secret mapping

A deployer who keeps credentials as repository or organisation secrets (a
single instance with no reviewers, say) maps each one explicitly by name:

```yaml
jobs:
  production:
    uses: standardgauge/m365-mcp/.github/workflows/deploy.yml@v1.4.0
    with:
      environment: production
      resource-group: rg-m365-mcp
      container-app: m365-mcp
      acr-name: m365mcpacr
    secrets:
      AZURE_CREDENTIALS: ${{ secrets.PRODUCTION_AZURE_CREDENTIALS }}
      SMOKE_AGENT_KEY: ${{ secrets.PRODUCTION_SMOKE_AGENT_KEY }}
```

The `environment` input is still required, so the run still waits on that
environment's reviewers if it has any, but the secret itself is readable,
without that approval, by any workflow run in the deploy repository. And per
the precedence rule above, an environment secret of the same name silently
replaces the mapped one. Use this form only where that is acceptable, and
never hold the same secret both ways.

### Not depending on a stored credential

`AZURE_CREDENTIALS` is a long-lived service principal secret. `azure/login`
also accepts OpenID Connect: the job presents a GitHub-issued token and Azure
exchanges it for an access token through a federated credential on the app
registration, so no Azure secret is stored in GitHub at all. The token's
subject names the caller's repository and environment
(`repo:<owner>/<deploy-repo>:environment:<name>`), so one federated credential
per environment binds an Azure identity to exactly one instance's environment
and its reviewers. The client, tenant and subscription IDs are not secrets and
can be inputs. It is worth adding to the interface as the preferred
credential, alongside `AZURE_CREDENTIALS` for existing deployers, when
`workflow_call` is built; it needs `id-token: write` on the calling job.

**Upgrading** is a pull request in the deploy repository that changes
`@v1.4.0` to `@v1.5.0`. Dependabot's `github-actions` ecosystem already opens
that pull request for reusable workflow references, with the release notes in
its body, so the Upgrade notes are in front of the reviewer at the moment they
matter. Merging it is the change-control approval, and the deploy runs on the
merge.

**What upstream would change** (none of it built yet):

- `deploy.yml` gains a `workflow_call` trigger alongside its current ones. It
  checks out this repository at the called ref rather than the caller's
  repository, takes the resource names and the caller's environment as
  inputs, declares `AZURE_CREDENTIALS` and `SMOKE_AGENT_KEY` as named secrets
  (never relying on `secrets: inherit`), sets `environment` on every job that
  reads them, and keeps every proof step. Those inputs and secret names become
  a public interface under semver.
- The gate on CI becomes "the tag is a release": `release.yml` refuses to
  build or sign a tag that is not on `main` or whose commit has no passing CI
  push run ([releasing](releasing.md)), so the caller does not need CI of its
  own.
- Optionally, the reusable workflow applies the Bicep from the caller's
  parameters file (with `what-if` on the pull request), which brings
  infrastructure changes under the same pinned version. Today nothing applies
  Bicep after the first deploy; the deploy workflow converges the ingress port
  and the probes by hand because of that.
- A template deploy repository, so the first one is a "Use this template"
  click.

**What it removes:** the sync job, the cross-repository credential and its
annual renewal, the workflow-push problem (the deploy repository's own workflow
is twenty lines that rarely change; upstream workflow changes arrive through
the ref), and the drift question. There is no code in the deploy repository to
drift, and the version an instance runs is written in a file.

**What it costs:** the reusable workflow's inputs become an interface that has
to stay compatible within a major version. Existing forks have to migrate to
benefit, though nothing forces them to. Liveness moves rather than vanishing:
an unmerged Dependabot pull request is an instance falling behind, which is at
least visible in the place the deployer already looks, and is the deployer's
decision rather than a silent failure.

## Option B: no deploy repository at all

The deployer keeps no repository. This one publishes, per release, a signed,
tenant-neutral container image and a compiled ARM template. The image already
carries no tenant configuration: the app reads all of it at runtime.

```bash
az deployment group create -g "$RG" \
  --template-uri https://github.com/standardgauge/m365-mcp/releases/download/v1.4.0/main.json \
  --parameters @production.parameters.json \
    azureClientSecret=... mcpSessionHmacKey="$HMAC" mcpDataEncryptionKey="$DEK"

# upgrading
az containerapp update -n "$APP" -g "$RG" \
  --image ghcr.io/standardgauge/m365-mcp@sha256:<digest of v1.5.0>
```

The parameters file lives wherever the deployer keeps configuration.

**What upstream would change:**

- `release.yml` pushes the image to a public registry and attests its digest.
  Today it builds the image only to catalogue it, and never pushes it.
- The Bicep takes an external image and registry. Today it creates a registry
  per instance, composes the image name from it, and grants pull on it.
- The release attaches the compiled ARM JSON.
- A script that runs [`scripts/post-deploy-smoke.js`](../scripts/post-deploy-smoke.js)
  and the `/health` and admin UI checks after an upgrade, since no workflow
  runs them.

**What it removes:** GitHub from the deployer's side entirely, and every image
build. One image serves every instance, so a provider rolls a release across
all its tenants without building anything. It is also the only shape a "Deploy
to Azure" button can use (see [First deploy](#first-deploy)).

**What it costs:** the deploy proof becomes something the deployer has to
remember to run. An upgrade is an imperative command with no review artefact
unless the deployer builds one, and configuration changes move from a reviewed
file to the portal, unversioned. An instance pulls its image from a public
registry on every new revision and scale-out, and a tenant whose policy
requires images in its own registry is back to an import step per release. And
it rebuilds, per deployer, the change control that Option A gets from a pull
request.

## Comparison

| | Today: tracking fork | A: thin deploy repo | B: no repo |
|---|---|---|---|
| Unit of update | a commit on `main` | a release tag | a release tag |
| What the deployer runs to stay current | sync job, credential, alarm | Dependabot | nothing; they must check for releases |
| Where change control happens | `AUTO_UPDATE` and a dispatch | a pull request with the notes in it | outside the system |
| Deploy proof | every deploy | every deploy | if they run the script |
| Can drift | yes, through the update path | no code to drift; version is in a file | no code; config in the portal |
| Instances per provider | one fork each | one environment each | one parameters file each |
| Upstream work | none | `workflow_call` with named secrets, template repo | image publishing, Bicep change, ARM asset, smoke script |
| Deploy to Azure button | no | no | yes |

## Recommendation

**Option A, with Option B's image publishing as a later step that A can use.**

A fixes what actually went wrong. Every failure above was in the update path,
and A replaces that path with a version string in a file and a pull request
that carries the release notes. That is the change-control model a deployer
under audit already has, and it keeps the deploy proof on every deploy, which
B gives up. B's gains are first-deploy friction and build time. Both are real,
but neither is what drifted.

The two compose. Once releases publish an image, A's reusable workflow can
import the released image by digest into the instance's registry instead of
building it, and B's button becomes possible for a first deploy that then
hands over to a deploy repository.

Order, if A is chosen:

1. Releases and the changelog. Done.
2. `workflow_call` on `deploy.yml`, additive, with the existing
   `workflow_run` and `workflow_dispatch` triggers unchanged so tracking forks
   are untouched, plus a template deploy repository.
3. Existing forks move one at a time, canary first. The fork sync retires when
   the last one has moved, not before.
4. Image publishing, and the Deploy to Azure evaluation below, revisited.

## First deploy

The go-live checklist ([`customer-go-live.md`](customer-go-live.md)) is the
supported first deploy. Two shortcuts were evaluated against it.

**A "Deploy to Azure" button: not now.** It needs a compiled ARM template at a
public URL, which is easy. It is blocked on three things that are not:

- **There is no image to deploy.** The template creates an empty registry and
  points the Container App at it, so the first revision fails until the fork's
  deploy workflow pushes an image (go-live item 1.2 says so). A button would
  deploy a broken app. It needs a published image first, which is Option B's
  first step.
- **The two application keys are generated fresh per instance**, in the shell,
  never from a file (go-live 1.1, checked at item 6). A portal form would have
  the operator paste them in, and generating them inside the template changes a
  guarantee the checklist verifies. That needs its own review.
- **The Entra registration and admin consent come first** (go-live 1.0 and 2),
  and they are the highest-trust step in the deployment. A button that hides
  that ordering works against the checklist's point.

**A single-command Bicep path: as part of Option A, not separately.** A script
wrapping items 1.1 and 1.2 and the first image would be a third deploy path
beside the README's manual steps and the deploy workflow, without the
workflow's proof. Under Option A the deploy repository's first workflow run
does it: apply the Bicep from the parameters file, import or build the image,
deploy, prove. The template repository makes that first run one click after the
Entra steps.

## The AGPL and hosted deployments

This server is licensed under the GNU AGPL, version 3. Section 13 of the
[licence](../LICENSE), "Remote Network Interaction; Use with the GNU General
Public License", applies when you **modify**
the program and users interact with it remotely over a network: it requires
offering those users the Corresponding Source of your version.

As facts relevant to the options above, not legal advice:

- An instance deployed from an unmodified release runs source published at that
  release's tag.
- A fork carrying its own commits, or an image built from patched source, runs a
  modified version.
- For an instance a provider hosts for a client tenant, the users interacting
  with it over the network are the client's staff.

Read section 13 itself, and ask your own counsel how it applies to your
deployment.
