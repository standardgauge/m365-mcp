# Working in this repository

Notes for anyone, human or agent, changing this code. The README covers what
the server is and how to deploy it; this file covers how changes are made.

## Build and check

```bash
npm ci
npx tsc --noEmit          # type check
npx eslint . --max-warnings 0   # lint; warnings fail, see below
npm test                  # Jest, the whole suite
bash scripts/identifier-scan.sh   # no deployment-specific names in the tree
```

Run all four before opening a pull request. CI runs the same set plus a
dependency audit, a secrets scan, CodeQL, and a container boot test that fails
if the image runs anything as root or if its installer surface fails its
release gates (`docs/installer-release-gates.md`).

## Lint warnings are failures

`npx eslint .` reports zero warnings on `main`, and CI runs it with
`--max-warnings 0` so it stays that way. The security rules
(`eslint-plugin-security`) are configured at `warn` because they are
heuristics, not proofs: `security/detect-object-injection` fires on every
`obj[key]` where `key` is a variable, and most of those index a lookup map by a
key from a constant list or type union.

When one fires on your change, do one of two things:

- Fix it. Look a caller-supplied name up in a `Map`, not a plain object, so
  there is no prototype chain to walk. Type a value instead of reaching for
  `any`.
- Suppress it on that one line with the reason attached:
  `// eslint-disable-next-line security/detect-object-injection -- key iterates
  the FIELDS constant`. The reason is the review artefact; a bare disable, a
  file-wide disable, or a disable without `--` is a finding.

A disable that no longer suppresses anything is itself a warning and fails the
build. Delete it rather than leave it.

## How a change lands

- One topic per pull request. A reviewer asked for a verdict on a stack of
  unrelated changes can only decline.
- The pull request title or body carries the tracker identifier the work is
  done under (for example `PROJ-123`). Automation discovers pull requests by that
  string; without it the change is not reviewed.
- `main` is protected: pull request required, the CI checks required, squash
  merges only, no force pushes. Nothing lands on `main` without a pull request,
  including from maintainers.
- Every finding in a review gets an explicit answer on the thread: fixed (with
  the commit), rebutted (with the evidence), or deferred (with the ticket).
  Silence on a finding reads as not having read the review.

## Deployments track this repository exactly

Organisations run this server from their own fork, in their own Azure
subscription and Entra tenant. A deploying fork's `main` is always equal to this
repository's `main`: it carries no commits of its own. Everything that differs
between deployments lives outside the tree, in that fork's repository variables
and secrets (`.github/workflows/deploy.yml` documents them) and in the Container
App's own configuration.

Consequences for a change here:

- Anything that reaches the tree reaches every tracking deployment on its next
  sync, and a tracking deployment deploys it as soon as CI passes there. Treat
  `main` as production.
- Never add a scope to `GRAPH_SCOPES` for one deployment's benefit. A scope a
  tenant has not consented to makes `acquireTokenSilent` throw
  `interaction_required` and breaks every tool call for signed-in users on
  every other deployment. Permissions are granted at the app registration.
- The container port is pinned in three places (the Dockerfile, the Bicep
  `targetPort`, each deploy workflow's `CONTAINER_PORT`) and
  `src/__tests__/containerPortInvariant.test.ts` holds them together. Change all
  of them or none.
- Nothing about any particular deployment belongs in the tree: no hostnames,
  resource names, people, mail domains or tracker ids in docs, fixtures or
  comments. `scripts/identifier-scan.sh` is the gate; fixtures use
  `example.com`, `fabrikam.com` and the like.

## Docs

`docs/` explains the code and how to operate an instance. A design note that
only makes sense with the history of one deployment does not go here.
