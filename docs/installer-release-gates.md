# Installer release gates

The install routes (`/install`, `/install.mcpb`, `/install.sh`, `/install.ps1`,
`/install/m365-mcp-shim.js`, `/api/extension-version` and
`/api/extension-update`) are public by design. They serve code that runs on
every user's machine and goes on to hold that user's session token, and the
desktop extension overwrites itself with whatever the update route returns.
That makes them the surface with the widest blast radius this server has, even
though nobody needs to be signed in to reach them.

Three properties keep that safe ([threat model](threat-model.md) rows 3.5 and
4.3). Each was already true and each already had a test. The gates described
here are what make them conditions a release has to pass rather than
properties that happen to hold: they run on every pull request, against the
source and against the built image, and a regression in any of them turns CI
red.

| Gate | What must hold | What a regression looks like |
|---|---|---|
| Origin | Every URL in served code is the origin of `OAUTH_REDIRECT_URI`, or one of two prerequisite links (`nodejs.org`, `learn.microsoft.com`). Nothing from the request's `Host`, `X-Forwarded-Host`, `X-Forwarded-Proto`, `X-Forwarded-Server`, `X-Original-Host` or `Forwarded` headers reaches a client, and with no origin configured the routes serve no URL at all rather than guess one. | A handler that falls back to a forwarded host, so whoever can set a header picks the server, and so the update source, of every client installed from that response. |
| Signing | `/api/extension-update` carries an Ed25519 signature that verifies, under Node's own `crypto`, against the `UPDATE_PUBLIC_KEY` baked into the served `server/index.js`; the bundle and the update bake in the same key; the signed version and files are the ones advertised; the served extension verifies before it applies and never applies the unsigned block; without `MCP_SESSION_HMAC_KEY` the instance serves nothing. | An update response without a `signed` block, a key that is not Ed25519, an extension that applies `update.files`, or a build that serves unsigned code when the key is missing. |
| No token in config | The entries the installers write into the MCP client configs name the credential store (`--token-store`) and hold no token; the backup of an earlier config has that installer's bearer header removed, and is mode 0600 from `install.sh`; the extension manifest starts the bridge with an empty `env` and no token in `args`. | A config writer that puts the session token back on the command line, where it sits in a plain JSON file and in every process listing. |

## Where the gates run

**On every `npm test`.**
[`src/__tests__/installReleaseGates.test.ts`](../src/__tests__/installReleaseGates.test.ts)
reads every route the install module registers back from its registrations and
calls each with forged host headers, so a route added later is covered the day
it lands. It also checks the source: the install module may not read a host
header or the request URL, and the installer templates take the server URL
from the rendered placeholder and nothing else, so a header read that does not
yet change the served text still fails. For the config gate it extracts the
config-writing function from the served installer and runs it, under `bash`
and `python3` for `install.sh` and under `pwsh` for `install.ps1`, with a
session token in scope and against a config an earlier installer wrote. The
PowerShell run is skipped on a machine without `pwsh` and required in CI.
After `npm run build` it also checks that `dist/install`, which is what the
image carries, is byte for byte `src/install`, which is what the tests
exercised.

**On every pull request, against the image.** The `container-boot` job in
[`ci.yml`](../.github/workflows/ci.yml) boots the image it just built with
`OAUTH_REDIRECT_URI` set to `https://mcp.example.com/api/auth/callback`, then
runs
[`scripts/installer-release-gates.js`](../scripts/installer-release-gates.js)
against it through `localhost` with forged `Host`, `X-Forwarded-Host`,
`X-Forwarded-Proto` and `Forwarded` headers on every request. The script is
the same code the test imports. The test feeds it the handlers' output; the job
feeds it what the container actually serves.

**On a deployment.** The same script, pointed at the instance:

```bash
node scripts/installer-release-gates.js https://your-mcp-host.example.com https://your-mcp-host.example.com
```

This is the automated form of items 9.1 and 9.2 of the
[go-live checklist](customer-go-live.md#9-client-surface-served-origin-and-update-signing).
Exit 0 means every gate passed, and the output lists each check. It needs
`node`, `bash` and `python3`, plus `pwsh` for the PowerShell config writer,
which is skipped without it unless `CI` is set.

## What the gates do not cover

- Clients installed before 2.11.0 do not verify update signatures (threat
  model 4.3). The gate holds what the instance serves today.
- A config an earlier installer wrote keeps its token until that installer is
  run again (threat model 3.5). The gate checks the writer, not the files
  already on users' machines.
- The CI container has no storage account and no Entra configuration, so the
  sign-in handoff itself (go-live item 9.3) is not exercised there;
  `installConfirm.test.ts` and `installPoll.test.ts` cover it.
