# Changelog

What changed in each release of this server, and what a deployment has to do
before it moves to it. The format is [Keep a Changelog 1.1.0](https://keepachangelog.com/en/1.1.0/),
with one group added for deployers, and versions follow
[Semantic Versioning](https://semver.org/spec/v2.0.0.html) as defined for this
project in [docs/releasing.md](docs/releasing.md). `scripts/release-check.sh`
checks this file's structure on every pull request.

Read **Upgrade notes** first. Anything there is an action a deployment takes
before or while it rolls forward: a Graph permission to consent, a setting to
add or rename, a Bicep change to apply. A release with no Upgrade notes rolls
forward on its image alone.

## [Unreleased]

### Added

- A release process: `vX.Y.Z` tags, this changelog, and
  `scripts/release-check.sh`, which CI runs on every pull request and the
  release workflow runs against the tag before it builds anything
  ([docs/releasing.md](docs/releasing.md)).
- Issue forms for bug reports, feature or control requests, and deployment
  questions.
- [docs/deployment-pattern.md](docs/deployment-pattern.md): the options for
  running an instance that stays current with this repository without a fork
  that drifts.

[Unreleased]: https://github.com/standardgauge/m365-mcp/commits/main
