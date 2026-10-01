# Contributing

## Pull requests are not accepted right now

Development happens in a private repository and this repository is published as
a snapshot of it. A pull request here cannot be merged; it would have to be
replayed by hand upstream, and we would rather say so than leave you to find
out after writing the patch.

This is a current constraint, not a position on outside contribution. If that
changes, this file changes with it.

## What is useful

**Issues are welcome and are read.** Bug reports, reproductions, documentation
that is wrong or missing, and questions about deploying this into a tenant are
all worth filing.

**Security issues go to [SECURITY.md](SECURITY.md), not to the issue tracker.**

**Forks are expected.** The deployment model is that you run your own instance
from your own fork, in your own Azure subscription and Entra tenant. The
licence permits that and nothing here discourages it.

## If you are filing a bug

Include the commit the running instance reports at `/health`, what you called,
what you expected, and what happened. If it involves a permission or a deny
rule, say which, because most surprising behaviour in this server is a policy
evaluating exactly as configured.

## Licence

This project is AGPL-3.0. By filing an issue you are not assigning anything;
by opening a pull request, were we to accept one, you would be contributing
under that licence.
