# Contributing

## Pull requests are open

This is the development repository: every change lands here first and every
deployment tracks it. Outside pull requests are welcome.

Contributions arrive under the licence this project already uses. There is no
contributor licence agreement to sign and nothing is assigned to anyone. The
project is AGPL-3.0, your contribution is AGPL-3.0, and your copyright stays
yours. That is possible because there is no proprietary edition of this server
and no plan to relicense it, so no one needs rights beyond the AGPL to keep the
project's options open.

What stands in place of an agreement is a sign-off on each commit.

### Sign your commits

```bash
git commit -s -m "Fix the thing"
```

which appends a trailer naming you:

```text
Signed-off-by: Your Name <you@example.com>
```

The address has to match the commit author. By adding it you certify the
[Developer Certificate of Origin 1.1](https://developercertificate.org/): that
you wrote the change, or otherwise have the right to submit it under this
project's licence.

Check a range before you push:

```bash
bash scripts/dco-check.sh origin/main HEAD
```

**If your employer owns the code you write, the sign-off says you have their
permission to send it.** Some employers restrict contributing to AGPL projects
specifically. Find that out before you write the change rather than after.

An unsigned commit is a mechanical problem with a mechanical fix: `git commit
--amend -s` for the last one, or rebase and sign the range. It is not a
judgement about the change.

### How a change lands

[AGENTS.md](AGENTS.md) has the rest: one topic per pull request, the checks to
run first, and how review findings get answered. It applies to everyone,
including maintainers.

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

This project is AGPL-3.0. By filing an issue you are not assigning anything; by
opening a pull request you are contributing under that licence, and saying so
with the sign-off above.
