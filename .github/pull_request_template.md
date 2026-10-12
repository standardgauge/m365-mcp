## Description

<!-- What does this PR do? Why? -->

## Sign-off

- [ ] Every commit carries `Signed-off-by` with the address matching its author (`git commit -s`), per [CONTRIBUTING.md](../CONTRIBUTING.md)

## Security Checklist

- [ ] Does this change touch auth or middleware? (`src/functions/auth/`, `src/middleware/`)
- [ ] Does this change touch the deny list? (`src/services/denyList.ts`, `getDenyList`, `DenyListManager`)
- [ ] Does this change touch the OAuth flow? (MSAL config, token cache, callback handler)
- [ ] Does this change add, remove, or upgrade dependencies? (run `npm audit` locally first)

## Testing

- [ ] Existing tests pass (`npm test`)
- [ ] New behavior is covered by tests (or explain why not)

## Notes for reviewers

<!-- Anything else the reviewer should know -->
