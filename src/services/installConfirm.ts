import { createHash, timingSafeEqual } from 'crypto';

export const INSTALL_CONFIRM_PATH = '/api/auth/install-confirm';
export const INSTALL_HANDOFF_COOKIE = 'install_handoff';

// Confirmation code for the install sign-in handoff.
//
// An installer opens /api/auth/login?install_nonce=<challenge> and polls
// install-poll with the verifier behind that challenge. Whoever wrote the link
// holds the verifier, so a link written by someone else and opened by a signed-in
// user would hand that user's session to its author. The callback therefore
// never attaches a session on its own: the browser has to enter the code the
// installer shows, and only a person looking at the installer has it.
//
// The code is derived from the challenge so the server can check it without a
// second URL parameter. It is not a secret from the link's author, who can
// compute it too; its job is to make the signed-in user prove they are looking
// at an installer, which someone who was only sent a link is not. The domain
// prefix keeps it from being read off the challenge in the address bar.
//
// install-mcp.sh, install-mcp.ps1 and the extension's server/index.js derive
// the same value; installConfirm.test.ts holds them to this function.
export const CONFIRM_CODE_DOMAIN = 'm365-mcp-install-confirm:';

/** Eight uppercase hex characters, without the display hyphen. */
export function installConfirmationCode(challenge: string): string {
  return createHash('sha256')
    .update(CONFIRM_CODE_DOMAIN + challenge)
    .digest('hex')
    .slice(0, 8)
    .toUpperCase();
}

/** `ABCD1234` → `ABCD-1234`, the form installers print. */
export function formatConfirmationCode(code: string): string {
  return `${code.slice(0, 4)}-${code.slice(4)}`;
}

/**
 * Compare what the user typed against the code for this challenge. Spaces and
 * hyphens are ignored, case is folded, and the letter O is read as zero since
 * the code never contains it.
 */
export function confirmationCodeMatches(challenge: string, typed: string): boolean {
  const normalized = typed.replace(/[\s-]/g, '').toUpperCase().replace(/O/g, '0');
  const expected = installConfirmationCode(challenge);
  if (normalized.length !== expected.length) return false;
  return timingSafeEqual(Buffer.from(normalized), Buffer.from(expected));
}
