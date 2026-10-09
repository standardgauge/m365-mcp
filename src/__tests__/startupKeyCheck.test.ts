/**
 * Application keys are validated at startup, not on first use.
 *
 * credentialCrypto used to read MCP_SESSION_HMAC_KEY and MCP_DATA_ENCRYPTION_KEY
 * lazily, so an instance missing one booted, passed /health, and failed on the
 * first session write. The image's CMD now runs dist/startup/checkKeys.js ahead
 * of the Functions host; these tests hold the validator, the check's exit code,
 * and the Dockerfile wiring together.
 */
import { readFileSync } from 'fs';
import { join } from 'path';
import { randomBytes } from 'crypto';
import {
  cryptoKeyProblems,
  isCryptoConfigured,
  requireCryptoConfigured,
  hashSessionToken,
} from '../services/credentialCrypto.js';
import { runStartupKeyCheck } from '../startup/checkKeys.js';

const HMAC = 'MCP_SESSION_HMAC_KEY';
const DEK = 'MCP_DATA_ENCRYPTION_KEY';
const goodKey = (): string => randomBytes(32).toString('hex');
const validEnv = (): NodeJS.ProcessEnv => ({ [HMAC]: goodKey(), [DEK]: goodKey() });

function check(env: NodeJS.ProcessEnv): { code: number; output: string } {
  const lines: string[] = [];
  const code = runStartupKeyCheck(env, (l) => lines.push(l));
  return { code, output: lines.join('\n') };
}

describe('application key validation', () => {
  it('accepts two 64-hex-char keys', () => {
    expect(cryptoKeyProblems(validEnv())).toEqual([]);
    expect(isCryptoConfigured(validEnv())).toBe(true);
    expect(() => requireCryptoConfigured(validEnv())).not.toThrow();
  });

  it('accepts upper-case hex and ignores surrounding whitespace', () => {
    const env = { [HMAC]: `  ${goodKey().toUpperCase()}\n`, [DEK]: `${goodKey()}\r\n` };
    expect(cryptoKeyProblems(env)).toEqual([]);
  });

  describe.each([HMAC, DEK])('%s', (name) => {
    const withKey = (value: string | undefined): NodeJS.ProcessEnv => {
      const env = validEnv();
      if (value === undefined) delete env[name];
      else env[name] = value;
      return env;
    };

    it.each([
      ['missing', undefined, /is not set/],
      ['empty', '', /is not set/],
      ['whitespace only', '   \n', /is not set/],
      ['short (16 bytes)', randomBytes(16).toString('hex'), /exactly 32 bytes .* got 32 hex chars/],
      ['long (33 bytes)', randomBytes(33).toString('hex'), /exactly 32 bytes .* got 66 hex chars/],
      // Buffer.from(hex, 'hex') silently drops a trailing odd nibble, so this
      // used to decode to a valid-looking 32 bytes.
      ['odd length (65 chars)', `${goodKey()}a`, /exactly 32 bytes .* got 65 hex chars/],
      // ...and stops at the first non-hex character, so junk after a good key
      // also used to pass.
      ['valid key with junk appended', `${goodKey()}zz`, /non-hex characters/],
      ['non-hex characters', 'g'.repeat(64), /non-hex characters/],
      ['base64 instead of hex', randomBytes(32).toString('base64'), /non-hex characters/],
      ['interior whitespace', `${goodKey().slice(0, 32)} ${goodKey().slice(0, 32)}`, /non-hex characters/],
    ])('rejects a %s key', (_label, value, pattern) => {
      const env = withKey(value);
      const problems = cryptoKeyProblems(env);
      expect(problems).toHaveLength(1);
      expect(problems[0]).toMatch(new RegExp(`^${name} `));
      expect(problems[0]).toMatch(pattern);
      expect(isCryptoConfigured(env)).toBe(false);
      expect(() => requireCryptoConfigured(env)).toThrow(pattern);
    });

    it('never echoes key material in the error', () => {
      const secret = `${goodKey()}zz`;
      const [problem] = cryptoKeyProblems(withKey(secret));
      expect(problem).not.toContain(secret.slice(0, 64));
    });
  });

  it('reports both keys at once when both are bad', () => {
    const problems = cryptoKeyProblems({});
    expect(problems).toHaveLength(2);
    expect(problems[0]).toMatch(/^MCP_SESSION_HMAC_KEY env var is not set/);
    expect(problems[1]).toMatch(/^MCP_DATA_ENCRYPTION_KEY env var is not set/);
  });

  it('the lazy path rejects what the startup check rejects', () => {
    const saved = process.env[HMAC];
    try {
      process.env[HMAC] = `${goodKey()}a`;
      expect(() => hashSessionToken('t')).toThrow(/got 65 hex chars/);
    } finally {
      if (saved === undefined) delete process.env[HMAC];
      else process.env[HMAC] = saved;
    }
  });
});

describe('startup key check', () => {
  it('exits 0 with valid keys', () => {
    const { code, output } = check(validEnv());
    expect(code).toBe(0);
    expect(output).toMatch(/validated/);
  });

  it.each([
    ['missing', {}],
    ['short', { [HMAC]: randomBytes(16).toString('hex'), [DEK]: goodKey() }],
    ['malformed', { [HMAC]: goodKey(), [DEK]: 'not-a-key' }],
  ])('exits non-zero when a key is %s, naming the key', (_label, env) => {
    const { code, output } = check(env as NodeJS.ProcessEnv);
    expect(code).toBe(1);
    expect(output).toMatch(/FATAL: MCP_(SESSION_HMAC|DATA_ENCRYPTION)_KEY/);
    expect(output).toMatch(/Refusing to start/);
  });
});

describe('the image runs the key check before the host', () => {
  const root = join(__dirname, '..', '..');

  it.each(['Dockerfile', 'Dockerfile.backend.dev'])('%s', (file) => {
    const dockerfile = readFileSync(join(root, file), 'utf8');
    const cmds = dockerfile.match(/^CMD .*$/gm) ?? [];
    expect(cmds).toEqual([
      'CMD ["/bin/sh", "-c", "node dist/startup/checkKeys.js && exec /opt/startup/start_nonappservice.sh"]',
    ]);
    // No ENTRYPOINT that would turn the CMD into arguments of something else.
    expect(dockerfile).not.toMatch(/^ENTRYPOINT /m);
  });

  it('the build compiles src/startup so dist/startup/checkKeys.js exists', () => {
    const tsconfig = JSON.parse(readFileSync(join(root, 'tsconfig.json'), 'utf8'));
    expect(tsconfig.include).toContain('src/startup/**/*');
    expect(tsconfig.compilerOptions.outDir).toBe('dist');
    expect(tsconfig.compilerOptions.rootDir).toBe('src');
  });
});
