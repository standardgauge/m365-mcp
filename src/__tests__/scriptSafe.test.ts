/**
 * Tests for jsonForScript (F13).
 *
 * serveAdmin injects runtime config into an inline <script> via
 * `window.__X__=${jsonForScript(value)}`. A value containing `</script>` must
 * not be able to close the script element early.
 */

import { jsonForScript } from '../services/scriptSafe.js';

describe('jsonForScript', () => {
  test('round-trips an ordinary value as valid JSON', () => {
    expect(JSON.parse(jsonForScript('abc-123'))).toBe('abc-123');
  });

  test('does not emit a literal </script> for a value containing it', () => {
    const out = jsonForScript('</script><script>alert(1)</script>');
    expect(out.toLowerCase()).not.toContain('</script');
    // The escaped form still parses back to the original string.
    expect(JSON.parse(out)).toBe('</script><script>alert(1)</script>');
  });

  test('escapes every < to \\u003c', () => {
    const out = jsonForScript('a<b<c');
    expect(out).not.toContain('<');
    expect(out).toContain('\\u003c');
  });

  test('leaves values with no < unchanged from plain JSON.stringify', () => {
    expect(jsonForScript('11111111-2222-3333-4444-555555555555')).toBe(
      JSON.stringify('11111111-2222-3333-4444-555555555555')
    );
  });
});
