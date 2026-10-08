/**
 * Tests for sanitizeKqlPhrase (F9).
 *
 * Graph mail $search embeds the query inside a double-quoted KQL phrase
 * (`.search(`"${q}"`)`). An embedded `"` closes the phrase early and lets the
 * remainder be parsed as additional KQL operators, widening the query scope.
 * sanitizeKqlPhrase strips the double-quotes so the whole input stays a single
 * phrase.
 */

import { sanitizeKqlPhrase } from '../services/kqlSearch.js';

describe('sanitizeKqlPhrase', () => {
  test('leaves an ordinary query unchanged', () => {
    expect(sanitizeKqlPhrase('subject:budget from:alice')).toBe('subject:budget from:alice');
  });

  test('strips a double-quote that would break out of the KQL phrase', () => {
    // Without sanitisation this becomes: "" OR from:ceo@corp.com " — a scope-widening breakout.
    const malicious = '" OR from:ceo@corp.com "';
    const cleaned = sanitizeKqlPhrase(malicious);
    expect(cleaned).not.toContain('"');
  });

  test('removes every double-quote, not just the first', () => {
    expect(sanitizeKqlPhrase('a"b"c"d')).not.toContain('"');
  });

  test('the sanitised value re-embeds as a single closed phrase', () => {
    const embedded = `"${sanitizeKqlPhrase('report "urgent" q3')}"`;
    // Exactly the opening and closing quotes remain — no interior quote.
    expect((embedded.match(/"/g) ?? []).length).toBe(2);
  });

  test('preserves non-quote punctuation used by KQL (colon, @, parens)', () => {
    expect(sanitizeKqlPhrase('from:a@b.com (budget)')).toBe('from:a@b.com (budget)');
  });

  test('strips a backslash, which would escape the closing quote of the $search string', () => {
    expect(sanitizeKqlPhrase('alice\\')).not.toContain('\\');
  });

  test('empty string stays empty', () => {
    expect(sanitizeKqlPhrase('')).toBe('');
  });
});
