/**
 * Serialize a value for safe embedding inside an inline <script> block.
 *
 * Plain JSON.stringify does not escape `<`, so a value containing the literal
 * `</script>` (or `<!--`) would close the script element early and inject
 * markup into the page. Escaping `<` to its `<` unicode form keeps the
 * value a valid JS string literal while making an early `</script>`
 * impossible. Only exploitable by whoever controls the deployment env vars, so
 * this is defense-in-depth hardening (F13).
 */
export function jsonForScript(value: string): string {
  return JSON.stringify(value).replace(/</g, '\\u003c');
}
