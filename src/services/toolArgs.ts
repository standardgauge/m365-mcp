/**
 * Tool-argument validation for the native MCP endpoint.
 *
 * The connector's core invariant is: **a parameter is honored, or the call
 * fails — never silently ignored.** Every MCP tool already declares the exact
 * set of arguments it implements in its own `inputSchema.properties`, so that
 * declaration is the ground truth for what "honored" means. This helper compares
 * the caller-supplied arguments against that declared set and reports any that
 * the tool does not implement, so the dispatcher can reject them instead of
 * dropping them on the floor.
 *
 * Why central rather than per-handler: a silently-ignored parameter is the
 * single most common defect class on this surface ( date filters, the
 * calendar timezone in, and others), and it looks identical to success
 * from the caller's side. Enforcing it once in the dispatch path — grounded in
 * each tool's own schema — makes the failure impossible for every tool at once,
 * including tools added later, with no per-tool code to remember to write.
 *
 * This deliberately does NOT type-check or range-check argument values; the
 * individual handlers own semantic validation. It answers exactly one question:
 * did the caller pass a parameter this tool never declared?
 */

/**
 * A minimal view of a tool's JSON-Schema `inputSchema` (or of any nested
 * object/array schema reached during recursion). Only the members this check
 * reads are modeled: `properties` (object shape), `type`, and `items` (array
 * element shape).
 */
export interface ToolInputSchema {
  type?: unknown;
  properties?: Record<string, unknown>;
  items?: unknown;
}

/**
 * Return the caller-supplied argument names that the tool does not declare —
 * recursing into declared object and array-item schemas so a nested field is
 * held to the same rule as a top-level one. An empty array means every
 * argument, at every depth, is recognized.
 *
 * Nested rejections are reported as paths so the caller can find the offending
 * field: `homeAddress.building`, `contacts[0].assistantName`,
 * `attachments[1].retentionLabel`. Spreads in the schema
 * (e.g. `{ ...CONTACT_FIELD_PROPS }`) are already expanded at runtime, so
 * reading the live object's keys is correct.
 *
 * Recursion is bounded by the schema, not the input: it only descends where the
 * declared child schema itself declares `properties` (an object) or an array
 * whose `items` declare `properties`. Scalars, freeform objects with no
 * declared `properties`, and `string[]`-style arrays are left to the handlers'
 * own semantic validation — this check answers exactly one question at every
 * level: did the caller pass a field this schema never declared?
 *
 * `args` is tolerant of a non-object value (returns `[]`) so a malformed
 * `params.arguments` is handled by the handler's own required-field checks
 * rather than misreported here.
 */
export function findUnsupportedArgs(
  args: unknown,
  schema: ToolInputSchema | undefined,
  pathPrefix = '',
): string[] {
  if (!args || typeof args !== 'object' || Array.isArray(args)) return [];
  const props = (schema?.properties ?? {}) as Record<string, unknown>;
  const declared = new Set(Object.keys(props));
  const unsupported: string[] = [];
  for (const key of Object.keys(args as Record<string, unknown>)) {
    const path = pathPrefix ? `${pathPrefix}.${key}` : key;
    if (!declared.has(key)) {
      unsupported.push(path);
      continue;
    }
    unsupported.push(
      ...findUnsupportedInValue((args as Record<string, unknown>)[key], props[key], path),
    );
  }
  return unsupported;
}

/**
 * Descend into a single declared value against its declared schema. Recurses
 * only where there is a declared shape to check against: an object schema with
 * `properties`, or an array schema whose `items` declare `properties`. Anything
 * else (scalar, `string[]`, freeform object) yields no findings here.
 */
function findUnsupportedInValue(value: unknown, schema: unknown, path: string): string[] {
  if (!schema || typeof schema !== 'object') return [];
  const s = schema as ToolInputSchema;

  // Declared object: hold its sub-fields to the same rule.
  if (s.properties && value && typeof value === 'object' && !Array.isArray(value)) {
    return findUnsupportedArgs(value, s, path);
  }

  // Declared array whose items are a structured object: check each element.
  if (s.type === 'array' && s.items && typeof s.items === 'object' && Array.isArray(value)) {
    const itemSchema = s.items as ToolInputSchema;
    if (itemSchema.properties) {
      const out: string[] = [];
      value.forEach((el, i) => {
        out.push(...findUnsupportedArgs(el, itemSchema, `${path}[${i}]`));
      });
      return out;
    }
  }

  return [];
}

/**
 * Build the caller-facing message for a rejected call. Names the offending
 * parameter(s) and lists what the tool actually accepts, so the fix is obvious
 * from the error alone rather than requiring the caller to re-read the schema.
 */
export function unsupportedArgsMessage(
  toolName: string,
  unsupported: string[],
  schema: ToolInputSchema | undefined,
): string {
  const accepted = Object.keys(schema?.properties ?? {});
  const acceptedList = accepted.length > 0 ? accepted.join(', ') : '(this tool takes no parameters)';
  return (
    `Unsupported parameter(s) for ${toolName}: ${unsupported.join(', ')}. ` +
    `This connector rejects parameters it does not implement rather than ignoring them, ` +
    `so a call that looks accepted always did what was asked. ` +
    `Accepted parameters: ${acceptedList}.`
  );
}
