// Runtime configuration injected by serveAdmin as a JSON data block
// (<script type="application/json" id="runtime-config">). A data block is never
// executed, so the admin CSP can refuse inline script entirely.
export interface RuntimeConfig {
  instanceName: string;
}

export function readRuntimeConfig(doc: Document = document): RuntimeConfig {
  const fallback: RuntimeConfig = { instanceName: 'M365 MCP' };
  const el = doc.getElementById('runtime-config');
  if (!el?.textContent) return fallback;
  try {
    const parsed = JSON.parse(el.textContent) as Partial<RuntimeConfig>;
    return {
      instanceName: typeof parsed.instanceName === 'string' && parsed.instanceName ? parsed.instanceName : fallback.instanceName,
    };
  } catch {
    return fallback;
  }
}
