#!/usr/bin/env node
/**
 * post-deploy-smoke.js: a functional probe of a running instance, driven by a
 * read-only agent key.
 *
 * The unit suite mocks the SDK and Graph layers, so it cannot see a defect that
 * only shows up when a real call reaches a real tenant: a storage client built
 * the wrong way, a Graph permission the tenant never consented to, a query
 * option an endpoint rejects. Each of those has shipped green and failed on the
 * first live call. The deploy workflow's own smoke step proves the right image
 * is serving; this proves it works, by making one real read per service the key
 * holds and checking that what a read-only key must be refused is refused.
 *
 * The key is an ordinary agent key (docs/agent-api-keys.md), minted by the
 * tenant's administrator under My Settings -> Agent Keys with every service the
 * tenant enables, every one of them read-only, and stored as the deploy fork's
 * SMOKE_AGENT_KEY repository secret. It acts as its owner, so every call here
 * is a delegated call audited under the key's device label.
 *
 *   SMOKE_AGENT_KEY=m365k_... node scripts/post-deploy-smoke.js https://mcp.example.com [expected-sha]
 *
 * Optional environment:
 *   SMOKE_EXPECT_SERVICES  comma-separated services the key should reach
 *                          (mail,calendar,...). When set, a service missing
 *                          from tools/list is a failure, not a quieter run.
 *   SMOKE_REPORT_FILE      write a Markdown report of the run here.
 *
 * Exit 0 when every check passes, or when SMOKE_AGENT_KEY is unset (the probe
 * is skipped with a notice); 1 when a check fails; 2 on usage.
 *
 * Nothing here can write. Write tools are called only to prove they are
 * refused, only when tools/list has already shown the key is read-only, and
 * with an id the server's own argument validation rejects before any Graph
 * call, so a broken read-only gate still writes nothing. The logout probe runs
 * last: if that gate were broken it would end the key's own session.
 */
'use strict';

/** Every service the server knows, in the order the probe walks them. */
const SERVICES = ['mail', 'calendar', 'sharepoint', 'onedrive', 'onenote', 'contacts', 'teams'];

/**
 * One read per service: the tool, and arguments that make it a real Graph call.
 * list_events takes a window so it goes through calendarView, the path
 * assistants use for "what is on today".
 */
const READ_PROBES = {
  mail: { tool: 'list_folders_mail', args: () => ({}) },
  calendar: {
    tool: 'list_events',
    args: (now = new Date()) => {
      const day = (d) => d.toISOString().slice(0, 10);
      return { startDateTime: day(now), endDateTime: day(new Date(now.getTime() + 86400000)) };
    },
  },
  sharepoint: { tool: 'list_sites', args: () => ({}) },
  onedrive: { tool: 'list_onedrive', args: () => ({}) },
  onenote: { tool: 'list_notebooks', args: () => ({}) },
  contacts: { tool: 'list_contact_folders', args: () => ({}) },
  teams: { tool: 'list_teams', args: () => ({}) },
};

/**
 * An id the server's opaque-id validation refuses (it contains a slash), so a
 * write that got past the read-only gate would still stop before Graph.
 */
const REFUSED_ID = 'post-deploy-smoke/refused';

/** One write per service, each with arguments that can never reach Graph. */
const WRITE_PROBES = {
  mail: { tool: 'delete_message', args: { messageId: REFUSED_ID } },
  calendar: { tool: 'delete_event', args: { eventId: REFUSED_ID } },
  sharepoint: { tool: 'delete_sharepoint_file', args: { siteId: REFUSED_ID, itemId: REFUSED_ID } },
  onedrive: { tool: 'delete_onedrive_item', args: { itemId: REFUSED_ID } },
  onenote: { tool: 'create_section', args: { notebookId: REFUSED_ID, displayName: 'post-deploy smoke' } },
  contacts: { tool: 'delete_contact', args: { contactId: REFUSED_ID } },
  teams: { tool: 'send_channel_message', args: { teamId: REFUSED_ID, channelId: REFUSED_ID, content: 'post-deploy smoke' } },
};

/**
 * The server's WRITE_TOOLS (src/functions/mcp/mcpEndpoint.ts). A read-only key
 * must see none of them in tools/list. postDeploySmoke.test.ts holds the two
 * lists equal.
 */
const WRITE_TOOLS = [
  'write_sharepoint_file', 'delete_sharepoint_file', 'create_sharepoint_folder',
  'create_sharepoint_list_item', 'update_sharepoint_list_item', 'delete_sharepoint_list_item',
  'create_mail_folder', 'rename_mail_folder', 'move_mail_folder',
  'create_draft', 'update_message', 'send_mail', 'move_message', 'delete_message',
  'reply_to_message', 'reply_all_to_message', 'forward_message', 'send_draft',
  'write_onedrive_file', 'move_onedrive_item', 'create_onedrive_folder', 'delete_onedrive_item',
  'create_calendar', 'create_event', 'update_event', 'delete_event', 'move_event', 'respond_to_event',
  'create_contact', 'update_contact', 'delete_contact',
  'create_contact_folder', 'create_contacts_batch',
  'create_notebook', 'create_section', 'create_onenote_page',
  'send_chat_message', 'send_channel_message',
];

/**
 * The server's KEY_REFUSED_TOOLS. tools/list hides these from every agent key,
 * so seeing one means the server did not take the request as a key at all.
 */
const KEY_REFUSED_TOOLS = ['set_email_output_mode'];

/** What the server says when a write is refused on a read-only service. */
const READ_ONLY_REFUSAL = /read-only/;

/** What the server says when a tool's service is out of reach. */
const SERVICE_REFUSAL = /not permitted for this agent key|is not enabled for this tenant|is disabled for your account|Mail access is disabled/;

/** What the server says when the bearer credential resolved to no session. */
const NO_SESSION = /Session expired\. Please re-authenticate/;

function parseServices(value) {
  if (!value) return null;
  const out = value.split(',').map((s) => s.trim().toLowerCase()).filter(Boolean);
  const unknown = out.filter((s) => !SERVICES.includes(s));
  if (unknown.length > 0) throw new Error(`SMOKE_EXPECT_SERVICES names unknown service(s): ${unknown.join(', ')}`);
  return out;
}

function toolText(result) {
  return (result?.content ?? []).map((c) => (c && typeof c.text === 'string' ? c.text : '')).join('\n');
}

/** Short and single-line, for a log or an issue body. Never includes the key. */
function excerpt(text, max = 200) {
  const one = String(text ?? '').replace(/\s+/g, ' ').trim();
  return one.length > max ? `${one.slice(0, max)}...` : one;
}

/**
 * Run every check against baseUrl. Returns { checks, failed }, where each check
 * is { name, ok, detail } and ok is true, false, or 'skip'.
 */
async function runSmoke({ baseUrl, key, expectedSha, expectServices, fetchImpl = fetch, now = new Date() }) {
  const base = baseUrl.replace(/\/+$/, '');
  const origin = new URL(base).origin;
  const checks = [];
  const record = (name, ok, detail = '') => checks.push({ name, ok, detail });
  const auth = { authorization: `Bearer ${key}` };

  async function http(path, init = {}) {
    try {
      const res = await fetchImpl(`${base}${path}`, { redirect: 'manual', signal: AbortSignal.timeout(60000), ...init });
      const text = await res.text();
      let json = null;
      try { json = JSON.parse(text); } catch { /* not JSON */ }
      return { status: res.status, type: res.headers.get('content-type') ?? '', text, json };
    } catch (err) {
      return { status: 0, type: '', text: err instanceof Error ? err.message : String(err), json: null };
    }
  }

  let rpcId = 0;
  async function rpc(method, params) {
    const res = await http('/api/mcp', {
      method: 'POST',
      headers: { ...auth, 'content-type': 'application/json' },
      body: JSON.stringify({ jsonrpc: '2.0', id: ++rpcId, method, ...(params ? { params } : {}) }),
    });
    if (res.status !== 200 || !res.json) {
      return { transport: `HTTP ${res.status}: ${excerpt(res.text)}` };
    }
    if (res.json.error) return { transport: `JSON-RPC error ${res.json.error.code}: ${excerpt(res.json.error.message)}` };
    return { result: res.json.result };
  }

  const callTool = (name, args) => rpc('tools/call', { name, arguments: args });

  // 1. /health serves JSON naming the commit that was deployed.
  {
    const res = await http('/health');
    const sha = res.json?.sha;
    if (res.status !== 200 || !res.type.startsWith('application/json') || res.json?.status !== 'ok') {
      record('/health', false, `HTTP ${res.status}, ${res.type || 'no content type'}: ${excerpt(res.text)}`);
    } else if (expectedSha && sha !== expectedSha) {
      record('/health', false, `serving ${sha || 'no sha'}, expected ${expectedSha}`);
    } else {
      record('/health', true, `serving ${sha || 'no sha'}`);
    }
  }

  // 2. The MCP handshake.
  {
    const r = await rpc('initialize', { protocolVersion: '2024-11-05', capabilities: {}, clientInfo: { name: 'post-deploy-smoke', version: '1' } });
    if (r.transport) record('initialize', false, r.transport);
    else if (!r.result?.protocolVersion) record('initialize', false, 'no protocolVersion in the result');
    else record('initialize', true, `protocol ${r.result.protocolVersion}`);
  }

  // 3. tools/list, which also says whether the key was recognised and whether
  //    it is read-only. Everything after this depends on both.
  const list = await rpc('tools/list');
  const names = new Set((list.result?.tools ?? []).map((t) => t.name));
  if (list.transport) {
    record('tools/list', false, list.transport);
    return finish();
  }
  if (names.size === 0) {
    record('tools/list', false, 'the server listed no tools');
    return finish();
  }
  const ownerTools = KEY_REFUSED_TOOLS.filter((t) => names.has(t));
  if (ownerTools.length > 0) {
    // tools/list answers an unauthenticated caller with the whole catalogue, so
    // this is what a revoked, expired or mistyped key looks like.
    record('key recognised', false,
      `tools/list includes ${ownerTools.join(', ')}, which no agent key is shown: the server did not take SMOKE_AGENT_KEY as a key. ` +
      'It may be revoked or expired, or its owner signed out; mint a new one and update the secret. ' +
      'The server also shows the whole catalogue when it cannot load tenant settings, so read the console log if the key is good.');
    return finish();
  }
  record('key recognised', true, `${names.size} tools listed`);

  const held = SERVICES.filter((s) => names.has(READ_PROBES[s].tool));
  const listedWrites = WRITE_TOOLS.filter((t) => names.has(t));
  record('tools/list', held.length > 0, held.length > 0 ? `services: ${held.join(', ')}` : 'no service read tool is listed');
  record('no write tool listed', listedWrites.length === 0,
    listedWrites.length === 0 ? 'read-only on every service it holds' : `listed: ${listedWrites.join(', ')}`);
  if (expectServices) {
    const missing = expectServices.filter((s) => !held.includes(s));
    const extra = held.filter((s) => !expectServices.includes(s));
    record('services as expected', missing.length === 0 && extra.length === 0,
      [missing.length ? `missing: ${missing.join(', ')}` : '', extra.length ? `unexpected: ${extra.join(', ')}` : '']
        .filter(Boolean).join('; ') || expectServices.join(', '));
  }

  // 4. One real read per service the key holds.
  for (const svc of held) {
    const { tool, args } = READ_PROBES[svc];
    const r = await callTool(tool, args(now));
    const text = toolText(r.result);
    if (r.transport) record(`read ${tool}`, false, r.transport);
    else if (r.result?.isError && NO_SESSION.test(text)) {
      // The key resolved but its owner's Graph session could not be renewed:
      // a credential to replace, not a code regression, but the probe is blind
      // until it is.
      record(`read ${tool}`, false, `the key's owner must sign in again, or mint a new key: ${excerpt(text)}`);
    } else if (r.result?.isError) record(`read ${tool}`, false, excerpt(text));
    else {
      let parsed = null;
      try { parsed = JSON.parse(text); } catch { /* reported below */ }
      record(`read ${tool}`, parsed !== null && typeof parsed === 'object', parsed !== null ? 'ok' : `not JSON: ${excerpt(text)}`);
    }
  }

  // 5. A write on each held service is refused as read-only. Skipped outright
  //    if tools/list already showed the key can write: the probe is about the
  //    gate, and a key that is not read-only is the wrong key for this job.
  for (const svc of held) {
    const { tool, args } = WRITE_PROBES[svc];
    if (listedWrites.length > 0) {
      record(`write ${tool} refused`, 'skip', 'not attempted: the key is not read-only');
      continue;
    }
    const r = await callTool(tool, args);
    const text = toolText(r.result);
    if (r.transport) record(`write ${tool} refused`, false, r.transport);
    else if (r.result?.isError && READ_ONLY_REFUSAL.test(text)) record(`write ${tool} refused`, true, excerpt(text, 100));
    else record(`write ${tool} refused`, false, `not refused as read-only: ${excerpt(text)}`);
  }

  // 6. A service outside the key is refused at tools/call, not only hidden.
  const outside = SERVICES.filter((s) => !held.includes(s));
  if (outside.length === 0) {
    record('service outside the key refused', 'skip', 'the key holds every service');
  }
  for (const svc of outside) {
    const { tool, args } = READ_PROBES[svc];
    const r = await callTool(tool, args(now));
    const text = toolText(r.result);
    if (r.transport) record(`${svc} outside the key refused`, false, r.transport);
    else if (r.result?.isError && SERVICE_REFUSAL.test(text)) record(`${svc} outside the key refused`, true, excerpt(text, 100));
    else record(`${svc} outside the key refused`, false, `${tool} was not refused: ${excerpt(text)}`);
  }

  // 7. A key may not change its owner's email output mode. The body only ever
  //    tightens, so even a broken gate cannot loosen anything.
  {
    const res = await http('/api/mail/settings', {
      method: 'POST',
      headers: { ...auth, 'content-type': 'application/json' },
      body: JSON.stringify({ emailOutputMode: 'draft' }),
    });
    const ok = res.status === 403 && /agent key/.test(res.json?.error ?? '');
    record('POST /api/mail/settings refused', ok, `HTTP ${res.status}: ${excerpt(res.text, 120)}`);
  }

  // 8. The admin API takes a browser console session only.
  {
    const res = await http('/api/manage/agent-keys', { headers: auth });
    record('admin API refused', res.status === 401, `GET /api/manage/agent-keys: HTTP ${res.status}`);
  }

  // 9. A key may not sign its owner out. Sent as the instance's own page would
  //    send it, so the browser-origin check passes and the key check is what
  //    answers; a 403 here is that refusal. Last, because a broken gate ends
  //    the key's session.
  {
    const res = await http('/api/auth/logout', {
      method: 'POST',
      headers: { ...auth, origin, 'sec-fetch-site': 'same-origin' },
    });
    record('logout refused', res.status === 403, `POST /api/auth/logout: HTTP ${res.status}`);
  }

  return finish();

  function finish() {
    return { checks, failed: checks.filter((c) => c.ok === false).length };
  }
}

function report({ checks, failed }, { baseUrl, expectedSha }) {
  const mark = (ok) => (ok === 'skip' ? 'skip' : ok ? 'pass' : '**FAIL**');
  const lines = [
    `Post-deploy smoke against ${baseUrl}${expectedSha ? ` at \`${expectedSha}\`` : ''}: ` +
      (failed === 0 ? 'every check passed.' : `${failed} check(s) failed.`),
    '',
    '| Check | Result | Detail |',
    '|---|---|---|',
    ...checks.map((c) => `| ${c.name} | ${mark(c.ok)} | ${String(c.detail).replace(/\|/g, '\\|')} |`),
  ];
  return lines.join('\n');
}

async function main(argv, env = process.env) {
  const [baseUrl, expectedSha] = argv;
  if (!baseUrl || !/^https?:\/\//.test(baseUrl)) {
    console.error('usage: SMOKE_AGENT_KEY=... post-deploy-smoke.js <base-url> [expected-sha]');
    return 2;
  }
  const key = (env.SMOKE_AGENT_KEY ?? '').trim();
  if (!key) {
    console.log('::notice::SMOKE_AGENT_KEY is not set, so the functional smoke is skipped. See docs/customer-go-live.md item 10.5.');
    return 0;
  }
  let expectServices;
  try {
    expectServices = parseServices(env.SMOKE_EXPECT_SERVICES);
  } catch (err) {
    console.error(`::error::${err.message}`);
    return 2;
  }

  const result = await runSmoke({ baseUrl, key, expectedSha, expectServices });
  for (const c of result.checks) {
    const line = `${c.name}: ${c.detail}`;
    if (c.ok === false) console.log(`::error::${line}`);
    else console.log(`${c.ok === 'skip' ? 'skip' : 'ok  '} ${line}`);
  }
  const md = report(result, { baseUrl, expectedSha });
  if (env.SMOKE_REPORT_FILE) require('fs').writeFileSync(env.SMOKE_REPORT_FILE, `${md}\n`);
  if (env.GITHUB_STEP_SUMMARY) require('fs').appendFileSync(env.GITHUB_STEP_SUMMARY, `\n### Functional smoke\n\n${md}\n`);
  return result.failed === 0 ? 0 : 1;
}

if (require.main === module) {
  main(process.argv.slice(2)).then(
    (code) => process.exit(code),
    (err) => {
      console.error(err);
      process.exit(2);
    },
  );
}

module.exports = {
  SERVICES,
  READ_PROBES,
  WRITE_PROBES,
  WRITE_TOOLS,
  KEY_REFUSED_TOOLS,
  REFUSED_ID,
  parseServices,
  runSmoke,
  report,
  main,
};
