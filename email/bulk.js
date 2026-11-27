/**
 * EA bulk tools that run inside the MCP server (which already holds the tokens),
 * so any Claude chat can do large mailbox jobs and only gets a compact summary back.
 *
 *  - mailbox-stats:     scan N messages, group by sender/domain, return top-N counts
 *  - bulk-email-action: move / archive / mark-read every message from a sender or
 *                       @domain. Dry run by default; executes only with dryRun:false.
 */
const config = require('../config');
const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');
const { WELL_KNOWN_FOLDERS, getFolderIdByName } = require('./folder-utils');

const PAGE = 500;            // Graph page size for scans
const BATCH = 20;            // Graph $batch limit
const BATCH_URL = `${config.GRAPH_API_ENDPOINT}$batch`;
const WELL_KNOWN_IDS = { inbox: 'inbox', archive: 'archive', deleted: 'deleteditems', junk: 'junkemail', sent: 'sentitems', drafts: 'drafts' };

const text = (t) => ({ content: [{ type: 'text', text: t }] });
const clamp = (v, def, max) => Math.min(Math.max(Number(v) || def, 1), max);

/** Resolve a source folder to a messages path; unlike resolveFolderPath, never silently falls back. */
async function sourcePath(token, folder) {
  const name = (folder || 'inbox').toLowerCase();
  if (WELL_KNOWN_FOLDERS[name]) return WELL_KNOWN_FOLDERS[name];
  const id = await getFolderIdByName(token, folder);
  if (!id) throw new Error(`Folder "${folder}" not found (top-level folders only).`);
  return `me/mailFolders/${id}/messages`;
}

async function destinationId(token, folder) {
  const name = (folder || '').toLowerCase();
  if (WELL_KNOWN_IDS[name]) return WELL_KNOWN_IDS[name];
  const id = await getFolderIdByName(token, folder);
  if (!id) throw new Error(`Destination folder "${folder}" not found (top-level folders only). Create it first with create-folder.`);
  return id;
}

const addrOf = (m) => ((m.from && m.from.emailAddress && m.from.emailAddress.address) || 'unknown').toLowerCase();
const domainOf = (a) => (a.includes('@') ? a.split('@').pop() : a);

/** Normalise "a@x.com" | "@x.com, @y.com" | ["@x.com", "b@y.com"] into a list of filters. */
function parseFrom(from) {
  const list = Array.isArray(from) ? from : String(from || '').split(',');
  return [...new Set(list.map((f) => String(f).trim().toLowerCase()).filter(Boolean))];
}

/** Does a message match a sender filter? "@domain.com" / "domain.com" match the domain, otherwise exact address. */
function matcher(from) {
  const f = String(from).trim().toLowerCase();
  if (f.startsWith('@')) return (a) => domainOf(a) === f.slice(1) || domainOf(a).endsWith(`.${f.slice(1)}`);
  if (!f.includes('@')) return (a) => domainOf(a) === f || domainOf(a).endsWith(`.${f}`);
  return (a) => a === f;
}

/** Iterate messages in a folder, newest first, selecting only the fields we need. */
async function* scan(token, path, select, max) {
  let url = path;
  let params = { $select: select, $top: PAGE, $orderby: 'receivedDateTime desc' };
  let n = 0;
  while (url) {
    const page = await callGraphAPI(token, 'GET', url, null, params);
    for (const m of page.value || []) {
      yield m;
      if (++n >= max) return;
    }
    url = page['@odata.nextLink'] || null;
    params = {};
  }
}

async function handleMailboxStats(args = {}) {
  try {
    const token = await ensureAuthenticated();
    const maxScan = clamp(args.maxScan, 5000, 100000);
    const top = clamp(args.top, 25, 100);
    const byDomain = args.groupBy === 'domain';
    const olderMs = args.olderThanDays ? Date.now() - Number(args.olderThanDays) * 864e5 : null;
    const path = await sourcePath(token, args.folder);

    const groups = new Map();
    let scanned = 0, unread = 0, counted = 0;
    for await (const m of scan(token, path, 'from,isRead,receivedDateTime', maxScan)) {
      scanned++;
      if (args.unreadOnly && m.isRead) continue;
      if (olderMs && new Date(m.receivedDateTime).getTime() > olderMs) continue;
      counted++;
      if (!m.isRead) unread++;
      const a = addrOf(m);
      const key = byDomain ? domainOf(a) : a;
      const g = groups.get(key) || { key, n: 0, unread: 0, newest: m.receivedDateTime };
      g.n++;
      if (!m.isRead) g.unread++;
      groups.set(key, g);
    }

    const rows = [...groups.values()].sort((x, y) => y.n - x.n).slice(0, top);
    const lines = rows.map((g, i) => `${String(i + 1).padStart(3)}. ${g.key} — ${g.n} (${g.unread} unread, latest ${String(g.newest).slice(0, 10)})`);
    return text(
      `Mailbox stats (${config.AUTH_CONFIG.account}, folder: ${args.folder || 'inbox'}): scanned ${scanned}${scanned >= maxScan ? ' (hit maxScan)' : ''}, ` +
      `matched ${counted}, ${unread} unread, ${groups.size} unique ${byDomain ? 'domains' : 'senders'}.\n` +
      `Top ${rows.length} by volume:\n${lines.join('\n')}`
    );
  } catch (e) {
    return text(e.message === 'Authentication required' ? "Authentication required. Use the 'authenticate' tool first." : `Error: ${e.message}`);
  }
}

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
// Graph already fans each $batch out to up to 4 parallel requests, and Outlook allows
// 4 concurrent requests per mailbox — so send one batch at a time to stay under the limit.
const CONCURRENCY = 1;
const jobs = new Map();  // in-memory; lives as long as the server process
let jobSeq = 0;

/** Send one $batch of up to 20 requests, retrying throttled (429/503) items with backoff. */
async function sendBatch(token, reqs, attempt = 0) {
  const chunk = reqs.map((r, j) => ({ id: String(j), ...r }));
  const res = await callGraphAPI(token, 'POST', BATCH_URL, { requests: chunk });
  let ok = 0, failed = 0, wait = 0;
  const retry = [];
  for (const r of res.responses || []) {
    if (r.status >= 200 && r.status < 300) ok++;
    else if ((r.status === 429 || r.status === 503) && attempt < 4) {
      retry.push(reqs[Number(r.id)]);
      wait = Math.max(wait, Number(r.headers && r.headers['Retry-After']) || 2 ** attempt * 2);
    } else failed++;
  }
  if (retry.length) {
    await sleep(wait * 1000);
    const again = await sendBatch(token, retry, attempt + 1);
    ok += again.ok; failed += again.failed;
  }
  return { ok, failed };
}

/** Run all requests in $batch chunks with limited concurrency; reports progress via onProgress. */
async function runBatches(token, requests, onProgress = () => {}) {
  const chunks = [];
  for (let i = 0; i < requests.length; i += BATCH) chunks.push(requests.slice(i, i + BATCH));
  let ok = 0, failed = 0, next = 0;
  async function worker() {
    while (next < chunks.length) {
      const c = chunks[next++];
      try {
        const r = await sendBatch(token, c);
        ok += r.ok; failed += r.failed;
      } catch (e) {
        failed += c.length;
      }
      onProgress(ok, failed);
    }
  }
  await Promise.all(Array.from({ length: Math.min(CONCURRENCY, chunks.length) }, worker));
  return { ok, failed };
}

const BACKGROUND_THRESHOLD = 200; // larger runs return a job id immediately

async function handleBulkEmailAction(args = {}) {
  const action = args.action;
  if (!['move', 'archive', 'mark-read'].includes(action)) return text("action must be 'move', 'archive' or 'mark-read'.");
  const filters = parseFrom(args.from);
  if (!filters.length) return text("from is required: sender address(es) and/or domain(s), e.g. [\"@x.com\", \"a@y.com\"].");
  if (action === 'move' && !args.destinationFolder) return text("destinationFolder is required for action 'move'.");
  const dryRun = args.dryRun !== false;
  const maxItems = clamp(args.maxItems, 50000, 100000);
  const maxScan = clamp(args.maxScan, 100000, 200000);

  try {
    const token = await ensureAuthenticated();
    const path = await sourcePath(token, args.folder);
    const dest = action === 'move' ? await destinationId(token, args.destinationFolder)
      : action === 'archive' ? 'archive' : null;
    const matchers = filters.map((f) => [f, matcher(f)]);

    // One pass over the folder for every sender.
    const hits = [];
    const perFilter = new Map(filters.map((f) => [f, { n: 0, sample: null }]));
    let scanned = 0;
    for await (const m of scan(token, path, 'id,from,subject,isRead,receivedDateTime', maxScan)) {
      scanned++;
      if (action === 'mark-read' && m.isRead) continue;
      const a = addrOf(m);
      const hit = matchers.find(([, fn]) => fn(a));
      if (!hit) continue;
      const pf = perFilter.get(hit[0]);
      pf.n++;
      if (!pf.sample) pf.sample = `${String(m.receivedDateTime).slice(0, 10)} ${m.subject || '(no subject)'}`;
      hits.push(m);
      if (hits.length >= maxItems) break;
    }

    const verb = action === 'mark-read' ? 'mark as read' : action === 'archive' ? 'archive' : `move to "${args.destinationFolder}"`;
    const where = `${args.folder || 'inbox'}${scanned >= maxScan ? ` (first ${scanned} scanned)` : ` (${scanned} scanned)`}`;
    if (!hits.length) return text(`No messages from ${filters.join(', ')} to ${verb} in ${where}.`);

    const breakdown = [...perFilter].map(([f, v]) => `  - ${f}: ${v.n}${v.sample ? ` (e.g. ${v.sample.slice(0, 70)})` : ''}`).join('\n');
    if (dryRun) {
      return text(
        `DRY RUN — nothing changed. Would ${verb} ${hits.length} message(s) in ${where}` +
        `${hits.length >= maxItems ? ` (capped at maxItems=${maxItems})` : ''}:\n${breakdown}\n` +
        `Get the user's explicit OK, then call again with dryRun:false.`
      );
    }

    const requests = hits.map((m) => action === 'mark-read'
      ? { method: 'PATCH', url: `/me/messages/${encodeURIComponent(m.id)}`, headers: { 'Content-Type': 'application/json' }, body: { isRead: true } }
      : { method: 'POST', url: `/me/messages/${encodeURIComponent(m.id)}/move`, headers: { 'Content-Type': 'application/json' }, body: { destinationId: dest } });

    if (requests.length <= BACKGROUND_THRESHOLD) {
      const { ok, failed } = await runBatches(token, requests);
      return text(`Done: ${verb} — ${ok} succeeded${failed ? `, ${failed} failed (re-run to retry)` : ''}.\n${breakdown}`);
    }

    const id = `job-${++jobSeq}`;
    const job = { id, verb, total: requests.length, ok: 0, failed: 0, state: 'running', started: Date.now(), breakdown };
    jobs.set(id, job);
    runBatches(token, requests, (ok, failed) => { job.ok = ok; job.failed = failed; })
      .then(({ ok, failed }) => Object.assign(job, { ok, failed, state: 'done', finished: Date.now() }))
      .catch((e) => Object.assign(job, { state: 'error', error: e.message, finished: Date.now() }));
    return text(`Started ${id}: ${verb} ${requests.length} message(s) in the background.\n${breakdown}\nCheck progress with bulk-job-status (id: ${id}).`);
  } catch (e) {
    return text(e.message === 'Authentication required' ? "Authentication required. Use the 'authenticate' tool first." : `Error: ${e.message}`);
  }
}

async function handleBulkJobStatus(args = {}) {
  const list = args.id ? [jobs.get(args.id)].filter(Boolean) : [...jobs.values()].slice(-5);
  if (!list.length) return text(args.id ? `No job ${args.id} (jobs reset when the server restarts).` : 'No bulk jobs yet.');
  return text(list.map((j) => {
    const secs = Math.round(((j.finished || Date.now()) - j.started) / 1000);
    return `${j.id}: ${j.state} — ${j.verb}: ${j.ok}/${j.total} done${j.failed ? `, ${j.failed} failed` : ''} (${secs}s)${j.error ? ` error: ${j.error}` : ''}`;
  }).join('\n'));
}

const bulkTools = [
  {
    name: 'mailbox-stats',
    description: 'Scan many messages server-side and return a compact summary: top senders or domains by volume with unread counts. Use this instead of listing emails for any inbox overview or cleanup planning (cheap in tokens).',
    inputSchema: {
      type: 'object',
      properties: {
        folder: { type: 'string', description: "Folder to scan (default 'inbox'; well-known or top-level folder name)" },
        maxScan: { type: 'number', description: 'Max messages to scan, newest first (default 5000, max 100000)' },
        groupBy: { type: 'string', enum: ['sender', 'domain'], description: "Group by sender address or domain (default 'sender')" },
        top: { type: 'number', description: 'How many groups to return (default 25, max 100)' },
        unreadOnly: { type: 'boolean', description: 'Count only unread messages' },
        olderThanDays: { type: 'number', description: 'Count only messages older than N days' }
      },
      required: []
    },
    handler: handleMailboxStats
  },
  {
    name: 'bulk-email-action',
    description: "Move, archive or mark-read every message from one or more senders (a@x.com) or domains (@x.com) in ONE pass over the folder. DRY RUN by default — returns per-sender counts + samples and changes nothing. Only call with dryRun:false after the user explicitly approves that exact action. Runs over 200 messages continue in the background; poll bulk-job-status.",
    inputSchema: {
      type: 'object',
      properties: {
        action: { type: 'string', enum: ['move', 'archive', 'mark-read'], description: 'What to do' },
        from: { type: 'array', items: { type: 'string' }, description: 'Sender addresses (a@x.com) and/or domains (@x.com): an array, or a comma-separated string' },
        folder: { type: 'string', description: "Source folder (default 'inbox')" },
        destinationFolder: { type: 'string', description: "For 'move': target folder name (well-known or top-level)" },
        maxItems: { type: 'number', description: 'Max messages to act on (default 50000, max 100000)' },
        maxScan: { type: 'number', description: 'Max messages to scan to find matches (default 100000, max 200000)' },
        dryRun: { type: 'boolean', description: 'Default true. Set false only after user approval.' }
      },
      required: ['action', 'from']
    },
    handler: handleBulkEmailAction
  },
  {
    name: 'bulk-job-status',
    description: 'Progress of background bulk-email-action jobs (one by id, or the last 5).',
    inputSchema: { type: 'object', properties: { id: { type: 'string', description: 'Job id, e.g. job-1' } }, required: [] },
    handler: handleBulkJobStatus
  }
];

module.exports = { bulkTools, handleMailboxStats, handleBulkEmailAction, handleBulkJobStatus, matcher, parseFrom };
