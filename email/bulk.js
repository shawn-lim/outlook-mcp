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

/** Run Graph $batch requests 20 at a time; returns { ok, failed }. */
async function runBatches(token, requests) {
  let ok = 0, failed = 0;
  for (let i = 0; i < requests.length; i += BATCH) {
    const chunk = requests.slice(i, i + BATCH).map((r, j) => ({ id: String(j), ...r }));
    const res = await callGraphAPI(token, 'POST', BATCH_URL, { requests: chunk });
    for (const r of res.responses || []) (r.status >= 200 && r.status < 300 ? ok++ : failed++);
  }
  return { ok, failed };
}

async function handleBulkEmailAction(args = {}) {
  const action = args.action;
  if (!['move', 'archive', 'mark-read'].includes(action)) return text("action must be 'move', 'archive' or 'mark-read'.");
  if (!args.from) return text("from is required: a sender address (a@x.com) or a domain (@x.com).");
  if (action === 'move' && !args.destinationFolder) return text("destinationFolder is required for action 'move'.");
  const dryRun = args.dryRun !== false;
  const maxItems = clamp(args.maxItems, 1000, 20000);
  const maxScan = clamp(args.maxScan, 20000, 100000);

  try {
    const token = await ensureAuthenticated();
    const path = await sourcePath(token, args.folder);
    const dest = action === 'move' ? await destinationId(token, args.destinationFolder)
      : action === 'archive' ? 'archive' : null;
    const match = matcher(args.from);

    const hits = [];
    let scanned = 0;
    for await (const m of scan(token, path, 'id,from,subject,isRead,receivedDateTime', maxScan)) {
      scanned++;
      if (!match(addrOf(m))) continue;
      if (action === 'mark-read' && m.isRead) continue;
      hits.push(m);
      if (hits.length >= maxItems) break;
    }

    const verb = action === 'mark-read' ? 'mark as read' : action === 'archive' ? 'archive' : `move to "${args.destinationFolder}"`;
    const where = `${args.folder || 'inbox'}${scanned >= maxScan ? ` (first ${scanned} scanned)` : ''}`;
    if (!hits.length) return text(`No messages from ${args.from} to ${verb} in ${where}.`);

    if (dryRun) {
      const sample = hits.slice(0, 3).map((m) => `  - ${String(m.receivedDateTime).slice(0, 10)} ${m.subject || '(no subject)'}`).join('\n');
      return text(
        `DRY RUN — nothing changed. Would ${verb}: ${hits.length} message(s) from ${args.from} in ${where}` +
        `${hits.length >= maxItems ? ` (capped at maxItems=${maxItems})` : ''}.\nSample:\n${sample}\n` +
        `Get the user's explicit OK, then call again with dryRun:false.`
      );
    }

    const requests = hits.map((m) => action === 'mark-read'
      ? { method: 'PATCH', url: `/me/messages/${encodeURIComponent(m.id)}`, headers: { 'Content-Type': 'application/json' }, body: { isRead: true } }
      : { method: 'POST', url: `/me/messages/${encodeURIComponent(m.id)}/move`, headers: { 'Content-Type': 'application/json' }, body: { destinationId: dest } });
    const { ok, failed } = await runBatches(token, requests);
    return text(`Done: ${verb} — ${ok} succeeded${failed ? `, ${failed} failed (re-run to retry)` : ''} (from ${args.from}, ${where}).`);
  } catch (e) {
    return text(e.message === 'Authentication required' ? "Authentication required. Use the 'authenticate' tool first." : `Error: ${e.message}`);
  }
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
    description: "Move, archive or mark-read every message from a sender (a@x.com) or domain (@x.com) in a folder. DRY RUN by default — returns count + sample and changes nothing. Only call with dryRun:false after the user explicitly approves that exact action.",
    inputSchema: {
      type: 'object',
      properties: {
        action: { type: 'string', enum: ['move', 'archive', 'mark-read'], description: 'What to do' },
        from: { type: 'string', description: 'Sender address (a@x.com) or domain (@x.com)' },
        folder: { type: 'string', description: "Source folder (default 'inbox')" },
        destinationFolder: { type: 'string', description: "For 'move': target folder name (well-known or top-level)" },
        maxItems: { type: 'number', description: 'Max messages to act on (default 1000, max 20000)' },
        maxScan: { type: 'number', description: 'Max messages to scan to find matches (default 20000, max 100000)' },
        dryRun: { type: 'boolean', description: 'Default true. Set false only after user approval.' }
      },
      required: ['action', 'from']
    },
    handler: handleBulkEmailAction
  }
];

module.exports = { bulkTools, handleMailboxStats, handleBulkEmailAction, matcher };
