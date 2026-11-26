#!/usr/bin/env node
/**
 * Example code-execution script: scan up to N inbox messages and print a compact
 * summary (top senders, unread counts). Thousands of messages in, ~20 lines out.
 *
 *   node scripts/inbox-top-senders.js --account personal [--max 2000] [--top 20]
 */
const { createClient } = require('../lib');

function arg(name, fallback) {
  const i = process.argv.indexOf(`--${name}`);
  return i !== -1 && process.argv[i + 1] ? process.argv[i + 1] : fallback;
}

(async () => {
  const account = arg('account');
  const max = Number(arg('max', 2000));
  const top = Number(arg('top', 20));
  const mail = createClient({ account });

  const bySender = new Map();
  let total = 0;
  let unread = 0;
  for await (const m of mail.paginate('me/mailFolders/inbox/messages',
      { $select: 'from,isRead', $top: 100 }, { max })) {
    total++;
    if (!m.isRead) unread++;
    const addr = (m.from && m.from.emailAddress && m.from.emailAddress.address || 'unknown').toLowerCase();
    const s = bySender.get(addr) || { sender: addr, count: 0, unread: 0 };
    s.count++;
    if (!m.isRead) s.unread++;
    bySender.set(addr, s);
  }

  const topSenders = [...bySender.values()].sort((a, b) => b.count - a.count).slice(0, top);
  console.log(JSON.stringify({ account: mail.account, scanned: total, unread, uniqueSenders: bySender.size, topSenders }, null, 1));
})().catch((e) => {
  console.error(`ERROR: ${e.message}`);
  process.exit(1);
});
