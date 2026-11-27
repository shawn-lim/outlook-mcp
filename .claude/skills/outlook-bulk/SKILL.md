---
name: outlook-bulk
description: Use for large Outlook jobs (scanning, summarizing, or filing hundreds+ emails or events across Shawn's Outlook accounts) — write a Node script against outlook-mcp/lib instead of many MCP tool calls.
---

# Outlook bulk work via code execution

## Pick the path
| Job | Use |
|---|---|
| A handful of items (read 3 emails, one event) | Direct tools: `list-emails`, `read-email`, ... |
| Inbox overview / top senders / cleanup planning | `mailbox-stats` (server-side scan, compact output) — works from any chat |
| Move / archive / mark-read everything from one or many senders or @domains | `bulk-email-action` with an array of senders (one pass) — dry run first, `dryRun:false` only after explicit approval; poll `bulk-job-status` for big runs |
| Anything else large, on the PC itself (Claude Code) | A script against `lib/` (below) |

## How
1. Repo: `C:\Users\shawn\repos\outlook-mcp` (Node). Accounts: `node -e "console.log(require('./lib').listAccounts())"`.
2. Write a script under `scripts/` (or a temp file) that uses:
   ```js
   const { createClient } = require('../lib');
   const mail = createClient({ account: 'personal' }); // read-only
   for await (const m of mail.paginate('me/mailFolders/inbox/messages',
       { $select: 'id,from,subject,receivedDateTime,isRead', $top: 100 }, { max: 5000 })) { /* aggregate */ }
   ```
   Always `$select` only needed fields. Use `getAll()` only for small sets.
3. Print a **compact JSON summary** to stdout (counts, top-N, IDs only when needed next). Never print full bodies or raw pages.
4. Run it, read the summary, report to Shawn.

## Token budget
- Script output: aim for under ~2k tokens (top-N lists, counts). Never dump raw pages, bodies or full ID lists.
- Direct tools: `list-emails`/`search-emails` max 50 per call; `read-email` bodies are capped at 6000 chars (`maxChars` to raise). Read bodies only when needed.
- Over ~50 items, or anything that needs aggregation, write a script.

## Writes (move, delete, send, PATCH, rules)
- `post/patch/delete` throw unless `createClient({ account, allowWrites: true })`.
- Before enabling writes: run a read-only dry run, show Shawn exactly what will change (count, sample, destination), and get an explicit yes in chat. Sending mail and permanent deletes always need that confirmation.
- Prefer moving to a folder over deleting.

## Errors
- "not signed in" → tell Shawn to run `node outlook-auth-server.js --account <name>` and open http://localhost:3333/auth.
- Graph 429 → slow down (sequential pages, smaller $top) and retry.
