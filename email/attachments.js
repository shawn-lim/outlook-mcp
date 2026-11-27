/**
 * Attachment tools: list a message's attachments and save one to disk on this PC.
 *
 * download-attachment writes ONLY under allow-listed folders (OUTLOOK_ATTACHMENT_ROOTS,
 * default: ~/.outlook-mcp/attachments and ~/iCloudDrive), never overwrites unless asked,
 * and rejects path traversal — so a scheduled task can file statements into
 * e.g. ~/iCloudDrive/Legal/Taxes/2026/Salt Spring without broader disk access.
 */
const fs = require('fs');
const os = require('os');
const path = require('path');
const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');

const text = (t) => ({ content: [{ type: 'text', text: t }] });
const home = () => process.env.USERPROFILE || process.env.HOME || os.homedir();

function allowedRoots() {
  const raw = process.env.OUTLOOK_ATTACHMENT_ROOTS;
  const list = raw
    ? raw.split(path.delimiter === ';' ? /;/ : /[;:]/).filter(Boolean)
    : [path.join(home(), '.outlook-mcp', 'attachments'), path.join(home(), 'iCloudDrive')];
  return list.map((p) => path.resolve(p.startsWith('~') ? path.join(home(), p.slice(1)) : p));
}

/** True if target is inside one of the roots (case-insensitive on Windows). */
function isInside(target, roots) {
  const norm = (p) => (process.platform === 'win32' ? p.toLowerCase() : p);
  const t = norm(path.resolve(target));
  return roots.some((r) => {
    const root = norm(r);
    return t === root || t.startsWith(root.endsWith(path.sep) ? root : root + path.sep);
  });
}

/** Strip characters Windows/macOS forbid in filenames and any directory parts. */
function safeFileName(name) {
  const base = path.basename(String(name || 'attachment')).replace(/[<>:"/\\|?*\x00-\x1f]/g, '_').trim();
  return base.replace(/^\.+$/, '_') || 'attachment';
}

async function listAttachments(token, messageId) {
  const res = await callGraphAPI(token, 'GET', `me/messages/${messageId}/attachments`, null, {
    $select: 'id,name,contentType,size,isInline',
  });
  return res.value || [];
}

async function handleListAttachments(args = {}) {
  if (!args.id) return text('id (message ID) is required.');
  try {
    const token = await ensureAuthenticated();
    const items = await listAttachments(token, args.id);
    if (!items.length) return text('No attachments on this message.');
    return text(items.map((a, i) =>
      `${i + 1}. ${a.name} — ${a.contentType || 'unknown'}, ${Math.round((a.size || 0) / 1024)} KB${a.isInline ? ' (inline)' : ''}\n   attachmentId: ${a.id}`
    ).join('\n'));
  } catch (e) {
    return text(e.message === 'Authentication required' ? "Authentication required. Use the 'authenticate' tool first." : `Error listing attachments: ${e.message}`);
  }
}

async function handleDownloadAttachment(args = {}) {
  const { id, attachmentId, name, saveDir, fileName, overwrite = false } = args;
  if (!id) return text('id (message ID) is required.');
  if (!saveDir) return text('saveDir is required (an absolute folder inside an allowed root).');

  const roots = allowedRoots();
  const dir = path.resolve(saveDir.startsWith('~') ? path.join(home(), saveDir.slice(1)) : saveDir);
  if (!isInside(dir, roots)) {
    return text(`Refusing to save outside allowed folders. saveDir must be inside: ${roots.join(' | ')} (set OUTLOOK_ATTACHMENT_ROOTS to change).`);
  }

  try {
    const token = await ensureAuthenticated();
    let target = attachmentId;
    if (!target) {
      const items = (await listAttachments(token, id)).filter((a) => !a.isInline);
      const pick = name ? items.filter((a) => a.name && a.name.toLowerCase().includes(String(name).toLowerCase())) : items;
      if (pick.length !== 1) {
        return text(pick.length
          ? `Multiple attachments match; pass attachmentId. Candidates:\n${pick.map((a) => `- ${a.name} (${a.id})`).join('\n')}`
          : 'No matching (non-inline) attachment found.');
      }
      target = pick[0].id;
    }

    const att = await callGraphAPI(token, 'GET', `me/messages/${id}/attachments/${target}`);
    if (!att || !att.contentBytes) {
      return text(`Attachment "${(att && att.name) || target}" has no downloadable content (it may be a link or an embedded item).`);
    }

    const outName = safeFileName(fileName || att.name);
    const outPath = path.join(dir, outName);
    if (!isInside(outPath, roots)) return text('Refusing: resolved file path escapes the allowed folders.');
    fs.mkdirSync(dir, { recursive: true });
    if (fs.existsSync(outPath) && !overwrite) {
      const existing = fs.statSync(outPath).size;
      const incoming = Buffer.byteLength(att.contentBytes, 'base64');
      return text(existing === incoming
        ? `Already saved (same size): ${outPath}`
        : `A different file already exists at ${outPath}. Pass overwrite:true or a different fileName.`);
    }
    const buf = Buffer.from(att.contentBytes, 'base64');
    fs.writeFileSync(outPath, buf);
    return text(`Saved ${outName} (${Math.round(buf.length / 1024)} KB) to ${outPath}`);
  } catch (e) {
    return text(e.message === 'Authentication required' ? "Authentication required. Use the 'authenticate' tool first." : `Error downloading attachment: ${e.message}`);
  }
}

const attachmentTools = [
  {
    name: 'list-attachments',
    description: 'List the attachments on an email (name, type, size, attachmentId).',
    inputSchema: {
      type: 'object',
      properties: { id: { type: 'string', description: 'Message ID' } },
      required: ['id'],
    },
    handler: handleListAttachments,
  },
  {
    name: 'download-attachment',
    description: "Save an email attachment to a folder on this PC (only inside allowed folders: ~/.outlook-mcp/attachments and ~/iCloudDrive by default). Never overwrites unless overwrite:true. Pick the attachment by attachmentId, or by a name substring when there's exactly one match.",
    inputSchema: {
      type: 'object',
      properties: {
        id: { type: 'string', description: 'Message ID' },
        attachmentId: { type: 'string', description: 'Attachment ID from list-attachments (optional if name matches exactly one)' },
        name: { type: 'string', description: 'Case-insensitive substring of the attachment file name' },
        saveDir: { type: 'string', description: 'Absolute destination folder, e.g. C:\\Users\\shawn\\iCloudDrive\\Legal\\Taxes\\2026\\Salt Spring (created if missing)' },
        fileName: { type: 'string', description: 'Optional file name to save as (defaults to the attachment name)' },
        overwrite: { type: 'boolean', description: 'Replace an existing different file (default false)' },
      },
      required: ['id', 'saveDir'],
    },
    handler: handleDownloadAttachment,
  },
];

module.exports = { attachmentTools, handleListAttachments, handleDownloadAttachment, isInside, safeFileName, allowedRoots };
