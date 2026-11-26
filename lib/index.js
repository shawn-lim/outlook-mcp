/**
 * outlook-mcp client library — for "code execution with MCP" style scripts.
 *
 * Instead of streaming hundreds of tool results through the model, an agent
 * writes a short Node script against this library, runs it, and prints only a
 * compact summary. Same auth/token files as the MCP server, one per account.
 *
 *   const { createClient } = require('./lib');
 *   const mail = createClient({ account: 'personal' });            // read-only
 *   for await (const m of mail.paginate('me/mailFolders/inbox/messages',
 *       { $select: 'from,subject,receivedDateTime', $top: 100 }, { max: 2000 })) { ... }
 *
 * Writes (send, move, delete, PATCH, rules...) throw unless the client is created
 * with { allowWrites: true } — only do that after the user has confirmed the action.
 */
const path = require('path');
const fs = require('fs');
const config = require('../config');
const TokenStorage = require('../auth/token-storage');
const { callGraphAPI } = require('../utils/graph-api');

const WRITE_METHODS = new Set(['POST', 'PATCH', 'PUT', 'DELETE']);

function createClient({ account, tokenPath, allowWrites = false } = {}) {
  const resolvedPath = tokenPath
    ? config.resolveTokenPath({ OUTLOOK_TOKEN_PATH: tokenPath })
    : account
      ? config.resolveTokenPath({ OUTLOOK_ACCOUNT: account })
      : config.AUTH_CONFIG.tokenStorePath;
  const label = account || config.AUTH_CONFIG.account;
  const storage = new TokenStorage({ tokenStorePath: resolvedPath });

  async function token() {
    const t = await storage.getValidAccessToken();
    if (!t) {
      throw new Error(
        `Account "${label}" is not signed in (no valid token at ${resolvedPath}). ` +
        `Run: node outlook-auth-server.js --account ${label}`
      );
    }
    return t;
  }

  async function request(method, apiPath, body = null, query = {}) {
    const m = method.toUpperCase();
    if (WRITE_METHODS.has(m) && !allowWrites) {
      throw new Error(`Refusing ${m} ${apiPath}: client is read-only. Confirm with the user, then createClient({ ..., allowWrites: true }).`);
    }
    return callGraphAPI(await token(), m, apiPath, body, { ...query });
  }

  /** Async-iterate every item across @odata.nextLink pages (GET only). */
  async function* paginate(apiPath, query = {}, { max = 0 } = {}) {
    let next = apiPath;
    let params = { ...query };
    let count = 0;
    while (next) {
      const page = await callGraphAPI(await token(), 'GET', next, null, params);
      for (const item of page.value || []) {
        yield item;
        if (max && ++count >= max) return;
      }
      next = page['@odata.nextLink'] || null;
      params = {};
    }
  }

  async function getAll(apiPath, query = {}, opts = {}) {
    const out = [];
    for await (const item of paginate(apiPath, query, opts)) out.push(item);
    return out;
  }

  return {
    account: label,
    tokenPath: resolvedPath,
    get: (p, q) => request('GET', p, null, q),
    post: (p, b, q) => request('POST', p, b, q),
    patch: (p, b, q) => request('PATCH', p, b, q),
    delete: (p, q) => request('DELETE', p, null, q),
    request,
    paginate,
    getAll,
  };
}

/** Accounts that have a token file under ~/.outlook-mcp/ */
function listAccounts() {
  const dir = path.dirname(config.resolveTokenPath({ OUTLOOK_ACCOUNT: 'x' }));
  try {
    return fs.readdirSync(dir).filter((f) => f.endsWith('.json')).map((f) => f.slice(0, -5));
  } catch (_) {
    return [];
  }
}

module.exports = { createClient, listAccounts };
