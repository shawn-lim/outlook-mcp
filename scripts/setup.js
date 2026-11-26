#!/usr/bin/env node
/**
 * One-command setup per Outlook account.
 *
 *   npm run setup -- personal        sign in (if needed) + register in Claude Desktop & Claude Code
 *   npm run setup -- personal --reauth   force a fresh sign-in
 *   npm run status                    show every account's sign-in + registration state
 *
 * Needs OUTLOOK_CLIENT_ID / OUTLOOK_CLIENT_SECRET in .env (one Azure app serves all accounts).
 * Config files are backed up (*.bak) before they are changed.
 */
const fs = require('fs');
const os = require('os');
const path = require('path');
const { spawn } = require('child_process');

const REPO = path.resolve(__dirname, '..');
const INDEX = path.join(REPO, 'index.js');

/**
 * Claude Desktop config file(s). The Microsoft Store (MSIX) build of Claude on Windows
 * reads from its sandbox (%LOCALAPPDATA%\\Packages\\Claude_*\\LocalCache\\Roaming\\Claude),
 * not %APPDATA%\\Claude, so we target every location that exists.
 */
function desktopConfigPaths() {
  if (process.env.CLAUDE_DESKTOP_CONFIG) return [process.env.CLAUDE_DESKTOP_CONFIG];
  const home = os.homedir();
  const file = 'claude_desktop_config.json';
  if (process.platform === 'darwin') return [path.join(home, 'Library', 'Application Support', 'Claude', file)];
  if (process.platform !== 'win32') return [path.join(home, '.config', 'Claude', file)];

  const paths = [];
  const pkgs = path.join(process.env.LOCALAPPDATA || path.join(home, 'AppData', 'Local'), 'Packages');
  try {
    for (const d of fs.readdirSync(pkgs)) {
      const dir = path.join(pkgs, d, 'LocalCache', 'Roaming', 'Claude');
      if (/^Claude_/i.test(d) && fs.existsSync(dir)) paths.push(path.join(dir, file));
    }
  } catch (_) { /* no Packages dir */ }
  const classic = path.join(process.env.APPDATA || path.join(home, 'AppData', 'Roaming'), 'Claude', file);
  if (!paths.length || fs.existsSync(classic)) paths.push(classic);
  return paths;
}
const codeConfigPath = () => process.env.CLAUDE_CODE_CONFIG || path.join(os.homedir(), '.claude.json');

function readJson(file) {
  try { return JSON.parse(fs.readFileSync(file, 'utf8')); }
  catch (e) {
    if (e.code === 'ENOENT') return null;
    throw new Error(`${file} is not valid JSON (${e.message}); fix it or move it aside, then re-run.`);
  }
}

/** Add/refresh mcpServers[name] in a JSON config, keeping everything else. Returns 'added' | 'updated' | 'unchanged'. */
function upsertServer(file, name, entry, { createIfMissing }) {
  const cfg = readJson(file);
  if (!cfg && !createIfMissing) return 'skipped (not installed)';
  const next = cfg || {};
  next.mcpServers = next.mcpServers || {};
  const before = JSON.stringify(next.mcpServers[name]);
  if (before === JSON.stringify(entry)) return 'unchanged';
  if (cfg) fs.copyFileSync(file, `${file}.bak`);
  fs.mkdirSync(path.dirname(file), { recursive: true });
  next.mcpServers[name] = entry;
  fs.writeFileSync(file, JSON.stringify(next, null, 2));
  return before ? 'updated' : 'added';
}

function openBrowser(url) {
  const [cmd, args] = process.platform === 'win32' ? ['explorer.exe', [url]]
    : process.platform === 'darwin' ? ['open', [url]] : ['xdg-open', [url]];
  try {
    const c = spawn(cmd, args, { stdio: 'ignore', detached: true });
    c.on('error', () => { /* no browser launcher; printed URL is the fallback */ });
    c.unref();
  } catch (_) { /* printed URL is the fallback */ }
}

/** Run the auth server for one account, open the browser, wait until the token file is (re)written. */
function signIn(account, tokenPath, timeoutMs = 5 * 60 * 1000) {
  return new Promise((resolve, reject) => {
    const startMtime = fs.existsSync(tokenPath) ? fs.statSync(tokenPath).mtimeMs : 0;
    const child = spawn(process.execPath, [path.join(REPO, 'outlook-auth-server.js'), '--account', account],
      { cwd: REPO, env: { ...process.env, OUTLOOK_ACCOUNT: account }, stdio: ['ignore', 'pipe', 'pipe'] });
    let opened = false;
    let errText = '';
    const done = (err) => { clearInterval(poll); clearTimeout(timer); child.kill(); err ? reject(err) : resolve(); };
    child.stdout.on('data', (d) => {
      if (!opened && /running at/i.test(String(d))) {
        opened = true;
        console.log(`→ Opening browser. Sign in with the Microsoft account for "${account}".`);
        console.log('  (If it does not open: http://localhost:3333/auth)');
        openBrowser('http://localhost:3333/auth');
      }
    });
    child.stderr.on('data', (d) => { errText += d; });
    child.on('exit', (code) => {
      if (code && code !== 0 && !child.killed) {
        done(new Error(/EADDRINUSE/.test(errText) ? 'Port 3333 is busy (an old auth server?). Close it or run: npx kill-port 3333' : `Auth server exited: ${errText.trim().slice(-300)}`));
      }
    });
    const poll = setInterval(() => {
      try {
        const st = fs.statSync(tokenPath);
        if (st.mtimeMs > startMtime && JSON.parse(fs.readFileSync(tokenPath, 'utf8')).refresh_token) done();
      } catch (_) { /* not written yet */ }
    }, 1000);
    const timer = setTimeout(() => done(new Error('Timed out waiting for sign-in (5 min).')), timeoutMs);
  });
}

async function tokenState(account) {
  const { createClient } = require('../lib');
  const client = createClient({ account });
  const quiet = [console.error, console.warn];
  console.error = console.warn = () => {}; // TokenStorage chatter would clutter the output
  try {
    const me = await client.get('me', { $select: 'mail,userPrincipalName,displayName' });
    return { ok: true, who: me.mail || me.userPrincipalName || me.displayName };
  } catch (e) {
    return { ok: false, why: e.message.split('\n')[0].slice(0, 160) };
  } finally {
    [console.error, console.warn] = quiet;
  }
}

async function status() {
  const { listAccounts } = require('../lib');
  const desk = Object.assign({}, ...desktopConfigPaths().map((f) => (readJson(f) || {}).mcpServers || {}));
  const code = (readJson(codeConfigPath()) || {}).mcpServers || {};
  const accounts = new Set(listAccounts());
  for (const n of [...Object.keys(desk), ...Object.keys(code)]) if (n.startsWith('outlook-')) accounts.add(n.slice(8));
  if (!accounts.size) { console.log('No accounts yet. Run: npm run setup -- <account-name>'); return; }
  const rows = [];
  for (const a of [...accounts].sort()) {
    const t = await tokenState(a);
    rows.push({ account: a, signedIn: t.ok ? `yes (${t.who})` : `no: ${t.why}`, desktop: desk[`outlook-${a}`] ? 'yes' : 'no', claudeCode: code[`outlook-${a}`] ? 'yes' : 'no' });
  }
  console.table(rows);
}

async function setup(account, { reauth }) {
  if (!/^[A-Za-z0-9._-]+$/.test(account)) throw new Error('Account name: letters, numbers, . _ - only (e.g. personal, hotmail).');
  process.env.OUTLOOK_ACCOUNT = account;
  const config = require('../config');
  const { clientId, clientSecret } = config.AUTH_CONFIG;
  if (!clientId || !clientSecret) {
    throw new Error(`Missing credentials. Create ${path.join(REPO, '.env')} from .env.example with OUTLOOK_CLIENT_ID and OUTLOOK_CLIENT_SECRET.`);
  }
  const tokenPath = config.resolveTokenPath({ OUTLOOK_ACCOUNT: account });

  console.log(`\n[1/3] Sign-in: ${account}`);
  let t = reauth ? { ok: false } : await tokenState(account);
  if (t.ok) {
    console.log(`  already signed in as ${t.who} (use --reauth to switch)`);
  } else {
    await signIn(account, tokenPath);
    t = await tokenState(account);
    if (!t.ok) throw new Error(`Signed in but Graph check failed: ${t.why}`);
    console.log(`  ✓ signed in as ${t.who}`);
  }

  const name = `outlook-${account}`;
  const entry = { command: 'node', args: [INDEX], env: { OUTLOOK_ACCOUNT: account } };
  for (const f of desktopConfigPaths()) {
    console.log('[2/3] Claude Desktop:', upsertServer(f, name, entry, { createIfMissing: true }), `(${name}) -> ${f}`);
  }
  console.log('[3/3] Claude Code:   ', upsertServer(codeConfigPath(), name, { type: 'stdio', ...entry }, { createIfMissing: false }), `(${name})`);
  console.log('\nDone. Fully quit Claude Desktop (tray icon → Quit) and reopen it to load the new server.');
}

module.exports = { upsertServer, desktopConfigPaths, codeConfigPath };
if (require.main !== module) return;

(async () => {
  const args = process.argv.slice(2);
  if (args[0] === 'status' || args.length === 0) return status();
  return setup(args[0], { reauth: args.includes('--reauth') });
})().catch((e) => { console.error(`\n✗ ${e.message}`); process.exit(1); });
