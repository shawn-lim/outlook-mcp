const fs = require('fs');
const os = require('os');
const path = require('path');

const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'omcp-'));
const tokenFile = path.join(tmp, 'acct.json');
process.env.OUTLOOK_TOKEN_PATH = tokenFile;
process.env.OUTLOOK_CLIENT_ID = 'cid';
process.env.OUTLOOK_CLIENT_SECRET = 'secret';

const TokenStorage = require('../../auth/token-storage');
const { ensureAuthenticated } = require('../../auth');

const write = (t) => fs.writeFileSync(tokenFile, JSON.stringify(t));

describe('ensureAuthenticated', () => {
  beforeAll(() => jest.spyOn(console, 'error').mockImplementation(() => {}));
  afterEach(() => jest.restoreAllMocks());

  it('throws when no token file exists', async () => {
    await expect(ensureAuthenticated()).rejects.toThrow('Authentication required');
  });

  it('returns a still-valid access token', async () => {
    write({ access_token: 'A1', refresh_token: 'R1', expires_at: Date.now() + 3600e3 });
    await expect(ensureAuthenticated()).resolves.toBe('A1');
  });

  it('refreshes an expired token instead of failing (was: re-auth every hour)', async () => {
    write({ access_token: 'OLD', refresh_token: 'R1', expires_at: Date.now() - 1000 });
    fs.utimesSync(tokenFile, new Date(), new Date(Date.now() + 5000)); // force reload
    const spy = jest.spyOn(TokenStorage.prototype, 'refreshAccessToken').mockResolvedValue('NEW');
    await expect(ensureAuthenticated()).resolves.toBe('NEW');
    expect(spy).toHaveBeenCalledTimes(1);
  });

  it('picks up a fresh sign-in written to disk without a restart', async () => {
    write({ access_token: 'FRESH', refresh_token: 'R2', expires_at: Date.now() + 3600e3 });
    fs.utimesSync(tokenFile, new Date(), new Date(Date.now() + 10000));
    await expect(ensureAuthenticated()).resolves.toBe('FRESH');
  });
});
