jest.mock('../../utils/graph-api', () => ({ callGraphAPI: jest.fn() }));
const { callGraphAPI } = require('../../utils/graph-api');
const TokenStorage = require('../../auth/token-storage');
const { createClient } = require('../../lib');

describe('lib createClient', () => {
  beforeEach(() => {
    callGraphAPI.mockReset();
    jest.spyOn(TokenStorage.prototype, 'getValidAccessToken').mockResolvedValue('TOK');
  });
  afterEach(() => jest.restoreAllMocks());

  it('uses a per-account token file', () => {
    expect(createClient({ account: 'work2' }).tokenPath).toMatch(/\.outlook-mcp[\\/]work2\.json$/);
  });

  it('is read-only by default', async () => {
    const c = createClient({ account: 'a' });
    await expect(c.post('me/sendMail', {})).rejects.toThrow(/read-only/);
    await expect(c.delete('me/messages/1')).rejects.toThrow(/read-only/);
    expect(callGraphAPI).not.toHaveBeenCalled();
  });

  it('allows writes only when explicitly enabled', async () => {
    callGraphAPI.mockResolvedValue({});
    await createClient({ account: 'a', allowWrites: true }).post('me/sendMail', { m: 1 });
    expect(callGraphAPI).toHaveBeenCalledWith('TOK', 'POST', 'me/sendMail', { m: 1 }, {});
  });

  it('paginates across nextLink pages and honours max', async () => {
    callGraphAPI
      .mockResolvedValueOnce({ value: [1, 2], '@odata.nextLink': 'https://graph/next' })
      .mockResolvedValueOnce({ value: [3, 4] });
    const c = createClient({ account: 'a' });
    expect(await c.getAll('me/messages', { $top: 2 })).toEqual([1, 2, 3, 4]);
    expect(callGraphAPI).toHaveBeenNthCalledWith(2, 'TOK', 'GET', 'https://graph/next', null, {});
    callGraphAPI.mockResolvedValueOnce({ value: [1, 2], '@odata.nextLink': 'x' });
    expect(await c.getAll('me/messages', {}, { max: 1 })).toEqual([1]);
  });

  it('gives an actionable error when the account is not signed in', async () => {
    TokenStorage.prototype.getValidAccessToken.mockResolvedValue(null);
    await expect(createClient({ account: 'nope' }).get('me')).rejects.toThrow(/--account nope/);
  });
});
