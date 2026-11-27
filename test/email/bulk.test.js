jest.mock('../../auth', () => ({ ensureAuthenticated: jest.fn().mockResolvedValue('TOK') }));
jest.mock('../../utils/graph-api', () => ({ callGraphAPI: jest.fn() }));
jest.mock('../../email/folder-utils', () => ({
  WELL_KNOWN_FOLDERS: { inbox: 'me/mailFolders/inbox/messages', archive: 'me/mailFolders/archive/messages' },
  getFolderIdByName: jest.fn(async (t, name) => (name === 'Promos' ? 'PROMO_ID' : null)),
}));
const { callGraphAPI } = require('../../utils/graph-api');
const { handleMailboxStats, handleBulkEmailAction, handleBulkJobStatus, matcher, parseFrom } = require('../../email/bulk');

const msg = (id, addr, isRead = false) => ({ id, from: { emailAddress: { address: addr } }, isRead, subject: `s${id}`, receivedDateTime: '2026-11-01T00:00:00Z' });
const pages = (...ps) => { ps.forEach((p, i) => callGraphAPI.mockResolvedValueOnce({ value: p, ...(i < ps.length - 1 ? { '@odata.nextLink': `next${i}` } : {}) })); };
const out = (r) => r.content[0].text;

beforeAll(() => jest.spyOn(console, 'error').mockImplementation(() => {}));
beforeEach(() => callGraphAPI.mockReset());

describe('matcher', () => {
  it('matches exact address, @domain and subdomains', () => {
    expect(matcher('a@x.com')('a@x.com')).toBe(true);
    expect(matcher('a@x.com')('b@x.com')).toBe(false);
    expect(matcher('@x.com')('news@mail.x.com')).toBe(true);
    expect(matcher('x.com')('b@x.com')).toBe(true);
    expect(matcher('@x.com')('a@notx.com')).toBe(false);
  });
});

describe('mailbox-stats', () => {
  it('aggregates across pages and returns a compact top list', async () => {
    pages([msg(1, 'a@x.com'), msg(2, 'a@x.com', true)], [msg(3, 'b@y.com')]);
    const t = out(await handleMailboxStats({ top: 5 }));
    expect(t).toMatch(/scanned 3, matched 3, 2 unread, 2 unique senders/);
    expect(t).toMatch(/1\. a@x\.com — 2 \(1 unread/);
    expect(callGraphAPI.mock.calls[1][2]).toBe('next0');
  });
  it('groups by domain', async () => {
    pages([msg(1, 'a@x.com'), msg(2, 'b@x.com'), msg(3, 'c@y.com')]);
    expect(out(await handleMailboxStats({ groupBy: 'domain' }))).toMatch(/1\. x\.com — 2/);
  });
});

describe('bulk-email-action', () => {
  it('is a dry run by default and changes nothing', async () => {
    pages([msg(1, 'n@shop.com'), msg(2, 'me@friend.com'), msg(3, 'n@shop.com')]);
    const t = out(await handleBulkEmailAction({ action: 'archive', from: '@shop.com' }));
    expect(t).toMatch(/DRY RUN.*Would archive 2 message/s);
    expect(t).toMatch(/@shop\.com: 2/);
    expect(callGraphAPI.mock.calls.every((c) => c[1] === 'GET')).toBe(true);
  });
  it('executes via $batch only with dryRun:false', async () => {
    pages([msg(1, 'n@shop.com'), msg(2, 'n@shop.com')]);
    callGraphAPI.mockResolvedValueOnce({ responses: [{ status: 201 }, { status: 201 }] });
    const t = out(await handleBulkEmailAction({ action: 'move', from: '@shop.com', destinationFolder: 'Promos', dryRun: false }));
    expect(t).toMatch(/2 succeeded/);
    const [, method, url, body] = callGraphAPI.mock.calls.at(-1);
    expect(method).toBe('POST');
    expect(url).toMatch(/\$batch$/);
    expect(body.requests[0]).toMatchObject({ method: 'POST', url: '/me/messages/1/move', body: { destinationId: 'PROMO_ID' } });
  });
  it('mark-read skips already-read messages', async () => {
    pages([msg(1, 'n@shop.com', true), msg(2, 'n@shop.com')]);
    expect(out(await handleBulkEmailAction({ action: 'mark-read', from: 'n@shop.com' }))).toMatch(/Would mark as read 1 message/);
  });
  it('refuses unknown destination folders instead of guessing', async () => {
    pages([msg(1, 'n@shop.com')]);
    expect(out(await handleBulkEmailAction({ action: 'move', from: '@shop.com', destinationFolder: 'Nope' }))).toMatch(/not found/);
  });
  it('validates inputs', async () => {
    expect(out(await handleBulkEmailAction({ action: 'delete', from: 'x' }))).toMatch(/action must be/);
    expect(out(await handleBulkEmailAction({ action: 'archive' }))).toMatch(/from is required/);
    expect(out(await handleBulkEmailAction({ action: 'archive', from: [] }))).toMatch(/from is required/);
  });
});

describe('multi-sender + background jobs', () => {
  it('parses arrays and comma strings, dedupes', () => {
    expect(parseFrom(['@A.com', ' b@y.com ', '@a.com'])).toEqual(['@a.com', 'b@y.com']);
    expect(parseFrom('@a.com, @b.com')).toEqual(['@a.com', '@b.com']);
  });

  it('handles many senders in ONE scan with per-sender counts', async () => {
    pages([msg(1, 'n@shop.com'), msg(2, 'x@ads.io'), msg(3, 'me@friend.com')], [msg(4, 'n@shop.com')]);
    const t = out(await handleBulkEmailAction({ action: 'archive', from: ['@shop.com', '@ads.io'] }));
    expect(t).toMatch(/Would archive 3 message/);
    expect(t).toMatch(/@shop\.com: 2/);
    expect(t).toMatch(/@ads\.io: 1/);
    expect(callGraphAPI).toHaveBeenCalledTimes(2); // two pages, one pass
  });

  it('retries throttled batch items', async () => {
    pages([msg(1, 'n@shop.com'), msg(2, 'n@shop.com')]);
    callGraphAPI
      .mockResolvedValueOnce({ responses: [{ id: '0', status: 201 }, { id: '1', status: 429, headers: { 'Retry-After': '0' } }] })
      .mockResolvedValueOnce({ responses: [{ id: '0', status: 201 }] });
    const t = out(await handleBulkEmailAction({ action: 'archive', from: '@shop.com', dryRun: false }));
    expect(t).toMatch(/2 succeeded/);
    expect(callGraphAPI.mock.calls.at(-1)[3].requests).toHaveLength(1);
  });

  it('runs large jobs in the background and reports progress', async () => {
    const many = Array.from({ length: 250 }, (_, i) => msg(i, 'n@shop.com'));
    pages(many);
    callGraphAPI.mockImplementation(async (t, method, url, body) => ({ responses: body.requests.map((r) => ({ id: r.id, status: 201 })) }));
    const t = out(await handleBulkEmailAction({ action: 'archive', from: '@shop.com', dryRun: false }));
    const id = t.match(/Started (job-\d+)/)[1];
    await new Promise((r) => setTimeout(r, 50));
    expect(out(await handleBulkJobStatus({ id }))).toMatch(new RegExp(`${id}: done .*250/250`));
  });
});
