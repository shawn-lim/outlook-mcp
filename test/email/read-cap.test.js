jest.mock('../../auth', () => ({ ensureAuthenticated: jest.fn().mockResolvedValue('TOK') }));
jest.mock('../../utils/graph-api', () => ({ callGraphAPI: jest.fn() }));
const { callGraphAPI } = require('../../utils/graph-api');
const handleReadEmail = require('../../email/read');

const email = (text) => ({
  subject: 'S', receivedDateTime: '2026-11-26T00:00:00Z',
  from: { emailAddress: { name: 'A', address: 'a@x.com' } }, toRecipients: [],
  body: { contentType: 'text', content: text },
});

describe('read-email body cap', () => {
  beforeAll(() => jest.spyOn(console, 'error').mockImplementation(() => {}));
  it('truncates long bodies to the default cap with a note', async () => {
    callGraphAPI.mockResolvedValue(email('x'.repeat(50000)));
    const out = (await handleReadEmail({ id: '1' })).content[0].text;
    expect(out.length).toBeLessThan(8000);
    expect(out).toMatch(/Body truncated: showing 6000 of \d+ characters/);
  });
  it('honours a per-call maxChars', async () => {
    callGraphAPI.mockResolvedValue(email('y'.repeat(50000)));
    const out = (await handleReadEmail({ id: '1', maxChars: 200 })).content[0].text;
    expect(out).toMatch(/showing 200 of/);
  });
  it('leaves short bodies alone', async () => {
    callGraphAPI.mockResolvedValue(email('short body'));
    const out = (await handleReadEmail({ id: '1' })).content[0].text;
    expect(out).not.toMatch(/truncated/);
  });
});
