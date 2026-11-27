jest.mock('../../auth', () => ({ ensureAuthenticated: jest.fn().mockResolvedValue('TOK') }));
jest.mock('../../utils/graph-api', () => ({ callGraphAPI: jest.fn() }));
jest.mock('../../email/folder-utils', () => ({ getFolderIdByName: jest.fn().mockResolvedValue('ARCH') }));
const { callGraphAPI } = require('../../utils/graph-api');
const handleCreateRule = require('../../rules/create');

beforeAll(() => jest.spyOn(console, 'error').mockImplementation(() => {}));

it('creates a domain rule from senderContains', async () => {
  callGraphAPI.mockImplementation(async (t, method) => (method === 'GET' ? { value: [] } : { id: 'r1' }));
  await handleCreateRule({ name: 'Promos', senderContains: '@shop.com, news.x.ca', moveToFolder: 'Archive' });
  const post = callGraphAPI.mock.calls.find((c) => c[1] === 'POST');
  expect(post[3].conditions.senderContains).toEqual(['@shop.com', 'news.x.ca']);
  expect(post[3].actions.moveToFolder).toBe('ARCH');
});

it('accepts senderContains as the only condition', async () => {
  callGraphAPI.mockImplementation(async (t, method) => (method === 'GET' ? { value: [] } : { id: 'r2' }));
  const out = await handleCreateRule({ name: 'X', senderContains: '@a.com', markAsRead: true });
  expect(out.content[0].text).not.toMatch(/At least one condition/);
});
