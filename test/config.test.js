const path = require('path');
const { resolveTokenPath } = require('../config');

describe('resolveTokenPath', () => {
  const home = '/h';
  it('defaults to the legacy single-account file', () => {
    expect(resolveTokenPath({}, home)).toBe(path.join(home, '.outlook-mcp-tokens.json'));
  });
  it('uses one file per OUTLOOK_ACCOUNT', () => {
    expect(resolveTokenPath({ OUTLOOK_ACCOUNT: 'personal' }, home)).toBe(path.join(home, '.outlook-mcp', 'personal.json'));
  });
  it('sanitizes account names so they cannot escape the token dir', () => {
    expect(resolveTokenPath({ OUTLOOK_ACCOUNT: '../evil' }, home)).toBe(path.join(home, '.outlook-mcp', '.._evil.json'));
  });
  it('OUTLOOK_TOKEN_PATH wins and expands ~', () => {
    expect(resolveTokenPath({ OUTLOOK_TOKEN_PATH: '~/x/t.json', OUTLOOK_ACCOUNT: 'a' }, home)).toBe(path.resolve(home, 'x/t.json'));
  });
});
