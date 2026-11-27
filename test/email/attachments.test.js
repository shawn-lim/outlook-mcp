const fs = require('fs');
const os = require('os');
const path = require('path');

jest.mock('../../auth', () => ({ ensureAuthenticated: jest.fn().mockResolvedValue('TOK') }));
jest.mock('../../utils/graph-api', () => ({ callGraphAPI: jest.fn() }));
const { callGraphAPI } = require('../../utils/graph-api');

const root = fs.mkdtempSync(path.join(os.tmpdir(), 'att-'));
process.env.OUTLOOK_ATTACHMENT_ROOTS = root;
const { handleListAttachments, handleDownloadAttachment, isInside, safeFileName } = require('../../email/attachments');

const pdf = Buffer.from('%PDF-1.4 test').toString('base64');
const out = (r) => r.content[0].text;
const listing = { value: [{ id: 'A1', name: 'statement.pdf', contentType: 'application/pdf', size: 2048, isInline: false }, { id: 'A2', name: 'logo.png', isInline: true }] };

beforeAll(() => jest.spyOn(console, 'error').mockImplementation(() => {}));
beforeEach(() => callGraphAPI.mockReset());

it('lists attachments with ids', async () => {
  callGraphAPI.mockResolvedValueOnce(listing);
  const t = out(await handleListAttachments({ id: 'M1' }));
  expect(t).toMatch(/statement\.pdf — application\/pdf, 2 KB/);
  expect(t).toMatch(/attachmentId: A1/);
});

it('saves a uniquely-named attachment into an allowed folder with a custom name', async () => {
  callGraphAPI.mockResolvedValueOnce(listing).mockResolvedValueOnce({ name: 'statement.pdf', contentBytes: pdf });
  const dir = path.join(root, 'Taxes', '2026', 'Salt Spring');
  const t = out(await handleDownloadAttachment({ id: 'M1', name: 'statement', saveDir: dir, fileName: 'Income Statement - 1 Test Ave - May 2026.pdf' }));
  expect(t).toMatch(/^Saved/);
  expect(fs.readFileSync(path.join(dir, 'Income Statement - 1 Test Ave - May 2026.pdf'), 'utf8')).toBe('%PDF-1.4 test');
});

it('does not overwrite; reports identical file as already saved', async () => {
  callGraphAPI.mockResolvedValue({ name: 'x.pdf', contentBytes: pdf });
  await handleDownloadAttachment({ id: 'M1', attachmentId: 'A1', saveDir: root, fileName: 'x.pdf' });
  expect(out(await handleDownloadAttachment({ id: 'M1', attachmentId: 'A1', saveDir: root, fileName: 'x.pdf' }))).toMatch(/Already saved/);
  callGraphAPI.mockResolvedValue({ name: 'x.pdf', contentBytes: Buffer.from('different!').toString('base64') });
  expect(out(await handleDownloadAttachment({ id: 'M1', attachmentId: 'A1', saveDir: root, fileName: 'x.pdf' }))).toMatch(/different file already exists/);
});

it('refuses folders outside the allowed roots and traversal in file names', async () => {
  expect(out(await handleDownloadAttachment({ id: 'M1', attachmentId: 'A1', saveDir: os.tmpdir() }))).toMatch(/Refusing to save outside/);
  expect(out(await handleDownloadAttachment({ id: 'M1', attachmentId: 'A1', saveDir: path.join(root, '..') }))).toMatch(/Refusing/);
  expect(callGraphAPI).not.toHaveBeenCalled();
  expect(safeFileName('../../evil.pdf')).toBe('evil.pdf');
  expect(safeFileName('a:b?.pdf')).toBe('a_b_.pdf');
  expect(isInside(path.join(root, 'sub'), [root])).toBe(true);
  expect(isInside(root + '-sibling', [root])).toBe(false);
});

it('asks for attachmentId when several match', async () => {
  callGraphAPI.mockResolvedValueOnce({ value: [{ id: 'A1', name: 'a.pdf' }, { id: 'A3', name: 'b.pdf' }] });
  expect(out(await handleDownloadAttachment({ id: 'M1', saveDir: root }))).toMatch(/Multiple attachments match/);
});
