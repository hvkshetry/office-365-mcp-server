const fs = require('fs');
const os = require('os');
const path = require('path');

jest.mock('../auth', () => ({ ensureAuthenticated: jest.fn().mockResolvedValue('fixture-token') }));
jest.mock('../utils/graph-api', () => ({ callGraphAPI: jest.fn() }));
jest.mock('../policy', () => ({ enforceMailPolicy: jest.fn(), recordSideEffect: jest.fn() }));
jest.mock('../email/attachments', () => ({ cleanupOldAttachments: jest.fn() }));

const { callGraphAPI } = require('../utils/graph-api');
const { enforceMailPolicy } = require('../policy');
const mail = require('../email').emailTools[0].handler;
let directory;
let attachment;

beforeEach(() => {
  jest.clearAllMocks();
  enforceMailPolicy.mockResolvedValue(null);
  callGraphAPI.mockResolvedValue({ id: 'draft-fixture' });
  directory = fs.mkdtempSync(path.join(os.tmpdir(), 'office-draft-test-'));
  attachment = path.join(directory, 'fixture.txt');
  fs.writeFileSync(attachment, 'Attachment fixture');
});
afterEach(() => fs.rmSync(directory, { recursive: true, force: true }));

it('adds an attachment to a draft without an empty PATCH or transmission', async () => {
  const result = await mail({ operation: 'update_draft', draftId: 'draft-fixture', attachments: [attachment] });
  expect(callGraphAPI).toHaveBeenCalledTimes(1);
  expect(callGraphAPI).toHaveBeenCalledWith('fixture-token', 'POST',
    'me/messages/draft-fixture/attachments', expect.objectContaining({
      name: 'fixture.txt', contentBytes: Buffer.from('Attachment fixture').toString('base64')
    }), null);
  expect(result.content[0].text).toContain('Attachments added: 1');
});

it('keeps recipient policy enforcement before any draft mutation', async () => {
  enforceMailPolicy.mockResolvedValue('Blocked by fixture policy');
  const result = await mail({ operation: 'update_draft', draftId: 'draft-fixture', attachments: [attachment] });
  expect(result.content[0].text).toContain('Blocked');
  expect(callGraphAPI).not.toHaveBeenCalled();
});

it('validates attachments before changing the draft subject', async () => {
  await mail({ operation: 'update_draft', draftId: 'draft-fixture', subject: 'Updated', attachments: [path.join(directory, 'missing.txt')] });
  expect(callGraphAPI).not.toHaveBeenCalled();
});

it('reads a draft with no sender, recipients, date or body', async () => {
  const result = await mail({ operation: 'read', emailId: 'draft-fixture' });
  expect(result.content[0].text).toContain('(Draft - not yet sent)');
  expect(result.content[0].text).toContain('(No recipients)');
  expect(result.content[0].text).toContain('(empty body)');
});
