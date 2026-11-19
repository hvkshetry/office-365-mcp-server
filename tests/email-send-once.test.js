// HOUSE-OFFICE-MCP-SEND-ONCE-20261119 (for HOUSE-TWIN-EMAILS-EXACTLY-ONCE-20261119): a send whose dispatch carries
// policy.send_once goes out exactly once. Before sending, the sent ledger and then the mailbox's Sent Items (the key
// stamped on every keyed send as a filterable extended property) are asked; either holding the key answers "already
// sent". A stamped send Graph refuses with a definite 400 is sent once more unstamped; a ledger that cannot be written
// never fails a send Graph accepted. Graph is mocked: nothing is sent.
const fs = require('fs');
const os = require('os');
const path = require('path');
const { afterEach, describe, expect, it } = require('@jest/globals');

const MAILBOX = 'digital.twin.agent@circleh2o.com';

function loadMailWithMocks(policy) {
  jest.resetModules();
  const callGraphAPI = jest.fn();
  const recordSideEffect = jest.fn().mockResolvedValue(undefined);
  jest.doMock('../auth', () => ({ ensureAuthenticated: jest.fn().mockResolvedValue('access-token') }));
  jest.doMock('../utils/graph-api', () => ({ callGraphAPI }));
  jest.doMock('../policy', () => ({
    enforceMailPolicy: jest.fn().mockResolvedValue(null),
    recipientGateRefusal: jest.fn().mockResolvedValue(null),
    getExecutionPolicy: jest.fn().mockReturnValue(policy),
    recordSideEffect
  }));
  const { emailTools } = require('../email');
  return { mail: emailTools[0].handler, callGraphAPI, recordSideEffect };
}

afterEach(() => {
  jest.dontMock('../auth');
  jest.dontMock('../utils/graph-api');
  jest.dontMock('../policy');
  jest.resetModules();
});

function sendArgs() {
  return { operation: 'send', mailbox: MAILBOX, to: ['owner@example.com'], subject: 'Meerut liquid train 2026-11-20',
           body: '<p>report</p>' };
}

function onceFor(dir) {
  return { send_once: { job_id: 'liquid_train_nightly', key: '2026-11-20',
                        ledger_path: path.join(dir, 'liquid_train_nightly', '2026-11-20.sent.json') } };
}

function text(result) {
  return JSON.stringify(result);
}

function graphError(statusCode, code) {
  // What utils/graph-api.js createGraphAPIError raises for an HTTP error answer.
  const error = new Error(`API call failed with status ${statusCode}: ${code}`);
  error.statusCode = statusCode;
  error.graphError = { statusCode, code, message: code };
  return error;
}

function sends(callGraphAPI) {
  return callGraphAPI.mock.calls.filter(call => call[1] === 'POST' && String(call[2]).endsWith('/sendMail'));
}

describe('mail send exactly once per send key', () => {
  it('stamps the key, sends once and records the ledger; a second send of the key goes nowhere', async () => {
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'send-once-'));
    const policy = onceFor(dir);
    let mocks = loadMailWithMocks(policy);
    mocks.callGraphAPI.mockResolvedValueOnce({ value: [] }).mockResolvedValueOnce({});
    await mocks.mail(sendArgs());
    const [get] = mocks.callGraphAPI.mock.calls;
    expect(get[1]).toBe('GET');
    expect(get[2]).toBe(`users/${MAILBOX}/mailFolders/sentitems/messages`);
    expect(get[4].$filter).toContain("ep/value eq 'liquid_train_nightly/2026-11-20'");
    const [sent] = sends(mocks.callGraphAPI);
    const message = sent[3].message;
    expect(message.singleValueExtendedProperties).toEqual([
      { id: expect.stringContaining('Name PuranosSendKey'), value: 'liquid_train_nightly/2026-11-20' }]);
    expect(message.internetMessageHeaders).toEqual([
      { name: 'X-Puranos-Send-Key', value: 'liquid_train_nightly/2026-11-20' }]);
    const ledger = JSON.parse(fs.readFileSync(policy.send_once.ledger_path, 'utf8'));
    expect(ledger).toMatchObject({ key: 'liquid_train_nightly/2026-11-20', witness: 'send' });
    expect(mocks.recordSideEffect).toHaveBeenCalledWith('mail.send', 'success', MAILBOX,
      expect.objectContaining({ send_key: 'liquid_train_nightly/2026-11-20' }));

    mocks = loadMailWithMocks(policy);                       // the resumed dispatch, a new office-mcp process
    const again = await mocks.mail(sendArgs());
    expect(mocks.callGraphAPI).not.toHaveBeenCalled();       // the ledger answers: no Graph call at all
    expect(text(again)).toContain('already sent for liquid_train_nightly/2026-11-20 (ledger)');
    expect(mocks.recordSideEffect).toHaveBeenCalledWith('mail.send', 'success', MAILBOX,
      expect.objectContaining({ already_sent: true, witness: 'ledger' }));
  });

  it('finds a send Graph accepted before the ledger was written (the process killed between) in Sent Items', async () => {
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'send-once-'));
    const policy = onceFor(dir);
    const mocks = loadMailWithMocks(policy);
    mocks.callGraphAPI.mockResolvedValueOnce({ value: [{ id: 'm-1', subject: 'Meerut liquid train 2026-11-20',
                                                         sentDateTime: '2026-11-20T07:03:00Z' }] });
    const result = await mocks.mail(sendArgs());
    expect(sends(mocks.callGraphAPI)).toEqual([]);
    expect(text(result)).toContain('(sent_items); not sent again');
    expect(JSON.parse(fs.readFileSync(policy.send_once.ledger_path, 'utf8')))
      .toMatchObject({ message_id: 'm-1', witness: 'sent_items' });
  });

  it('a Sent Items query that fails does not block the send (a missing daily email is the worse failure)', async () => {
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'send-once-'));
    const mocks = loadMailWithMocks(onceFor(dir));
    mocks.callGraphAPI.mockRejectedValueOnce(new Error('403')).mockResolvedValueOnce({});
    await mocks.mail(sendArgs());
    expect(sends(mocks.callGraphAPI)).toHaveLength(1);
  });

  it('a stamped send Graph refuses with a definite 400 is sent once more, unstamped, and says so', async () => {
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'send-once-'));
    const policy = onceFor(dir);
    const mocks = loadMailWithMocks(policy);
    mocks.callGraphAPI
      .mockResolvedValueOnce({ value: [] })                                   // Sent Items: not there
      .mockRejectedValueOnce(graphError(400, 'ErrorInvalidProperty'))         // the stamped send: refused outright
      .mockResolvedValueOnce({});                                             // the unstamped send: accepted
    const result = await mocks.mail(sendArgs());
    const posted = sends(mocks.callGraphAPI);
    expect(posted).toHaveLength(2);
    expect(posted[0][3].message.singleValueExtendedProperties).toHaveLength(1);
    expect(posted[1][3].message.singleValueExtendedProperties).toBeUndefined();
    expect(posted[1][3].message.internetMessageHeaders).toBeUndefined();
    expect(posted[1][3].message.subject).toBe(posted[0][3].message.subject);
    expect(text(result)).toContain('Email sent successfully');
    expect(JSON.parse(fs.readFileSync(policy.send_once.ledger_path, 'utf8')))
      .toMatchObject({ key: 'liquid_train_nightly/2026-11-20', witness: 'send', stamp_rejected: true });
    expect(mocks.recordSideEffect).toHaveBeenCalledTimes(1);
    expect(mocks.recordSideEffect).toHaveBeenCalledWith('mail.send', 'success', MAILBOX,
      expect.objectContaining({ send_key: 'liquid_train_nightly/2026-11-20', stamp_rejected: true,
                                stamp_rejection: expect.stringContaining('status 400') }));
  });

  it.each([
    ['a 503', () => graphError(503, 'ServiceUnavailable')],
    ['a 500', () => graphError(500, 'InternalServerError')],
    ['a 429', () => graphError(429, 'TooManyRequests')],
    ['a 403', () => graphError(403, 'ErrorAccessDenied')],
    ['a timeout (no status)', () => new Error('Network error during API call: socket hang up')]
  ])('never retries a stamped send that failed with %s: the first send may have gone out', async (_name, makeError) => {
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'send-once-'));
    const policy = onceFor(dir);
    const mocks = loadMailWithMocks(policy);
    mocks.callGraphAPI.mockResolvedValueOnce({ value: [] }).mockRejectedValueOnce(makeError());
    const result = await mocks.mail(sendArgs());
    expect(sends(mocks.callGraphAPI)).toHaveLength(1);
    expect(text(result)).toContain('Email send error');
    expect(fs.existsSync(policy.send_once.ledger_path)).toBe(false);
    expect(mocks.recordSideEffect).toHaveBeenCalledWith('mail.send', 'failed', MAILBOX, expect.anything());
  });

  it('an unstamped retry that also fails is a failed send, tried no third time', async () => {
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'send-once-'));
    const policy = onceFor(dir);
    const mocks = loadMailWithMocks(policy);
    mocks.callGraphAPI.mockResolvedValueOnce({ value: [] })
      .mockRejectedValueOnce(graphError(400, 'ErrorInvalidRecipients'))
      .mockRejectedValueOnce(graphError(400, 'ErrorInvalidRecipients'));
    const result = await mocks.mail(sendArgs());
    expect(sends(mocks.callGraphAPI)).toHaveLength(2);
    expect(text(result)).toContain('Email send error');
    expect(fs.existsSync(policy.send_once.ledger_path)).toBe(false);
  });

  it('a ledger that cannot be written after Graph accepted the send never fails the call', async () => {
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'send-once-'));
    const blocker = path.join(dir, 'not-a-directory');
    fs.writeFileSync(blocker, '');
    const policy = { send_once: { job_id: 'liquid_train_nightly', key: '2026-11-20',
                                  ledger_path: path.join(blocker, 'liquid_train_nightly', '2026-11-20.sent.json') } };
    const mocks = loadMailWithMocks(policy);
    mocks.callGraphAPI.mockResolvedValueOnce({ value: [] }).mockResolvedValueOnce({});
    const result = await mocks.mail(sendArgs());
    expect(sends(mocks.callGraphAPI)).toHaveLength(1);                    // sent once, and not sent again
    expect(text(result)).toContain('Email sent successfully');
    expect(mocks.recordSideEffect).toHaveBeenCalledTimes(1);
    expect(mocks.recordSideEffect).toHaveBeenCalledWith('mail.send', 'success', MAILBOX,
      expect.objectContaining({ send_key: 'liquid_train_nightly/2026-11-20', ledger_write_failed: true,
                                ledger_error: expect.any(String) }));
  });

  it('a Sent Items hit stays "already sent" when its ledger copy cannot be written', async () => {
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'send-once-'));
    const blocker = path.join(dir, 'not-a-directory');
    fs.writeFileSync(blocker, '');
    const policy = { send_once: { job_id: 'liquid_train_nightly', key: '2026-11-20',
                                  ledger_path: path.join(blocker, 'x', '2026-11-20.sent.json') } };
    const mocks = loadMailWithMocks(policy);
    mocks.callGraphAPI.mockResolvedValueOnce({ value: [{ id: 'm-1', subject: 's', sentDateTime: '2026-11-20T07:03:00Z' }] });
    const result = await mocks.mail(sendArgs());
    expect(sends(mocks.callGraphAPI)).toEqual([]);                        // never a second email for want of a ledger
    expect(text(result)).toContain('(sent_items); not sent again');
    expect(mocks.recordSideEffect).toHaveBeenCalledWith('mail.send', 'success', MAILBOX,
      expect.objectContaining({ already_sent: true, witness: 'sent_items', ledger_write_failed: true }));
  });

  it('a send without a send key is sent exactly as before, with no Sent Items query and no stamp', async () => {
    const mocks = loadMailWithMocks({ mail: { send: true } });
    mocks.callGraphAPI.mockResolvedValueOnce({});
    await mocks.mail(sendArgs());
    expect(mocks.callGraphAPI.mock.calls).toHaveLength(1);
    const [sent] = sends(mocks.callGraphAPI);
    expect(sent[3].message.singleValueExtendedProperties).toBeUndefined();
    expect(sent[3].message.internetMessageHeaders).toBeUndefined();
  });
});
