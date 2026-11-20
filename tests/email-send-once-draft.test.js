// HOUSE-OFFICE-MCP-KEYED-SEND-COVERS-SENT-DRAFT-20261120: a keyed send covers an email written as a draft and then
// sent, with the contract of a keyed `mail send` (HOUSE-OFFICE-MCP-SEND-ONCE-20261119): the sent ledger, the Sent
// Items witness, the stamp, and the two fallbacks (a refused stamp never costs the email; a ledger that cannot be
// written never fails a sent one). The key attaches when the draft is written and again, by PATCH, when it is sent.
// The kill points of the keyed-send matrix are played as a second process: before the draft, after the draft, after
// Graph accepted the send (before the ledger), after the ledger (before the audit row), after the audit row.
// Graph is mocked: nothing is sent.
const fs = require('fs');
const os = require('os');
const path = require('path');
const { afterEach, describe, expect, it } = require('@jest/globals');

const { createGraphAPIError } = jest.requireActual('../utils/graph-api');
const sendOnce = jest.requireActual('../email/send-once');

const MAILBOX = 'maintenance.agent@example.com';
const PREFIX = `users/${MAILBOX}`;
const KEY = 'maintenance_ensaras_digest/2026-11-20';
const RULING = 'HOUSE-OFFICE-MCP-KEYED-SEND-COVERS-SENT-DRAFT-20261120';

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

function keyedPolicy() {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'send-once-draft-'));
  const ledger = path.join(dir, 'maintenance_ensaras_digest', '2026-11-20.sent.json');
  return { policy: { send_once: { job_id: 'maintenance_ensaras_digest', key: '2026-11-20', ledger_path: ledger } }, ledger };
}

function draftArgs() {
  return { operation: 'draft', mailbox: MAILBOX, to: ['promoters@example.com'], subject: 'Facility digest 2026-11-20',
           body: '<p>digest</p>' };
}

function sendDraftArgs(draftId = 'draft-1') {
  return { operation: 'send_draft', mailbox: MAILBOX, draftId };
}

function calls(callGraphAPI, method, ending) {
  return callGraphAPI.mock.calls.filter(call => call[1] === method && String(call[2]).endsWith(ending));
}

const sentItemsQueries = callGraphAPI => calls(callGraphAPI, 'GET', '/mailFolders/sentitems/messages');
const draftsWritten = callGraphAPI => calls(callGraphAPI, 'POST', `${PREFIX}/messages`);
const stamps = callGraphAPI => callGraphAPI.mock.calls.filter(call => call[1] === 'PATCH');
const sends = callGraphAPI => callGraphAPI.mock.calls.filter(call => call[1] === 'POST' && /\/messages\/[^/]+\/send$/.test(call[2]));
const said = result => result.content[0].text;
const refused = () => createGraphAPIError(400, '{}', { error: { code: 'ErrorInvalidProperty', message: 'bad property' } });
const unconfirmed = status => createGraphAPIError(status, '{}', { error: { code: 'ServiceUnavailable', message: 'busy' } },
  {}, { unconfirmedSend: 'sendDraft' });

describe('a keyed email written as a draft and then sent', () => {
  it('the key attaches when the draft is written and again when it is sent; it goes out once', async () => {
    const { policy, ledger } = keyedPolicy();
    let mocks = loadMailWithMocks(policy);
    mocks.callGraphAPI.mockResolvedValueOnce({ value: [] }).mockResolvedValueOnce({ id: 'draft-1', subject: 'Facility digest 2026-11-20' });
    const drafted = await mocks.mail(draftArgs());
    expect(sentItemsQueries(mocks.callGraphAPI)).toHaveLength(1);        // the witnesses are asked before writing
    const [written] = draftsWritten(mocks.callGraphAPI);
    expect(written[3].singleValueExtendedProperties).toEqual([{ id: sendOnce.SEND_KEY_PROPERTY, value: KEY }]);
    expect(written[3].internetMessageHeaders).toEqual([{ name: 'X-Puranos-Send-Key', value: KEY }]);
    expect(said(drafted)).toContain('Draft ID: draft-1');
    expect(said(drafted)).toContain(`Keyed send ${KEY}: send this draft with mail { operation: "send_draft", draftId: "draft-1" }`);
    expect(mocks.recordSideEffect).toHaveBeenCalledWith('mail.draft', 'success', MAILBOX,
      expect.objectContaining({ draftId: 'draft-1', send_key: KEY }));
    expect(fs.existsSync(ledger)).toBe(false);                           // a draft is not a send

    mocks = loadMailWithMocks(policy);
    mocks.callGraphAPI.mockResolvedValueOnce({ value: [] }).mockResolvedValueOnce({}).mockResolvedValueOnce({});
    const sent = await mocks.mail(sendDraftArgs());
    expect(mocks.callGraphAPI.mock.calls.map(call => call[1])).toEqual(['GET', 'PATCH', 'POST']);
    expect(stamps(mocks.callGraphAPI)[0].slice(2, 4)).toEqual([`${PREFIX}/messages/draft-1`, sendOnce.draftStampPatch(KEY)]);
    expect(sends(mocks.callGraphAPI)).toHaveLength(1);
    expect(said(sent)).toContain('Draft sent successfully');
    const record = JSON.parse(fs.readFileSync(ledger, 'utf8'));
    expect(record).toEqual(expect.objectContaining({ key: KEY, job_id: 'maintenance_ensaras_digest', witness: 'send',
                                                     via: 'send_draft', draft_id: 'draft-1', ruling: RULING }));
    expect(mocks.recordSideEffect).toHaveBeenCalledWith('mail.send_draft', 'success', MAILBOX,
      { draftId: 'draft-1', send_key: KEY });

    // any repeat, of this draft or another, goes nowhere: the ledger holds the key (killed after the ledger or after
    // the audit row, then retried)
    mocks = loadMailWithMocks(policy);
    const again = await mocks.mail(sendDraftArgs('draft-2'));
    expect(mocks.callGraphAPI).not.toHaveBeenCalled();
    expect(said(again)).toBe(`Email already sent for ${KEY} (ledger); the draft was not sent (${RULING}).`);
    expect(mocks.recordSideEffect).toHaveBeenCalledWith('mail.send_draft', 'success', MAILBOX,
      expect.objectContaining({ draftId: 'draft-2', send_key: KEY, already_sent: true, witness: 'ledger' }));
    mocks = loadMailWithMocks(policy);
    const redraft = await mocks.mail(draftArgs());
    expect(mocks.callGraphAPI).not.toHaveBeenCalled();
    expect(said(redraft)).toBe(`Email already sent for ${KEY} (ledger); no draft was written and there is nothing to send (${RULING}).`);
  });

  it('killed after the draft was written and before its send: the retry drafts and sends, one email', async () => {
    const { policy, ledger } = keyedPolicy();
    let mocks = loadMailWithMocks(policy);
    mocks.callGraphAPI.mockResolvedValueOnce({ value: [] }).mockResolvedValueOnce({ id: 'draft-1' });
    await mocks.mail(draftArgs());                                       // ...and the process is killed here
    mocks = loadMailWithMocks(policy);                                   // the retried dispatch
    mocks.callGraphAPI.mockResolvedValueOnce({ value: [] }).mockResolvedValueOnce({ id: 'draft-2' });
    expect(said(await mocks.mail(draftArgs()))).toContain('Draft ID: draft-2');   // a second draft, not a second email
    mocks.callGraphAPI.mockResolvedValueOnce({ value: [] }).mockResolvedValueOnce({}).mockResolvedValueOnce({});
    await mocks.mail(sendDraftArgs('draft-2'));
    expect(sends(mocks.callGraphAPI)).toHaveLength(1);
    expect(JSON.parse(fs.readFileSync(ledger, 'utf8')).draft_id).toBe('draft-2');
    // the earlier draft, sent by mistake afterwards, is refused
    mocks = loadMailWithMocks(policy);
    expect(said(await mocks.mail(sendDraftArgs('draft-1')))).toContain('already sent');
    expect(mocks.callGraphAPI).not.toHaveBeenCalled();
  });

  it('killed after Graph accepted the send and before the ledger: Sent Items answers, for a send and for a draft', async () => {
    const { policy, ledger } = keyedPolicy();
    const sentCopy = { value: [{ id: 'AAMk-sent', subject: 'Facility digest 2026-11-20', sentDateTime: '2026-11-20T18:14:02Z' }] };
    let mocks = loadMailWithMocks(policy);
    mocks.callGraphAPI.mockResolvedValueOnce(sentCopy);
    const again = await mocks.mail(sendDraftArgs('draft-2'));
    expect(mocks.callGraphAPI).toHaveBeenCalledTimes(1);                 // the Sent Items query; no stamp, no send
    expect(mocks.callGraphAPI.mock.calls[0][4].$filter).toContain(`ep/value eq '${KEY}'`);
    expect(said(again)).toBe(`Email already sent for ${KEY} (sent_items); the draft was not sent (${RULING}).`);
    expect(JSON.parse(fs.readFileSync(ledger, 'utf8'))).toEqual(expect.objectContaining({ key: KEY, witness: 'sent_items' }));
    fs.rmSync(ledger);
    mocks = loadMailWithMocks(policy);
    mocks.callGraphAPI.mockResolvedValueOnce(sentCopy);
    const redraft = await mocks.mail(draftArgs());
    expect(draftsWritten(mocks.callGraphAPI)).toHaveLength(0);
    expect(said(redraft)).toContain('no draft was written');
  });

  it('a stamp Graph refuses never costs the email, at the draft or at the send', async () => {
    const { policy, ledger } = keyedPolicy();
    let mocks = loadMailWithMocks(policy);
    mocks.callGraphAPI.mockResolvedValueOnce({ value: [] }).mockRejectedValueOnce(refused()).mockResolvedValueOnce({ id: 'draft-1' });
    const drafted = await mocks.mail(draftArgs());
    const written = draftsWritten(mocks.callGraphAPI);
    expect(written).toHaveLength(2);                                     // stamped, refused; then the same draft unstamped
    expect(written[1][3]).not.toHaveProperty('singleValueExtendedProperties');
    expect(said(drafted)).toContain('Draft ID: draft-1');
    expect(mocks.recordSideEffect).toHaveBeenCalledWith('mail.draft', 'success', MAILBOX,
      expect.objectContaining({ send_key: KEY, stamp_rejected: true }));

    for (const stampFailure of [refused(), unconfirmed(503), new Error('Network error during API call: timeout')]) {
      if (fs.existsSync(ledger)) fs.rmSync(ledger);
      mocks = loadMailWithMocks(policy);
      mocks.callGraphAPI.mockResolvedValueOnce({ value: [] }).mockRejectedValueOnce(stampFailure).mockResolvedValueOnce({});
      const sent = await mocks.mail(sendDraftArgs());
      expect(sends(mocks.callGraphAPI)).toHaveLength(1);                 // the stamp step sends nothing: still sent, once
      expect(said(sent)).toContain('Draft sent successfully');
      expect(JSON.parse(fs.readFileSync(ledger, 'utf8')).stamp_rejected).toBe(true);
      expect(mocks.recordSideEffect).toHaveBeenCalledWith('mail.send_draft', 'success', MAILBOX,
        expect.objectContaining({ send_key: KEY, stamp_rejected: true }));
    }
  });

  it('a ledger that cannot be written never fails a draft that was sent', async () => {
    const { policy } = keyedPolicy();
    const blocked = fs.mkdtempSync(path.join(os.tmpdir(), 'send-once-draft-blocked-'));
    fs.writeFileSync(path.join(blocked, 'not-a-directory'), 'x');
    policy.send_once.ledger_path = path.join(blocked, 'not-a-directory', '2026-11-20.sent.json');
    const mocks = loadMailWithMocks(policy);
    mocks.callGraphAPI.mockResolvedValueOnce({ value: [] }).mockResolvedValueOnce({}).mockResolvedValueOnce({});
    const sent = await mocks.mail(sendDraftArgs());
    expect(said(sent)).toContain('Draft sent successfully');
    expect(sends(mocks.callGraphAPI)).toHaveLength(1);
    expect(mocks.recordSideEffect).toHaveBeenCalledWith('mail.send_draft', 'success', MAILBOX,
      expect.objectContaining({ send_key: KEY, ledger_write_failed: true }));
  });

  it('a send Graph answers 503 or 504 is not repeated, writes no ledger, and its repeat asks Sent Items', async () => {
    for (const status of [503, 504]) {
      const { policy, ledger } = keyedPolicy();
      let mocks = loadMailWithMocks(policy);
      mocks.callGraphAPI.mockResolvedValueOnce({ value: [] }).mockResolvedValueOnce({}).mockRejectedValueOnce(unconfirmed(status));
      const first = await mocks.mail(sendDraftArgs());
      expect(sends(mocks.callGraphAPI)).toHaveLength(1);
      expect(said(first)).toMatch(/^Draft send UNCONFIRMED: /);
      expect(said(first)).toContain(`This send is keyed (${KEY}): a repeated send_draft asks Sent Items for that key first`);
      expect(fs.existsSync(ledger)).toBe(false);
      expect(mocks.recordSideEffect).toHaveBeenCalledWith('mail.send_draft', 'failed', MAILBOX,
        expect.objectContaining({ send_unconfirmed: true, status_code: status, draftId: 'draft-1' }));
      mocks = loadMailWithMocks(policy);                                 // Graph had accepted it: the repeat finds the key
      mocks.callGraphAPI.mockResolvedValueOnce({ value: [{ id: 'AAMk-sent', sentDateTime: '2026-11-20T18:14:02Z' }] });
      expect(said(await mocks.mail(sendDraftArgs()))).toContain('already sent');
      expect(sends(mocks.callGraphAPI)).toHaveLength(0);
    }
  });

  it('a Sent Items query that fails does not block the draft or its send', async () => {
    const { policy } = keyedPolicy();
    const mocks = loadMailWithMocks(policy);
    mocks.callGraphAPI.mockRejectedValueOnce(new Error('Sent Items unavailable')).mockResolvedValueOnce({ id: 'draft-1' });
    expect(said(await mocks.mail(draftArgs()))).toContain('Draft ID: draft-1');
    mocks.callGraphAPI.mockRejectedValueOnce(new Error('Sent Items unavailable')).mockResolvedValueOnce({}).mockResolvedValueOnce({});
    expect(said(await mocks.mail(sendDraftArgs()))).toContain('Draft sent successfully');
    expect(sends(mocks.callGraphAPI)).toHaveLength(1);
  });

  it('a draft and its send without a send key are exactly as before: no query, no stamp', async () => {
    const mocks = loadMailWithMocks(null);
    mocks.callGraphAPI.mockResolvedValueOnce({ id: 'draft-1' });
    const drafted = await mocks.mail(draftArgs());
    expect(mocks.callGraphAPI).toHaveBeenCalledTimes(1);
    expect(draftsWritten(mocks.callGraphAPI)[0][3]).not.toHaveProperty('singleValueExtendedProperties');
    expect(said(drafted)).not.toContain('Keyed send');
    expect(mocks.recordSideEffect.mock.calls[0][3]).not.toHaveProperty('send_key');
    mocks.callGraphAPI.mockResolvedValueOnce({});
    await mocks.mail(sendDraftArgs());
    expect(mocks.callGraphAPI).toHaveBeenCalledTimes(2);
    expect(stamps(mocks.callGraphAPI)).toHaveLength(0);
    expect(mocks.recordSideEffect).toHaveBeenLastCalledWith('mail.send_draft', 'success', MAILBOX, { draftId: 'draft-1' });
  });

  it('the capability is readable from the module', () => {
    expect(sendOnce.KEYED_DRAFT_RULING_ID).toBe(RULING);
    expect(sendOnce.draftStampPatch(KEY)).toEqual({ singleValueExtendedProperties: [{ id: sendOnce.SEND_KEY_PROPERTY, value: KEY }] });
    expect(fs.readFileSync(path.join(__dirname, '..', 'email', 'send-once.js'), 'utf8')).toContain('draftStampPatch');
  });
});
