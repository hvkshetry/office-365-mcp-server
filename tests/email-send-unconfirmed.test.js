// HOUSE-OFFICE-MCP-NO-SENDMAIL-AUTO-RETRY-20261119, at the mail tool: a sendMail answered 503/504 (or not answered) is
// not sent again by office-mcp. The agent is told the send is unconfirmed and how to check Sent Items before trying
// again, the audit row says so, and a keyed send that is repeated asks its witnesses first. Graph is mocked: nothing
// is sent.
const fs = require('fs');
const os = require('os');
const path = require('path');
const { afterEach, describe, expect, it } = require('@jest/globals');

const { createGraphAPIError, NO_SENDMAIL_AUTO_RETRY_RULING_ID } = jest.requireActual('../utils/graph-api');

const MAILBOX = 'digital.twin.agent@example.com';

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
  return { operation: 'send', mailbox: MAILBOX, to: ['owner@example.com'], subject: 'Facility digest 2026-11-20',
           body: '<p>digest</p>' };
}

function unconfirmed(statusCode) {
  // what utils/graph-api.js raises for a sendMail answered 5xx, after no automatic retry
  return createGraphAPIError(statusCode, '{}', { error: { code: 'ServiceUnavailable', message: 'busy' } }, {},
    { unconfirmedSend: true });
}

function sends(callGraphAPI) {
  return callGraphAPI.mock.calls.filter(call => call[1] === 'POST' && String(call[2]).endsWith('/sendMail'));
}

function said(result) {
  return result.content[0].text;
}

describe('a sendMail whose outcome is unknown', () => {
  it.each([503, 504])('answered %i: asked of Graph once, reported unconfirmed, audited as such', async status => {
    const { mail, callGraphAPI, recordSideEffect } = loadMailWithMocks(null);
    callGraphAPI.mockRejectedValueOnce(unconfirmed(status));
    const result = await mail(sendArgs());
    expect(sends(callGraphAPI)).toHaveLength(1);
    expect(said(result)).toMatch(/^Email send UNCONFIRMED: /);
    expect(said(result)).toContain(`sendMail was answered with HTTP ${status}`);
    expect(said(result)).toContain('Do NOT send it again yet');
    expect(said(result)).toContain('folderName: "sent"');
    expect(said(result)).toContain(NO_SENDMAIL_AUTO_RETRY_RULING_ID);
    expect(said(result)).not.toContain('This send is keyed');
    expect(recordSideEffect).toHaveBeenCalledTimes(1);
    expect(recordSideEffect).toHaveBeenCalledWith('mail.send', 'failed', MAILBOX,
      expect.objectContaining({ send_unconfirmed: true, status_code: status, subject: 'Facility digest 2026-11-20' }));
  });

  it('a send refused outright stays a plain send error', async () => {
    const { mail, callGraphAPI, recordSideEffect } = loadMailWithMocks(null);
    callGraphAPI.mockRejectedValueOnce(createGraphAPIError(403, '{}', { error: { code: 'ErrorAccessDenied', message: 'no' } }));
    const result = await mail(sendArgs());
    expect(said(result)).toMatch(/^Email send error: /);
    expect(said(result)).not.toContain('UNCONFIRMED');
    expect(recordSideEffect.mock.calls[0][3]).not.toHaveProperty('send_unconfirmed');
  });

  it('a keyed send answered 503 is not sent again unstamped, and its repeat asks Sent Items first', async () => {
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'send-unconfirmed-'));
    const ledger = path.join(dir, 'liquid_train_nightly', '2026-11-20.sent.json');
    const policy = { send_once: { job_id: 'liquid_train_nightly', key: '2026-11-20', ledger_path: ledger } };

    let mocks = loadMailWithMocks(policy);
    mocks.callGraphAPI.mockResolvedValueOnce({ value: [] }).mockRejectedValueOnce(unconfirmed(503));
    const first = await mocks.mail(sendArgs());
    expect(sends(mocks.callGraphAPI)).toHaveLength(1);      // a 503 is no refused stamp: no unstamped second send
    expect(said(first)).toMatch(/^Email send UNCONFIRMED: /);
    expect(said(first)).toContain('This send is keyed (liquid_train_nightly/2026-11-20)');
    expect(fs.existsSync(ledger)).toBe(false);              // nothing claims the email went out
    expect(mocks.recordSideEffect).toHaveBeenCalledWith('mail.send', 'failed', MAILBOX,
      expect.objectContaining({ send_unconfirmed: true, status_code: 503 }));

    // Graph had accepted it after all: the agent sends again, and office-mcp finds the key in Sent Items
    mocks = loadMailWithMocks(policy);
    mocks.callGraphAPI.mockResolvedValueOnce({ value: [{ id: 'AAMk-sent', sentDateTime: '2026-11-20T07:04:11Z' }] });
    const second = await mocks.mail(sendArgs());
    expect(sends(mocks.callGraphAPI)).toHaveLength(0);
    expect(said(second)).toContain('already sent');
    expect(JSON.parse(fs.readFileSync(ledger, 'utf8')).key).toBe('liquid_train_nightly/2026-11-20');
  });
});
