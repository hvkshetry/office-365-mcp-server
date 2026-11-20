// HOUSE-OFFICE-MCP-NO-AUTO-RETRY-ALL-MAIL-TRANSMITS-20261119, at the mail tool: sending a draft, reply and reply-all
// answered 503/504 (or not answered) are not sent again by office-mcp. The agent is told the send is unconfirmed and
// how to check before trying again; a reply that was being sent as a draft names that draft, so it is finished by
// sending the draft and never by a second reply. Graph is mocked: nothing is sent.
const fs = require('fs');
const os = require('os');
const path = require('path');
const { afterEach, describe, expect, it } = require('@jest/globals');

const { createGraphAPIError, NO_AUTO_RETRY_ALL_MAIL_TRANSMITS_RULING_ID } = jest.requireActual('../utils/graph-api');

const MAILBOX = 'engineering.agent@example.com';
const PREFIX = `users/${MAILBOX}`;

function loadMailWithMocks() {
  jest.resetModules();
  const callGraphAPI = jest.fn();
  const recordSideEffect = jest.fn().mockResolvedValue(undefined);
  jest.doMock('../auth', () => ({ ensureAuthenticated: jest.fn().mockResolvedValue('access-token') }));
  jest.doMock('../utils/graph-api', () => ({ callGraphAPI }));
  jest.doMock('../policy', () => ({
    enforceMailPolicy: jest.fn().mockResolvedValue(null),
    recipientGateRefusal: jest.fn().mockResolvedValue(null),
    getExecutionPolicy: jest.fn().mockReturnValue(null),
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

function unconfirmed(kind, statusCode) {
  // what utils/graph-api.js raises for that call answered 5xx, after no automatic retry
  return createGraphAPIError(statusCode, '{}', { error: { code: 'ServiceUnavailable', message: 'busy' } }, {},
    { unconfirmedSend: kind });
}

function routingRefusal() {
  // the one 503 that is a definite refusal: the reply action was not routed, nothing was sent
  const error = new Error('API call failed with status 503: Mailbox move in progress. Cross Server access is not allowed.');
  error.statusCode = 503;
  error.graphError = { statusCode: 503, code: 'ErrorMailboxMoveInProgress',
                       message: 'Mailbox move in progress. Cross Server access is not allowed.' };
  return error;
}

function posts(callGraphAPI, ending) {
  return callGraphAPI.mock.calls.filter(call => call[1] === 'POST' && String(call[2]).endsWith(ending));
}

function said(result) {
  return result.content[0].text;
}

describe('a draft send whose outcome is unknown', () => {
  it.each([503, 504])('send_draft answered %i: asked of Graph once, reported unconfirmed, audited as such', async status => {
    const { mail, callGraphAPI, recordSideEffect } = loadMailWithMocks();
    callGraphAPI.mockRejectedValueOnce(unconfirmed('sendDraft', status));
    const result = await mail({ operation: 'send_draft', draftId: 'draft-7', mailbox: MAILBOX });
    expect(callGraphAPI).toHaveBeenCalledTimes(1);
    expect(posts(callGraphAPI, '/messages/draft-7/send')).toHaveLength(1);
    expect(said(result)).toMatch(/^Draft send UNCONFIRMED: /);
    expect(said(result)).toContain(`the send of the draft was answered with HTTP ${status}`);
    expect(said(result)).toContain('folderName: "drafts"');
    expect(said(result)).toContain(NO_AUTO_RETRY_ALL_MAIL_TRANSMITS_RULING_ID);
    expect(recordSideEffect).toHaveBeenCalledTimes(1);
    expect(recordSideEffect).toHaveBeenCalledWith('mail.send_draft', 'failed', MAILBOX,
      expect.objectContaining({ send_unconfirmed: true, status_code: status, draftId: 'draft-7' }));
  });

  it('a draft Graph refuses outright stays a plain error', async () => {
    const { mail, callGraphAPI, recordSideEffect } = loadMailWithMocks();
    callGraphAPI.mockRejectedValueOnce(createGraphAPIError(404, '{}', { error: { code: 'ErrorItemNotFound', message: 'gone' } }));
    const result = await mail({ operation: 'send_draft', draftId: 'draft-7', mailbox: MAILBOX });
    expect(said(result)).toMatch(/^Error sending draft: /);
    expect(said(result)).not.toContain('UNCONFIRMED');
    expect(recordSideEffect.mock.calls[0][3]).not.toHaveProperty('send_unconfirmed');
  });
});

describe('a reply whose outcome is unknown', () => {
  it.each([['reply', false, 503], ['reply', false, 504], ['replyAll', true, 503], ['replyAll', true, 504]])(
    '%s answered %i: asked of Graph once, no draft fallback, reported unconfirmed', async (kind, replyAll, _unused) => {
      const status = _unused;
      const { mail, callGraphAPI, recordSideEffect } = loadMailWithMocks();
      callGraphAPI.mockRejectedValueOnce(unconfirmed(kind, status));
      const result = await mail({ operation: 'reply', emailId: 'message-1', body: '<p>Noted.</p>', mailbox: MAILBOX,
                                  replyAll });
      expect(callGraphAPI).toHaveBeenCalledTimes(1);                  // no second reply, and no draft written
      expect(posts(callGraphAPI, `/messages/message-1/${kind}`)).toHaveLength(1);
      expect(said(result)).toMatch(/^Email reply UNCONFIRMED: /);
      expect(said(result)).toContain(`was answered with HTTP ${status}`);
      expect(said(result)).toContain('the replying mailbox\'s Sent Items');
      expect(said(result)).not.toContain('was being sent as draft');
      expect(recordSideEffect).toHaveBeenCalledWith('mail.reply', 'failed', MAILBOX,
        expect.objectContaining({ send_unconfirmed: true, status_code: status, emailId: 'message-1' }));
    });

  it('a reply with an attachment names its draft: it is finished by sending that draft, not by replying again', async () => {
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'reply-unconfirmed-'));
    const file = path.join(dir, 'note.txt');
    fs.writeFileSync(file, 'attached');
    const { mail, callGraphAPI, recordSideEffect } = loadMailWithMocks();
    callGraphAPI
      .mockResolvedValueOnce({ id: 'draft-9' })                       // createReply
      .mockResolvedValueOnce({})                                      // the attachment
      .mockRejectedValueOnce(unconfirmed('sendDraft', 503));          // the draft's send
    const result = await mail({ operation: 'reply', emailId: 'message-1', body: '<p>See attached.</p>',
                                mailbox: MAILBOX, attachments: [file] });
    expect(callGraphAPI).toHaveBeenCalledTimes(3);
    expect(posts(callGraphAPI, '/createReply')).toHaveLength(1);      // one draft, never a second
    expect(posts(callGraphAPI, '/messages/draft-9/send')).toHaveLength(1);
    expect(said(result)).toMatch(/^Email reply UNCONFIRMED: /);
    expect(said(result)).toContain('The reply was being sent as draft draft-9.');
    expect(said(result)).toContain('mail { operation: "send_draft", draftId: "draft-9" }');
    expect(said(result)).toContain('do not reply again');
    expect(recordSideEffect).toHaveBeenCalledWith('mail.reply', 'failed', MAILBOX,
      expect.objectContaining({ send_unconfirmed: true, status_code: 503, draftId: 'draft-9' }));
  });

  it('the routing refusal still falls back to a draft, and that draft\'s unconfirmed send names it', async () => {
    const { mail, callGraphAPI, recordSideEffect } = loadMailWithMocks();
    callGraphAPI
      .mockRejectedValueOnce(routingRefusal())                        // the reply action: definitely not sent
      .mockResolvedValueOnce({ subject: 'Site visit', from: { emailAddress: { address: 'client@example.com' } },
                               toRecipients: [{ emailAddress: { address: MAILBOX } }], ccRecipients: [],
                               internetMessageId: '<m1@example.com>', conversationId: 'c-1' })
      .mockResolvedValueOnce({ id: 'draft-fallback-1' })              // the threaded draft
      .mockRejectedValueOnce(unconfirmed('sendDraft', 504));          // its send
    const result = await mail({ operation: 'reply', emailId: 'message-1', body: '<p>Noted.</p>', mailbox: MAILBOX });
    expect(posts(callGraphAPI, '/messages/message-1/reply')).toHaveLength(1);
    expect(posts(callGraphAPI, `${PREFIX}/messages`)).toHaveLength(1);
    expect(posts(callGraphAPI, '/messages/draft-fallback-1/send')).toHaveLength(1);
    expect(said(result)).toMatch(/^Email reply UNCONFIRMED: /);
    expect(said(result)).toContain('The reply was being sent as draft draft-fallback-1.');
    expect(recordSideEffect).toHaveBeenCalledWith('mail.reply', 'failed', MAILBOX,
      expect.objectContaining({ send_unconfirmed: true, status_code: 504, draftId: 'draft-fallback-1' }));
  });

  it('a reply Graph refuses outright stays a plain error', async () => {
    const { mail, callGraphAPI, recordSideEffect } = loadMailWithMocks();
    callGraphAPI.mockRejectedValueOnce(createGraphAPIError(403, '{}', { error: { code: 'ErrorAccessDenied', message: 'no' } }));
    const result = await mail({ operation: 'reply', emailId: 'message-1', body: '<p>Noted.</p>', mailbox: MAILBOX });
    expect(said(result)).toMatch(/^Email reply error: /);
    expect(recordSideEffect.mock.calls[0][3]).not.toHaveProperty('send_unconfirmed');
  });
});
