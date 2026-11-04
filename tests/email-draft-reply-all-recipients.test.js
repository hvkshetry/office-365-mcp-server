// draft_reply_all applies the caller's cc/bcc to the draft it creates and refuses any parameter it does not act on.
// Found 2026-11-04: { operation: "draft_reply_all", emailId, cc: [...] } created the draft and silently dropped the cc
// (the handler posted only { comment } to createReplyAll); the cc had to be added afterwards with update_draft.
// Graph is mocked: nothing is created or sent.
const { afterEach, describe, expect, it } = require('@jest/globals');

const MAILBOX = 'engineering.agent@circleh2o.com';
const PREFIX = `users/${MAILBOX}`;

function routingError() {
  const message = 'Mailbox move in progress. Cross Server access is not allowed.';
  const error = new Error(`API call failed with status 503: ${message}`);
  error.statusCode = 503;
  error.graphError = { statusCode: 503, code: 'ErrorMailboxMoveInProgress', message };
  return error;
}

function replyAllDraft(overrides = {}) {
  // What Graph's createReplyAll returns: the draft message, its reply-all recipients already filled.
  return {
    id: 'draft-1',
    toRecipients: [{ emailAddress: { name: 'Sender', address: 'sender@example.com' } }],
    ccRecipients: [{ emailAddress: { name: 'Existing', address: 'existing.cc@example.com' } }],
    bccRecipients: [],
    ...overrides
  };
}

function loadMailWithMocks() {
  jest.resetModules();
  const callGraphAPI = jest.fn();
  const recordSideEffect = jest.fn().mockResolvedValue(undefined);
  jest.doMock('../auth', () => ({ ensureAuthenticated: jest.fn().mockResolvedValue('access-token') }));
  jest.doMock('../utils/graph-api', () => ({ callGraphAPI }));
  jest.doMock('../policy', () => ({ enforceMailPolicy: jest.fn().mockResolvedValue(null), recordSideEffect }));
  const { emailTools } = require('../email');
  return { mail: emailTools[0].handler, callGraphAPI, recordSideEffect };
}

afterEach(() => {
  jest.dontMock('../auth');
  jest.dontMock('../utils/graph-api');
  jest.dontMock('../policy');
  jest.resetModules();
});

function never(callGraphAPI, suffix) {
  return callGraphAPI.mock.calls.some(call => String(call[2]).endsWith(suffix));
}

describe('mail draft_reply_all recipients', () => {
  it('adds cc and bcc to the reply-all recipients of the created draft, never sending', async () => {
    const { mail, callGraphAPI, recordSideEffect } = loadMailWithMocks();
    callGraphAPI.mockResolvedValueOnce(replyAllDraft()).mockResolvedValueOnce({});

    const result = await mail({
      operation: 'draft_reply_all',
      emailId: 'message-1',
      body: '<p>Reply</p>',
      mailbox: MAILBOX,
      cc: ['added.cc@example.com', 'EXISTING.CC@example.com', 'sender@example.com'],
      bcc: ['audit@example.com']
    });

    expect(callGraphAPI).toHaveBeenNthCalledWith(1, 'access-token', 'POST',
      `${PREFIX}/messages/message-1/createReplyAll`, { comment: '<p>Reply</p>' }, null);
    expect(callGraphAPI).toHaveBeenNthCalledWith(2, 'access-token', 'PATCH', `${PREFIX}/messages/draft-1`, {
      // the reply-all cc kept, the new one added, duplicates of To/Cc (any case) not repeated
      ccRecipients: [
        { emailAddress: { name: 'Existing', address: 'existing.cc@example.com' } },
        { emailAddress: { address: 'added.cc@example.com' } }
      ],
      bccRecipients: [{ emailAddress: { address: 'audit@example.com' } }]
    }, null);
    expect(callGraphAPI).toHaveBeenCalledTimes(2);
    expect(never(callGraphAPI, '/send')).toBe(false);
    expect(result).toEqual(expect.objectContaining({
      draftId: 'draft-1',
      sent: false,
      cc: ['existing.cc@example.com', 'added.cc@example.com'],
      bcc: ['audit@example.com']
    }));
    expect(result.structuredContent.cc).toEqual(['existing.cc@example.com', 'added.cc@example.com']);
    expect(result.content[0].text).toContain('Cc: existing.cc@example.com, added.cc@example.com.');
    expect(recordSideEffect).toHaveBeenCalledWith('mail.draft_reply_all', 'success', MAILBOX,
      expect.objectContaining({ cc: 2, bcc: 1, sent: false }));
  });

  it('reads the draft recipients when createReplyAll returns none, then patches', async () => {
    const { mail, callGraphAPI } = loadMailWithMocks();
    callGraphAPI
      .mockResolvedValueOnce({ id: 'draft-2' })
      .mockResolvedValueOnce(replyAllDraft({ id: 'draft-2' }))
      .mockResolvedValueOnce({});

    await mail({ operation: 'draft_reply_all', emailId: 'message-2', body: '<p>x</p>', mailbox: MAILBOX,
      cc: ['added.cc@example.com'] });

    expect(callGraphAPI).toHaveBeenNthCalledWith(2, 'access-token', 'GET', `${PREFIX}/messages/draft-2`, null,
      { $select: 'toRecipients,ccRecipients,bccRecipients' });
    expect(callGraphAPI.mock.calls[2][1]).toBe('PATCH');
    expect(callGraphAPI.mock.calls[2][3]).toEqual({ ccRecipients: [
      { emailAddress: { name: 'Existing', address: 'existing.cc@example.com' } },
      { emailAddress: { address: 'added.cc@example.com' } }
    ] });
  });

  it('applies cc on the threaded-draft fallback too, on top of the derived reply-all cc', async () => {
    const { mail, callGraphAPI, recordSideEffect } = loadMailWithMocks();
    const original = {
      subject: 'Thread', sender: { emailAddress: { address: 'sender@example.com' } }, replyTo: [],
      toRecipients: [{ emailAddress: { address: MAILBOX } }],
      ccRecipients: [{ emailAddress: { address: 'existing.cc@example.com' } }],
      internetMessageId: '<m@example.com>', conversationId: 'conv-1'
    };
    callGraphAPI
      .mockRejectedValueOnce(routingError())                        // createReplyAll
      .mockResolvedValueOnce(original)                              // fallback reads the source message
      .mockResolvedValueOnce({ id: 'draft-3' })                     // fallback POST /messages (no recipients echoed)
      .mockResolvedValueOnce(replyAllDraft({ id: 'draft-3',         // GET the draft's recipients
        ccRecipients: [{ emailAddress: { address: 'existing.cc@example.com' } }] }))
      .mockResolvedValueOnce({});                                   // PATCH

    const result = await mail({ operation: 'draft_reply_all', emailId: 'message-3', body: '<p>x</p>',
      mailbox: MAILBOX, cc: ['added.cc@example.com'] });

    const fallbackPost = callGraphAPI.mock.calls[2];
    expect(fallbackPost[1]).toBe('POST');
    expect(fallbackPost[2]).toBe(`${PREFIX}/messages`);
    // the fallback derived the reply-all cc itself rather than replacing it with the caller's list
    expect(fallbackPost[3].ccRecipients).toEqual([{ emailAddress: { address: 'existing.cc@example.com' } }]);
    expect(callGraphAPI.mock.calls[4][1]).toBe('PATCH');
    expect(callGraphAPI.mock.calls[4][3].ccRecipients.map(r => r.emailAddress.address))
      .toEqual(['existing.cc@example.com', 'added.cc@example.com']);
    expect(result).toEqual(expect.objectContaining({ draftId: 'draft-3', fallback: 'threaded-draft', sent: false }));
    expect(never(callGraphAPI, '/send')).toBe(false);
    expect(recordSideEffect).toHaveBeenCalledWith('mail.draft_reply_all', 'success', MAILBOX,
      expect.objectContaining({ cc: 2, fallback: 'threaded-draft' }));
  });

  it('refuses a parameter it does not act on by name, before any Graph call', async () => {
    const { mail, callGraphAPI } = loadMailWithMocks();
    const result = await mail({ operation: 'draft_reply_all', emailId: 'message-4', body: '<p>x</p>',
      mailbox: MAILBOX, to: ['someone@example.com'], subject: 'New subject' });
    expect(callGraphAPI).not.toHaveBeenCalled();
    expect(result.content[0].text).toMatch(/draft_reply_all does not support: to, subject\. Nothing was created\./);
  });

  it('refuses replyAll=false and accepts the redundant replyAll=true', async () => {
    const { mail, callGraphAPI } = loadMailWithMocks();
    const refused = await mail({ operation: 'draft_reply_all', emailId: 'message-5', body: '<p>x</p>',
      mailbox: MAILBOX, replyAll: false });
    expect(refused.content[0].text).toContain('replyAll=false');
    expect(callGraphAPI).not.toHaveBeenCalled();

    callGraphAPI.mockResolvedValueOnce(replyAllDraft({ id: 'draft-5' }));
    const accepted = await mail({ operation: 'draft_reply_all', emailId: 'message-5', body: '<p>x</p>',
      mailbox: MAILBOX, replyAll: true, cc: [] });
    expect(accepted.draftId).toBe('draft-5');
    expect(callGraphAPI).toHaveBeenCalledTimes(1);          // no cc/bcc named: no PATCH
  });

  it('says so when the draft was created but cc could not be applied', async () => {
    const { mail, callGraphAPI, recordSideEffect } = loadMailWithMocks();
    callGraphAPI.mockResolvedValueOnce(replyAllDraft({ id: 'draft-6' })).mockRejectedValueOnce(new Error('403'));
    const result = await mail({ operation: 'draft_reply_all', emailId: 'message-6', body: '<p>x</p>',
      mailbox: MAILBOX, cc: ['added.cc@example.com'] });
    expect(result.draftId).toBe('draft-6');
    expect(result.content[0].text).toContain('adding cc/bcc FAILED');
    expect(recordSideEffect).toHaveBeenCalledWith('mail.draft_reply_all', 'failed', MAILBOX,
      expect.objectContaining({ draftId: 'draft-6' }));
  });
});
