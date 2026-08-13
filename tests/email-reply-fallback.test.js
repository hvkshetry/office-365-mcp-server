const fs = require('fs');
const os = require('os');
const path = require('path');
const { afterEach, describe, expect, it } = require('@jest/globals');

function routingError(message = 'Mailbox move in progress. Cross Server access is not allowed.') {
  const error = new Error(`API call failed with status 503: ${message}`);
  error.statusCode = 503;
  error.graphError = {
    statusCode: 503,
    code: 'ErrorMailboxMoveInProgress',
    message,
    innerError: {
      'request-id': 'request-123'
    },
    requestId: 'request-123'
  };
  return error;
}

function originalMessage(overrides = {}) {
  return {
    subject: 'Re: Existing subject',
    sender: {
      emailAddress: { name: 'Original Sender', address: 'sender@example.com' }
    },
    from: {
      emailAddress: { name: 'From User', address: 'from@example.com' }
    },
    replyTo: [{
      emailAddress: { name: 'Reply To', address: 'replyto@example.com' }
    }],
    toRecipients: [
      { emailAddress: { name: 'Agent', address: 'engineering.agent@circleh2o.com' } },
      { emailAddress: { name: 'Peer', address: 'peer@example.com' } }
    ],
    ccRecipients: [
      { emailAddress: { name: 'CC User', address: 'cc@example.com' } },
      { emailAddress: { name: 'Agent', address: 'Engineering.Agent@circleh2o.com' } }
    ],
    internetMessageId: '<original-message@example.com>',
    conversationId: 'conversation-1',
    ...overrides
  };
}

function loadMailWithMocks() {
  jest.resetModules();

  const ensureAuthenticated = jest.fn().mockResolvedValue('access-token');
  const callGraphAPI = jest.fn();
  const enforceMailPolicy = jest.fn().mockResolvedValue(null);
  const recordSideEffect = jest.fn().mockResolvedValue(undefined);

  jest.doMock('../auth', () => ({ ensureAuthenticated }));
  jest.doMock('../utils/graph-api', () => ({ callGraphAPI }));
  jest.doMock('../policy', () => ({
    enforceMailPolicy,
    recordSideEffect
  }));

  const { emailTools } = require('../email');
  return {
    mail: emailTools[0].handler,
    callGraphAPI,
    enforceMailPolicy,
    recordSideEffect,
    ensureAuthenticated
  };
}

afterEach(() => {
  jest.dontMock('../auth');
  jest.dontMock('../utils/graph-api');
  jest.dontMock('../policy');
  jest.resetModules();
});

describe('mail reply draft-send fallback', () => {
  it('falls back from replyAll action 503 to a threaded draft send', async () => {
    const mailbox = 'engineering.agent@circleh2o.com';
    const { mail, callGraphAPI, recordSideEffect } = loadMailWithMocks();

    callGraphAPI
      .mockRejectedValueOnce(routingError())
      .mockResolvedValueOnce(originalMessage())
      .mockResolvedValueOnce({ id: 'draft-1' })
      .mockResolvedValueOnce({});

    const result = await mail({
      operation: 'reply',
      emailId: 'message-1',
      body: '<p>Fallback body</p>',
      mailbox,
      replyAll: true
    });

    expect(result.fallback).toBe('draft-send');
    expect(result.content[0].text).toContain('fallback: draft-send');

    expect(callGraphAPI).toHaveBeenNthCalledWith(
      1,
      'access-token',
      'POST',
      `${'users/' + mailbox}/messages/message-1/replyAll`,
      expect.objectContaining({ comment: '<p>Fallback body</p>' }),
      null
    );
    expect(callGraphAPI).toHaveBeenNthCalledWith(
      2,
      'access-token',
      'GET',
      `${'users/' + mailbox}/messages/message-1`,
      null,
      {
        $select: 'subject,from,sender,replyTo,toRecipients,ccRecipients,internetMessageId,conversationId'
      }
    );

    const draftPayload = callGraphAPI.mock.calls[2][3];
    expect(draftPayload.subject).toBe('Re: Existing subject');
    expect(draftPayload.toRecipients.map(r => r.emailAddress.address)).toEqual([
      'sender@example.com',
      'replyto@example.com',
      'peer@example.com'
    ]);
    expect(draftPayload.ccRecipients.map(r => r.emailAddress.address)).toEqual(['cc@example.com']);
    expect(draftPayload.internetMessageHeaders).toEqual([
      { name: 'In-Reply-To', value: '<original-message@example.com>' },
      { name: 'References', value: '<original-message@example.com>' }
    ]);
    expect(draftPayload.body.content).toContain('<p>Fallback body</p>');

    expect(callGraphAPI).toHaveBeenNthCalledWith(
      4,
      'access-token',
      'POST',
      `${'users/' + mailbox}/messages/draft-1/send`,
      null,
      null
    );
    expect(recordSideEffect).toHaveBeenCalledWith(
      'mail.reply',
      'success',
      mailbox,
      expect.objectContaining({
        fallback: 'draft-send',
        draftId: 'draft-1',
        conversationId: 'conversation-1'
      })
    );
  });

  it('attaches files to the fallback draft before sending', async () => {
    const mailbox = 'engineering.agent@circleh2o.com';
    const tempFile = path.join(os.tmpdir(), `office-mcp-reply-${Date.now()}.txt`);
    fs.writeFileSync(tempFile, 'attachment content');

    try {
      const { mail, callGraphAPI } = loadMailWithMocks();
      callGraphAPI
        .mockRejectedValueOnce(routingError('Cross Server access is not allowed.'))
        .mockResolvedValueOnce(originalMessage({ subject: 'Fresh subject', replyTo: [] }))
        .mockResolvedValueOnce({ id: 'draft-2' })
        .mockResolvedValueOnce({})
        .mockResolvedValueOnce({});

      const result = await mail({
        operation: 'reply',
        emailId: 'message-2',
        body: '<p>With attachment</p>',
        mailbox,
        attachments: [tempFile]
      });

      expect(result.fallback).toBe('draft-send');

      const draftPayload = callGraphAPI.mock.calls[2][3];
      expect(draftPayload.subject).toBe('Re: Fresh subject');
      expect(draftPayload.toRecipients.map(r => r.emailAddress.address)).toEqual(['sender@example.com']);

      expect(callGraphAPI).toHaveBeenNthCalledWith(
        4,
        'access-token',
        'POST',
        `${'users/' + mailbox}/messages/draft-2/attachments`,
        expect.objectContaining({
          '@odata.type': '#microsoft.graph.fileAttachment',
          name: path.basename(tempFile),
          contentType: 'text/plain',
          contentBytes: Buffer.from('attachment content').toString('base64')
        }),
        null
      );
      expect(callGraphAPI).toHaveBeenNthCalledWith(
        5,
        'access-token',
        'POST',
        `${'users/' + mailbox}/messages/draft-2/send`,
        null,
        null
      );
    } finally {
      fs.unlinkSync(tempFile);
    }
  });
});
