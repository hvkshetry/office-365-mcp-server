const { describe, it, expect, beforeEach } = require('@jest/globals');
const { EventEmitter } = require('events');
const https = require('https');
const {
  callGraphAPI, createGraphAPIError, getAnchorMailboxFromPath, isRetryableStatus, isSendMailCall, mailTransmitCall,
  NO_SENDMAIL_AUTO_RETRY_RULING_ID, NO_AUTO_RETRY_ALL_MAIL_TRANSMITS_RULING_ID
} = require('../utils/graph-api');

jest.mock('https');

function mockResponse(statusCode, headers, chunks = []) {
  return { statusCode, headers, chunks };
}

function installHttpsResponses(responses, requests) {
  https.request.mockImplementation((url, options, callback) => {
    requests.push({ url: url.toString(), options });
    const responseSpec = responses.shift();

    const req = new EventEmitter();
    req.write = jest.fn();
    req.end = () => {
      process.nextTick(() => {
        const res = new EventEmitter();
        res.statusCode = responseSpec.statusCode;
        res.headers = responseSpec.headers || {};
        res.resume = jest.fn();
        callback(res);

        for (const chunk of responseSpec.chunks || []) {
          res.emit('data', chunk);
        }
        res.emit('end');
      });
    };
    return req;
  });
}

describe('Graph API helper', () => {
  beforeEach(() => {
    jest.clearAllMocks();
  });

  it('should follow Graph download redirects and return binary buffers', async () => {
    const requests = [];
    const responses = [
      mockResponse(302, { location: 'https://purposeenergy0-my.sharepoint.com/download.aspx?tempauth=abc' }),
      mockResponse(200, { 'content-type': 'video/mp4' }, [Buffer.from([0, 1, 2, 3])])
    ];
    installHttpsResponses(responses, requests);

    const { callGraphAPI } = require('../utils/graph-api');
    const result = await callGraphAPI('token123', 'GET', '/drives/drive123/items/file123/content');

    expect(Buffer.isBuffer(result)).toBe(true);
    expect(result).toEqual(Buffer.from([0, 1, 2, 3]));
    expect(requests).toHaveLength(2);
    expect(requests[0].options.headers.Authorization).toBe('Bearer token123');
    expect(requests[1].url).toContain('purposeenergy0-my.sharepoint.com');
    expect(requests[1].options.headers.Authorization).toBeUndefined();
  });
});

describe('Graph API helper behavior', () => {
  describe('getAnchorMailboxFromPath', () => {
    it('extracts anchor mailbox for mailbox-scoped mail paths', () => {
      expect(getAnchorMailboxFromPath('users/engineering.agent@circleh2o.com/messages')).toBe('engineering.agent@circleh2o.com');
      expect(getAnchorMailboxFromPath('/users/shared%40example.com/mailFolders/inbox/messages')).toBe('shared@example.com');
      expect(getAnchorMailboxFromPath('users/shared@example.com/sendMail')).toBe('shared@example.com');
    });

    it('does not anchor non-mailbox user paths', () => {
      expect(getAnchorMailboxFromPath('me/messages')).toBeNull();
      expect(getAnchorMailboxFromPath('users/person@example.com/manager')).toBeNull();
      expect(getAnchorMailboxFromPath('users')).toBeNull();
    });
  });

  describe('createGraphAPIError', () => {
    it('preserves structured Graph error details on the thrown Error', () => {
      const parsedError = {
        error: {
          code: 'ErrorMailboxMoveInProgress',
          message: 'Mailbox move in progress. Cross Server access is not allowed.',
          innerError: {
            date: '2026-05-16T10:00:00',
            'request-id': 'request-123',
            'client-request-id': 'client-456'
          }
        }
      };

      const error = createGraphAPIError(503, JSON.stringify(parsedError), parsedError);

      expect(error).toBeInstanceOf(Error);
      expect(error.statusCode).toBe(503);
      expect(error.message).toContain('API call failed with status 503');
      expect(error.graphError).toEqual({
        statusCode: 503,
        code: 'ErrorMailboxMoveInProgress',
        message: 'Mailbox move in progress. Cross Server access is not allowed.',
        innerError: parsedError.error.innerError,
        requestId: 'request-123',
        date: '2026-05-16T10:00:00'
      });
    });
  });
});

describe('mislabelled binary downloads', () => {
  beforeEach(() => {
    // clearAllMocks only. resetModules() tears down the https mock wiring that
    // installHttpsResponses depends on, and the failure surfaces as an opaque
    // "Cannot read properties of undefined (reading 'on')".
    jest.clearAllMocks();
  });

  it('returns a Buffer when Content-Type lies and says text/plain', async () => {
    // SharePoint's pre-authenticated download host labels .xlsx as text/plain.
    // Decoding that as UTF-8 replaces every invalid byte with U+FFFD, which is
    // irreversible -- observed on a real workbook, where the MCP path produced
    // mangled text and the same URL fetched directly produced a valid PK zip.
    const XLSX = Buffer.concat([
      Buffer.from([0x50, 0x4b, 0x03, 0x04]),
      Buffer.from([0xde, 0xad, 0xbe, 0xef, 0x00, 0x01]),
    ]);
    const requests = [];
    installHttpsResponses([mockResponse(200, { 'content-type': 'text/plain' }, [XLSX])], requests);

    const { callGraphAPI } = require('../utils/graph-api');
    const result = await callGraphAPI('tok', 'GET', '/drives/d/items/i/content');

    expect(Buffer.isBuffer(result)).toBe(true);
    expect(result).toEqual(XLSX);
  });

  it('still returns genuine text/plain as a string', () => {
    // The guard must not turn every text response into a Buffer.
    const { looksBinary } = require('../utils/graph-api');
    if (typeof looksBinary === 'function') {
      expect(looksBinary(Buffer.from('WEBVTT\n\n00:00.000 --> 00:01.000\nhello\n'))).toBe(false);
      expect(looksBinary(Buffer.from([0x50, 0x4b, 0x03, 0x04]))).toBe(true);
    }
  });

  it('does not misread a plain-text transcript as binary', async () => {
    const VTT = Buffer.from('WEBVTT\n\n00:00.000 --> 00:02.000\nSpeaker: hello\n');
    const requests = [];
    installHttpsResponses([mockResponse(200, { 'content-type': 'text/vtt' }, [VTT])], requests);

    const { callGraphAPI } = require('../utils/graph-api');
    const result = await callGraphAPI('tok', 'GET', '/me/onlineMeetings/x/transcripts/y/content');

    expect(typeof result).toBe('string');
    expect(result).toContain('WEBVTT');
  });
});

describe('Teams transcript content policy', () => {
  const format = 'application/vnd.microsoft.graph.transcript+text';
  const metadata = () => mockResponse(200, { 'content-type': 'application/json' }, [Buffer.from('{}')]);
  const forbidden = code => mockResponse(403, { 'content-type': 'application/json' }, [
    Buffer.from(JSON.stringify({ error: { code: 'Forbidden', message: code, innerError: { code } } }))
  ]);

  it.each(['SpeakerAttributionNotAllowed', 'GraphAccessToTranscriptsDisabled'])(
    'handles %s without losing the transcript or retrying a tenant access refusal', async code => {
      jest.spyOn(require('../auth'), 'ensureAuthenticated').mockResolvedValue('token');
      try {
        const meeting = require('../teams/consolidated/teams_meeting');
        const requests = [];
        const body = '00:00:02.000 --> 00:00:04.000\nReview the proposal on Friday.';
        const responses = [metadata(), forbidden(code)];
        if (code === 'SpeakerAttributionNotAllowed') {
          responses.push(mockResponse(200, { 'content-type': format }, [Buffer.from(body)]));
        }
        installHttpsResponses(responses, requests);
        const result = await meeting({ operation: 'get_transcript', meetingId: 'm', transcriptId: 't' });
        if (code === 'SpeakerAttributionNotAllowed') {
          expect(requests).toHaveLength(3);
          expect(requests[2].options.headers.Accept).toEqual(format);
          expect(requests[2].url).not.toContain('?');
          expect(result.content[0].text).toContain(body);
        } else {
          expect(requests).toHaveLength(2);
          expect(result.content[0].text).toContain(code);
        }
      } finally {
        jest.restoreAllMocks();
      }
    }
  );
});

// HOUSE-OFFICE-MCP-NO-SENDMAIL-AUTO-RETRY-20261119: a sendMail Graph answers with 503 or 504 may have gone out, so it
// is never retried automatically; a 429 (throttled, nothing sent) still retries. The error says the send is
// unconfirmed and how to check Sent Items before sending again. Every other call keeps its retry.
describe('sendMail is never retried automatically on 503/504', () => {
  const SEND = 'users/agent@example.com/sendMail';
  const busy = status => mockResponse(status, { 'content-type': 'application/json' },
    [Buffer.from(JSON.stringify({ error: { code: 'ServiceUnavailable', message: 'The service is unavailable.' } }))]);
  const accepted = () => mockResponse(202, {}, []);    // Graph answers sendMail 202 Accepted with no body
  const ok = () => mockResponse(200, { 'content-type': 'application/json' }, [Buffer.from('{"value":[]}')]);

  beforeEach(() => {
    jest.clearAllMocks();
  });

  it.each([503, 504])('a sendMail answered %i is sent to Graph once and reported unconfirmed', async status => {
    const requests = [];
    installHttpsResponses([busy(status), accepted()], requests);
    const error = await callGraphAPI('access-token', 'POST', SEND, { message: {} }, null).catch(e => e);
    expect(requests).toHaveLength(1);                       // the 202 waiting behind it was never asked for
    expect(error).toBeInstanceOf(Error);
    expect(error.sendUnconfirmed).toBe(true);
    expect(error.statusCode).toBe(status);
    expect(error.ruling).toBe(NO_SENDMAIL_AUTO_RETRY_RULING_ID);
    expect(error.message).toContain(`The send is UNCONFIRMED: sendMail was answered with HTTP ${status}`);
    expect(error.message).toContain(NO_SENDMAIL_AUTO_RETRY_RULING_ID);
    expect(error.message).toContain('Do NOT send it again yet');
    expect(error.message).toContain('mail { operation: "search", mailbox: "<the sending mailbox>", folderName: "sent"');
    expect(error.message).toContain('If the email is in Sent Items it went out: do not send again');
    expect(error.message).not.toMatch(/retried automatically\.|try again later/i);
    expect(error.graphError.code).toBe('ServiceUnavailable');   // the structured Graph error is kept
  });

  it('a throttled sendMail (429) still retries, and the email goes out once', async () => {
    const requests = [];
    installHttpsResponses([busy(429), accepted()], requests);
    await callGraphAPI('access-token', 'POST', SEND, { message: {} }, null);
    expect(requests).toHaveLength(2);
  });

  it('a sendMail whose connection fails is unconfirmed too, and is not retried', async () => {
    let asked = 0;
    https.request.mockImplementation(() => {
      asked += 1;
      const req = new EventEmitter();
      req.write = jest.fn();
      req.end = () => process.nextTick(() => req.emit('error', new Error('socket hang up')));
      return req;
    });
    const error = await callGraphAPI('access-token', 'POST', 'me/sendMail', { message: {} }, null).catch(e => e);
    expect(asked).toBe(1);
    expect(error.sendUnconfirmed).toBe(true);
    expect(error.statusCode).toBeUndefined();               // so a keyed send never takes it for a refused stamp
    expect(error.message).toContain('Network error during API call: socket hang up');
    expect(error.message).toContain('The send is UNCONFIRMED: the connection failed before sendMail was answered');
  });

  it('a sendMail refused outright (400, 403) is a plain error, and a 500 is unconfirmed without any retry', async () => {
    for (const status of [400, 403]) {
      const requests = [];
      installHttpsResponses([busy(status)], requests);
      const error = await callGraphAPI('access-token', 'POST', SEND, { message: {} }, null).catch(e => e);
      expect(requests).toHaveLength(1);
      expect(error.statusCode).toBe(status);
      expect(error.sendUnconfirmed).toBeUndefined();
    }
    const requests = [];
    installHttpsResponses([busy(500)], requests);
    const error = await callGraphAPI('access-token', 'POST', SEND, { message: {} }, null).catch(e => e);
    expect(requests).toHaveLength(1);
    expect(error.sendUnconfirmed).toBe(true);
    expect(error.message).not.toContain('Please try again later');
  });

  it('every call that transmits nothing keeps its automatic retry', async () => {
    const reads = [];
    installHttpsResponses([busy(503), ok()], reads);
    await expect(callGraphAPI('access-token', 'GET', 'users/agent@example.com/messages', null, {})).resolves.toEqual({ value: [] });
    expect(reads).toHaveLength(2);
    // writing a draft, a reply draft or an attachment sends nothing: retried as before
    for (const path of ['me/messages', 'me/messages/AAMk/createReply', 'me/messages/AAMk/createReplyAll',
                        'me/messages/AAMk/attachments']) {
      expect(mailTransmitCall('POST', path)).toBeNull();
      expect([429, 503, 504].map(status => isRetryableStatus('POST', path, status))).toEqual([true, true, true]);
    }
    expect(isSendMailCall('POST', 'me/sendMail')).toBe(true);
    expect(isSendMailCall('POST', 'users/agent@example.com/sendMail')).toBe(true);
    expect(isSendMailCall('GET', 'me/sendMail')).toBe(false);
    expect([429, 503, 504].map(status => isRetryableStatus('POST', SEND, status))).toEqual([true, false, false]);
  });
});

// HOUSE-OFFICE-MCP-NO-AUTO-RETRY-ALL-MAIL-TRANSMITS-20261119: sending a draft, reply and reply-all get sendMail's
// rule. Answered 503 or 504 they may have gone out, so they are never retried automatically; a 429 still retries. The
// error's Sent Items check is worded for what was being sent.
describe('send a draft, reply and reply-all are never retried automatically on 503/504', () => {
  const busy = status => mockResponse(status, { 'content-type': 'application/json' },
    [Buffer.from(JSON.stringify({ error: { code: 'ServiceUnavailable', message: 'The service is unavailable.' } }))]);
  const accepted = () => mockResponse(202, {}, []);
  const TRANSMITS = [
    ['sendDraft', 'users/agent@example.com/messages/AAMkDraft=/send', 'the send of the draft', 'folderName: "drafts"',
      'send the same draft once more (mail { operation: "send_draft", draftId: "<the draft>" }), never a new copy'],
    ['reply', 'users/agent@example.com/messages/AAMkOrig=/reply', 'the reply', 'the replying mailbox\'s Sent Items',
      'reply once more'],
    ['replyAll', 'users/agent@example.com/messages/AAMkOrig=/replyAll', 'the reply-all',
      'the replying mailbox\'s Sent Items', 'reply once more']
  ];
  const CASES = TRANSMITS.flatMap(transmit => [503, 504].map(status => [...transmit, status]));

  beforeEach(() => {
    jest.clearAllMocks();
  });

  it.each(CASES)('%s (%s) answered as "%s": asked of Graph once, reported unconfirmed',
    async (kind, path, call, check, again, status) => {
      const requests = [];
      installHttpsResponses([busy(status), accepted()], requests);
      const error = await callGraphAPI('access-token', 'POST', path, null, null).catch(e => e);
      expect(requests).toHaveLength(1);                     // the 202 waiting behind it was never asked for
      expect(mailTransmitCall('POST', path)).toBe(kind);
      expect(error.sendUnconfirmed).toBe(true);
      expect(error.transmit).toBe(kind);
      expect(error.statusCode).toBe(status);
      expect(error.ruling).toBe(NO_AUTO_RETRY_ALL_MAIL_TRANSMITS_RULING_ID);
      expect(error.message).toContain(`The send is UNCONFIRMED: ${call} was answered with HTTP ${status}`);
      expect(error.message).toContain(NO_AUTO_RETRY_ALL_MAIL_TRANSMITS_RULING_ID);
      expect(error.message).toContain('Do NOT send it again yet');
      expect(error.message).toContain(check);
      expect(error.message).toContain(`Only if it is still absent, ${again}.`);
      expect(error.message).not.toMatch(/retried automatically\.|try again later/i);
    });

  it.each(TRANSMITS)('%s throttled (429) still retries, and goes out once', async (kind, path) => {
    const requests = [];
    installHttpsResponses([busy(429), accepted()], requests);
    await callGraphAPI('access-token', 'POST', path, null, null);
    expect(requests).toHaveLength(2);
  });

  it('a draft send whose connection fails is unconfirmed too, and a refused one (404, 400) is a plain error', async () => {
    let asked = 0;
    https.request.mockImplementation(() => {
      asked += 1;
      const req = new EventEmitter();
      req.write = jest.fn();
      req.end = () => process.nextTick(() => req.emit('error', new Error('socket hang up')));
      return req;
    });
    const lost = await callGraphAPI('access-token', 'POST', 'me/messages/AAMkDraft=/send', null, null).catch(e => e);
    expect(asked).toBe(1);
    expect(lost.sendUnconfirmed).toBe(true);
    expect(lost.transmit).toBe('sendDraft');
    expect(lost.message).toContain('the connection failed before the send of the draft was answered');
    for (const status of [400, 404]) {
      const requests = [];
      installHttpsResponses([busy(status)], requests);
      const error = await callGraphAPI('access-token', 'POST', 'me/messages/AAMkDraft=/send', null, null).catch(e => e);
      expect(requests).toHaveLength(1);
      expect(error.sendUnconfirmed).toBeUndefined();
    }
  });

  it('the sendMail wording is unchanged', () => {
    const error = createGraphAPIError(503, '{}', { error: { code: 'ServiceUnavailable', message: 'busy' } }, {},
      { unconfirmedSend: true });
    expect(error.transmit).toBe('sendMail');
    expect(error.ruling).toBe(NO_SENDMAIL_AUTO_RETRY_RULING_ID);
    expect(error.message).toContain('The send is UNCONFIRMED: sendMail was answered with HTTP 503');
    expect(error.message).toContain('First check the sending mailbox\'s Sent Items');
    expect(error.message).toContain('Only if it is still absent, send it once more.');
  });
});
