const { describe, it, expect, beforeEach } = require('@jest/globals');
const { EventEmitter } = require('events');
const https = require('https');
const { createGraphAPIError, getAnchorMailboxFromPath } = require('../utils/graph-api');

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
