const { describe, it, expect, beforeEach } = require('@jest/globals');
const { EventEmitter } = require('events');
const https = require('https');

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
