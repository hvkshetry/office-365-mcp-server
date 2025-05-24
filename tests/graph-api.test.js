const { describe, it, expect, jest } = require('@jest/globals');
const EventEmitter = require('events');

jest.mock('https');

const https = require('https');
const { callGraphAPI } = require('../utils/graph-api');
const config = require('../config');

describe('graph-api path encoding', () => {
  it('preserves colon slash sequences', async () => {
    const path = '/me/drive/items/123:/foo bar.txt:/content';
    const expectedUrl = `${config.GRAPH_API_ENDPOINT}me/drive/items/123:/foo%20bar.txt:/content`;

    https.request.mockImplementation((url, options, callback) => {
      const res = new EventEmitter();
      res.statusCode = 200;
      res.headers = { 'content-type': 'application/json' };
      callback(res);
      process.nextTick(() => {
        res.emit('data', JSON.stringify({ ok: true }));
        res.emit('end');
      });
      return {
        on: jest.fn(),
        write: jest.fn(),
        end: jest.fn()
      };
    });

    const result = await callGraphAPI('token', 'GET', path);
    expect(https.request).toHaveBeenCalledWith(expectedUrl, expect.any(Object), expect.any(Function));
    expect(result.ok).toBe(true);
  });
});
