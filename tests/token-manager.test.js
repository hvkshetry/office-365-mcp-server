const { describe, it, expect, beforeEach, afterEach, jest } = require('@jest/globals');
const fs = require('fs');
const path = require('path');

const tokenFile = path.join(__dirname, 'temp-tokens.json');
let tokenManager;
let config;

function loadModule() {
  jest.resetModules();
  config = require('../config');
  config.AUTH_CONFIG.tokenStorePath = tokenFile;
  tokenManager = require('../auth/token-manager');
}

describe('Token Manager getAccessToken', () => {
  beforeEach(() => {
    if (fs.existsSync(tokenFile)) {
      fs.unlinkSync(tokenFile);
    }
    loadModule();
  });

  afterEach(() => {
    if (fs.existsSync(tokenFile)) {
      fs.unlinkSync(tokenFile);
    }
  });

  it('should reload from disk when cached token is expired', () => {
    const expired = { access_token: 'expired', refresh_token: 'r', expires_at: Date.now() - 1000 };
    const valid = { access_token: 'valid', refresh_token: 'r2', expires_at: Date.now() + 3600000 };

    tokenManager.saveTokenCache(expired);
    fs.writeFileSync(tokenFile, JSON.stringify(valid));

    const token = tokenManager.getAccessToken();
    expect(token).toBe('valid');
  });

  it('should return null when both cached and stored tokens are expired', () => {
    const expired = { access_token: 'expired', refresh_token: 'r', expires_at: Date.now() - 1000 };

    tokenManager.saveTokenCache(expired);
    const token = tokenManager.getAccessToken();
    expect(token).toBeNull();
  });
});
