const { describe, it, expect } = require('@jest/globals');
const { authTools } = require('../auth');
const tokenManager = require('../auth/token-manager');

jest.mock('../auth/token-manager');
jest.mock('../auth/auto-refresh', () => ({
  needsRefresh: jest.fn().mockReturnValue(false)
}));

// The consolidated system tool
const systemHandler = authTools[0].handler;

describe('Auth Module (Consolidated System Tool)', () => {
  beforeEach(() => {
    jest.clearAllMocks();
  });

  describe('routing', () => {
    it('should require operation parameter', async () => {
      const result = await systemHandler({});
      expect(result.content[0].text).toContain('Missing required parameter: operation');
    });

    it('should reject invalid operation', async () => {
      const result = await systemHandler({ operation: 'invalid' });
      expect(result.content[0].text).toContain('Invalid operation');
    });
  });

  describe('about operation', () => {
    it('should return server information', async () => {
      const result = await systemHandler({ operation: 'about' });
      expect(result.content[0].type).toBe('text');
      expect(result.content[0].text).toContain('Office MCP Server');
    });
  });

  describe('authenticate operation', () => {
    it('should return authentication URL', async () => {
      const result = await systemHandler({ operation: 'authenticate' });
      expect(result.content[0].type).toBe('text');
      // Should contain either auth URL or test mode success
      expect(result.content[0].text).toBeTruthy();
    });
  });

  describe('check_status operation', () => {
    it('should return not authenticated when no tokens', async () => {
      tokenManager.loadTokenCache.mockReturnValue(null);

      const result = await systemHandler({ operation: 'check_status' });
      expect(result.content[0].text).toContain('Not authenticated');
    });

    it('should return authenticated with valid tokens', async () => {
      tokenManager.loadTokenCache.mockReturnValue({
        access_token: 'mock-token',
        expires_at: Date.now() + 3600000
      });

      const result = await systemHandler({ operation: 'check_status' });
      expect(result.content[0].text).toContain('Authenticated');
    });
  });
});
