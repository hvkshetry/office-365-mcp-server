const { describe, it, expect, jest } = require('@jest/globals');
const { handleAuthenticate, handleGetAuthStatus, handleLogout } = require('../auth');
const tokenManager = require('../auth/token-manager');

jest.mock('../auth/token-manager');

describe('Authentication Module', () => {
  beforeEach(() => {
    jest.clearAllMocks();
  });

  describe('handleAuthenticate', () => {
    it('should initiate authentication successfully', async () => {
      const result = await handleAuthenticate({});
      
      expect(result.content[0].type).toBe('text');
      expect(result.content[0].text).toContain('Please visit the following URL');
    });
  });

  describe('handleGetAuthStatus', () => {
    it('should return authenticated status with valid tokens', async () => {
      const mockTokens = {
        access_token: 'mock-access-token',
        email: 'user@example.com'
      };
      
      tokenManager.loadTokenCache.mockReturnValue(mockTokens);
      tokenManager.checkAuthStatus.mockResolvedValue(true);
      
      const result = await handleGetAuthStatus({});
      
      expect(result.content[0].text).toContain('Authenticated as user@example.com');
    });

    it('should return not authenticated status with no tokens', async () => {
      tokenManager.loadTokenCache.mockReturnValue(null);
      
      const result = await handleGetAuthStatus({});
      
      expect(result.content[0].text).toBe('Not authenticated. Use the authenticate tool to log in.');
    });
  });

  describe('handleLogout', () => {
    it('should successfully logout', async () => {
      tokenManager.loadTokenCache.mockReturnValue({ access_token: 'mock-token' });
      tokenManager.clearTokenCache.mockImplementation(() => {});
      
      const result = await handleLogout({});
      
      expect(tokenManager.clearTokenCache).toHaveBeenCalled();
      expect(result.content[0].text).toBe('Successfully logged out.');
    });

    it('should handle logout when not authenticated', async () => {
      tokenManager.loadTokenCache.mockReturnValue(null);
      
      const result = await handleLogout({});
      
      expect(result.content[0].text).toBe('Not currently authenticated.');
    });
  });
});