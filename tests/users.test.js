const { describe, it, expect, jest } = require('@jest/globals');
const {
  handleGetMyProfile,
  handleGetUserProfile,
  handleSearchUsers,
  handleListDirectReports,
  handleGetPresence,
  handleUpdatePresence
} = require('../users');
const tokenManager = require('../auth/token-manager');
const { callGraphAPI } = require('../utils/graph-api');

jest.mock('../auth/token-manager');
jest.mock('../utils/graph-api');

describe('Users Module', () => {
  const mockTokens = {
    access_token: 'mock-access-token',
    email: 'user@example.com'
  };

  beforeEach(() => {
    jest.clearAllMocks();
    tokenManager.loadTokenCache.mockReturnValue(mockTokens);
  });

  describe('handleGetMyProfile', () => {
    it('should get current user profile', async () => {
      const mockProfile = {
        id: 'user123',
        displayName: 'John Doe',
        userPrincipalName: 'john.doe@example.com',
        mail: 'john.doe@example.com',
        jobTitle: 'Software Engineer'
      };
      
      callGraphAPI.mockResolvedValue(mockProfile);
      
      const result = await handleGetMyProfile({});
      
      expect(callGraphAPI).toHaveBeenCalledWith(
        'mock-access-token',
        'GET',
        '/me',
        null
      );
      expect(result.content[0].text).toContain('John Doe');
      expect(result.content[0].text).toContain('Software Engineer');
    });
  });

  describe('handleGetUserProfile', () => {
    it('should get specific user profile by ID', async () => {
      const mockProfile = {
        id: 'user456',
        displayName: 'Jane Smith',
        userPrincipalName: 'jane.smith@example.com'
      };
      
      callGraphAPI.mockResolvedValue(mockProfile);
      
      const result = await handleGetUserProfile({ userId: 'user456' });
      
      expect(callGraphAPI).toHaveBeenCalledWith(
        'mock-access-token',
        'GET',
        '/users/user456',
        null
      );
      expect(result.content[0].text).toContain('Jane Smith');
    });

    it('should get user profile by email', async () => {
      const mockProfile = {
        displayName: 'Jane Smith'
      };
      
      callGraphAPI.mockResolvedValue(mockProfile);
      
      const result = await handleGetUserProfile({ userId: 'jane.smith@example.com' });
      
      expect(callGraphAPI).toHaveBeenCalledWith(
        'mock-access-token',
        'GET',
        '/users/jane.smith@example.com',
        null
      );
    });
  });

  describe('handleSearchUsers', () => {
    it('should search users by query', async () => {
      const mockUsers = {
        value: [
          { displayName: 'John Doe', mail: 'john.doe@example.com' },
          { displayName: 'John Smith', mail: 'john.smith@example.com' }
        ]
      };
      
      callGraphAPI.mockResolvedValue(mockUsers);
      
      const result = await handleSearchUsers({ query: 'John' });
      
      expect(callGraphAPI).toHaveBeenCalledWith(
        'mock-access-token',
        'GET',
        "/users?$filter=startswith(displayName,'John') or startswith(userPrincipalName,'John')",
        null
      );
      expect(result.content[0].text).toContain('Found 2 users');
    });

    it('should show select fields when specified', async () => {
      const mockUsers = { value: [] };
      
      callGraphAPI.mockResolvedValue(mockUsers);
      
      const result = await handleSearchUsers({ 
        query: 'test',
        selectFields: ['displayName', 'mail']
      });
      
      expect(callGraphAPI).toHaveBeenCalledWith(
        'mock-access-token',
        'GET',
        "/users?$filter=startswith(displayName,'test') or startswith(userPrincipalName,'test')&$select=displayName,mail",
        null
      );
    });
  });

  describe('handleListDirectReports', () => {
    it('should list direct reports for a manager', async () => {
      const mockReports = {
        value: [
          { displayName: 'Report 1', mail: 'report1@example.com' },
          { displayName: 'Report 2', mail: 'report2@example.com' }
        ]
      };
      
      callGraphAPI.mockResolvedValue(mockReports);
      
      const result = await handleListDirectReports({ userId: 'manager123' });
      
      expect(callGraphAPI).toHaveBeenCalledWith(
        'mock-access-token',
        'GET',
        '/users/manager123/directReports',
        null
      );
      expect(result.content[0].text).toContain('Found 2 direct reports');
    });
  });

  describe('handleGetPresence', () => {
    it('should get user presence status', async () => {
      const mockPresence = {
        availability: 'Available',
        activity: 'Available'
      };
      
      callGraphAPI.mockResolvedValue(mockPresence);
      
      const result = await handleGetPresence({ userId: 'user123' });
      
      expect(callGraphAPI).toHaveBeenCalledWith(
        'mock-access-token',
        'GET',
        '/users/user123/presence',
        null
      );
      expect(result.content[0].text).toContain('Availability: Available');
    });
  });

  describe('handleUpdatePresence', () => {
    it('should update user presence status', async () => {
      callGraphAPI.mockResolvedValue({});
      
      const result = await handleUpdatePresence({
        availability: 'Busy',
        activity: 'InACall'
      });
      
      expect(callGraphAPI).toHaveBeenCalledWith(
        'mock-access-token',
        'PATCH',
        '/me/presence/setUserPreferredPresence',
        {
          availability: 'Busy',
          activity: 'InACall'
        }
      );
      expect(result.content[0].text).toBe('Successfully updated presence status.');
    });

    it('should validate required parameters', async () => {
      const result = await handleUpdatePresence({});
      
      expect(result.content[0].text).toContain('Missing required parameter: availability');
    });
  });
});