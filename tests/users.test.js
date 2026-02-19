const { describe, it, expect } = require('@jest/globals');
const { directoryTools } = require('../directory');
const { ensureAuthenticated } = require('../auth');
const { callGraphAPI } = require('../utils/graph-api');

jest.mock('../auth', () => ({
  ensureAuthenticated: jest.fn()
}));
jest.mock('../utils/graph-api');

// The consolidated directory tool
const directoryHandler = directoryTools[0].handler;

describe('Directory Module', () => {
  const mockAccessToken = 'mock-access-token';

  beforeEach(() => {
    jest.clearAllMocks();
    ensureAuthenticated.mockResolvedValue(mockAccessToken);
  });

  describe('routing', () => {
    it('should require operation parameter', async () => {
      const result = await directoryHandler({});
      expect(result.content[0].text).toContain('Missing required parameter: operation');
    });

    it('should reject invalid operation', async () => {
      const result = await directoryHandler({ operation: 'invalid' });
      expect(result.content[0].text).toContain('Invalid operation');
    });
  });

  describe('lookup_user operation', () => {
    it('should look up user by email', async () => {
      const mockUser = {
        id: 'user123',
        displayName: 'John Doe',
        mail: 'john.doe@example.com',
        jobTitle: 'Software Engineer',
        department: 'Engineering'
      };

      callGraphAPI.mockResolvedValue(mockUser);

      const result = await directoryHandler({
        operation: 'lookup_user',
        email: 'john.doe@example.com'
      });

      expect(callGraphAPI).toHaveBeenCalledWith(
        mockAccessToken,
        'GET',
        'users/john.doe@example.com',
        null,
        expect.any(Object)
      );
      expect(result.content[0].text).toContain('John Doe');
      expect(result.content[0].text).toContain('Software Engineer');
    });

    it('should require email or userId', async () => {
      const result = await directoryHandler({ operation: 'lookup_user' });
      expect(result.content[0].text).toContain('Missing required parameter');
    });
  });

  describe('get_profile operation', () => {
    it('should get current user profile', async () => {
      const mockProfile = {
        id: 'user123',
        displayName: 'John Doe',
        mail: 'john.doe@example.com',
        jobTitle: 'Engineer'
      };

      callGraphAPI.mockResolvedValue(mockProfile);

      const result = await directoryHandler({ operation: 'get_profile' });

      expect(callGraphAPI).toHaveBeenCalledWith(
        mockAccessToken, 'GET', 'me', null, expect.any(Object)
      );
      expect(result.content[0].text).toContain('John Doe');
    });

    it('should get specific user profile', async () => {
      const mockProfile = {
        id: 'user456',
        displayName: 'Jane Smith',
        mail: 'jane.smith@example.com'
      };

      callGraphAPI.mockResolvedValue(mockProfile);

      const result = await directoryHandler({
        operation: 'get_profile',
        userId: 'user456'
      });

      expect(callGraphAPI).toHaveBeenCalledWith(
        mockAccessToken, 'GET', 'users/user456', null, expect.any(Object)
      );
      expect(result.content[0].text).toContain('Jane Smith');
    });
  });

  describe('get_manager operation', () => {
    it('should get user manager', async () => {
      const mockManager = {
        id: 'mgr123',
        displayName: 'Manager Name',
        mail: 'manager@example.com',
        jobTitle: 'Director',
        department: 'Engineering'
      };

      callGraphAPI.mockResolvedValue(mockManager);

      const result = await directoryHandler({
        operation: 'get_manager',
        userId: 'user123'
      });

      expect(callGraphAPI).toHaveBeenCalledWith(
        mockAccessToken, 'GET', 'users/user123/manager', null, expect.any(Object)
      );
      expect(result.content[0].text).toContain('Manager Name');
    });
  });

  describe('get_reports operation', () => {
    it('should list direct reports', async () => {
      const mockReports = {
        value: [
          { id: 'r1', displayName: 'Report 1', mail: 'report1@example.com', jobTitle: 'Engineer', department: 'Eng' },
          { id: 'r2', displayName: 'Report 2', mail: 'report2@example.com', jobTitle: 'Designer', department: 'Design' }
        ]
      };

      callGraphAPI.mockResolvedValue(mockReports);

      const result = await directoryHandler({
        operation: 'get_reports',
        userId: 'manager123'
      });

      expect(result.content[0].text).toContain('2 direct reports');
      expect(result.content[0].text).toContain('Report 1');
      expect(result.content[0].text).toContain('Report 2');
    });

    it('should handle no direct reports', async () => {
      callGraphAPI.mockResolvedValue({ value: [] });

      const result = await directoryHandler({
        operation: 'get_reports',
        userId: 'user123'
      });

      expect(result.content[0].text).toContain('No direct reports');
    });
  });

  describe('search_users operation', () => {
    it('should search users by query', async () => {
      const mockUsers = {
        value: [
          { id: 'u1', displayName: 'John Doe', mail: 'john@example.com', jobTitle: 'Engineer', department: 'Eng' },
          { id: 'u2', displayName: 'John Smith', mail: 'jsmith@example.com', jobTitle: 'PM', department: 'Product' }
        ]
      };

      callGraphAPI.mockResolvedValue(mockUsers);

      const result = await directoryHandler({
        operation: 'search_users',
        query: 'John'
      });

      expect(result.content[0].text).toContain('Found 2 users');
      expect(result.content[0].text).toContain('John Doe');
      expect(result.content[0].text).toContain('John Smith');
    });

    it('should require query parameter', async () => {
      const result = await directoryHandler({ operation: 'search_users' });
      expect(result.content[0].text).toContain('Missing required parameter: query');
    });
  });
});
