const { describe, it, expect, jest } = require('@jest/globals');
const {
  handleCreateSubscription,
  handleListSubscriptions,
  handleDeleteSubscription,
  handleUpdateSubscription,
  handleManageWebhook
} = require('../notifications');
const tokenManager = require('../auth/token-manager');
const { callGraphAPI } = require('../utils/graph-api');

jest.mock('../auth/token-manager');
jest.mock('../utils/graph-api');

describe('Notifications Module', () => {
  const mockTokens = {
    access_token: 'mock-access-token',
    email: 'user@example.com'
  };

  beforeEach(() => {
    jest.clearAllMocks();
    tokenManager.loadTokenCache.mockReturnValue(mockTokens);
  });

  describe('handleCreateSubscription', () => {
    it('should create a subscription successfully', async () => {
      const mockResponse = {
        id: 'sub123',
        resource: '/users/user123/events',
        changeType: 'created,updated',
        expirationDateTime: '2024-01-15T00:00:00Z'
      };
      
      callGraphAPI.mockResolvedValue(mockResponse);
      
      const result = await handleCreateSubscription({
        resource: '/users/user123/events',
        changeType: 'created,updated',
        notificationUrl: 'https://webhook.example.com/events',
        expirationMinutes: 60
      });
      
      expect(callGraphAPI).toHaveBeenCalledWith(
        'mock-access-token',
        'POST',
        '/subscriptions',
        expect.objectContaining({
          changeType: 'created,updated',
          notificationUrl: 'https://webhook.example.com/events',
          resource: '/users/user123/events',
          clientState: 'secretClientState'
        })
      );
      expect(result.content[0].text).toContain('Subscription created successfully');
    });

    it('should validate required parameters', async () => {
      const result = await handleCreateSubscription({
        resource: '/users/user123/events'
      });
      
      expect(result.content[0].text).toContain('Missing required parameters');
    });

    it('should handle includeResourceData option', async () => {
      const mockResponse = { id: 'sub123' };
      
      callGraphAPI.mockResolvedValue(mockResponse);
      
      await handleCreateSubscription({
        resource: '/users/user123/events',
        changeType: 'created',
        notificationUrl: 'https://webhook.example.com/events',
        includeResourceData: true
      });
      
      expect(callGraphAPI).toHaveBeenCalledWith(
        'mock-access-token',
        'POST',
        '/subscriptions',
        expect.objectContaining({
          includeResourceData: true
        })
      );
    });
  });

  describe('handleListSubscriptions', () => {
    it('should list all subscriptions', async () => {
      const mockSubscriptions = {
        value: [
          { id: 'sub1', resource: '/users/user1/events', changeType: 'created' },
          { id: 'sub2', resource: '/teams/team1/channels', changeType: 'updated' }
        ]
      };
      
      callGraphAPI.mockResolvedValue(mockSubscriptions);
      
      const result = await handleListSubscriptions({});
      
      expect(callGraphAPI).toHaveBeenCalledWith(
        'mock-access-token',
        'GET',
        '/subscriptions',
        null
      );
      expect(result.content[0].text).toContain('Found 2 subscriptions');
    });
  });

  describe('handleDeleteSubscription', () => {
    it('should delete a subscription successfully', async () => {
      callGraphAPI.mockResolvedValue(null);
      
      const result = await handleDeleteSubscription({
        subscriptionId: 'sub123'
      });
      
      expect(callGraphAPI).toHaveBeenCalledWith(
        'mock-access-token',
        'DELETE',
        '/subscriptions/sub123',
        null
      );
      expect(result.content[0].text).toBe('Successfully deleted subscription.');
    });

    it('should validate required parameters', async () => {
      const result = await handleDeleteSubscription({});
      
      expect(result.content[0].text).toContain('Missing required parameter: subscriptionId');
    });
  });

  describe('handleUpdateSubscription', () => {
    it('should update subscription expiration', async () => {
      const mockResponse = {
        id: 'sub123',
        expirationDateTime: '2024-01-16T00:00:00Z'
      };
      
      callGraphAPI.mockResolvedValue(mockResponse);
      
      const result = await handleUpdateSubscription({
        subscriptionId: 'sub123',
        expirationMinutes: 120
      });
      
      expect(callGraphAPI).toHaveBeenCalledWith(
        'mock-access-token',
        'PATCH',
        '/subscriptions/sub123',
        expect.objectContaining({
          expirationDateTime: expect.any(String)
        })
      );
      expect(result.content[0].text).toContain('Successfully updated subscription');
    });
  });

  describe('handleManageWebhook', () => {
    it('should validate a webhook', async () => {
      const result = await handleManageWebhook({
        action: 'validate',
        validationToken: 'test-token-123'
      });
      
      expect(result.content[0].text).toBe('test-token-123');
    });

    it('should process a webhook notification', async () => {
      const notification = {
        resource: '/users/user123/events/event123',
        changeType: 'created',
        clientState: 'secretClientState',
        resourceData: {
          id: 'event123',
          subject: 'Team Meeting'
        }
      };
      
      const result = await handleManageWebhook({
        action: 'process',
        notifications: [notification]
      });
      
      expect(result.content[0].text).toContain('Processed 1 notifications');
    });

    it('should validate required parameters', async () => {
      const result = await handleManageWebhook({});
      
      expect(result.content[0].text).toContain('Missing required parameter: action');
    });
  });
});