const { describe, it, expect } = require('@jest/globals');
const { notificationTools } = require('../notifications');
const { ensureAuthenticated } = require('../auth');
const { callGraphAPI } = require('../utils/graph-api');

jest.mock('../auth', () => ({
  ensureAuthenticated: jest.fn()
}));
jest.mock('../utils/graph-api');

// The consolidated notifications tool
const notificationsHandler = notificationTools[0].handler;

describe('Notifications Module (Consolidated)', () => {
  const mockAccessToken = 'mock-access-token';

  beforeEach(() => {
    jest.clearAllMocks();
    ensureAuthenticated.mockResolvedValue(mockAccessToken);
  });

  describe('routing', () => {
    it('should require operation parameter', async () => {
      const result = await notificationsHandler({});
      expect(result.content[0].text).toContain('Missing required parameter: operation');
    });

    it('should reject invalid operation', async () => {
      const result = await notificationsHandler({ operation: 'invalid' });
      expect(result.content[0].text).toContain('Invalid operation');
    });
  });

  describe('create operation', () => {
    it('should create a subscription successfully', async () => {
      const mockResponse = {
        id: 'sub123',
        resource: '/users/user123/events',
        expirationDateTime: '2024-01-15T00:00:00Z'
      };

      callGraphAPI.mockResolvedValue(mockResponse);

      const result = await notificationsHandler({
        operation: 'create',
        resource: '/users/user123/events',
        changeType: 'created,updated',
        notificationUrl: 'https://webhook.example.com/events',
        expirationMinutes: 60
      });

      expect(callGraphAPI).toHaveBeenCalledWith(
        mockAccessToken,
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
      const result = await notificationsHandler({
        operation: 'create',
        resource: '/users/user123/events'
      });

      expect(result.content[0].text).toContain('Missing required parameters');
    });
  });

  describe('list operation', () => {
    it('should list all subscriptions', async () => {
      const mockSubscriptions = {
        value: [
          { id: 'sub1', resource: '/users/user1/events', changeType: 'created', expirationDateTime: '2024-01-15T00:00:00Z' },
          { id: 'sub2', resource: '/teams/team1/channels', changeType: 'updated', expirationDateTime: '2024-01-16T00:00:00Z' }
        ]
      };

      callGraphAPI.mockResolvedValue(mockSubscriptions);

      const result = await notificationsHandler({ operation: 'list' });

      expect(callGraphAPI).toHaveBeenCalledWith(
        mockAccessToken,
        'GET',
        '/subscriptions',
        null
      );
      expect(result.content[0].text).toContain('Found 2');
    });

    it('should handle no subscriptions', async () => {
      callGraphAPI.mockResolvedValue({ value: [] });

      const result = await notificationsHandler({ operation: 'list' });
      expect(result.content[0].text).toContain('No active subscriptions');
    });
  });

  describe('renew operation', () => {
    it('should renew a subscription', async () => {
      callGraphAPI.mockResolvedValue({});

      const result = await notificationsHandler({
        operation: 'renew',
        subscriptionId: 'sub123',
        expirationMinutes: 120
      });

      expect(callGraphAPI).toHaveBeenCalledWith(
        mockAccessToken,
        'PATCH',
        '/subscriptions/sub123',
        expect.objectContaining({
          expirationDateTime: expect.any(String)
        })
      );
      expect(result.content[0].text).toContain('renewed successfully');
    });

    it('should validate required parameters', async () => {
      const result = await notificationsHandler({ operation: 'renew' });
      expect(result.content[0].text).toContain('Missing required parameter: subscriptionId');
    });
  });

  describe('delete operation', () => {
    it('should delete a subscription', async () => {
      callGraphAPI.mockResolvedValue(null);

      const result = await notificationsHandler({
        operation: 'delete',
        subscriptionId: 'sub123'
      });

      expect(callGraphAPI).toHaveBeenCalledWith(
        mockAccessToken,
        'DELETE',
        '/subscriptions/sub123',
        null
      );
      expect(result.content[0].text).toContain('deleted successfully');
    });

    it('should validate required parameters', async () => {
      const result = await notificationsHandler({ operation: 'delete' });
      expect(result.content[0].text).toContain('Missing required parameter: subscriptionId');
    });
  });
});
