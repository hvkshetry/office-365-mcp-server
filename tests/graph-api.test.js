const { describe, it, expect } = require('@jest/globals');
const { createGraphAPIError, getAnchorMailboxFromPath } = require('../utils/graph-api');

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
