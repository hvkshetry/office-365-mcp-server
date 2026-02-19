const { describe, it, expect } = require('@jest/globals');
const { formatError, safeTool, getHint } = require('../utils/errors');

describe('Error Utilities', () => {
  describe('getHint', () => {
    it('should return auth hint for 401 errors', () => {
      const error = new Error('Request failed with status 401');
      expect(getHint(error)).toContain('Authentication expired');
    });

    it('should return permission hint for 403 errors', () => {
      const error = new Error('Request failed with status 403');
      expect(getHint(error)).toContain('Permission denied');
    });

    it('should return not found hint for 404 errors', () => {
      const error = new Error('Request failed with status 404');
      expect(getHint(error)).toContain('Resource not found');
    });

    it('should return rate limit hint for 429 errors', () => {
      const error = new Error('Request failed with status 429');
      expect(getHint(error)).toContain('Rate limit');
    });

    it('should return bad request hint for 400 errors', () => {
      const error = new Error('Request failed with status 400');
      expect(getHint(error)).toContain('Invalid request');
    });

    it('should return timeout hint for timeout errors', () => {
      const error = new Error('Request timeout');
      expect(getHint(error)).toContain('timed out');
    });

    it('should return connection hint for ECONNREFUSED', () => {
      const error = new Error('connect ECONNREFUSED 127.0.0.1:443');
      expect(getHint(error)).toContain('Connection refused');
    });

    it('should return null for unknown errors', () => {
      const error = new Error('Something unexpected happened');
      expect(getHint(error)).toBeNull();
    });
  });

  describe('formatError', () => {
    it('should return isError: true', () => {
      const result = formatError(new Error('test error'));
      expect(result.isError).toBe(true);
    });

    it('should include error message in content', () => {
      const result = formatError(new Error('test error'));
      expect(result.content[0].type).toBe('text');
      expect(result.content[0].text).toContain('test error');
    });

    it('should include context prefix when provided', () => {
      const result = formatError(new Error('test error'), 'calendar.list');
      expect(result.content[0].text).toContain('[calendar.list]');
    });

    it('should include hint for known error codes', () => {
      const result = formatError(new Error('Request failed with status 401'));
      expect(result.content[0].text).toContain('Hint:');
      expect(result.content[0].text).toContain('Authentication expired');
    });

    it('should not include hint for unknown errors', () => {
      const result = formatError(new Error('unknown'));
      expect(result.content[0].text).not.toContain('Hint:');
    });

    it('should handle string errors', () => {
      const result = formatError('string error');
      expect(result.content[0].text).toContain('string error');
      expect(result.isError).toBe(true);
    });
  });

  describe('safeTool', () => {
    it('should pass through successful results', async () => {
      const handler = async (args) => ({
        content: [{ type: 'text', text: 'success' }]
      });

      const wrapped = safeTool('test', handler);
      const result = await wrapped({ operation: 'list' });
      expect(result.content[0].text).toBe('success');
      expect(result.isError).toBeUndefined();
    });

    it('should catch and format thrown errors', async () => {
      const handler = async () => {
        throw new Error('something broke');
      };

      const wrapped = safeTool('test', handler);
      const result = await wrapped({ operation: 'create' });
      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain('[test.create]');
      expect(result.content[0].text).toContain('something broke');
    });

    it('should include hint for Graph API errors', async () => {
      const handler = async () => {
        throw new Error('Request failed with status 401');
      };

      const wrapped = safeTool('mail', handler);
      const result = await wrapped({ operation: 'list' });
      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain('Hint:');
      expect(result.content[0].text).toContain('Authentication expired');
    });

    it('should use entity field when operation is absent', async () => {
      const handler = async () => {
        throw new Error('fail');
      };

      const wrapped = safeTool('planner', handler);
      const result = await wrapped({ entity: 'task' });
      expect(result.content[0].text).toContain('[planner.task]');
    });

    it('should handle missing args gracefully', async () => {
      const handler = async () => {
        throw new Error('fail');
      };

      const wrapped = safeTool('test', handler);
      const result = await wrapped();
      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain('[test.unknown]');
    });
  });
});
