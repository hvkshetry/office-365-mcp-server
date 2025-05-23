const { jest } = require('@jest/globals');

/**
 * Creates a mock token object for testing
 */
function createMockTokens() {
  return {
    access_token: 'mock-access-token',
    email: 'user@example.com',
    refresh_token: 'mock-refresh-token',
    expires_in: 3600
  };
}

/**
 * Creates a mock Graph API response with pagination
 */
function createMockPagedResponse(items, nextLink = null) {
  const response = {
    value: items
  };
  
  if (nextLink) {
    response['@odata.nextLink'] = nextLink;
  }
  
  return response;
}

/**
 * Creates a mock error response from Graph API
 */
function createMockGraphError(code, message, statusCode = 400) {
  const error = new Error(message);
  error.response = {
    data: {
      error: {
        code: code,
        message: message
      }
    },
    status: statusCode
  };
  return error;
}

/**
 * Resets all mocks between tests
 */
function resetAllMocks() {
  jest.clearAllMocks();
  jest.resetAllMocks();
}

module.exports = {
  createMockTokens,
  createMockPagedResponse,
  createMockGraphError,
  resetAllMocks
};