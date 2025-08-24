/**
 * Authentication module for Office MCP server
 */
const tokenManager = require('./token-manager');
const { authTools } = require('./tools');

/**
 * Ensures the user is authenticated and returns an access token
 * @param {boolean} forceNew - Whether to force a new authentication
 * @returns {Promise<string>} - Access token
 * @throws {Error} - If authentication fails
 */
async function ensureAuthenticated(forceNew = false) {
  if (forceNew) {
    // Force re-authentication
    throw new Error('Authentication required');
  }
  
  try {
    // Check for existing token with auto-refresh
    const accessToken = await tokenManager.getAccessToken(true);
    if (!accessToken) {
      throw new Error('Authentication required');
    }
    
    return accessToken;
  } catch (error) {
    console.error('[AUTH] Error ensuring authentication:', error.message);
    
    // If refresh failed, throw authentication required error
    if (error.message.includes('refresh') || error.message.includes('authentication')) {
      throw new Error('Authentication required - please run the authenticate tool');
    }
    
    throw error;
  }
}

module.exports = {
  tokenManager,
  authTools,
  ensureAuthenticated
};
