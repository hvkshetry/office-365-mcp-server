/**
 * Token management for Microsoft Graph API authentication
 */
const fs = require('fs');
const config = require('../config');
const { refreshAccessToken, needsRefresh } = require('./auto-refresh');

// Global variable to store tokens
let cachedTokens = null;

/**
 * Loads authentication tokens from the token file
 * @returns {object|null} - The loaded tokens or null if not available
 */
function loadTokenCache() {
  try {
    const tokenPath = config.AUTH_CONFIG.tokenStorePath;

    if (!fs.existsSync(tokenPath)) {
      console.error('[TOKEN-MANAGER] Token file not found');
      return null;
    }

    const tokenData = fs.readFileSync(tokenPath, 'utf8');

    try {
      const tokens = JSON.parse(tokenData);

      // Safe logging - only confirm presence, never log content
      console.error('[TOKEN-MANAGER] Token loaded - has access_token:', !!tokens.access_token);
      console.error('[TOKEN-MANAGER] Token loaded - has refresh_token:', !!tokens.refresh_token);

      if (!tokens.access_token) {
        console.error('[TOKEN-MANAGER] No access_token found in token file');
        return null;
      }

      // Check token expiration
      const now = Date.now();
      const expiresAt = tokens.expires_at || 0;

      if (now > expiresAt) {
        console.error('[TOKEN-MANAGER] Token expired - will need refresh');
      } else {
        const expiresIn = Math.round((expiresAt - now) / 1000 / 60);
        console.error(`[TOKEN-MANAGER] Token valid for ~${expiresIn} minutes`);
      }

      cachedTokens = tokens;
      return tokens;
    } catch (parseError) {
      console.error('[TOKEN-MANAGER] Error parsing token file');
      return null;
    }
  } catch (error) {
    console.error('[TOKEN-MANAGER] Error loading tokens:', error.message);
    return null;
  }
}

/**
 * Saves authentication tokens to the token file
 * @param {object} tokens - The tokens to save
 * @returns {boolean} - Whether the save was successful
 */
function saveTokenCache(tokens) {
  try {
    const tokenPath = config.AUTH_CONFIG.tokenStorePath;
    const tempPath = tokenPath + '.tmp';

    // Atomic write: write to temp file then rename to avoid corruption
    fs.writeFileSync(tempPath, JSON.stringify(tokens, null, 2), { mode: 0o600 });
    fs.renameSync(tempPath, tokenPath);

    // Also ensure correct permissions on final file (for pre-existing files)
    try {
      fs.chmodSync(tokenPath, 0o600);
    } catch (chmodError) {
      // Windows may not support chmod, that's OK
    }

    console.error('[TOKEN-MANAGER] Tokens saved securely');

    cachedTokens = tokens;
    return true;
  } catch (error) {
    console.error('[TOKEN-MANAGER] Error saving tokens:', error.message);
    return false;
  }
}

/**
 * Gets the current access token, loading from cache if necessary
 * @param {boolean} autoRefresh - Whether to automatically refresh expired tokens
 * @returns {Promise<string>|string|null} - The access token or null if not available
 */
async function getAccessToken(autoRefresh = true) {
  // First check cache
  if (cachedTokens && cachedTokens.access_token) {
    if (autoRefresh && needsRefresh(cachedTokens)) {
      console.error('[TOKEN-MANAGER] Cached token needs refresh');
      try {
        const newTokens = await refreshAccessToken();
        cachedTokens = newTokens;
        return newTokens.access_token;
      } catch (error) {
        console.error('[TOKEN-MANAGER] Auto-refresh failed:', error.message);
        // Return existing token if refresh fails (might still work briefly)
        return cachedTokens.access_token;
      }
    }
    return cachedTokens.access_token;
  }
  
  // Load from file
  const tokens = loadTokenCache();
  if (!tokens || !tokens.access_token) {
    return null;
  }
  
  // Check if refresh needed
  if (autoRefresh && needsRefresh(tokens)) {
    console.error('[TOKEN-MANAGER] Token needs refresh');
    try {
      const newTokens = await refreshAccessToken();
      cachedTokens = newTokens;
      return newTokens.access_token;
    } catch (error) {
      console.error('[TOKEN-MANAGER] Auto-refresh failed:', error.message);
      // Return existing token if refresh fails
      return tokens.access_token;
    }
  }
  
  return tokens.access_token;
}

/**
 * Creates a test access token for use in test mode
 * @returns {object} - The test tokens
 */
function createTestTokens() {
  const testTokens = {
    access_token: "test_access_token_" + Date.now(),
    refresh_token: "test_refresh_token_" + Date.now(),
    expires_at: Date.now() + (3600 * 1000) // 1 hour
  };
  
  saveTokenCache(testTokens);
  return testTokens;
}

module.exports = {
  loadTokenCache,
  saveTokenCache,
  getAccessToken,
  createTestTokens
};
