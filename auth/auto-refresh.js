/**
 * Automatic token refresh functionality for Office MCP
 * Handles refreshing access tokens using refresh tokens without user interaction
 */
const https = require('https');
const fs = require('fs');
const config = require('../config');

/**
 * Refreshes the access token using the stored refresh token
 * @returns {Promise<object>} - New token data including access_token and refresh_token
 */
async function refreshAccessToken() {
  console.error('[AUTO-REFRESH] Starting token refresh process');
  
  try {
    // Load existing tokens
    const tokenPath = config.AUTH_CONFIG.tokenStorePath;
    if (!fs.existsSync(tokenPath)) {
      throw new Error('Token file not found. Initial authentication required.');
    }
    
    const tokens = JSON.parse(fs.readFileSync(tokenPath, 'utf8'));
    
    if (!tokens.refresh_token) {
      throw new Error('No refresh token available. Re-authentication required.');
    }
    
    console.error('[AUTO-REFRESH] Found refresh token, attempting refresh');
    
    // Prepare refresh request
    const clientId = process.env.OFFICE_CLIENT_ID || config.AUTH_CONFIG.clientId;
    const clientSecret = process.env.OFFICE_CLIENT_SECRET || config.AUTH_CONFIG.clientSecret;
    const tenantId = process.env.OFFICE_TENANT_ID || 'common';
    
    if (!clientId || !clientSecret) {
      throw new Error('Missing client credentials. Check OFFICE_CLIENT_ID and OFFICE_CLIENT_SECRET.');
    }
    
    const refreshData = new URLSearchParams({
      client_id: clientId,
      client_secret: clientSecret,
      refresh_token: tokens.refresh_token,
      grant_type: 'refresh_token',
      scope: config.AUTH_CONFIG.scopes.join(' ')
    }).toString();
    
    // Make refresh request to Microsoft
    return new Promise((resolve, reject) => {
      const options = {
        hostname: 'login.microsoftonline.com',
        path: `/${tenantId}/oauth2/v2.0/token`,
        method: 'POST',
        headers: {
          'Content-Type': 'application/x-www-form-urlencoded',
          'Content-Length': Buffer.byteLength(refreshData)
        }
      };
      
      const req = https.request(options, (res) => {
        let data = '';
        
        res.on('data', (chunk) => {
          data += chunk;
        });
        
        res.on('end', () => {
          try {
            if (res.statusCode >= 200 && res.statusCode < 300) {
              const newTokens = JSON.parse(data);
              
              // Calculate expiration time
              const expiresAt = Date.now() + (newTokens.expires_in * 1000);
              newTokens.expires_at = expiresAt;
              
              // Preserve email if it exists
              if (tokens.email) {
                newTokens.email = tokens.email;
              }
              
              // Save refreshed tokens
              fs.writeFileSync(tokenPath, JSON.stringify(newTokens, null, 2), 'utf8');
              console.error('[AUTO-REFRESH] Token refresh successful');
              console.error(`[AUTO-REFRESH] New token expires at: ${new Date(expiresAt).toLocaleString()}`);
              
              resolve(newTokens);
            } else {
              const error = JSON.parse(data);
              console.error('[AUTO-REFRESH] Token refresh failed:', error);
              reject(new Error(`Token refresh failed: ${error.error_description || error.error}`));
            }
          } catch (parseError) {
            console.error('[AUTO-REFRESH] Error parsing response:', parseError);
            reject(new Error(`Failed to parse token response: ${parseError.message}`));
          }
        });
      });
      
      req.on('error', (error) => {
        console.error('[AUTO-REFRESH] Network error during refresh:', error);
        reject(new Error(`Network error during token refresh: ${error.message}`));
      });
      
      req.write(refreshData);
      req.end();
    });
  } catch (error) {
    console.error('[AUTO-REFRESH] Error in refresh process:', error);
    throw error;
  }
}

/**
 * Checks if a token needs refresh
 * @param {object} tokens - Token object with expires_at property
 * @param {number} bufferMinutes - Minutes before expiry to trigger refresh (default 5)
 * @returns {boolean} - True if token needs refresh
 */
function needsRefresh(tokens, bufferMinutes = 5) {
  if (!tokens || !tokens.expires_at) {
    return true;
  }
  
  const now = Date.now();
  const expiresAt = tokens.expires_at;
  const bufferMs = bufferMinutes * 60 * 1000;
  
  // Refresh if expired or will expire within buffer time
  const shouldRefresh = now > (expiresAt - bufferMs);
  
  if (shouldRefresh) {
    const timeLeft = Math.max(0, (expiresAt - now) / 1000 / 60);
    console.error(`[AUTO-REFRESH] Token expires in ${timeLeft.toFixed(1)} minutes, refreshing`);
  }
  
  return shouldRefresh;
}

/**
 * Gets a valid access token, refreshing if necessary
 * @returns {Promise<string>} - Valid access token
 */
async function getValidAccessToken() {
  const tokenPath = config.AUTH_CONFIG.tokenStorePath;
  
  if (!fs.existsSync(tokenPath)) {
    throw new Error('No tokens found. Initial authentication required.');
  }
  
  let tokens = JSON.parse(fs.readFileSync(tokenPath, 'utf8'));
  
  if (needsRefresh(tokens)) {
    console.error('[AUTO-REFRESH] Token needs refresh, initiating refresh');
    tokens = await refreshAccessToken();
  }
  
  if (!tokens.access_token) {
    throw new Error('No valid access token available');
  }
  
  return tokens.access_token;
}

module.exports = {
  refreshAccessToken,
  needsRefresh,
  getValidAccessToken
};