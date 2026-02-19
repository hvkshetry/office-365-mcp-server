/**
 * Authentication-related tools for the Office MCP server
 */
const config = require('../config');
const tokenManager = require('./token-manager');
const { safeTool } = require('../utils/errors');

/**
 * About tool handler
 * @returns {object} - MCP response
 */
async function handleAbout() {
  return {
    content: [{
      type: "text",
      text: `🖥️ Office MCP Server v${config.SERVER_VERSION} 🖥️\n\nProvides access to Microsoft 365 services including Outlook, Teams, OneDrive, and more through Microsoft Graph API.\nImplemented with a modular architecture for improved maintainability.`
    }]
  };
}

/**
 * Authentication tool handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleAuthenticate(args) {
  const force = args && args.force === true;
  
  // For test mode, create a test token
  if (config.USE_TEST_MODE) {
    // Create a test token with a 1-hour expiry
    tokenManager.createTestTokens();
    
    return {
      content: [{
        type: "text",
        text: 'Successfully authenticated with Microsoft Graph API (test mode)'
      }]
    };
  }
  
  // For real authentication, generate an auth URL and instruct the user to visit it
  const authUrl = `${config.AUTH_CONFIG.authServerUrl}/auth?client_id=${config.AUTH_CONFIG.clientId}`;
  
  return {
    content: [{
      type: "text",
      text: `Authentication required. Please visit the following URL to authenticate with Microsoft: ${authUrl}\n\nAfter authentication, you will be redirected back to this application.`
    }]
  };
}

/**
 * Check authentication status tool handler
 * @returns {object} - MCP response
 */
async function handleCheckAuthStatus() {
  console.error('[CHECK-AUTH-STATUS] Starting authentication status check');
  
  const tokens = tokenManager.loadTokenCache();
  
  console.error(`[CHECK-AUTH-STATUS] Tokens loaded: ${tokens ? 'YES' : 'NO'}`);
  
  if (!tokens || !tokens.access_token) {
    console.error('[CHECK-AUTH-STATUS] No valid access token found');
    return {
      content: [{ type: "text", text: "Not authenticated" }]
    };
  }
  
  console.error('[CHECK-AUTH-STATUS] Access token present');
  console.error(`[CHECK-AUTH-STATUS] Token expires at: ${tokens.expires_at}`);
  console.error(`[CHECK-AUTH-STATUS] Current time: ${Date.now()}`);
  
  // Check if token needs refresh
  const { needsRefresh } = require('./auto-refresh');
  if (needsRefresh(tokens)) {
    const timeLeft = Math.max(0, (tokens.expires_at - Date.now()) / 1000 / 60);
    return {
      content: [{ 
        type: "text", 
        text: `Authenticated - token expires in ${timeLeft.toFixed(1)} minutes (will auto-refresh)` 
      }]
    };
  }
  
  const timeLeft = (tokens.expires_at - Date.now()) / 1000 / 60;
  return {
    content: [{ 
      type: "text", 
      text: `Authenticated and ready - token valid for ${timeLeft.toFixed(1)} minutes` 
    }]
  };
}

/**
 * Unified system handler - single entry point for auth/system operations
 */
async function handleSystem(args) {
  const { operation, ...params } = args || {};

  if (!operation) {
    return {
      content: [{
        type: "text",
        text: "Missing required parameter: operation. Valid operations: about, authenticate, check_status"
      }]
    };
  }

  switch (operation) {
    case 'about':
      return await handleAbout();
    case 'authenticate':
      return await handleAuthenticate(params);
    case 'check_status':
      return await handleCheckAuthStatus();
    default:
      return {
        content: [{
          type: "text",
          text: `Invalid operation: ${operation}. Valid operations: about, authenticate, check_status`
        }]
      };
  }
}

// Tool definitions - consolidated from 3 tools to 1
const authTools = [
  {
    name: "system",
    description: "System operations: get server info, authenticate with Microsoft Graph API, or check authentication status",
    inputSchema: {
      type: "object",
      properties: {
        operation: {
          type: "string",
          enum: ["about", "authenticate", "check_status"],
          description: "Operation to perform: about (server info), authenticate (connect to MS Graph), check_status (verify auth)"
        },
        force: {
          type: "boolean",
          description: "Force re-authentication even if already authenticated (for authenticate operation)"
        }
      },
      required: ["operation"]
    },
    handler: safeTool('system', handleSystem)
  }
];

module.exports = {
  authTools,
  handleAbout,
  handleAuthenticate,
  handleCheckAuthStatus
};
