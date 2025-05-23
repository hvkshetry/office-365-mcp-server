const express = require('express');
const crypto = require('crypto');
const fs = require('fs');
const path = require('path');
const os = require('os');
const https = require('https');
require('dotenv').config();

// MCP Server Auth Helper for Office MCP
// This server handles the OAuth2 redirect callback from Microsoft

const app = express();
const PORT = process.env.PORT || 3333;
const TOKEN_FILE = path.join(os.homedir(), '.office-mcp-tokens.json');

// Store auth codes temporarily in memory
const authCodes = new Map();

// Generate a random state parameter for security
function generateState() {
  return crypto.randomBytes(16).toString('hex');
}

// Root route with authentication instructions
app.get('/', (req, res) => {
  res.send(`
    <html>
      <head>
        <title>Office MCP Authentication Server</title>
        <style>
          body { font-family: Arial, sans-serif; max-width: 800px; margin: 0 auto; padding: 20px; }
          code { background: #f4f4f4; padding: 2px 4px; border-radius: 3px; }
          pre { background: #f4f4f4; padding: 10px; border-radius: 5px; overflow-x: auto; }
        </style>
      </head>
      <body>
        <h1>Office MCP Authentication Server</h1>
        <p>This server handles OAuth2 authentication for the Office MCP server.</p>
        
        <h2>Authentication Steps:</h2>
        <ol>
          <li>Use the <code>authenticate</code> tool in Claude to start the auth flow</li>
          <li>Visit the provided URL to sign in with Microsoft</li>
          <li>Grant the requested permissions</li>
          <li>You'll be redirected back here with the authorization code</li>
        </ol>
        
        <h2>Current Status:</h2>
        <p>Authentication server is running on port ${PORT}</p>
        <p>Redirect URI: <code>http://localhost:${PORT}/auth/callback</code></p>
        
        ${fs.existsSync(TOKEN_FILE) ? '<p>✅ Tokens file exists</p>' : '<p>❌ No tokens file found</p>'}
      </body>
    </html>
  `);
});

// Handle auth route - redirect to Microsoft's OAuth endpoint
app.get('/auth', (req, res) => {
  console.log('Auth request received, redirecting to Microsoft login...');
  
  // Load environment variables or use config
  const clientId = process.env.OFFICE_CLIENT_ID || '';
  const clientSecret = process.env.OFFICE_CLIENT_SECRET || '';
  const tenantId = process.env.OFFICE_TENANT_ID || 'common';
  
  // Verify credentials are set
  if (!clientId || !clientSecret) {
    res.send(`
      <html>
        <head>
          <title>Configuration Error</title>
          <style>
            body { font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto; padding: 20px; }
            h1 { color: #d9534f; }
            .error-box { background-color: #f8d7da; border: 1px solid #f5c6cb; padding: 15px; border-radius: 4px; }
            code { background: #f4f4f4; padding: 2px 4px; border-radius: 4px; }
          </style>
        </head>
        <body>
          <h1>Configuration Error</h1>
          <div class="error-box">
            <p>Microsoft Graph API credentials are not set. Please set the following environment variables:</p>
            <ul>
              <li><code>OFFICE_CLIENT_ID</code></li>
              <li><code>OFFICE_CLIENT_SECRET</code></li>
            </ul>
          </div>
        </body>
      </html>
    `);
    return;
  }
  
  // Get client_id from query parameters or use the default
  const requestedClientId = req.query.client_id || clientId;
  
  // Generate a secure state parameter
  const state = crypto.randomBytes(16).toString('hex');
  authCodes.set('state', state);
  
  // Build the authorization URL
  const authParams = new URLSearchParams({
    client_id: requestedClientId,
    response_type: 'code',
    redirect_uri: `http://localhost:${PORT}/auth/callback`,
    scope: 'offline_access User.Read User.ReadWrite User.ReadBasic.All Mail.Read Mail.ReadWrite Mail.Send Calendars.Read Calendars.ReadWrite Files.Read Files.ReadWrite Files.ReadWrite.All Team.ReadBasic.All Team.Create Chat.Read Chat.ReadWrite ChannelMessage.Read.All ChannelMessage.Send OnlineMeetingTranscript.Read.All OnlineMeetings.ReadWrite Tasks.Read Tasks.ReadWrite Group.Read.All Directory.Read.All Presence.Read Presence.ReadWrite',
    response_mode: 'query',
    state: state,
    tenant: tenantId
  });
  
  const authUrl = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/authorize?${authParams}`;
  console.log(`Redirecting to: ${authUrl}`);
  
  // Redirect to Microsoft's login page
  res.redirect(authUrl);
});

// Handle OAuth callback
app.get('/auth/callback', async (req, res) => {
  const { code, state, error, error_description } = req.query;
  
  // Handle errors
  if (error) {
    return res.send(`
      <html>
        <head>
          <title>Authentication Error</title>
          <style>
            body { font-family: Arial, sans-serif; max-width: 800px; margin: 0 auto; padding: 20px; }
            .error { color: red; background: #ffeeee; padding: 10px; border-radius: 5px; }
          </style>
        </head>
        <body>
          <h1>Authentication Error</h1>
          <div class="error">
            <p><strong>Error:</strong> ${error}</p>
            <p><strong>Description:</strong> ${error_description || 'No description provided'}</p>
          </div>
          <p><a href="/">Back to home</a></p>
        </body>
      </html>
    `);
  }
  
  // Validate state parameter
  const storedState = authCodes.get('state');
  if (state !== storedState) {
    return res.send(`
      <html>
        <head>
          <title>Invalid State</title>
          <style>
            body { font-family: Arial, sans-serif; max-width: 800px; margin: 0 auto; padding: 20px; }
            .error { color: red; background: #ffeeee; padding: 10px; border-radius: 5px; }
          </style>
        </head>
        <body>
          <h1>Invalid State Parameter</h1>
          <div class="error">
            <p>The state parameter doesn't match. This could be a security issue.</p>
          </div>
          <p><a href="/">Back to home</a></p>
        </body>
      </html>
    `);
  }
  
  // Exchange authorization code for tokens
  console.log('Authorization code received, exchanging for tokens...');
  
  exchangeCodeForTokens(code)
    .then((tokens) => {
      console.log('Token exchange successful');
      res.send(`
        <html>
          <head>
            <title>Authentication Successful</title>
            <style>
              body { font-family: Arial, sans-serif; max-width: 800px; margin: 0 auto; padding: 20px; }
              .success { color: green; background: #eeffee; padding: 10px; border-radius: 5px; }
              code { background: #f4f4f4; padding: 2px 4px; border-radius: 3px; }
            </style>
          </head>
          <body>
            <h1>Authentication Successful!</h1>
            <div class="success">
              <p>✅ You have successfully authenticated with Microsoft Graph API.</p>
              <p>✅ Access tokens have been saved securely.</p>
            </div>
            
            <p>You can now close this window and return to Claude.</p>
          </body>
        </html>
      `);
    })
    .catch((error) => {
      console.error(`Token exchange error: ${error.message}`);
      res.status(500).send(`
        <html>
          <head>
            <title>Authentication Error</title>
            <style>
              body { font-family: Arial, sans-serif; max-width: 800px; margin: 0 auto; padding: 20px; }
              .error { color: red; background: #ffeeee; padding: 10px; border-radius: 5px; }
            </style>
          </head>
          <body>
            <h1>Token Exchange Error</h1>
            <div class="error">
              <p>Failed to exchange authorization code for tokens.</p>
              <p>Error: ${error.message}</p>
            </div>
          </body>
        </html>
      `);
    });
});

// API endpoint to get the stored auth code
app.get('/api/auth-code', (req, res) => {
  const code = authCodes.get('code');
  if (code) {
    authCodes.delete('code'); // Clear after retrieving
    res.json({ code });
  } else {
    res.status(404).json({ error: 'No authorization code available' });
  }
});

// API endpoint to set the state parameter
app.post('/api/state', express.json(), (req, res) => {
  const { state } = req.body;
  authCodes.set('state', state);
  res.json({ status: 'ok' });
});

// API endpoint to check server status
app.get('/api/status', (req, res) => {
  res.json({
    status: 'running',
    port: PORT,
    tokenFileExists: fs.existsSync(TOKEN_FILE),
    hasPendingCode: authCodes.has('code')
  });
});

// Function to exchange authorization code for tokens
function exchangeCodeForTokens(code) {
  return new Promise((resolve, reject) => {
    const clientId = process.env.OFFICE_CLIENT_ID || '';
    const clientSecret = process.env.OFFICE_CLIENT_SECRET || '';
    const tenantId = process.env.OFFICE_TENANT_ID || 'common';
    
    const postData = new URLSearchParams({
      client_id: clientId,
      client_secret: clientSecret,
      code: code,
      redirect_uri: `http://localhost:${PORT}/auth/callback`,
      grant_type: 'authorization_code',
      scope: 'offline_access User.Read User.ReadWrite User.ReadBasic.All Mail.Read Mail.ReadWrite Mail.Send Calendars.Read Calendars.ReadWrite Files.Read Files.ReadWrite Files.ReadWrite.All Team.ReadBasic.All Team.Create Chat.Read Chat.ReadWrite ChannelMessage.Read.All ChannelMessage.Send OnlineMeetingTranscript.Read.All OnlineMeetings.ReadWrite Tasks.Read Tasks.ReadWrite Group.Read.All Directory.Read.All Presence.Read Presence.ReadWrite'
    }).toString();
    
    const options = {
      hostname: 'login.microsoftonline.com',
      path: `/${tenantId}/oauth2/v2.0/token`,
      method: 'POST',
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded',
        'Content-Length': Buffer.byteLength(postData)
      }
    };
    
    const req = https.request(options, (res) => {
      let data = '';
      
      res.on('data', (chunk) => {
        data += chunk;
      });
      
      res.on('end', () => {
        if (res.statusCode >= 200 && res.statusCode < 300) {
          try {
            const tokenResponse = JSON.parse(data);
            
            // Add email from the access token (decode JWT to get user info)
            try {
              const idToken = tokenResponse.id_token;
              if (idToken) {
                const payload = JSON.parse(Buffer.from(idToken.split('.')[1], 'base64').toString());
                tokenResponse.email = payload.preferred_username || payload.email || payload.upn;
              }
            } catch (e) {
              console.error('Error decoding ID token:', e);
            }
            
            // Calculate expiration time (current time + expires_in seconds)
            const expiresAt = Date.now() + (tokenResponse.expires_in * 1000);
            
            // Add expires_at for easier expiration checking
            tokenResponse.expires_at = expiresAt;
            
            // Save tokens to file
            fs.writeFileSync(TOKEN_FILE, JSON.stringify(tokenResponse, null, 2), 'utf8');
            console.log(`Tokens saved to ${TOKEN_FILE}`);
            
            resolve(tokenResponse);
          } catch (error) {
            reject(new Error(`Error parsing token response: ${error.message}`));
          }
        } else {
          reject(new Error(`Token exchange failed with status ${res.statusCode}: ${data}`));
        }
      });
    });
    
    req.on('error', (error) => {
      reject(error);
    });
    
    req.write(postData);
    req.end();
  });
}

// Start the server
app.listen(PORT, () => {
  console.log(`Office MCP Authentication Server running on http://localhost:${PORT}`);
  console.log(`Redirect URI: http://localhost:${PORT}/auth/callback`);
  console.log('');
  console.log('Use the "authenticate" tool in Claude to start the authentication flow.');
});

// Handle graceful shutdown
process.on('SIGINT', () => {
  console.log('\nShutting down authentication server...');
  process.exit(0);
});

process.on('SIGTERM', () => {
  console.log('\nShutting down authentication server...');
  process.exit(0);
});
