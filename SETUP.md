# Office MCP Server Setup Guide

This guide will walk you through setting up the Office MCP Server with Claude Desktop.

## Prerequisites

- Node.js installed on your system
- A Microsoft account (personal or work/school)
- Access to Azure Portal for app registration

## Step 1: Azure App Registration

1. Go to [Azure Portal](https://portal.azure.com/)
2. Navigate to "App registrations"
3. Click "New registration"
4. Configure your app:
   - Name: `Office MCP Server`
   - Supported account types: "Accounts in any organizational directory and personal Microsoft accounts"
   - Redirect URI: Select "Web" and enter `http://localhost:3333/auth/callback`
5. Click "Register"

## Step 2: Configure App Permissions

1. In your app registration, go to "API permissions"
2. Click "Add a permission" > "Microsoft Graph" > "Delegated permissions"
3. Add all the permissions listed in the README.md file
4. Click "Grant admin consent" if you have admin privileges

## Step 3: Create Client Secret

1. Go to "Certificates & secrets"
2. Click "New client secret"
3. Add a description and select expiration time
4. Copy the secret value immediately (it won't be shown again)

## Step 4: Configure the MCP Server

### Option 1: Using Environment Variables

1. Copy `.env.example` to `.env`:
   ```bash
   cp .env.example .env
   ```

2. Edit `.env` with your values:
   ```
   OFFICE_CLIENT_ID=your-application-client-id
   OFFICE_CLIENT_SECRET=your-client-secret
   OFFICE_TENANT_ID=common
   OFFICE_REDIRECT_URI=http://localhost:3333/auth/callback
   ```

### Option 2: Using Claude Desktop Configuration

1. Copy the configuration from `claude_desktop_config.json`

2. Add it to your Claude Desktop configuration file:
   - macOS: `~/Library/Application Support/Claude/claude_desktop_config.json`
   - Windows: `%APPDATA%\Claude\claude_desktop_config.json`
   - Linux: `~/.config/Claude/claude_desktop_config.json`

3. Update the values:
   - Replace `your-azure-app-client-id` with your app's client ID
   - Replace `your-azure-app-client-secret` with your client secret
   - Update `/path/to/office-mcp/index.js` with the actual path to the server

Example configuration:
```json
{
  "mcpServers": {
    "office": {
      "command": "node",
      "args": [
        "/Users/yourname/office-mcp/index.js"
      ],
      "env": {
        "OFFICE_CLIENT_ID": "12345678-1234-1234-1234-123456789012",
        "OFFICE_CLIENT_SECRET": "your-secret-here",
        "OFFICE_TENANT_ID": "common",
        "OFFICE_REDIRECT_URI": "http://localhost:3333/auth/callback"
      }
    }
  }
}
```

## Step 5: Install Dependencies

```bash
cd office-mcp
npm install
```

## Step 6: Start the Authentication Server

The Office MCP server requires an authentication server to handle OAuth callbacks:

### Windows:
```bash
# Start both servers automatically
run-office-mcp.bat

# Or start just the auth server
start-auth-server.bat
```

### Unix/Linux/macOS:
```bash
# Start the auth server
./start-auth-server.sh
```

The authentication server will run on `http://localhost:3333`

## Step 7: Test the Server

1. With the auth server running, start the MCP server:
   ```bash
   npm start
   ```

2. Or use the test script:
   ```bash
   ./test-modular-server.sh
   ```

## Step 8: Restart Claude Desktop

After updating the configuration, restart Claude Desktop for the changes to take effect.

## Step 9: Authenticate

1. Make sure the authentication server is running (see Step 6)
2. In Claude, use the authenticate tool:
   ```
   Tool: authenticate
   ```
3. Visit the provided URL in your browser
4. Sign in with your Microsoft account
5. Grant the requested permissions
6. You'll be redirected to the local auth server

## Troubleshooting

### Authentication Issues
- Check that your redirect URI matches exactly: `http://localhost:3333/auth/callback`
- Ensure your client secret is correct and not expired
- Verify all required permissions are granted

### Connection Issues
- Make sure the path to `index.js` is absolute, not relative
- Check that Node.js is in your system PATH
- Look for error messages in Claude Desktop's logs

### Token Issues
- Tokens are stored in `~/.office-mcp-tokens.json`
- Delete this file to force re-authentication
- Check file permissions if token storage fails

## Security Notes

- Never commit your `.env` file to version control
- Keep your client secret secure
- Regularly rotate your client secrets
- Use the principle of least privilege for API permissions

## Next Steps

1. Review the [TOOLS_DOCUMENTATION.md](./TOOLS_DOCUMENTATION.md) for available tools
2. Start using Office MCP tools in Claude
3. Report issues on the project's GitHub page
