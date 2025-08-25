# Headless Setup Guide for Office MCP Server

This guide explains how to set up the Office MCP server for completely headless operation without browser popups or terminal windows.

## Overview

The headless setup provides:
- **No browser authentication after initial setup** - Automatic token refresh
- **No terminal windows** - Runs as a Windows service
- **Auto-start on boot** - Service starts with Windows
- **Claude Code integration** - Connects via HTTP transport

## Prerequisites

1. Node.js installed (v14 or higher)
2. Office MCP server cloned and dependencies installed
3. Azure App Registration completed (see main README)
4. Administrator privileges (for service installation)

## Setup Steps

### Step 1: Initial One-Time Authentication

You need to authenticate **once** to get the initial tokens. After this, the server will automatically refresh tokens.

1. Start the authentication server:
   ```bash
   npm run auth-server
   ```

2. In a new terminal, trigger authentication:
   ```bash
   # Use the stdio server temporarily
   npm start
   
   # In Claude Code, use the authenticate tool
   # Or visit: http://localhost:3000/auth
   ```

3. Complete the browser authentication flow once

4. Verify tokens are saved:
   ```bash
   # Windows
   dir %USERPROFILE%\.office-mcp-tokens.json
   
   # WSL/Linux
   ls -la ~/.office-mcp-tokens.json
   ```

5. Stop the auth server (Ctrl+C)

### Step 2: Configure Environment

1. Copy `.env.example` to `.env`:
   ```bash
   cp .env.example .env
   ```

2. Edit `.env` and set:
   ```env
   # Required
   OFFICE_CLIENT_ID=your-client-id
   OFFICE_CLIENT_SECRET=your-client-secret
   OFFICE_TENANT_ID=your-tenant-id-or-common
   
   # For headless operation
   TRANSPORT_TYPE=http
   HTTP_PORT=3333
   HTTP_HOST=127.0.0.1
   SERVICE_MODE=true
   ```

### Step 3: Test HTTP Server

Before installing as a service, test the HTTP server:

```bash
npm run start:http
```

Check that it's running:
```bash
# In a browser or curl
curl http://127.0.0.1:3333/health
```

You should see a JSON response with server status.

### Step 4: Install as Windows Service

**Run PowerShell or Command Prompt as Administrator:**

```powershell
# Navigate to the project directory
cd C:\path\to\office-mcp

# Install the service
node scripts\install-service.js
```

The service will:
- Install successfully
- Start automatically
- Show the Claude Code configuration

### Step 5: Configure Claude Code

Add to your Claude Code configuration:

#### Option A: Using Claude CLI
```bash
claude mcp add --transport http office365 http://127.0.0.1:3333/mcp
```

#### Option B: Manual Configuration
Edit your Claude Code settings file and add:

```json
{
  "mcpServers": {
    "office365": {
      "type": "http",
      "url": "http://127.0.0.1:3333/mcp"
    }
  }
}
```

### Step 6: Verify Setup

1. Check service status:
   ```powershell
   # In PowerShell
   Get-Service "Office MCP Server"
   ```

2. Check health endpoint:
   ```bash
   curl http://127.0.0.1:3333/health
   ```

3. Test in Claude Code:
   - Open Claude Code
   - The Office MCP tools should be available automatically
   - Try: "Check my Office 365 authentication status"

## Service Management

### Start Service
```bash
node scripts/start-service.js
```

### Stop Service
```bash
node scripts/stop-service.js
```

### Restart Service
```bash
node scripts/stop-service.js
node scripts/start-service.js
```

### Uninstall Service
```bash
node scripts/uninstall-service.js
```

### View Service Logs
Windows services log to the Event Viewer:
1. Open Event Viewer (eventvwr.msc)
2. Navigate to: Windows Logs > Application
3. Filter by Source: "Office MCP Server"

## Token Refresh Behavior

- **Access tokens** expire after 1 hour
- **Refresh tokens** typically last 90 days
- The server automatically refreshes tokens 5 minutes before expiry
- As long as the server runs at least once every 90 days, tokens stay valid
- Each refresh extends the 90-day window

## Troubleshooting

### Service Won't Start
1. Check Event Viewer for errors
2. Verify `.env` file has correct credentials
3. Ensure token file exists: `%USERPROFILE%\.office-mcp-tokens.json`
4. Run manually to see errors: `npm run start:http`

### Authentication Fails
1. Token may have expired (>90 days unused)
2. Re-run initial authentication (Step 1)
3. Check credentials in `.env`

### Claude Code Can't Connect
1. Verify service is running: `Get-Service "Office MCP Server"`
2. Check health: `curl http://127.0.0.1:3333/health`
3. Ensure firewall isn't blocking port 3333
4. Check Claude Code logs for connection errors

### Token Refresh Fails
1. Check network connectivity to Microsoft
2. Verify client secret hasn't expired
3. Ensure refresh token hasn't been revoked
4. Check service logs in Event Viewer

## Security Considerations

1. **Token File Security**
   - Stored in user profile directory
   - Only accessible by the user account
   - Consider encrypting the drive

2. **Service Account**
   - Runs under SYSTEM account by default
   - Consider using a dedicated service account

3. **Network Security**
   - Only listens on localhost (127.0.0.1)
   - Not accessible from network
   - Port 3333 is local only

4. **Credential Storage**
   - Client secret in environment variables
   - Consider using Windows Credential Manager
   - Or Azure Key Vault for production

## Backup and Recovery

### Backup Token File
```powershell
# Create backup
copy %USERPROFILE%\.office-mcp-tokens.json %USERPROFILE%\.office-mcp-tokens.backup.json
```

### Restore Token File
```powershell
# Restore from backup
copy %USERPROFILE%\.office-mcp-tokens.backup.json %USERPROFILE%\.office-mcp-tokens.json
```

## Advanced Configuration

### Custom Port
Edit `.env`:
```env
HTTP_PORT=8080
```

Then reinstall service:
```bash
node scripts/uninstall-service.js
node scripts/install-service.js
```

### Multiple Instances
For multiple Office 365 tenants, run multiple instances on different ports:
1. Clone to different directories
2. Use different ports in `.env`
3. Install with different service names

## Monitoring

### Health Check Script
Create a simple monitoring script:
```powershell
# check-health.ps1
$response = Invoke-RestMethod -Uri "http://127.0.0.1:3333/health"
if ($response.status -eq "healthy") {
    Write-Host "✅ Service is healthy"
} else {
    Write-Host "❌ Service is unhealthy"
}
```

### Scheduled Health Check
Use Task Scheduler to run health checks periodically and alert on failures.

## Migration from Batch Files

If you were using the old batch file approach:

1. Stop any running batch file terminals
2. Follow this guide from Step 1
3. Delete old batch files (optional):
   - `start-auth-server.bat`
   - `start-office-mcp.bat`
   - `run-office-mcp.bat`

## Support

For issues or questions:
1. Check the [main README](README.md)
2. Review service logs in Event Viewer
3. Enable debug logging by setting `NODE_ENV=development`
4. Create an issue on GitHub with logs attached