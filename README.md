# Office MCP Server

This is a comprehensive implementation of the Office MCP (Model Context Protocol) server that connects Claude with Microsoft 365 services through the Microsoft Graph API.

> **🚀 Headless Operation!** Run without browser authentication after initial setup. Automatic token refresh and Windows Task Scheduler support for invisible background operation. See [TASK_SCHEDULER_SETUP.md](TASK_SCHEDULER_SETUP.md) for Windows setup guide.

## Features

- **Complete Microsoft 365 Integration**: Email, Calendar, Teams, OneDrive/SharePoint, and Planner
- **Headless Operation**: Run without browser after initial authentication
- **Automatic Token Management**: Persistent token storage with automatic refresh
- **Email Attachment Handling**: Download embedded attachments and map SharePoint URLs to local paths
- **Advanced Email Search**: Unified search with KQL support and automatic query optimization
- **Teams Meeting Management**: Access transcripts, recordings, and AI insights
- **File Management**: Full OneDrive and SharePoint file operations
- **Task Management**: Complete Microsoft Planner integration
- **Configurable Paths**: Environment variables for all local sync paths

## Quick Start

### Prerequisites
- Node.js 16 or higher
- Microsoft 365 account (personal or work/school)
- Azure App Registration (see below)

### Installation

1. Clone the repository:
```bash
git clone https://github.com/yourusername/office-mcp.git
cd office-mcp
```

2. Install dependencies:
```bash
npm install
```

3. Copy the environment template:
```bash
cp .env.example .env
```

4. Configure your `.env` file with:
   - Azure App credentials (see Azure Setup below)
   - Local file paths for SharePoint/OneDrive sync
   - Optional settings

5. Run initial authentication:
```bash
npm run auth-server
# Visit http://localhost:3000/auth and sign in
```

6. Configure Claude Desktop (see Claude Desktop Configuration below)

## Core Capabilities

### Email Operations
- **Unified Search**: Single `email_search` tool with automatic optimization
- **Attachment Handling**: Download embedded attachments, map SharePoint URLs to local paths
- **Advanced Features**: Categories, rules, focused inbox, folder management
- **Batch Operations**: Move multiple emails efficiently

### Calendar Management
- **Full CRUD Operations**: Create, read, update, delete events
- **Teams Integration**: Create meetings with Teams links
- **Recurrence Support**: Complex recurring event patterns
- **UTC Time Handling**: Proper timezone management

### Teams Features
- **Meeting Management**: Create, update, cancel meetings
- **Transcript Access**: Retrieve meeting transcripts
- **Recording Access**: Access meeting recordings
- **Channel Operations**: Messages, members, tabs
- **Chat Management**: Create, send, manage chat messages

### File Management
- **SharePoint Integration**: Local sync path mapping
- **OneDrive Support**: Full file operations
- **Batch Operations**: Upload/download multiple files
- **Search**: Content and metadata search

### Task Management (Planner)
- **Plan Operations**: Create and manage plans
- **Task Assignment**: User lookup and assignment
- **Bucket Organization**: Group tasks efficiently
- **Bulk Operations**: Update/delete multiple tasks

## Azure App Registration & Configuration

To use this MCP server you need to first register and configure an app in Azure Portal. The following steps will take you through the process of registering a new app, configuring its permissions, and generating a client secret.

### App Registration

1. Open [Azure Portal](https://portal.azure.com/) in your browser
2. Sign in with a Microsoft Work or Personal account
3. Search for or click on "App registrations"
4. Click on "New registration"
5. Enter a name for the app, for example "Office MCP Server"
6. Select the "Accounts in any organizational directory and personal Microsoft accounts" option
7. In the "Redirect URI" section, select "Web" from the dropdown and enter "http://localhost:3000/auth/callback" in the textbox
8. Click on "Register"
9. From the Overview section of the app settings page, copy the "Application (client) ID" and enter it as the OFFICE_CLIENT_ID in the .env file as well as in the claude-config-sample.json file

### App Permissions

1. From the app settings page in Azure Portal select the "API permissions" option under the Manage section
2. Click on "Add a permission"
3. Click on "Microsoft Graph"
4. Select "Delegated permissions"
5. Search for and select the checkbox next to each of these permissions:
    - offline_access
    - User.Read
    - User.ReadWrite
    - User.ReadBasic.All
    - Mail.Read
    - Mail.ReadWrite
    - Mail.Send
    - Calendars.Read
    - Calendars.ReadWrite
    - Files.Read
    - Files.ReadWrite
    - Files.ReadWrite.All
    - Team.ReadBasic.All
    - Team.Create
    - Chat.Read
    - Chat.ReadWrite
    - ChannelMessage.Read.All
    - ChannelMessage.Send
    - OnlineMeetingTranscript.Read.All
    - OnlineMeetings.ReadWrite
    - Tasks.Read
    - Tasks.ReadWrite
    - Group.Read.All
    - Directory.Read.All
    - Presence.Read
    - Presence.ReadWrite
6. Click on "Add permissions"

### Client Secret

1. From the app settings page in Azure Portal select the "Certificates & secrets" option under the Manage section
2. Switch to the "Client secrets" tab
3. Click on "New client secret"
4. Enter a description, for example "Client Secret"
5. Select the longest possible expiration time
6. Click on "Add"
7. Copy the secret value and enter it as the OFFICE_CLIENT_SECRET in the .env file as well as in the claude-config-sample.json file

## Environment Configuration

### Required Variables
```bash
# Azure App Registration
OFFICE_CLIENT_ID=your-azure-app-client-id
OFFICE_CLIENT_SECRET=your-azure-app-client-secret
OFFICE_TENANT_ID=common

# Authentication
OFFICE_REDIRECT_URI=http://localhost:3000/auth/callback
```

### Optional Variables
```bash
# Local file paths (customize to your system)
SHAREPOINT_SYNC_PATH=/path/to/your/sharepoint/sync
ONEDRIVE_SYNC_PATH=/path/to/your/onedrive/sync
TEMP_ATTACHMENTS_PATH=/path/to/temp/attachments
SHAREPOINT_SYMLINK_PATH=/path/to/sharepoint/symlink

# Server settings
USE_TEST_MODE=false
TRANSPORT_TYPE=stdio  # or 'http' for headless
HTTP_PORT=3333
HTTP_HOST=127.0.0.1
```

## Claude Desktop Configuration

1. Locate your Claude Desktop configuration file:
   - Windows: `%APPDATA%\Claude\claude_desktop_config.json`
   - macOS: `~/Library/Application Support/Claude/claude_desktop_config.json`
   - Linux: `~/.config/Claude/claude_desktop_config.json`

2. Add the MCP server configuration:
```json
{
  "mcpServers": {
    "office-mcp": {
      "command": "node",
      "args": ["/path/to/office-mcp/index.js"],
      "env": {
        "OFFICE_CLIENT_ID": "your-client-id",
        "OFFICE_CLIENT_SECRET": "your-client-secret",
        "SHAREPOINT_SYNC_PATH": "/path/to/sharepoint",
        "ONEDRIVE_SYNC_PATH": "/path/to/onedrive"
      }
    }
  }
}
```

3. Restart Claude Desktop

4. In Claude, use the `authenticate` tool to connect to Microsoft 365

## Testing

### MCP Inspector
Test the server directly using the MCP Inspector:
```bash
npx @modelcontextprotocol/inspector node index.js
```

### Test Mode
Enable test mode to use mock data without API calls:
```bash
USE_TEST_MODE=true node index.js
```

## Authentication Flow

1. Start the authentication server:
   - Windows: Run `start-auth-server.bat` or `run-office-mcp.bat`
   - Unix/Linux/macOS: Run `./start-auth-server.sh`
2. The auth server runs on port 3000 and handles OAuth callbacks
3. In Claude, use the `authenticate` tool to get an authentication URL
4. Complete the authentication in your browser
5. Tokens are stored in `~/.office-mcp-tokens.json`

## Headless Operation

### Automatic Token Refresh
After initial authentication, the server automatically refreshes tokens without user interaction.

### HTTP Transport Mode
For headless environments, use HTTP transport:
```bash
TRANSPORT_TYPE=http HTTP_PORT=3333 node index.js
```

### Windows Service (Optional)
For Windows background operation:
1. Complete initial authentication
2. Configure as Windows Task Scheduler task
3. Runs invisibly at system startup

## Troubleshooting

### Common Issues

1. **Authentication Errors**
   - Ensure Azure App has correct permissions
   - Check token file exists: `~/.office-mcp-tokens.json`
   - Verify redirect URI matches Azure configuration

2. **Email Attachment Issues**
   - Configure local sync paths in `.env`
   - Ensure temp directory has write permissions
   - Check SharePoint sync is active

3. **API Rate Limits**
   - Server includes automatic retry with exponential backoff
   - Reduce request frequency if persistent

4. **Permission Errors**
   - Verify all required Graph API permissions are granted
   - Admin consent may be required for some permissions

## Security Considerations

- **Token Storage**: Tokens are encrypted and stored locally
- **Environment Variables**: Never commit `.env` files
- **Client Secrets**: Rotate regularly and use Azure Key Vault in production
- **Local Paths**: Use environment variables instead of hardcoding paths
- **Audit Logging**: All API calls are logged for security monitoring

## Contributing

Contributions are welcome! Please:
1. Fork the repository
2. Create a feature branch
3. Submit a pull request

## License

MIT License - See LICENSE file for details
