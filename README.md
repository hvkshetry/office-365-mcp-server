# Office MCP Server

This is a comprehensive implementation of the Office MCP (Model Context Protocol) server that connects Claude with Microsoft 365 services through the Microsoft Graph API.

## Full Tool Documentation
For detailed documentation of all available tools and usage examples, see [TOOLS_DOCUMENTATION.md](./TOOLS_DOCUMENTATION.md)

## Directory Structure

```
/office-mcp/
├── index.js                 # Main entry point
├── config.js                # Configuration settings
├── auth/                    # Authentication modules
│   ├── index.js             # Authentication exports
│   ├── token-manager.js     # Token storage and refresh
│   └── tools.js             # Auth-related tools
├── calendar/                # Calendar functionality
│   ├── index.js             # Calendar exports
│   ├── list.js              # List events
│   ├── create.js            # Create event
│   ├── update.js            # Update event
│   └── ...                  # Other calendar operations
├── email/                   # Email functionality
│   ├── index.js             # Email exports
│   ├── list.js              # List emails
│   ├── search.js            # Search emails
│   ├── read.js              # Read email
│   └── send.js              # Send email
├── folder/                  # Email folder functionality
│   ├── index.js             # Folder exports
│   ├── list.js              # List folders
│   ├── create.js            # Create folder
│   └── ...                  # Other folder operations
├── teams/                   # Teams functionality
│   ├── index.js             # Teams exports
│   ├── list.js              # List teams
│   ├── channels.js          # Channel operations
│   ├── messages.js          # Message operations
│   └── ...                  # Other Teams operations
├── planner/                 # Planner functionality
│   ├── index.js             # Planner exports
│   ├── plans.js             # Plan operations
│   ├── tasks.js             # Task operations
│   └── ...                  # Other Planner operations
├── notifications/           # Notification functionality
│   ├── index.js             # Notification exports
│   └── subscriptions.js     # Webhook subscriptions
├── utils/                   # Utility functions
│   ├── graph-api.js         # Microsoft Graph API helper
│   └── mock-data.js         # Test mode mock data
├── office-auth-server.js    # OAuth2 authentication server
├── config.js                # Configuration settings
├── package.json             # Dependencies
└── README.md                # This file
```

## Features

- **Email Operations**: List, search, read, send, and manage emails and folders
- **Calendar Management**: Create, update, delete, and search calendar events
- **Teams Integration**: Access teams, channels, messages, and meeting transcripts
- **Planner Support**: Manage plans, tasks, buckets, and assignments
- **Notifications**: Set up webhook subscriptions for real-time updates
- **Authentication**: Secure OAuth2 flow with token management
- **Test Mode**: Built-in mock data for development and testing

## Prerequisites

1. **Node.js**: Version 16.x or higher
2. **Azure App Registration**: Required for Microsoft Graph API access
3. **npm**: For dependency management

## Setup

### 1. Register an Azure Application

1. Go to the [Azure Portal](https://portal.azure.com)
2. Navigate to "Azure Active Directory" → "App registrations"
3. Click "New registration"
4. Configure your app:
   - **Name**: "Office MCP Server" (or your preferred name)
   - **Supported account types**: Choose based on your needs:
     - "Single tenant" for organization-only access
     - "Multitenant" for broader access
     - "Multitenant + Personal Microsoft accounts" for maximum compatibility
   - **Redirect URI**: Set to `http://localhost:3333/auth/callback` (Web platform)

### 2. Configure API Permissions

After registration, configure the following permissions:

1. Go to "API permissions" in your app registration
2. Click "Add a permission" → "Microsoft Graph" → "Delegated permissions"
3. Add these permissions:
   - **Email**: Mail.Read, Mail.ReadWrite, Mail.Send, MailboxSettings.ReadWrite
   - **Calendar**: Calendars.Read, Calendars.ReadWrite
   - **Files**: Files.Read, Files.ReadWrite
   - **Teams**: Team.ReadBasic.All, Team.Create, Chat.Read, Chat.ReadWrite, ChannelMessage.Read.All, ChannelMessage.Send
   - **Meetings**: OnlineMeetings.ReadWrite, OnlineMeetingTranscript.Read.All
   - **Planner**: Tasks.Read, Tasks.ReadWrite
   - **User**: User.Read, User.ReadWrite

### 3. Create Client Secret

1. Go to "Certificates & secrets"
2. Click "New client secret"
3. Add a description and choose expiry period
4. Copy the secret value immediately (it won't be shown again)

### 4. Install and Configure

```bash
# Clone the repository
git clone https://github.com/hvkshetry/office-365-mcp-server.git
cd office-mcp-server

# Install dependencies
npm install

# Copy environment template
cp .env.example .env

# Edit .env with your Azure app details:
# - OFFICE_CLIENT_ID: Your Application (client) ID
# - OFFICE_CLIENT_SECRET: Your client secret value
```

### 5. Start the Authentication Server

The authentication server handles the OAuth2 flow:

```bash
# Windows
start-auth-server.bat

# macOS/Linux
./start-auth-server.sh
```

### 6. Configure Claude Desktop

Add to your Claude Desktop configuration:

```json
{
  "mcpServers": {
    "office-mcp": {
      "command": "node",
      "args": ["path/to/office-mcp-server/index.js"],
      "env": {
        "OFFICE_CLIENT_ID": "your-client-id",
        "OFFICE_CLIENT_SECRET": "your-client-secret"
      }
    }
  }
}
```

## Authentication Flow

1. Start the authentication server (runs on port 3333)
2. Use the `authenticate` tool in Claude to initiate OAuth2 flow
3. A browser window opens for Microsoft login
4. After successful login, tokens are stored locally
5. Tokens auto-refresh when needed

## Testing

Enable test mode to use mock data without Microsoft 365 connection:

```bash
# Set in .env or environment
USE_TEST_MODE=true
```

Run tests:

```bash
npm test
```

## Troubleshooting

### Authentication Issues

- **"Application not found"**: Ensure client ID is correct
- **"Invalid client secret"**: Check secret hasn't expired
- **"Redirect URI mismatch"**: Verify `http://localhost:3333/auth/callback` is configured in Azure

### Permission Errors

- **"Insufficient privileges"**: Add required permissions in Azure portal
- **"Consent required"**: Admin consent may be needed for some permissions

### Token Issues

- **"Token expired"**: Tokens should auto-refresh; try re-authenticating
- **"Invalid token"**: Delete `~/.office-mcp-tokens.json` and re-authenticate

## Security Notes

- Store credentials securely in environment variables
- Never commit `.env` files to version control
- Tokens are stored locally in `~/.office-mcp-tokens.json`
- Use appropriate Azure AD permissions for your use case

## Contributing

Contributions are welcome! Please:

1. Fork the repository
2. Create a feature branch
3. Add tests for new functionality
4. Submit a pull request

## License

MIT License - see LICENSE file for details