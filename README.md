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
│   └── ...                  # Other folder operations
├── rules/                   # Email rules functionality
│   ├── index.js             # Rules exports
│   ├── list.js              # List rules
│   └── ...                  # Other rules operations
├── teams/                   # Teams functionality
│   ├── index.js             # Teams exports
│   ├── list.js              # List teams
│   ├── channels.js          # Channel operations
│   ├── chats.js             # Chat operations
│   ├── meetings.js          # Meeting operations
│   ├── insights.js          # AI insights and recording operations
│   └── transcripts.js       # Transcript operations
├── drive/                   # OneDrive/SharePoint functionality
│   ├── index.js             # Drive exports
│   ├── list.js              # List files/folders
│   ├── content.js           # File content operations
│   ├── upload.js            # File upload and folder operations
│   └── recyclebin.js        # Recycle bin management
├── planner/                 # Planner/Tasks functionality
│   ├── index.js             # Planner exports
│   ├── plans.js             # Plan operations
│   ├── tasks.js             # Task operations
│   └── buckets.js           # Bucket operations
├── users/                   # User management functionality
│   ├── index.js             # User exports
│   ├── profile.js           # User profile operations
│   ├── directory.js         # Directory search operations
│   └── presence.js          # Presence status operations
├── notifications/           # Change notifications functionality
│   ├── index.js             # Notifications exports
│   ├── subscribe.js         # Create subscriptions
│   └── manage.js            # Manage subscriptions
└── utils/                   # Utility functions
    ├── graph-api.js         # Microsoft Graph API helper
    ├── odata-helpers.js     # OData query building
    └── mock-data.js         # Test mode data
```

## Features

- **Authentication**: OAuth 2.0 authentication with Microsoft Graph API
- **Email Management**: List, search, read, and send emails
- **Calendar Management**: Create, read, update, and delete calendar events
- **Teams Integration**: Access Teams chats, channels, meetings, transcripts, and AI insights
- **OneDrive/SharePoint Access**: File and folder management, upload/download, recycle bin
- **Planner/Tasks**: Plan, task, and bucket management capabilities
- **User Management**: Profile access, presence status, directory search
- **Change Notifications**: Subscribe to resource changes via webhooks
- **Modular Structure**: Clean separation of concerns for better maintainability
- **OData Filter Handling**: Proper escaping and formatting of OData queries
- **Test Mode**: Simulated responses for testing without real API calls

## Azure App Registration & Configuration

To use this MCP server you need to first register and configure an app in Azure Portal. The following steps will take you through the process of registering a new app, configuring its permissions, and generating a client secret.

### App Registration

1. Open [Azure Portal](https://portal.azure.com/) in your browser
2. Sign in with a Microsoft Work or Personal account
3. Search for or click on "App registrations"
4. Click on "New registration"
5. Enter a name for the app, for example "Office MCP Server"
6. Select the "Accounts in any organizational directory and personal Microsoft accounts" option
7. In the "Redirect URI" section, select "Web" from the dropdown and enter "http://localhost:3333/auth/callback" in the textbox
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

## Configuration

To configure the server, edit the `config.js` file to change:

- Server name and version
- Test mode settings
- Authentication parameters
- Field selections for various services
- API endpoints

## Usage with Claude Desktop

1. Copy the sample configuration from `claude-config-sample.json` to your Claude Desktop configuration
2. Restart Claude Desktop
3. Authenticate with Microsoft using the `authenticate` tool
4. Use the available tools to interact with Microsoft 365 services

## Running Standalone

You can test the server using:

```bash
./test-modular-server.sh
```

This will use the MCP Inspector to directly connect to the server and let you test the available tools.

## Authentication Flow

1. Start the authentication server:
   - Windows: Run `start-auth-server.bat` or `run-office-mcp.bat`
   - Unix/Linux/macOS: Run `./start-auth-server.sh`
2. The auth server runs on port 3333 and handles OAuth callbacks
3. In Claude, use the `authenticate` tool to get an authentication URL
4. Complete the authentication in your browser
5. Tokens are stored in `~/.office-mcp-tokens.json`

### Quick Start

For Windows users, simply run:
```bash
run-office-mcp.bat
```

This will start both the authentication server and the MCP server in separate windows.

## Troubleshooting

- **Authentication Issues**: Check the token file and authentication server logs
- **OData Filter Errors**: Look for escape sequences in the server logs
- **API Call Failures**: Check for detailed error messages in the response

## Extending the Server

To add more functionality:

1. Create new module directories (e.g., `teams/`)
2. Implement tool handlers in separate files
3. Export tool definitions from module index files
4. Import and add tools to `TOOLS` array in `index.js`
