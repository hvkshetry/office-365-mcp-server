# Office MCP Server

This is a comprehensive implementation of the Office MCP (Model Context Protocol) server that connects Claude with Microsoft 365 services through the Microsoft Graph API.

> **Headless Operation!** Run without browser authentication after initial setup. Automatic token refresh and Windows Task Scheduler support for invisible background operation. See [TASK_SCHEDULER_SETUP.md](TASK_SCHEDULER_SETUP.md) for Windows setup guide.


> **DEVELOPMENT STATUS: This project is under active development and is not yet production-ready. APIs, interfaces, and functionality may change without notice. Use at your own risk for evaluation and testing purposes only. Not recommended for production deployments.**

## Features

- **Complete Microsoft 365 Integration**: Email, Calendar, Teams, OneDrive/SharePoint, Contacts, Planner, To Do, Groups, and Directory
- **Consolidated Tool Architecture**: Operation-based routing reduces tool count for efficient LLM context usage
- **Headless Operation**: Run without browser after initial authentication
- **Automatic Token Management**: Persistent token storage with automatic refresh
- **Shared Mailbox Support**: Access shared mailboxes with `.Shared` scopes
- **Centralized Error Handling**: Consistent error formatting with actionable hints for Graph API errors
- **Email Attachment Handling**: Download embedded attachments and map SharePoint URLs to local paths
- **Advanced Email Search**: Unified search with KQL support and automatic query optimization
- **Teams Meeting Management**: Access transcripts, recordings, and AI insights
- **File Management**: Full OneDrive and SharePoint file operations
- **Contact Management**: Full CRUD operations for Outlook contacts with advanced search
- **Task Management**: Microsoft Planner and To Do integration
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

## Tool Architecture

The server uses a consolidated tool design where each Microsoft 365 domain is exposed as a single tool with `operation` (and optionally `entity`) routing. This minimizes LLM context overhead while providing full functionality.

| Tool | Domain | Operations |
|------|--------|------------|
| `system` | Auth & server info | `about`, `authenticate`, `check_status` |
| `mail` | Email | `list`, `read`, `send`, `reply`, `draft`, `search`, `move`, `folder`, `rules`, `categories`, `focused` |
| `calendar` | Calendar events | `list`, `get`, `create`, `update`, `delete`, `find_free_slots` |
| `teams_meeting` | Teams meetings | `create`, `update`, `cancel`, `find`, `list_transcripts`, `get_transcript`, `get_recordings` |
| `teams_channel` | Teams channels | `list`, `create`, `get`, `update`, `delete`, `send_message`, `list_messages`, `list_members` |
| `teams_chat` | Teams chat | `list`, `create`, `get`, `send_message`, `list_messages`, `list_members` |
| `files` | OneDrive/SharePoint | `list`, `get`, `search`, `upload`, `download`, `create_folder`, `delete`, `move`, `copy` |
| `search` | Unified search | Keyword-based search across mail, files, events |
| `contacts` | Outlook contacts | `list`, `get`, `create`, `update`, `delete`, `search`, `list_folders` |
| `planner` | Microsoft Planner | Entity+operation: `plan.list`, `task.create`, `bucket.get_tasks`, `user.lookup`, etc. |
| `todo` | Microsoft To Do | `list_lists`, `create_list`, `list_tasks`, `create_task`, `update_task`, `list_checklist`, etc. |
| `groups` | M365 Groups | `list`, `get`, `create`, `update`, `delete`, `list_members`, `add_member`, `remove_member` |
| `directory` | User directory | `lookup_user`, `get_profile`, `get_manager`, `get_reports`, `get_presence`, `search_users` |
| `notifications` | Webhooks | `create`, `list`, `renew`, `delete` |

### Error Handling

All tools are wrapped with `safeTool()` which provides:
- Consistent `isError: true` responses for failures
- Context-tagged messages (e.g., `[calendar.create] Error: ...`)
- Actionable hints for common Graph API errors (401, 403, 404, 429)

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
    - Mail.ReadWrite
    - Mail.Send
    - Mail.ReadWrite.Shared
    - Mail.Send.Shared
    - MailboxSettings.ReadWrite
    - Calendars.ReadWrite
    - Contacts.ReadWrite
    - Files.ReadWrite.All
    - Team.ReadBasic.All
    - Team.Create
    - Chat.ReadWrite
    - ChannelMessage.Read.All
    - ChannelMessage.Send
    - OnlineMeetingTranscript.Read.All
    - OnlineMeetings.ReadWrite
    - Tasks.ReadWrite
    - Group.Read.All
    - Directory.Read.All
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
TRANSPORT_TYPE=stdio  # or 'http' for SSE headless mode
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

4. In Claude, use the `system` tool with `operation: "authenticate"` to connect to Microsoft 365

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

### Unit Tests
```bash
npm test
```

## Authentication Flow

1. Start the authentication server:
   - Run `./start-auth-server.sh` (or use `npm run auth-server`)
2. The auth server runs on port 3000 and handles OAuth callbacks
3. In Claude, use the `system` tool with `operation: "authenticate"` to get an authentication URL
4. Complete the authentication in your browser
5. Tokens are stored in `~/.office-mcp-tokens.json`

## Headless Operation

### Automatic Token Refresh
After initial authentication, the server automatically refreshes tokens without user interaction.

### HTTP/SSE Transport Mode
For headless environments, use SSE transport:
```bash
TRANSPORT_TYPE=http HTTP_PORT=3333 node index.js
```

The server exposes `/sse` (GET) for SSE connections and `/message` (POST) for client messages.

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

2. **Email Search with Date Filters**
   - Date-filtered searches now route directly to $filter API for reliability
   - Use wildcard `*` for all emails in a date range
   - Both `startDate` and `endDate` support ISO format (2025-08-27) or relative (7d/1w/1m/1y)

3. **Email Attachment Issues**
   - Configure local sync paths in `.env`
   - Ensure temp directory has write permissions
   - Check SharePoint sync is active

4. **API Rate Limits**
   - Server includes automatic retry with exponential backoff
   - Reduce request frequency if persistent

5. **Permission Errors**
   - Verify all required Graph API permissions are granted
   - Admin consent may be required for some permissions

## Security Considerations

- **Secure Token Storage**: Tokens are stored with restricted file permissions (0o600) using atomic writes to prevent corruption
- **No Credential Logging**: Token content is never logged; only boolean presence checks are used
- **Sensitive Data Redaction**: Email bodies, recipients, and search queries are not logged; API URLs with query params are gated behind DEBUG_VERBOSE
- **Environment Variables**: Never commit `.env` files
- **Client Secrets**: Rotate regularly and use Azure Key Vault in production
- **Local Paths**: Use environment variables instead of hardcoding paths
- **Graceful Shutdown**: Proper SIGTERM/SIGINT handling for clean process termination

## Contributing

Contributions are welcome! Please:
1. Fork the repository
2. Create a feature branch
3. Submit a pull request

## License

MIT License - See LICENSE file for details
