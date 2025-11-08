# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Office MCP Server is a Model Context Protocol (MCP) server that provides Claude with access to Microsoft 365 services through the Microsoft Graph API. The server supports headless operation with automatic token refresh after initial authentication.

## Development Commands

### Build and Run
```bash
# Start server (stdio transport)
npm start

# Start auth server for initial OAuth flow
npm run auth-server

# Test mode with mock data
npm run test-mode

# MCP Inspector for testing
npm run inspect
```

### Testing
```bash
# Run all tests
npm test

# Run tests in watch mode
npm test:watch

# Run with coverage
npm test:coverage

# Run specific test suites
npm run test:auth
npm run test:teams
npm run test:drive
npm run test:planner
npm run test:users
npm run test:notifications
npm run test:contacts
npm run test:newsletter
```

## Architecture

### Module-Based Design

The server is organized into isolated modules in `/[module]/index.js`, each exporting a `[module]Tools` array. Modules are:

- **auth**: OAuth 2.0 flow, token management with auto-refresh
- **email**: Unified email search with KQL support, attachment handling, SharePoint URL mapping, contact extraction with newsletter filtering
- **calendar**: Full CRUD operations, Teams meeting integration, recurrence patterns
- **teams**: Consolidated tools for meetings, channels, and chats
- **contacts**: Full contact management with advanced search
- **planner**: Microsoft Planner task management
- **files**: OneDrive/SharePoint operations with local path mapping
- **search**: Unified Microsoft 365 search capabilities
- **notifications**: Graph API change notifications/webhooks
- **utils**: Shared utilities (Graph API client, batch operations, mock data, contact extraction, newsletter detection)

### Tool Registration Pattern

All modules follow this pattern:

```javascript
// /[module]/index.js
const { ensureAuthenticated } = require('../auth');
const { callGraphAPI } = require('../utils/graph-api');
const config = require('../config');

const [module]Tools = [
  {
    name: 'tool_name',
    description: 'Tool description',
    inputSchema: { /* JSON schema */ },
    handler: async (args) => {
      const accessToken = await ensureAuthenticated();
      // Implementation
    }
  }
];

module.exports = { [module]Tools };
```

Main entry point (`index.js`) imports all tools and registers them with the MCP Server SDK.

### Authentication Flow

1. **Initial auth**: User runs `office-auth-server.js` on port 3000, completes OAuth in browser
2. **Token storage**: Tokens saved to `~/.office-mcp-tokens.json` (configurable via `AUTH_CONFIG.tokenStorePath`)
3. **Auto-refresh**: `auth/auto-refresh.js` automatically refreshes tokens before expiry
4. **Headless operation**: After initial auth, server runs without browser interaction

All tools use `ensureAuthenticated()` which transparently handles token refresh.

### Graph API Client

`utils/graph-api.js` provides `callGraphAPI()` with:
- Automatic retry with exponential backoff for rate limits (429) and service errors (503, 504)
- Enhanced error messages with actionable suggestions
- Test mode support routing to `utils/mock-data.js`
- Proper OData filter encoding for Graph API

### Local Path Mapping

Email and file modules map SharePoint/OneDrive URLs to local sync paths:
- `SHAREPOINT_SYNC_PATH`: Local SharePoint sync folder
- `ONEDRIVE_SYNC_PATH`: Local OneDrive sync folder
- `TEMP_ATTACHMENTS_PATH`: Temp directory for email attachments

Pattern: `convertSharePointUrlToLocal()` in `email/index.js` handles URL parsing and path construction.

### Configuration

`config.js` centralizes all configuration:
- Environment variables with sensible defaults
- Graph API select field patterns for each module
- Retry/pagination settings
- Local path configuration

All modules import and reference `config` for consistency.

### Teams Module Structure

Teams has a consolidated structure (`teams/consolidated/`) with operation-based tools:
- `teams_meeting.js`: Complete meeting management (create, update, cancel, transcripts, recordings)
- `teams_channel.js`: Channel operations (messages, members, tabs)
- `teams_chat.js`: Chat management (create, send, members)

This reduces tool count from many individual tools to 3 consolidated, operation-based tools.

### Email Search Architecture

Email module uses a unified `email_search` tool that automatically routes queries:
- **KQL route**: Advanced queries with operators → `/me/messages?$search="..."`
- **Filter route**: Date-filtered queries → `/me/messages?$filter=receivedDateTime ge ...`
- **Smart fallback**: Switches routes based on query patterns and date requirements

### Contact Extraction and Newsletter Detection

The email module includes sophisticated contact extraction with intelligent newsletter filtering.

**Contact Extraction** (`tools/email-contact-extractor.js`):
- Extracts contacts from email metadata (From, To, CC) and body content
- Parses signatures, phone numbers, LinkedIn URLs, company names, job titles
- Supports bilingual extraction (French and English patterns)
- Cross-references with Outlook contacts to identify new contacts
- Deduplicates contacts with confidence scoring
- Exports to CSV with full contact information

**Utility Modules**:
- `utils/contact-parser.js`: Email, phone, LinkedIn extraction; name parsing; signature detection
- `utils/html-processor.js`: HTML stripping, signature extraction, body processing
- `utils/deduplicator.js`: Contact merging, deduplication, confidence scoring
- `utils/csv-generator.js`: CSV generation with proper escaping
- `utils/newsletter-detector.js`: Intelligent newsletter/marketing email detection

**Newsletter Detection** (`utils/newsletter-detector.js`):
- Multi-factor scoring system analyzing 14+ signals
- Bilingual pattern detection (French and English)
- Configurable via `/config/newsletter-rules.json` (whitelist, blacklist, custom patterns)
- Performance optimized with LRU caching (max 1000 entries)
- Default threshold: 60/100 (configurable)

**Detection Signals**:
- Headers: List-Unsubscribe, Precedence: bulk, ESP markers (Mailchimp, Sendgrid)
- Sender patterns: noreply@, newsletter@, ne-pas-repondre@, marketing@, bulletin@
- Body content: Unsubscribe links, "view in browser", preferences management
- Structure: High image/text ratio, tracking pixels, table-based layouts
- Recipients: BCC recipients, generic names ("Cher Client", "Valued Customer")

**Configuration** (`config/newsletter-rules.json`):
- Whitelist domains/senders (never filtered)
- Blacklist domains/senders (always filtered)
- Custom regex patterns for sender addresses and subject lines
- Configurable threshold and cache settings

**Integration Pattern**:
```javascript
const { detectNewsletter, applyWhitelistBlacklist } = require('../utils/newsletter-detector');

// Filter newsletters before processing
for (const email of emails) {
  const detection = await detectNewsletter(email, threshold);
  const finalDetection = applyWhitelistBlacklist(email, detection, rules);

  if (!finalDetection.isNewsletter) {
    // Process email for contact extraction
  }
}
```

## Testing Strategy

### Test Mode
Set `USE_TEST_MODE=true` to use mock data from `utils/mock-data.js` without real API calls. All Graph API calls route to `simulateGraphAPIResponse()`.

### Test Structure
- Tests in `/tests/*.test.js`
- Use `tests/test-utils.js` for shared test helpers
- Mock token manager and Graph API in unit tests
- Integration tests verify end-to-end tool execution

**Key Test Files**:
- `tests/contact-extraction.test.js`: Contact parsing, HTML processing, deduplication, CSV generation (42 tests)
- `tests/newsletter-detector.test.js`: Newsletter detection patterns, bilingual support, whitelist/blacklist (43 tests)
- Other module tests: auth, teams, drive, planner, users, notifications

### Running Specific Tests
Each module has a dedicated test script (`npm run test:[module]`) for focused testing during development.

### French Localization Testing
The contact extraction and newsletter detection modules include comprehensive French language support:
- French signature markers: Cordialement, Bien cordialement, Salutations, Amitiés
- French phone formats: +33, 01-09 (landline), 06-07 (mobile)
- French company types: SA, SARL, SAS, SASU, SNC
- French job titles: PDG, DG, Directeur, Responsable, Ingénieur
- French newsletter patterns: ne-pas-repondre@, désabonnement, bulletin@

All bilingual patterns are tested in both French and English.

## Important Patterns

### Error Handling
Always wrap Graph API calls in try-catch and provide user-friendly error messages. The Graph API client already enhances errors, but tools should add context:

```javascript
try {
  const result = await callGraphAPI(accessToken, 'GET', path);
  return { content: [{ type: "text", text: formatResult(result) }] };
} catch (error) {
  return {
    content: [{ type: "text", text: `Failed to fetch data: ${error.message}` }],
    isError: true
  };
}
```

### Batch Operations
Use `utils/batch.js` for operations on multiple items (e.g., deleting multiple emails, updating multiple tasks). Batch requests reduce API calls and improve performance.

### Date Handling
Graph API requires ISO 8601 format. Email and calendar modules support both:
- Absolute: `YYYY-MM-DD` (converted to ISO 8601)
- Relative: `7d`, `1w`, `1m`, `1y` (converted to dates, then ISO 8601)

### Tool Response Format
All tools return MCP-compliant responses:
```javascript
{
  content: [
    { type: "text", text: "Result message" }
  ],
  isError: false  // Set to true for errors
}
```

## Common Development Tasks

### Adding a New Tool
1. Add tool definition to appropriate module's tools array
2. Implement handler following the pattern above
3. Add tests in `/tests/[module].test.js`
4. Update tool count validation in tests if needed

### Adding a New Module
1. Create `/[module]/index.js` with tools array export
2. Import in main `index.js` and add to TOOLS array
3. Create test file `/tests/[module].test.js`
4. Add test script to `package.json`

### Debugging Authentication
- Check token file exists: `~/.office-mcp-tokens.json`
- Verify Azure app permissions match `config.js` scopes
- Use `npm run auth-server` to re-authenticate
- Check `office-auth-server.js` logs for OAuth callback errors

### Adding SharePoint Path Mapping
The `convertSharePointUrlToLocal()` function handles both organizational SharePoint sites and personal OneDrive. When adding new path patterns:
1. Test with real URLs from target environment
2. Handle URL decoding for special characters
3. Consider both Windows and Unix path formats

### Extending Contact Extraction
To add new extraction patterns or enhance existing ones:
1. Update pattern regexes in `utils/contact-parser.js` for new formats
2. Add corresponding tests in `tests/contact-extraction.test.js`
3. For French patterns, ensure both accented and non-accented variants are matched
4. Update confidence scoring in `utils/deduplicator.js` if adding new high-value fields

### Customizing Newsletter Detection
To adjust newsletter detection behavior:
1. Modify signal weights in `NEWSLETTER_SIGNALS` object in `utils/newsletter-detector.js`
2. Add new detection patterns to appropriate check functions (checkHeaders, checkSenderPatterns, etc.)
3. Update `/config/newsletter-rules.json` for domain-specific whitelist/blacklist
4. Add tests for new patterns in `tests/newsletter-detector.test.js`
5. Adjust default threshold in config or per-call if needed

### Adding Bilingual Support for Other Languages
The contact extraction and newsletter detection modules use a bilingual pattern approach:
1. Add new language patterns alongside existing French/English patterns
2. Use non-capturing groups for language alternatives: `(?:English|French|NewLang)`
3. Test all language variants independently
4. Document supported languages in README.md and tool descriptions
