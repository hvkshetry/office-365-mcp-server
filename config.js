/**
 * Configuration for Office MCP Server
 */
const path = require('path');
const os = require('os');

// Ensure we have a home directory path even if process.env.HOME is undefined
const homeDir = process.env.HOME || process.env.USERPROFILE || os.homedir() || '/tmp';

module.exports = {
  // Server information
  SERVER_NAME: "office-mcp",
  SERVER_VERSION: "1.0.0",
  
  // Transport configuration
  TRANSPORT_TYPE: process.env.TRANSPORT_TYPE || 'stdio', // 'stdio' or 'http'
  HTTP_PORT: process.env.HTTP_PORT || 3333,
  HTTP_HOST: process.env.HTTP_HOST || '127.0.0.1',
  SERVICE_MODE: process.env.SERVICE_MODE === 'true',
  
  // Test mode setting
  USE_TEST_MODE: process.env.USE_TEST_MODE === 'true',
  
  // Authentication configuration
  AUTH_CONFIG: {
    clientId: process.env.OFFICE_CLIENT_ID || '',
    clientSecret: process.env.OFFICE_CLIENT_SECRET || '',
    redirectUri: 'http://localhost:3000/auth/callback',
    scopes: [
      'offline_access',  // CRITICAL: Required for refresh tokens
      'User.Read', 'User.ReadWrite', 'User.ReadBasic.All',
      'Mail.ReadWrite', 'Mail.Send', 'Mail.ReadWrite.Shared', 'Mail.Send.Shared',
      'MailboxSettings.ReadWrite',
      'Calendars.ReadWrite',
      'Contacts.ReadWrite',
      'Files.ReadWrite.All',
      'Team.ReadBasic.All', 'Team.Create',
      'Chat.ReadWrite',
      'ChannelMessage.Read.All', 'ChannelMessage.Send',
      'OnlineMeetingTranscript.Read.All',
      'OnlineMeetings.ReadWrite',
      'Tasks.ReadWrite',
      'Group.Read.All', 'Directory.Read.All',
      'Presence.ReadWrite',
      'Sites.Read.All'
    ],
    tokenStorePath: path.join(homeDir, '.office-mcp-tokens.json'),
    authServerUrl: 'http://localhost:3000'
  },
  
  // Microsoft Graph API
  GRAPH_API_ENDPOINT: 'https://graph.microsoft.com/v1.0/',

  // Helper function to get mailbox prefix for Graph API calls
  // With .Shared scopes, use users/{mailbox} for shared mailbox access
  getMailboxPrefix(mailbox) {
    if (mailbox && mailbox !== 'me') {
      // Validate mailbox looks like an email address or simple identifier
      if (/[\/\\]|\.\./.test(mailbox)) {
        throw new Error('Invalid mailbox: contains disallowed characters');
      }
      return `users/${mailbox}`;
    }
    return 'me';
  },

  
  // Email constants
  EMAIL_SELECT_FIELDS: 'id,subject,from,toRecipients,ccRecipients,receivedDateTime,bodyPreview,hasAttachments,importance,isRead',
  EMAIL_DETAIL_FIELDS: 'id,subject,from,toRecipients,ccRecipients,bccRecipients,receivedDateTime,bodyPreview,body,hasAttachments,importance,isRead,internetMessageHeaders',
  
  // Calendar constants
  CALENDAR_SELECT_FIELDS: 'id,subject,bodyPreview,start,end,location,organizer,attendees,isAllDay,isCancelled',
  
  // Teams constants
  TEAMS_SELECT_FIELDS: 'id,displayName,description,isArchived,visibility',
  TEAMS_CHANNELS_SELECT_FIELDS: 'id,displayName,description,membershipType',
  TEAMS_MESSAGES_SELECT_FIELDS: 'id,messageType,createdDateTime,lastModifiedDateTime,body,from,importance',
  
  // OneDrive constants
  DRIVE_FILES_SELECT_FIELDS: 'id,name,size,createdDateTime,lastModifiedDateTime,webUrl,folder,file',
  
  // Pagination
  DEFAULT_PAGE_SIZE: 25,
  MAX_RESULT_COUNT: 50,
  
  // Local file paths - MUST be configured via environment variables for your specific setup
  SHAREPOINT_SYNC_PATH: process.env.SHAREPOINT_SYNC_PATH || path.join(homeDir, 'SharePoint'),
  ONEDRIVE_SYNC_PATH: process.env.ONEDRIVE_SYNC_PATH || path.join(homeDir, 'OneDrive'),
  TEMP_ATTACHMENTS_PATH: process.env.TEMP_ATTACHMENTS_PATH || path.join(homeDir, 'temp', 'email-attachments'),
  SHAREPOINT_SYMLINK_PATH: process.env.SHAREPOINT_SYMLINK_PATH || path.join(homeDir, 'temp', 'sharepoint'),

  // Timezone configuration — all calendar/todo operations use Eastern Time
  DEFAULT_TIMEZONE: process.env.DEFAULT_TIMEZONE || 'America/New_York',

  // IANA → Microsoft timezone name mapping
  IANA_TO_MS_TIMEZONE: {
    'America/New_York': 'Eastern Standard Time',
    'America/Chicago': 'Central Standard Time',
    'America/Denver': 'Mountain Standard Time',
    'America/Los_Angeles': 'Pacific Standard Time',
    'America/Phoenix': 'US Mountain Standard Time',
    'UTC': 'UTC',
  },

  /**
   * Get the Microsoft timezone name for the configured IANA timezone.
   * @returns {string} Microsoft timezone name (e.g. "Eastern Standard Time")
   */
  getMsTimezone() {
    return this.IANA_TO_MS_TIMEZONE[this.DEFAULT_TIMEZONE] || 'Eastern Standard Time';
  },

  /**
   * Format a datetime string for display using wall-clock time.
   * Pure regex parsing — no `new Date()` to avoid UTC conversion surprises.
   * @param {string} dateTimeStr - ISO-ish datetime string from Graph API
   * @param {string} [timeZoneName] - Microsoft timezone name (e.g. "Eastern Standard Time")
   * @returns {string} Formatted datetime string (e.g. "Feb 24, 2026 3:00 PM ET")
   */
  formatDateTime(dateTimeStr, timeZoneName) {
    if (!dateTimeStr) return 'N/A';

    // Derive short TZ abbreviation from MS timezone name
    const tzAbbrevMap = {
      'Eastern Standard Time': 'ET',
      'Central Standard Time': 'CT',
      'Mountain Standard Time': 'MT',
      'Pacific Standard Time': 'PT',
      'US Mountain Standard Time': 'MT',
      'UTC': 'UTC',
    };
    const tzAbbrev = tzAbbrevMap[timeZoneName] || tzAbbrevMap[this.getMsTimezone()] || 'ET';

    // Parse wall-clock components via regex (handles "2026-02-24T15:00:00.0000000", etc.)
    const match = dateTimeStr.match(/^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2})/);
    if (!match) return dateTimeStr; // Unparseable — return as-is

    const [, year, month, day, hour, minute] = match;
    const months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
    const monthName = months[parseInt(month, 10) - 1] || month;
    const h = parseInt(hour, 10);
    const ampm = h >= 12 ? 'PM' : 'AM';
    const h12 = h === 0 ? 12 : h > 12 ? h - 12 : h;

    return `${monthName} ${parseInt(day, 10)}, ${year} ${h12}:${minute} ${ampm} ${tzAbbrev}`;
  }
};
