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
      'Mail.Read', 'Mail.ReadWrite', 'Mail.Send', 'MailboxSettings.ReadWrite', 
      'User.Read', 'User.ReadWrite',
      'Calendars.Read', 'Calendars.ReadWrite',
      'Contacts.ReadWrite',
      'Files.Read', 'Files.ReadWrite',
      'Team.ReadBasic.All', 'Team.Create',
      'Chat.Read', 'Chat.ReadWrite',
      'ChannelMessage.Read.All', 'ChannelMessage.Send',
      'OnlineMeetingTranscript.Read.All',
      'OnlineMeetings.ReadWrite',
      'Tasks.Read', 'Tasks.ReadWrite'
    ],
    tokenStorePath: path.join(homeDir, '.office-mcp-tokens.json'),
    authServerUrl: 'http://localhost:3000'
  },
  
  // Microsoft Graph API
  GRAPH_API_ENDPOINT: 'https://graph.microsoft.com/v1.0/',
  
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
  SHAREPOINT_SYMLINK_PATH: process.env.SHAREPOINT_SYMLINK_PATH || path.join(homeDir, 'temp', 'sharepoint')
};
