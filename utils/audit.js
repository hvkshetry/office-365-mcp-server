/**
 * Audit logger for Office MCP Server.
 * Records what actions the AI agent performs for accountability.
 * Logs to ~/.office-mcp-audit.log as newline-delimited JSON.
 */
const fs = require('fs');
const path = require('path');
const os = require('os');

const AUDIT_LOG_PATH = process.env.OFFICE_AUDIT_LOG_PATH ||
  path.join(os.homedir(), '.office-mcp-audit.log');

// Keys worth extracting per tool for concise audit entries
const AUDIT_KEYS = {
  mail: ['operation', 'to', 'subject', 'emailId', 'mailbox', 'folderId'],
  calendar: ['operation', 'subject', 'start', 'end', 'attendees'],
  files: ['operation', 'fileId', 'path', 'driveId', 'siteId'],
  teams_meeting: ['operation', 'subject'],
  teams_channel: ['operation', 'teamId', 'channelId'],
  teams_chat: ['operation', 'chatId'],
  contacts: ['operation', 'contactId', 'displayName'],
  planner: ['operation', 'entity', 'planId', 'taskId'],
  todo: ['operation', 'listId', 'taskId'],
  groups: ['operation', 'groupId'],
  directory: ['operation', 'userId'],
  notifications: ['operation'],
  search: ['operation', 'query'],
  system: ['operation'],
};

/**
 * Log an audit entry for a tool call.
 * @param {string} toolName - The MCP tool name
 * @param {object} args - The arguments passed to the tool
 */
function auditLog(toolName, args) {
  try {
    const keys = AUDIT_KEYS[toolName] || ['operation'];
    const params = {};
    for (const key of keys) {
      if (args[key] !== undefined) {
        params[key] = args[key];
      }
    }
    const entry = {
      ts: new Date().toISOString(),
      tool: toolName,
      ...params,
    };
    fs.appendFileSync(AUDIT_LOG_PATH, JSON.stringify(entry) + '\n');
  } catch (err) {
    console.error('[AUDIT] Failed to write audit log:', err.message);
  }
}

module.exports = { auditLog };
