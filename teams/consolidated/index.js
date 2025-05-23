/**
 * Consolidated Teams Tools Export
 * 
 * This module exports the consolidated Teams tools:
 * - teams_meeting: All meeting operations
 * - teams_channel: All channel operations
 * - teams_chat: All chat operations
 * 
 * Each tool is operation-based, providing a unified interface
 * for all Teams functionality.
 */

const handleTeamsMeeting = require('./teams_meeting');
const handleTeamsChannel = require('./teams_channel');
const handleTeamsChat = require('./teams_chat');

// Define the tool schemas
const meetingToolSchema = {
  type: 'object',
  required: ['operation'],
  properties: {
    operation: {
      type: 'string',
      description: 'The operation to perform',
      enum: [
        'create', 'update', 'cancel', 'get', 'find_by_url',
        'list_transcripts', 'get_transcript', 'list_recordings', 
        'get_recording', 'get_participants', 'get_insights'
      ]
    }
  }
};

const channelToolSchema = {
  type: 'object',
  required: ['operation'],
  properties: {
    operation: {
      type: 'string',
      description: 'The operation to perform',
      enum: [
        'list', 'create', 'get', 'update', 'delete',
        'list_messages', 'get_message', 'create_message', 'reply_to_message',
        'list_members', 'add_member', 'remove_member', 'list_tabs'
      ]
    }
  }
};

const chatToolSchema = {
  type: 'object',
  required: ['operation'],
  properties: {
    operation: {
      type: 'string',
      description: 'The operation to perform',
      enum: [
        'list', 'create', 'get', 'update', 'delete',
        'list_messages', 'get_message', 'send_message', 'update_message', 'delete_message',
        'list_members', 'add_member', 'remove_member'
      ]
    }
  }
};

// Export the tools
module.exports = [
  {
    name: 'teams_meeting',
    description: 'Teams meeting operations: create, update, cancel, find, list transcripts, get recordings, and more',
    inputSchema: meetingToolSchema,
    handler: handleTeamsMeeting
  },
  {
    name: 'teams_channel',
    description: 'Teams channel operations: list, create, get, update, delete channels and manage messages, members, and tabs',
    inputSchema: channelToolSchema,
    handler: handleTeamsChannel
  },
  {
    name: 'teams_chat',
    description: 'Teams chat operations: list, create, get, update, delete chats and manage messages and members',
    inputSchema: chatToolSchema,
    handler: handleTeamsChat
  }
];