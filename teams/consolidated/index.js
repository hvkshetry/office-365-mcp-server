/**
 * Consolidated Teams Tools
 * 
 * This module consolidates all Teams functionality into three main tools:
 * - teams_meeting: Meeting and transcript operations
 * - teams_channel: Channel and message operations  
 * - teams_chat: Chat operations
 */

const teamsChannel = require('./teams_channel');
const teamsChat = require('./teams_chat');
const teamsMeeting = require('./teams_meeting');

// Export all consolidated teams tools
module.exports = [
  teamsChannel,
  teamsChat,
  teamsMeeting
];
