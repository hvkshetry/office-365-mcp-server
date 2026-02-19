const { describe, it, expect } = require('@jest/globals');
const teamsTools = require('../teams');
const { ensureAuthenticated } = require('../auth');
const { callGraphAPI } = require('../utils/graph-api');

jest.mock('../auth', () => ({
  ensureAuthenticated: jest.fn()
}));
jest.mock('../utils/graph-api');

describe('Teams Module - Consolidated Tools', () => {
  const mockAccessToken = 'mock-access-token';

  beforeEach(() => {
    jest.clearAllMocks();
    ensureAuthenticated.mockResolvedValue(mockAccessToken);
  });

  describe('Consolidated Tools Export', () => {
    it('should export exactly 3 consolidated tools', () => {
      expect(teamsTools).toHaveLength(3);

      const toolNames = teamsTools.map(tool => tool.name);
      expect(toolNames).toContain('teams_meeting');
      expect(toolNames).toContain('teams_channel');
      expect(toolNames).toContain('teams_chat');
    });

    it('should have proper tool schema structure', () => {
      teamsTools.forEach(tool => {
        expect(tool).toHaveProperty('name');
        expect(tool).toHaveProperty('description');
        expect(tool).toHaveProperty('inputSchema');
        expect(tool).toHaveProperty('handler');
      });
    });
  });

  describe('teams_meeting tool', () => {
    let meetingTool;

    beforeEach(() => {
      meetingTool = teamsTools.find(tool => tool.name === 'teams_meeting');
    });

    it('should have correct operations in schema', () => {
      const operations = meetingTool.inputSchema.properties.operation.enum;
      expect(operations).toContain('create');
      expect(operations).toContain('list_transcripts');
      expect(operations).toContain('get_transcript');
      expect(operations).toContain('get_participants');
    });

    it('should handle meeting creation', async () => {
      const mockMeeting = {
        id: 'meeting-id',
        joinWebUrl: 'https://teams.microsoft.com/l/meetup-join/...'
      };

      callGraphAPI.mockResolvedValue(mockMeeting);

      const result = await meetingTool.handler({
        operation: 'create',
        subject: 'Test Meeting',
        startDateTime: '2025-05-20T13:00:00Z',
        endDateTime: '2025-05-20T14:00:00Z'
      });

      expect(result.content[0].text).toContain('Meeting created successfully');
      expect(result.content[0].text).toContain(mockMeeting.joinWebUrl);
    });
  });

  describe('teams_channel tool', () => {
    let channelTool;

    beforeEach(() => {
      channelTool = teamsTools.find(tool => tool.name === 'teams_channel');
    });

    it('should have correct operations in schema', () => {
      const operations = channelTool.inputSchema.properties.operation.enum;
      expect(operations).toContain('list');
      expect(operations).toContain('create');
      expect(operations).toContain('list_messages');
      expect(operations).toContain('create_message');
    });

    it('should handle channel listing', async () => {
      const mockChannels = {
        value: [
          { id: 'channel1', displayName: 'General' }
        ]
      };

      callGraphAPI.mockResolvedValue(mockChannels);

      const result = await channelTool.handler({
        operation: 'list',
        teamId: 'team-id'
      });

      expect(result.content[0].text).toContain('Found 1 channel');
      expect(result.content[0].text).toContain('General');
    });
  });

  describe('teams_chat tool', () => {
    let chatTool;

    beforeEach(() => {
      chatTool = teamsTools.find(tool => tool.name === 'teams_chat');
    });

    it('should have correct operations in schema', () => {
      const operations = chatTool.inputSchema.properties.operation.enum;
      expect(operations).toContain('list');
      expect(operations).toContain('create');
      expect(operations).toContain('send_message');
      expect(operations).toContain('list_members');
    });

    it('should handle chat message sending', async () => {
      callGraphAPI.mockResolvedValue({ id: 'message-id' });

      const result = await chatTool.handler({
        operation: 'send_message',
        chatId: 'chat-id',
        content: 'Test message'
      });

      expect(result.content[0].text).toContain('Message sent successfully');
    });
  });
});
