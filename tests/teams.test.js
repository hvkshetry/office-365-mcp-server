const { describe, it, expect, jest } = require('@jest/globals');
const teamsTools = require('../teams');
const tokenManager = require('../auth/token-manager');
const { callGraphAPI } = require('../utils/graph-api');

jest.mock('../auth/token-manager');
jest.mock('../utils/graph-api');

describe('Teams Module - Consolidated Tools', () => {
  const mockTokens = {
    access_token: 'mock-access-token',
    email: 'user@example.com'
  };

  beforeEach(() => {
    jest.clearAllMocks();
    tokenManager.loadTokenCache.mockReturnValue(mockTokens);
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

      expect(result.content[0].text).toContain('Found 1 channels');
      expect(result.content[0].text).toContain('General');
    });

    it('should add channel member using email', async () => {
      // Mock user lookup
      callGraphAPI.mockResolvedValueOnce({ value: [{ id: 'user123' }] });
      // Mock add member response
      callGraphAPI.mockResolvedValueOnce({ id: 'member123' });

      const result = await channelTool.handler({
        operation: 'add_member',
        teamId: 'team-id',
        channelId: 'channel-id',
        email: 'member@example.com',
        roles: ['owner']
      });

      expect(callGraphAPI).toHaveBeenNthCalledWith(1,
        'mock-access-token',
        'GET',
        'users',
        null,
        {
          $filter: "mail eq 'member@example.com' or userPrincipalName eq 'member@example.com'",
          $select: 'id'
        }
      );

      expect(callGraphAPI).toHaveBeenNthCalledWith(2,
        'mock-access-token',
        'POST',
        'teams/team-id/channels/channel-id/members',
        {
          '@odata.type': '#microsoft.graph.aadUserConversationMember',
          'user@odata.bind': "https://graph.microsoft.com/v1.0/users('user123')",
          roles: ['owner']
        }
      );

      expect(result.content[0].text).toContain('Member added successfully');
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