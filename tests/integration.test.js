const { describe, it, expect } = require('@jest/globals');
const config = require('../config');

// Import all modules
const { authTools } = require('../auth');
const { emailTools } = require('../email');
const { calendarTools } = require('../calendar');
const teamsTools = require('../teams');
const { notificationTools } = require('../notifications');
const { plannerTools } = require('../planner');
const { filesTools } = require('../files');
const { searchTools } = require('../search');
const { contactsTools } = require('../contacts');
const { todoTools } = require('../todo');
const { groupsTools } = require('../groups');
const { directoryTools } = require('../directory');

describe('Office MCP Server Integration Tests', () => {
  describe('Tool Registration', () => {
    it('should have 1 consolidated auth/system tool', () => {
      expect(authTools).toHaveLength(1);
      expect(authTools[0].name).toBe('system');
    });

    it('should have 1 consolidated email tool', () => {
      expect(emailTools).toHaveLength(1);
      expect(emailTools[0].name).toBe('mail');
    });

    it('should have 3 consolidated teams tools', () => {
      expect(teamsTools).toHaveLength(3);
      const names = teamsTools.map(t => t.name);
      expect(names).toContain('teams_meeting');
      expect(names).toContain('teams_channel');
      expect(names).toContain('teams_chat');
    });

    it('should have 1 consolidated planner tool', () => {
      expect(plannerTools).toHaveLength(1);
      expect(plannerTools[0].name).toBe('planner');
    });

    it('should have 1 consolidated notifications tool', () => {
      expect(notificationTools).toHaveLength(1);
      expect(notificationTools[0].name).toBe('notifications');
    });

    it('should have 1 calendar tool', () => {
      expect(calendarTools).toHaveLength(1);
    });

    it('should have files tools (files + sharepoint path mapper)', () => {
      expect(filesTools.length).toBeGreaterThanOrEqual(1);
      expect(filesTools.find(t => t.name === 'files')).toBeDefined();
    });

    it('should have 1 search tool', () => {
      expect(searchTools).toHaveLength(1);
    });

    it('should have 1 contacts tool', () => {
      expect(contactsTools).toHaveLength(1);
    });

    it('should have 1 todo tool', () => {
      expect(todoTools).toHaveLength(1);
      expect(todoTools[0].name).toBe('todo');
    });

    it('should have 1 groups tool', () => {
      expect(groupsTools).toHaveLength(1);
      expect(groupsTools[0].name).toBe('groups');
    });

    it('should have 1 directory tool', () => {
      expect(directoryTools).toHaveLength(1);
      expect(directoryTools[0].name).toBe('directory');
    });
  });

  describe('Tool Collection', () => {
    it('should have correct total number of consolidated tools', () => {
      const allTools = [
        ...authTools,
        ...emailTools,
        ...calendarTools,
        ...teamsTools,
        ...notificationTools,
        ...plannerTools,
        ...filesTools,
        ...searchTools,
        ...contactsTools,
        ...todoTools,
        ...groupsTools,
        ...directoryTools
      ];

      // Verify we have the expected consolidated tool count
      // (exact count may vary as files module includes sharepoint path mapper)
      expect(allTools.length).toBeGreaterThanOrEqual(15);
    });

    it('should have unique tool names', () => {
      const allTools = [
        ...authTools,
        ...emailTools,
        ...calendarTools,
        ...teamsTools,
        ...notificationTools,
        ...plannerTools,
        ...filesTools,
        ...searchTools,
        ...contactsTools,
        ...todoTools,
        ...groupsTools,
        ...directoryTools
      ];

      const toolNames = allTools.map(tool => tool.name);
      const uniqueNames = [...new Set(toolNames)];

      expect(toolNames.length).toBe(uniqueNames.length);
    });

    it('should have handler and inputSchema on all tools', () => {
      const allTools = [
        ...authTools,
        ...emailTools,
        ...calendarTools,
        ...teamsTools,
        ...notificationTools,
        ...plannerTools,
        ...filesTools,
        ...searchTools,
        ...contactsTools,
        ...todoTools,
        ...groupsTools,
        ...directoryTools
      ];

      allTools.forEach(tool => {
        expect(tool).toHaveProperty('name');
        expect(tool).toHaveProperty('handler');
        expect(tool).toHaveProperty('inputSchema');
        expect(typeof tool.handler).toBe('function');
      });
    });
  });

  describe('Configuration Validation', () => {
    it('should have valid server configuration', () => {
      expect(config.SERVER_NAME).toBe('office-mcp');
      expect(config.SERVER_VERSION).toBeDefined();
      expect(config.AUTH_CONFIG).toBeDefined();
      expect(config.AUTH_CONFIG.scopes).toBeInstanceOf(Array);
      expect(config.AUTH_CONFIG.scopes.length).toBeGreaterThan(0);
    });

    it('should have required OAuth configuration', () => {
      expect(config.AUTH_CONFIG.clientId).toBeDefined();
      expect(config.AUTH_CONFIG.clientSecret).toBeDefined();
      expect(config.AUTH_CONFIG.redirectUri).toBeDefined();
    });

    it('should have correct OnlineMeetings permission', () => {
      const scopes = config.AUTH_CONFIG.scopes;
      expect(scopes).toContain('OnlineMeetings.ReadWrite');
    });
  });
});
