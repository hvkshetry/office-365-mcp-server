const { describe, it, expect, jest } = require('@jest/globals');
const { Server } = require("@modelcontextprotocol/sdk/server/index.js");
const config = require('../config');
const { createMockTokens } = require('./test-utils');

// Import all modules to test
const { authTools } = require('../auth');
const { emailTools } = require('../email');
const { calendarTools } = require('../calendar');
const { teamsTools } = require('../teams');
const { driveTools } = require('../drive');
const { plannerTools } = require('../planner');
const { userTools } = require('../users');
const { notificationTools } = require('../notifications');

describe('Office MCP Server Integration Tests', () => {
  let server;
  
  beforeEach(() => {
    // Create a new server instance for each test
    server = new Server(
      { name: config.SERVER_NAME, version: config.SERVER_VERSION },
      { 
        capabilities: { 
          tools: {}
        } 
      }
    );
  });
  
  describe('Server Configuration', () => {
    it('should initialize with correct name and version', () => {
      expect(server.name).toBe(config.SERVER_NAME);
      expect(server.version).toBe(config.SERVER_VERSION);
    });
  });
  
  describe('Tool Registration', () => {
    it('should register all authentication tools', () => {
      const toolCount = authTools.length;
      expect(toolCount).toBe(3);
      
      authTools.forEach(tool => {
        expect(tool).toHaveProperty('name');
        expect(tool).toHaveProperty('handler');
        expect(tool).toHaveProperty('schema');
      });
    });
    
    it('should register all teams tools', () => {
      const toolCount = teamsTools.length;
      expect(toolCount).toBe(19);
      
      teamsTools.forEach(tool => {
        expect(tool).toHaveProperty('name');
        expect(tool).toHaveProperty('handler');
      });
    });
    
    it('should register all drive tools', () => {
      const toolCount = driveTools.length;
      expect(toolCount).toBe(13);
    });
    
    it('should register all planner tools', () => {
      const toolCount = plannerTools.length;
      expect(toolCount).toBe(16);
    });
    
    it('should register all user tools', () => {
      const toolCount = userTools.length;
      expect(toolCount).toBe(16);
    });
    
    it('should register all notification tools', () => {
      const toolCount = notificationTools.length;
      expect(toolCount).toBe(5);
    });
  });
  
  describe('Tool Collection', () => {
    it('should have correct total number of tools', () => {
      const allTools = [
        ...authTools,
        ...emailTools,
        ...calendarTools,
        ...teamsTools,
        ...driveTools,
        ...plannerTools,
        ...userTools,
        ...notificationTools
      ];
      
      expect(allTools.length).toBe(81); // Total number of tools
    });
    
    it('should have unique tool names', () => {
      const allTools = [
        ...authTools,
        ...emailTools,
        ...calendarTools,
        ...teamsTools,
        ...driveTools,
        ...plannerTools,
        ...userTools,
        ...notificationTools
      ];
      
      const toolNames = allTools.map(tool => tool.name);
      const uniqueNames = [...new Set(toolNames)];
      
      expect(toolNames.length).toBe(uniqueNames.length);
    });
  });
  
  describe('Error Handling', () => {
    it('should handle missing authentication gracefully', async () => {
      // Test a tool that requires authentication
      const listTeamsTool = teamsTools.find(tool => tool.name === 'list_teams');
      const result = await listTeamsTool.handler({});
      
      expect(result.content[0].text).toContain('Not authenticated');
    });
    
    it('should validate required parameters', async () => {
      const createMeetingTool = teamsTools.find(tool => tool.name === 'create_meeting');
      const result = await createMeetingTool.handler({});
      
      expect(result.content[0].text).toContain('Missing required parameters');
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
      const hasCorrectPermission = scopes.includes('OnlineMeetings.ReadWrite');
      expect(hasCorrectPermission).toBe(true);
      
      // Ensure we don't have incorrect versions
      const hasMisspelledPermission1 = scopes.includes('OnlineMeeting.ReadWrite.All');
      const hasMisspelledPermission2 = scopes.includes('OnlineMeetings.ReadWrite.All');
      expect(hasMisspelledPermission1).toBe(false);
      expect(hasMisspelledPermission2).toBe(false);
    });
  });
  
  describe('Schema Validation', () => {
    it('should have valid JSON schemas for all tools', () => {
      const allTools = [
        ...authTools,
        ...emailTools,
        ...calendarTools,
        ...teamsTools,
        ...driveTools,
        ...plannerTools,
        ...userTools,
        ...notificationTools
      ];
      
      allTools.forEach(tool => {
        if (tool.schema) {
          expect(tool.schema).toHaveProperty('properties');
          expect(tool.schema).toHaveProperty('type', 'object');
        }
      });
    });
  });
});