#!/usr/bin/env node
/**
 * Office MCP Server - Main entry point
 * 
 * A Model Context Protocol server that provides access to
 * Microsoft 365 services through the Microsoft Graph API.
 */

// Load environment variables from .env file
// Use absolute path to ensure it loads regardless of working directory
const path = require('path');
require('dotenv').config({ path: path.join(__dirname, '.env') });

const { Server } = require("@modelcontextprotocol/sdk/server/index.js");
const { StdioServerTransport } = require("@modelcontextprotocol/sdk/server/stdio.js");
const config = require('./config');

// Import module tools
const { authTools } = require('./auth');
const { calendarTools } = require('./calendar');
const { emailTools } = require('./email');
const teamsTools = require('./teams');
const { notificationTools } = require('./notifications');
const { plannerTools } = require('./planner');
const { filesTools } = require('./files');
const { searchTools } = require('./search');
const { contactsTools } = require('./contacts');
// Future modules to be developed:
// const { adminTools } = require('./admin');

// Log startup information
console.error(`STARTING ${config.SERVER_NAME.toUpperCase()} MCP SERVER`);
console.error(`Test mode is ${config.USE_TEST_MODE ? 'enabled' : 'disabled'}`);
console.error(`Client ID: ${config.AUTH_CONFIG.clientId ? config.AUTH_CONFIG.clientId.substring(0, 8) + '...' : 'NOT SET'}`);
console.error(`Token path: ${config.AUTH_CONFIG.tokenStorePath}`);
console.error(`Token exists: ${require('fs').existsSync(config.AUTH_CONFIG.tokenStorePath)}`);

// Combine all tools
const TOOLS = [
  ...authTools,
  ...calendarTools,
  ...emailTools,
  ...teamsTools,
  ...notificationTools,
  ...plannerTools,
  ...filesTools,
  ...searchTools,
  ...contactsTools
  // Future modules will be added here:
  // ...adminTools
];

// Create server with tools capabilities
const server = new Server(
  { name: config.SERVER_NAME, version: config.SERVER_VERSION },
  { 
    capabilities: { 
      tools: TOOLS.reduce((acc, tool) => {
        acc[tool.name] = {};
        return acc;
      }, {})
    } 
  }
);

// Handle all requests
server.fallbackRequestHandler = async (request) => {
  try {
    const { method, params, id } = request;
    console.error(`REQUEST: ${method} [${id}]`);
    
    // Initialize handler
    if (method === "initialize") {
      console.error(`INITIALIZE REQUEST: ID [${id}]`);
      return {
        protocolVersion: "2024-11-05",
        capabilities: { 
          tools: TOOLS.reduce((acc, tool) => {
            acc[tool.name] = {};
            return acc;
          }, {})
        },
        serverInfo: { name: config.SERVER_NAME, version: config.SERVER_VERSION }
      };
    }
    
    // Tools list handler
    if (method === "tools/list") {
      console.error(`TOOLS LIST REQUEST: ID [${id}]`);
      console.error(`TOOLS COUNT: ${TOOLS.length}`);
      console.error(`TOOLS NAMES: ${TOOLS.map(t => t.name).join(', ')}`);
      
      return {
        tools: TOOLS.map(tool => ({
          name: tool.name,
          description: tool.description,
          inputSchema: tool.inputSchema
        }))
      };
    }
    
    // Required empty responses for other capabilities
    if (method === "resources/list") return { resources: [] };
    if (method === "prompts/list") return { prompts: [] };
    
    // Tool call handler
    if (method === "tools/call") {
      try {
        const { name, arguments: args = {} } = params || {};
        
        console.error(`TOOL CALL: ${name}`);
        
        // Find the tool handler
        const tool = TOOLS.find(t => t.name === name);
        
        if (tool && tool.handler) {
          return await tool.handler(args);
        }
        
        // Tool not found
        return {
          error: {
            code: -32601,
            message: `Tool not found: ${name}`
          }
        };
      } catch (error) {
        console.error(`Error in tools/call:`, error);
        return {
          error: {
            code: -32603,
            message: `Error processing tool call: ${error.message}`
          }
        };
      }
    }
    
    // For any other method, return method not found
    return {
      error: {
        code: -32601,
        message: `Method not found: ${method}`
      }
    };
  } catch (error) {
    console.error(`Error in fallbackRequestHandler:`, error);
    return {
      error: {
        code: -32603,
        message: `Error processing request: ${error.message}`
      }
    };
  }
};

// Make the script executable
process.on('SIGTERM', () => {
  console.error('SIGTERM received but staying alive');
});

// Start the server
const transport = new StdioServerTransport();
server.connect(transport)
  .then(() => console.error(`${config.SERVER_NAME} connected and listening`))
  .catch(error => {
    console.error(`Connection error: ${error.message}`);
    process.exit(1);
  });