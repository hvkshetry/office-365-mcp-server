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
const { ListToolsRequestSchema, CallToolRequestSchema, McpError, ErrorCode } = require("@modelcontextprotocol/sdk/types.js");
const config = require('./config');
const { auditLog } = require('./utils/audit');

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
const { todoTools } = require('./todo');
const { groupsTools } = require('./groups');
const { directoryTools } = require('./directory');

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
  ...contactsTools,
  ...todoTools,
  ...groupsTools,
  ...directoryTools
];

// Create server with tools capabilities
// SDK handles initialize and ping automatically via setRequestHandler
const server = new Server(
  { name: config.SERVER_NAME, version: config.SERVER_VERSION },
  { capabilities: { tools: {} } }
);

// tools/list — return tool metadata
server.setRequestHandler(ListToolsRequestSchema, async () => {
  console.error(`TOOLS LIST REQUEST — ${TOOLS.length} tools`);
  return {
    tools: TOOLS.map(tool => ({
      name: tool.name,
      description: tool.description,
      inputSchema: tool.inputSchema
    }))
  };
});

// tools/call — dispatch to the matching tool handler
server.setRequestHandler(CallToolRequestSchema, async (request) => {
  const { name, arguments: args = {} } = request.params;
  console.error(`TOOL CALL: ${name}`);
  auditLog(name, args);

  const tool = TOOLS.find(t => t.name === name);
  if (!tool) {
    throw new McpError(ErrorCode.MethodNotFound, `Tool not found: ${name}`);
  }
  return await tool.handler(args);
});

// Graceful shutdown handlers
let isShuttingDown = false;

function gracefulShutdown(signal) {
  if (isShuttingDown) return;
  isShuttingDown = true;

  console.error(`[SHUTDOWN] ${signal} received, shutting down gracefully`);

  // Give pending operations time to complete
  setTimeout(() => {
    console.error('[SHUTDOWN] Exiting');
    process.exit(0);
  }, 1000);
}

process.on('SIGTERM', () => gracefulShutdown('SIGTERM'));
process.on('SIGINT', () => gracefulShutdown('SIGINT'));

// Start the server
if (config.TRANSPORT_TYPE === 'http') {
  const { SSEServerTransport } = require("@modelcontextprotocol/sdk/server/sse.js");
  const express = require('express');
  const app = express();
  app.use(express.json());

  let sseTransport;

  app.get("/sse", async (req, res) => {
    sseTransport = new SSEServerTransport("/message", res);
    await server.connect(sseTransport);
  });

  app.post("/message", async (req, res) => {
    if (sseTransport) {
      await sseTransport.handlePostMessage(req, res, req.body);
    } else {
      res.status(503).json({ error: "No SSE connection" });
    }
  });

  app.listen(config.HTTP_PORT, config.HTTP_HOST, () => {
    console.error(`${config.SERVER_NAME} HTTP/SSE on ${config.HTTP_HOST}:${config.HTTP_PORT}`);
  });
} else {
  const transport = new StdioServerTransport();
  server.connect(transport)
    .then(() => console.error(`${config.SERVER_NAME} connected and listening`))
    .catch(error => {
      console.error(`Connection error: ${error.message}`);
      process.exit(1);
    });
}