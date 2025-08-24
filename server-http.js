#!/usr/bin/env node
/**
 * Office MCP Server - HTTP Transport Mode
 * 
 * This server provides HTTP/SSE transport for headless operation
 * Can be run as a Windows service or background process
 */
const { Server } = require("@modelcontextprotocol/sdk/server/index.js");
const express = require('express');
const config = require('./config');

// Import module tools (reuse all existing tools)
const { authTools } = require('./auth');
const { calendarTools } = require('./calendar');
const { emailTools } = require('./email');
const teamsTools = require('./teams');
const { notificationTools } = require('./notifications');
const { plannerTools } = require('./planner');
const { filesTools } = require('./files');
const { searchTools } = require('./search');

// Get port from environment or config
const PORT = process.env.HTTP_PORT || process.env.PORT || 3333;
const HOST = process.env.HTTP_HOST || '127.0.0.1';

// Log startup information
console.log(`Starting ${config.SERVER_NAME.toUpperCase()} MCP HTTP Server`);
console.log(`Server will listen on http://${HOST}:${PORT}/mcp`);
console.log(`Test mode is ${config.USE_TEST_MODE ? 'enabled' : 'disabled'}`);

// Combine all tools (exactly as in index.js)
const TOOLS = [
  ...authTools,
  ...calendarTools,
  ...emailTools,
  ...teamsTools,
  ...notificationTools,
  ...plannerTools,
  ...filesTools,
  ...searchTools
];

console.log(`Loaded ${TOOLS.length} tools`);

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

// Reuse the exact same fallback handler from index.js
server.fallbackRequestHandler = async (request) => {
  try {
    const { method, params, id } = request;
    console.log(`REQUEST: ${method} [${id}]`);
    
    // Initialize handler
    if (method === "initialize") {
      console.log(`INITIALIZE REQUEST: ID [${id}]`);
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
      console.log(`TOOLS LIST REQUEST: ID [${id}]`);
      console.log(`TOOLS COUNT: ${TOOLS.length}`);
      
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
        
        console.log(`TOOL CALL: ${name}`);
        
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

// Create Express app
const app = express();

// Health check endpoint
app.get('/health', (req, res) => {
  res.json({
    status: 'healthy',
    server: config.SERVER_NAME,
    version: config.SERVER_VERSION,
    uptime: process.uptime(),
    timestamp: new Date().toISOString()
  });
});

// Info endpoint
app.get('/info', (req, res) => {
  res.json({
    name: config.SERVER_NAME,
    version: config.SERVER_VERSION,
    transport: 'http',
    port: PORT,
    tools: TOOLS.length,
    capabilities: ['tools']
  });
});

// Enable JSON parsing for all requests
app.use(express.json());

// Always set up the MCP POST endpoint for HTTP transport
app.post('/mcp', async (req, res) => {
  console.log('Received MCP request:', req.body?.method || 'unknown method');
  
  try {
    const response = await server.fallbackRequestHandler(req.body);
    res.json(response);
  } catch (error) {
    console.error('Error handling MCP request:', error);
    res.status(500).json({
      error: {
        code: -32603,
        message: `Internal error: ${error.message}`
      }
    });
  }
});

// SSE stream endpoint for potential future use
app.get('/mcp/stream', (req, res) => {
  res.writeHead(200, {
    'Content-Type': 'text/event-stream',
    'Cache-Control': 'no-cache',
    'Connection': 'keep-alive'
  });
  
  // Send initial connection message
  res.write(`data: ${JSON.stringify({ type: 'connected' })}\n\n`);
  
  // Keep connection alive
  const keepAlive = setInterval(() => {
    res.write(': keep-alive\n\n');
  }, 30000);
  
  req.on('close', () => {
    clearInterval(keepAlive);
  });
});

// Try to import StreamableHTTPServerTransport if available
let StreamableHTTPServerTransport;
let transport;
try {
  // Try to import from the SDK
  const transportModule = require("@modelcontextprotocol/sdk/server/streamableHttp.js");
  StreamableHTTPServerTransport = transportModule.StreamableHTTPServerTransport || transportModule.default;
  
  if (StreamableHTTPServerTransport) {
    // Use official StreamableHTTPServerTransport if available
    transport = new StreamableHTTPServerTransport({
      app,
      path: '/mcp',
      port: PORT,
      host: HOST
    });
    console.log('Using official StreamableHTTPServerTransport');
  }
} catch (error) {
  console.log('Note: StreamableHTTPServerTransport not found in SDK, using HTTP POST endpoint');
}

// Handle graceful shutdown
process.on('SIGTERM', () => {
  console.log('SIGTERM received, shutting down gracefully');
  process.exit(0);
});

process.on('SIGINT', () => {
  console.log('SIGINT received, shutting down gracefully');
  process.exit(0);
});

// Start the Express server (handles both transport modes)
const httpServer = app.listen(PORT, HOST, () => {
  console.log(`${config.SERVER_NAME} HTTP server running`);
  console.log(`Server listening on http://${HOST}:${PORT}/mcp`);
  console.log(`Health check available at http://${HOST}:${PORT}/health`);
});

// If we have the official transport, connect it
if (StreamableHTTPServerTransport && transport) {
  server.connect(transport)
    .then(() => {
      console.log(`MCP transport connected successfully`);
    })
    .catch(error => {
      console.error(`MCP transport connection error: ${error.message}`);
      console.error('Falling back to simple HTTP endpoint mode');
    });
}

// Keep the server running
httpServer.on('error', (error) => {
  console.error('Server error:', error);
  if (error.code === 'EADDRINUSE') {
    console.error(`Port ${PORT} is already in use. Please stop any existing server or use a different port.`);
    process.exit(1);
  }
});