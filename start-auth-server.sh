#!/bin/bash

echo "Starting Office MCP Authentication Server..."
echo ""
echo "This server handles OAuth authentication for the Office MCP server."
echo "It must be running when you use the authenticate tool in Claude."
echo ""
echo "Server will run on: http://localhost:3000"
echo "Redirect URI: http://localhost:3000/auth/callback"
echo ""
echo "Press Ctrl+C to stop the server"
echo "========================================"
echo ""

# Start the authentication server
node office-auth-server.js

# If node command fails, provide helpful error message
if [ $? -ne 0 ]; then
    echo ""
    echo "Error: Failed to start authentication server"
    echo "Please ensure Node.js is installed and in your PATH"
    echo ""
    
    # Check if node is installed
    if ! command -v node &> /dev/null; then
        echo "Node.js is not installed or not in PATH"
        echo "Install Node.js from: https://nodejs.org/"
    else
        echo "Node.js is installed but failed to start the server"
        echo "Check for errors above"
    fi
fi