@echo off
echo Starting Office MCP Authentication Server...
echo.
echo This server handles OAuth authentication for the Office MCP server.
echo It must be running when you use the authenticate tool in Claude.
echo.
echo Server will run on: http://localhost:3333
echo Redirect URI: http://localhost:3333/auth/callback
echo.
echo Press Ctrl+C to stop the server
echo ========================================
echo.

:: Start the authentication server
node office-auth-server.js

:: If node command fails, try with full path
if errorlevel 1 (
    echo.
    echo Error: Node.js not found in PATH
    echo Trying with common Node.js installation paths...
    
    :: Try Program Files
    if exist "%ProgramFiles%\nodejs\node.exe" (
        "%ProgramFiles%\nodejs\node.exe" office-auth-server.js
    ) else if exist "%ProgramFiles(x86)%\nodejs\node.exe" (
        "%ProgramFiles(x86)%\nodejs\node.exe" office-auth-server.js
    ) else (
        echo.
        echo Error: Could not find Node.js
        echo Please ensure Node.js is installed and in your PATH
        echo.
        pause
    )
)

pause