@echo off
echo =========================================
echo Starting Office MCP Server Suite
echo =========================================
echo.

:: Check if Node.js is available
where node >nul 2>nul
if errorlevel 1 (
    echo Error: Node.js is not installed or not in PATH
    echo Please install Node.js from https://nodejs.org/
    pause
    exit /b 1
)

:: Start authentication server in a new window
echo Starting Authentication Server...
start "Office MCP Auth Server" /min cmd /c "node office-auth-server.js"

:: Wait a moment for the auth server to start
timeout /t 2 /nobreak > nul

:: Start MCP server in a new window
echo Starting MCP Server...
start "Office MCP Server" /min cmd /c "node index.js"

echo.
echo =========================================
echo Both servers are now running!
echo =========================================
echo.
echo Authentication Server: http://localhost:3000
echo.
echo To use with Claude Desktop:
echo 1. Make sure you've configured your .env file
echo 2. Update your Claude Desktop config
echo 3. Use the authenticate tool in Claude
echo.
echo Press any key to open both server windows...
pause > nul

:: Bring both windows to the foreground
powershell -Command "Get-Process | Where-Object {$_.MainWindowTitle -like '*Office MCP*'} | ForEach-Object { [Microsoft.VisualBasic.Interaction]::AppActivate($_.Id) }"

echo.
echo To stop the servers, close their windows or press Ctrl+C in each window.
pause