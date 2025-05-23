@echo off
title Office MCP Server Launcher

echo =========================================
echo Office MCP Server Launcher
echo =========================================
echo.
echo This will start both the authentication server and the main MCP server.
echo.
echo 1. Start authentication server only
echo 2. Start MCP server only
echo 3. Start both servers (in separate windows)
echo 4. Exit
echo.
choice /C 1234 /N /M "Select an option: "

if errorlevel 4 goto :exit
if errorlevel 3 goto :both
if errorlevel 2 goto :mcp
if errorlevel 1 goto :auth

:auth
echo Starting Authentication Server...
start "Office MCP Auth Server" cmd /k "node office-auth-server.js"
goto :end

:mcp
echo Starting MCP Server...
start "Office MCP Server" cmd /k "node index.js"
goto :end

:both
echo Starting Authentication Server...
start "Office MCP Auth Server" cmd /k "node office-auth-server.js"
timeout /t 2 /nobreak > nul
echo Starting MCP Server...
start "Office MCP Server" cmd /k "node index.js"
echo.
echo Both servers have been started in separate windows.
goto :end

:exit
echo Exiting...
goto :end

:end
echo.
pause