@echo off
REM Office MCP Headless Runner for Task Scheduler
REM This runs the stdio server without showing a window

cd /d "C:\Users\hvksh\mcp-servers\office-mcp"

REM Create logs directory if it doesn't exist
if not exist "logs" mkdir logs

REM Run the server with output redirected to log files
REM The server will run invisibly in the background
powershell -WindowStyle Hidden -Command "node index.js 2>&1 | Tee-Object -FilePath 'logs\stdio-server.log' -Append"