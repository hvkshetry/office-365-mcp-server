# Windows Task Scheduler Setup for Office MCP Server

Since the node-windows service isn't creating properly, we'll use Windows Task Scheduler instead. This achieves the same goal - automatic startup without "human in the loop" - using Windows' built-in task scheduling.

## Quick Setup (Recommended)

### Step 1: Run the Setup Script

Open PowerShell **as Administrator** and run:

```powershell
# Navigate to the office-mcp directory
cd C:\Users\hvksh\mcp-servers\office-mcp

# Run the setup script
.\scripts\setup-task-scheduler.ps1
```

This will:
- Create a scheduled task that starts at Windows boot
- Run the server invisibly in the background
- Configure automatic restart on failure
- Start the server immediately

### Step 2: Verify It's Working

```powershell
# Check if the task is running
Get-ScheduledTask "Office MCP Stdio Server" | Select-Object TaskName, State

# Should show: State = Running
```

## Manual Setup (Alternative)

If you prefer to set it up manually or the script doesn't work:

### Step 1: Open Task Scheduler
- Press `Win + R`, type `taskschd.msc`, press Enter

### Step 2: Create Basic Task
1. Click "Create Task..." (not "Create Basic Task")
2. **General Tab:**
   - Name: `Office MCP Stdio Server`
   - Description: `Office 365 MCP Server - Headless stdio operation`
   - Check: "Run whether user is logged on or not"
   - Check: "Run with highest privileges"

### Step 3: Configure Trigger
1. **Triggers Tab:**
   - Click "New..."
   - Begin the task: "At startup"
   - Check: "Enabled"
   - Click "OK"

### Step 4: Configure Action
1. **Actions Tab:**
   - Click "New..."
   - Action: "Start a program"
   - Program/script: `C:\Users\hvksh\mcp-servers\office-mcp\scripts\run-headless.bat`
   - Start in: `C:\Users\hvksh\mcp-servers\office-mcp`
   - Click "OK"

### Step 5: Configure Settings
1. **Settings Tab:**
   - Check: "Allow task to be run on demand"
   - Check: "If the task fails, restart every: 1 minute"
   - Attempt to restart up to: "3 times"
   - Check: "Start the task as soon as possible after a scheduled start is missed"
   - Uncheck: "Stop the task if it runs longer than"

### Step 6: Save and Test
1. Click "OK"
2. Enter your Windows password if prompted
3. Right-click the task and select "Run" to test

## How It Works

1. **At Windows Startup**: Task Scheduler launches `run-headless.bat`
2. **Invisible Operation**: The batch file runs Node.js with PowerShell in hidden mode
3. **Stdio Mode**: The server operates in stdio mode, compatible with Claude Code
4. **Auto-Refresh**: Tokens refresh automatically using stored refresh token
5. **Logging**: Output is saved to `logs\stdio-server.log`

## Benefits Over Windows Service

✅ **More Reliable**: Uses Windows' built-in scheduler (no third-party service wrapper)
✅ **Easier to Debug**: Can see task status in Task Scheduler GUI
✅ **Better Logging**: Direct control over log files
✅ **Simpler Management**: Standard Windows tool everyone knows

## Management Commands

### View Task Status
```powershell
Get-ScheduledTask "Office MCP Stdio Server"
```

### Start Task Manually
```powershell
Start-ScheduledTask "Office MCP Stdio Server"
```

### Stop Task
```powershell
Stop-ScheduledTask "Office MCP Stdio Server"
```

### Remove Task
```powershell
Unregister-ScheduledTask "Office MCP Stdio Server" -Confirm:$false
```

### View Logs
```powershell
Get-Content C:\Users\hvksh\mcp-servers\office-mcp\logs\stdio-server.log -Tail 50
```

### Follow Logs in Real-Time
```powershell
Get-Content C:\Users\hvksh\mcp-servers\office-mcp\logs\stdio-server.log -Wait
```

## Troubleshooting

### Task Won't Start

1. **Check PowerShell execution policy:**
   ```powershell
   Get-ExecutionPolicy
   # If Restricted, run:
   Set-ExecutionPolicy RemoteSigned -Scope CurrentUser
   ```

2. **Verify paths exist:**
   ```powershell
   Test-Path C:\Users\hvksh\mcp-servers\office-mcp\scripts\run-headless.bat
   Test-Path C:\Users\hvksh\.office-mcp-tokens.json
   ```

### Server Not Responding in Claude Code

1. **Check if Node process is running:**
   ```powershell
   Get-Process node -ErrorAction SilentlyContinue
   ```

2. **Check logs for errors:**
   ```powershell
   Get-Content C:\Users\hvksh\mcp-servers\office-mcp\logs\stdio-server.log -Tail 100
   ```

3. **Restart the task:**
   ```powershell
   Stop-ScheduledTask "Office MCP Stdio Server"
   Start-ScheduledTask "Office MCP Stdio Server"
   ```

### Task Runs but Claude Code Can't Connect

1. **Verify .mcp.json configuration** is correct (should already be set for stdio)
2. **Restart Claude Code** after setting up the task
3. **Check that only one instance** of the server is running

## Testing the Setup

1. **Immediate Test:**
   ```powershell
   # Start the task
   Start-ScheduledTask "Office MCP Stdio Server"
   
   # Wait a moment
   Start-Sleep -Seconds 5
   
   # Check if running
   Get-Process node
   ```

2. **Reboot Test:**
   - Restart Windows
   - After login, check if the server is running:
   ```powershell
   Get-ScheduledTask "Office MCP Stdio Server" | Select-Object State
   Get-Process node
   ```

3. **Claude Code Test:**
   - Open Claude Code
   - The Office MCP tools should be available immediately
   - No browser authentication needed

## Summary

This Task Scheduler approach provides the same benefits as a Windows service:
- ✅ Automatic startup at Windows boot
- ✅ Runs invisibly in background
- ✅ No "human in the loop" for authentication
- ✅ Automatic token refresh
- ✅ Restart on failure

But with better reliability and easier management using Windows' built-in tools.