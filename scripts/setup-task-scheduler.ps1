# Office MCP Task Scheduler Setup Script
# This creates a Windows scheduled task to run the Office MCP server at startup
# Run this script as Administrator in PowerShell

$taskName = "Office MCP Stdio Server"
$taskDescription = "Office 365 MCP Server - Runs at startup for headless operation with auto-refresh"
$scriptPath = "C:\Users\hvksh\mcp-servers\office-mcp\scripts\run-headless.bat"
$workingDirectory = "C:\Users\hvksh\mcp-servers\office-mcp"

Write-Host "Setting up Office MCP Server as scheduled task..." -ForegroundColor Cyan
Write-Host ""

# Check if running as Administrator
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Write-Host "ERROR: This script must be run as Administrator!" -ForegroundColor Red
    Write-Host "Please run PowerShell as Administrator and try again." -ForegroundColor Yellow
    exit 1
}

# Check if script exists
if (-not (Test-Path $scriptPath)) {
    Write-Host "ERROR: Batch script not found at: $scriptPath" -ForegroundColor Red
    exit 1
}

# Check if task already exists
$existingTask = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
if ($existingTask) {
    Write-Host "Task already exists. Removing old task..." -ForegroundColor Yellow
    Unregister-ScheduledTask -TaskName $taskName -Confirm:$false
}

# Create the scheduled task action
$action = New-ScheduledTaskAction -Execute $scriptPath -WorkingDirectory $workingDirectory

# Create the trigger (at system startup)
$trigger = New-ScheduledTaskTrigger -AtStartup

# Create principal (run whether user is logged in or not)
$principal = New-ScheduledTaskPrincipal -UserId "$env:USERDOMAIN\$env:USERNAME" `
    -LogonType ServiceAccount -RunLevel Highest

# Create settings
$settings = New-ScheduledTaskSettingsSet `
    -AllowStartIfOnBatteries `
    -DontStopIfGoingOnBatteries `
    -StartWhenAvailable `
    -RestartCount 3 `
    -RestartInterval (New-TimeSpan -Minutes 1) `
    -ExecutionTimeLimit (New-TimeSpan -Hours 0)

# Register the scheduled task
try {
    $task = Register-ScheduledTask `
        -TaskName $taskName `
        -Description $taskDescription `
        -Action $action `
        -Trigger $trigger `
        -Settings $settings `
        -User "$env:USERDOMAIN\$env:USERNAME" `
        -RunLevel Highest

    Write-Host "✅ Scheduled task created successfully!" -ForegroundColor Green
    Write-Host ""
    
    # Start the task immediately
    Write-Host "Starting the task now..." -ForegroundColor Cyan
    Start-ScheduledTask -TaskName $taskName
    
    Start-Sleep -Seconds 2
    
    # Check task status
    $taskInfo = Get-ScheduledTask -TaskName $taskName
    $taskState = (Get-ScheduledTask -TaskName $taskName).State
    
    Write-Host "✅ Task is now: $taskState" -ForegroundColor Green
    Write-Host ""
    Write-Host "Task Details:" -ForegroundColor Cyan
    Write-Host "  Name: $taskName"
    Write-Host "  Status: $taskState"
    Write-Host "  Trigger: At system startup"
    Write-Host "  Run as: $env:USERNAME"
    Write-Host ""
    Write-Host "The Office MCP Server will now:" -ForegroundColor Green
    Write-Host "  ✓ Start automatically when Windows boots"
    Write-Host "  ✓ Run invisibly in the background"
    Write-Host "  ✓ Auto-refresh tokens without browser"
    Write-Host "  ✓ Restart automatically if it crashes (up to 3 times)"
    Write-Host ""
    Write-Host "To manage the scheduled task:" -ForegroundColor Yellow
    Write-Host "  View in GUI: taskschd.msc"
    Write-Host "  Stop task: Stop-ScheduledTask -TaskName '$taskName'"
    Write-Host "  Start task: Start-ScheduledTask -TaskName '$taskName'"
    Write-Host "  Remove task: Unregister-ScheduledTask -TaskName '$taskName' -Confirm:`$false"
    Write-Host ""
    Write-Host "Check logs at: C:\Users\hvksh\mcp-servers\office-mcp\logs\stdio-server.log" -ForegroundColor Cyan
    
} catch {
    Write-Host "❌ Failed to create scheduled task: $_" -ForegroundColor Red
    exit 1
}