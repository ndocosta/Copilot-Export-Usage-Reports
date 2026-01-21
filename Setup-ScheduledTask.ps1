<#
.SYNOPSIS
    Sets up a Windows Scheduled Task to run the Copilot usage report export script daily.

.DESCRIPTION
    This script creates a scheduled task that runs Export-CopilotUsageReports.ps1 daily.
    The task can be configured to run at a specific time and under a specific user account.

.PARAMETER TaskName
    Name of the scheduled task. Default is "Export-CopilotUsageReports"

.PARAMETER ExecutionTime
    Time to run the task daily (HH:mm format). Default is "02:00" (2 AM)

.PARAMETER ScriptPath
    Path to the Export-CopilotUsageReports.ps1 script. Defaults to the same directory as this script.

.PARAMETER RunAsUser
    User account to run the task under. If not specified, uses the current user.

.EXAMPLE
    .\Setup-ScheduledTask.ps1
    .\Setup-ScheduledTask.ps1 -ExecutionTime "03:30" -TaskName "CopilotReports"

.NOTES
    Requires administrative privileges to create scheduled tasks
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$TaskName = "Export-CopilotUsageReports",
    
    [Parameter(Mandatory = $false)]
    [string]$ExecutionTime = "02:00",
    
    [Parameter(Mandatory = $false)]
    [string]$ScriptPath = (Join-Path $PSScriptRoot "Export-CopilotUsageReports.ps1"),
    
    [Parameter(Mandatory = $false)]
    [string]$RunAsUser = $env:USERNAME
)

function Test-Administrator {
    $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($currentUser)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

# Check for administrative privileges
if (-not (Test-Administrator)) {
    Write-Host "This script requires administrative privileges to create scheduled tasks." -ForegroundColor Red
    Write-Host "Please run PowerShell as Administrator and try again." -ForegroundColor Yellow
    exit 1
}

# Validate script path
if (-not (Test-Path $ScriptPath)) {
    Write-Host "Script not found: $ScriptPath" -ForegroundColor Red
    exit 1
}

Write-Host "=== Setting up Scheduled Task ===" -ForegroundColor Cyan
Write-Host "Task Name: $TaskName" -ForegroundColor Gray
Write-Host "Execution Time: $ExecutionTime daily" -ForegroundColor Gray
Write-Host "Script Path: $ScriptPath" -ForegroundColor Gray
Write-Host "Run As User: $RunAsUser" -ForegroundColor Gray
Write-Host ""

try {
    # Check if task already exists
    $existingTask = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
    
    if ($existingTask) {
        Write-Host "Scheduled task '$TaskName' already exists." -ForegroundColor Yellow
        $response = Read-Host "Do you want to replace it? (Y/N)"
        
        if ($response -eq 'Y' -or $response -eq 'y') {
            Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
            Write-Host "Existing task removed." -ForegroundColor Green
        }
        else {
            Write-Host "Setup cancelled." -ForegroundColor Yellow
            exit 0
        }
    }
    
    # Parse execution time
    try {
        $timeComponents = $ExecutionTime.Split(':')
        $hour = [int]$timeComponents[0]
        $minute = [int]$timeComponents[1]
        
        if ($hour -lt 0 -or $hour -gt 23 -or $minute -lt 0 -or $minute -gt 59) {
            throw "Invalid time format"
        }
    }
    catch {
        Write-Host "Invalid time format. Please use HH:mm format (e.g., 02:00)" -ForegroundColor Red
        exit 1
    }
    
    # Create action - PowerShell command to execute the script
    $action = New-ScheduledTaskAction `
        -Execute "powershell.exe" `
        -Argument "-ExecutionPolicy Bypass -NoProfile -WindowStyle Hidden -File `"$ScriptPath`""
    
    # Create trigger - Daily at specified time
    $trigger = New-ScheduledTaskTrigger -Daily -At $ExecutionTime
    
    # Create principal - Run whether user is logged on or not
    $principal = New-ScheduledTaskPrincipal `
        -UserId $RunAsUser `
        -LogonType S4U `
        -RunLevel Highest
    
    # Create settings
    $settings = New-ScheduledTaskSettingsSet `
        -AllowStartIfOnBatteries `
        -DontStopIfGoingOnBatteries `
        -StartWhenAvailable `
        -RunOnlyIfNetworkAvailable `
        -ExecutionTimeLimit (New-TimeSpan -Hours 2)
    
    # Register the scheduled task
    $task = Register-ScheduledTask `
        -TaskName $TaskName `
        -Action $action `
        -Trigger $trigger `
        -Principal $principal `
        -Settings $settings `
        -Description "Exports Microsoft 365 Copilot usage reports from Graph API and uploads to SharePoint Online"
    
    Write-Host ""
    Write-Host "Scheduled task created successfully!" -ForegroundColor Green
    Write-Host ""
    Write-Host "Task Details:" -ForegroundColor Cyan
    Write-Host "  Name: $TaskName" -ForegroundColor Gray
    Write-Host "  Status: $($task.State)" -ForegroundColor Gray
    Write-Host "  Next Run: $($trigger.StartBoundary)" -ForegroundColor Gray
    Write-Host "  User: $RunAsUser" -ForegroundColor Gray
    Write-Host ""
    
    # Test the task
    $response = Read-Host "Do you want to run the task now to test it? (Y/N)"
    
    if ($response -eq 'Y' -or $response -eq 'y') {
        Write-Host "Starting task..." -ForegroundColor Yellow
        Start-ScheduledTask -TaskName $TaskName
        
        Start-Sleep -Seconds 2
        $taskInfo = Get-ScheduledTaskInfo -TaskName $TaskName
        
        Write-Host "Task triggered. Last Run Time: $($taskInfo.LastRunTime)" -ForegroundColor Green
        Write-Host "Last Result: $($taskInfo.LastTaskResult)" -ForegroundColor Gray
        Write-Host ""
        Write-Host "Check the log files in your configured LocalExportPath for execution details." -ForegroundColor Cyan
    }
    
    Write-Host ""
    Write-Host "Setup complete! The task will run daily at $ExecutionTime" -ForegroundColor Green
    Write-Host ""
    Write-Host "To manage the task:" -ForegroundColor Cyan
    Write-Host "  - View: Get-ScheduledTask -TaskName '$TaskName'" -ForegroundColor Gray
    Write-Host "  - Run manually: Start-ScheduledTask -TaskName '$TaskName'" -ForegroundColor Gray
    Write-Host "  - Disable: Disable-ScheduledTask -TaskName '$TaskName'" -ForegroundColor Gray
    Write-Host "  - Remove: Unregister-ScheduledTask -TaskName '$TaskName' -Confirm:`$false" -ForegroundColor Gray
}
catch {
    Write-Host ""
    Write-Host "Error setting up scheduled task: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Stack trace: $($_.ScriptStackTrace)" -ForegroundColor Red
    exit 1
}
