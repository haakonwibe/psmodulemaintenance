<#
.SYNOPSIS
    Installs PSModuleMaintenance as a scheduled task.

.DESCRIPTION
    Creates a Windows Scheduled Task that runs Invoke-PSModuleMaintenance.ps1 weekly.
    The task runs as the current user with highest privileges.

.PARAMETER ScriptPath
    Path to Invoke-PSModuleMaintenance.ps1. Defaults to same directory as this script.

.PARAMETER TaskName
    Name for the scheduled task. Defaults to 'PSModuleMaintenance'.

.PARAMETER DayOfWeek
    Day of the week to run. Defaults to 'Sunday'.

.PARAMETER Time
    Time to run (24h format). Defaults to '03:00'.

.PARAMETER Uninstall
    Remove the scheduled task instead of creating it.

.EXAMPLE
    .\Install-ModuleMaintenance.ps1

.EXAMPLE
    .\Install-ModuleMaintenance.ps1 -DayOfWeek Saturday -Time "04:30"

.EXAMPLE
    .\Install-ModuleMaintenance.ps1 -Uninstall

.NOTES
    Author: Haakon Wibe
    Requires: Windows PowerShell 7+, Administrator privileges
#>

[CmdletBinding()]
param(
    [Parameter()]
    [string]$ScriptPath,

    [Parameter()]
    [string]$TaskName = 'PSModuleMaintenance',

    [Parameter()]
    [ValidateSet('Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday')]
    [string]$DayOfWeek = 'Sunday',

    [Parameter()]
    [ValidatePattern('^\d{1,2}:\d{2}$')]
    [string]$Time = '03:00',

    [Parameter()]
    [switch]$Uninstall
)

#Requires -Version 7.0
#Requires -RunAsAdministrator

$ErrorActionPreference = 'Stop'

# ============================================================================
# FUNCTIONS
# ============================================================================

function Test-ScheduledTaskExists {
    param([string]$Name)
    
    $task = Get-ScheduledTask -TaskName $Name -ErrorAction SilentlyContinue
    return $null -ne $task
}

function Uninstall-MaintenanceTask {
    param([string]$Name)

    if (Test-ScheduledTaskExists -Name $Name) {
        Write-Host "Removing scheduled task: $Name" -ForegroundColor Yellow
        Unregister-ScheduledTask -TaskName $Name -Confirm:$false
        Write-Host "Scheduled task removed successfully." -ForegroundColor Green
    }
    else {
        Write-Host "Scheduled task '$Name' does not exist." -ForegroundColor Yellow
    }
}

function Install-MaintenanceTask {
    param(
        [string]$Name,
        [string]$Script,
        [string]$Day,
        [string]$RunTime
    )

    # Validate script exists
    if (-not (Test-Path $Script)) {
        throw "Script not found: $Script"
    }

    $scriptFullPath = (Resolve-Path $Script).Path

    # Check if task already exists
    if (Test-ScheduledTaskExists -Name $Name) {
        Write-Host "Scheduled task '$Name' already exists. Updating..." -ForegroundColor Yellow
        Unregister-ScheduledTask -TaskName $Name -Confirm:$false
    }

    # Get pwsh.exe path
    $pwshPath = (Get-Command pwsh.exe -ErrorAction Stop).Source

    # Build the action (hidden window for silent background execution)
    $actionArgument = "-NoProfile -NonInteractive -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$scriptFullPath`""
    $action = New-ScheduledTaskAction -Execute $pwshPath -Argument $actionArgument

    # Build the trigger (weekly)
    $trigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek $Day -At $RunTime

    # Build the principal (current user, run with highest privileges)
    # Using Interactive logon - requires user to be logged in, but StartWhenAvailable will catch up
    $currentUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
    $principal = New-ScheduledTaskPrincipal -UserId $currentUser -LogonType Interactive -RunLevel Highest

    # Build settings
    $settings = New-ScheduledTaskSettingsSet `
        -AllowStartIfOnBatteries `
        -DontStopIfGoingOnBatteries `
        -StartWhenAvailable `
        -RunOnlyIfNetworkAvailable `
        -ExecutionTimeLimit (New-TimeSpan -Hours 4)

    # Register the task
    $taskParams = @{
        TaskName    = $Name
        Action      = $action
        Trigger     = $trigger
        Principal   = $principal
        Settings    = $settings
        Description = "Weekly PowerShell module maintenance - updates modules and prunes old versions. https://github.com/haakonwibe/PSModuleMaintenance"
    }

    Register-ScheduledTask @taskParams | Out-Null

    Write-Host ""
    Write-Host "============================================" -ForegroundColor Cyan
    Write-Host " Scheduled Task Created Successfully! " -ForegroundColor Green
    Write-Host "============================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Task Name    : $Name"
    Write-Host "Script       : $scriptFullPath"
    Write-Host "Schedule     : Every $Day at $RunTime"
    Write-Host "Run As       : $currentUser"
    Write-Host "Elevated     : Yes"
    Write-Host ""
    Write-Host "Logs will be written to: $env:ProgramData\PSModuleMaintenance\Logs" -ForegroundColor Gray
    Write-Host ""
}

function Test-ManualRun {
    param([string]$Name)

    $response = Read-Host "Would you like to run the task now to verify it works? (y/N)"
    
    if ($response -eq 'y' -or $response -eq 'Y') {
        Write-Host "Starting task..." -ForegroundColor Yellow
        Start-ScheduledTask -TaskName $Name
        Write-Host "Task started. Check logs at: $env:ProgramData\PSModuleMaintenance\Logs" -ForegroundColor Green
    }
}

# ============================================================================
# MAIN
# ============================================================================

try {
    if ($Uninstall) {
        Uninstall-MaintenanceTask -Name $TaskName
        exit 0
    }

    # Determine script path
    if ([string]::IsNullOrEmpty($ScriptPath)) {
        $ScriptPath = Join-Path $PSScriptRoot 'Invoke-PSModuleMaintenance.ps1'
    }

    Write-Host ""
    Write-Host "PSModuleMaintenance Installer" -ForegroundColor Cyan
    Write-Host "=============================" -ForegroundColor Cyan
    Write-Host ""

    # Install the task
    Install-MaintenanceTask -Name $TaskName -Script $ScriptPath -Day $DayOfWeek -RunTime $Time

    # Offer test run
    Test-ManualRun -Name $TaskName
}
catch {
    Write-Host "ERROR: $_" -ForegroundColor Red
    exit 1
}
