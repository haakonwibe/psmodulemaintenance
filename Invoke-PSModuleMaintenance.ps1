<#
.SYNOPSIS
    Automated PowerShell module maintenance - updates modules and prunes old versions.

.DESCRIPTION
    This script performs automated maintenance of PowerShell modules installed via PSResourceGet:
    - Updates all installed modules to their latest versions
    - Removes old versions while keeping the latest
    - Logs all operations with configurable retention
    - Supports module exclusions via config file

.PARAMETER ConfigPath
    Path to the JSON configuration file. Defaults to script directory's config.json.

.PARAMETER LogPath
    Base path for logs. Defaults to $env:ProgramData\PSModuleMaintenance\Logs

.PARAMETER UpdateOnly
    Only perform updates, skip pruning old versions.

.PARAMETER PruneOnly
    Only prune old versions, skip updates.

.PARAMETER WhatIf
    Show what would be done without making changes.

.EXAMPLE
    .\Invoke-PSModuleMaintenance.ps1

.EXAMPLE
    .\Invoke-PSModuleMaintenance.ps1 -PruneOnly -WhatIf

.NOTES
    Author: Haakon Wibe
    Requires: PowerShell 7+, Microsoft.PowerShell.PSResourceGet
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter()]
    [string]$ConfigPath,

    [Parameter()]
    [string]$LogPath = "$env:ProgramData\PSModuleMaintenance\Logs",

    [Parameter()]
    [switch]$UpdateOnly,

    [Parameter()]
    [switch]$PruneOnly
)

#Requires -Version 7.0
#Requires -Modules Microsoft.PowerShell.PSResourceGet

# ============================================================================
# CONFIGURATION
# ============================================================================

$script:Config = @{
    ExcludedModules = @()
    LogRetentionDays = 180
    TrustPSGallery = $true
    NotificationMode = 'Always'
}

function Import-MaintenanceConfig {
    [CmdletBinding()]
    param([string]$Path)

    if ([string]::IsNullOrEmpty($Path)) {
        $Path = Join-Path $PSScriptRoot 'config.json'
    }

    if (Test-Path $Path) {
        try {
            $jsonConfig = Get-Content $Path -Raw | ConvertFrom-Json
            
            if ($jsonConfig.ExcludedModules) {
                $script:Config.ExcludedModules = @($jsonConfig.ExcludedModules)
            }
            if ($null -ne $jsonConfig.LogRetentionDays) {
                $script:Config.LogRetentionDays = $jsonConfig.LogRetentionDays
            }
            if ($null -ne $jsonConfig.TrustPSGallery) {
                $script:Config.TrustPSGallery = $jsonConfig.TrustPSGallery
            }
            if ($null -ne $jsonConfig.NotificationMode) {
                $script:Config.NotificationMode = $jsonConfig.NotificationMode
            }

            Write-Verbose "Loaded configuration from: $Path"
        }
        catch {
            Write-Warning "Failed to parse config file: $_. Using defaults."
        }
    }
    else {
        Write-Verbose "No config file found at $Path. Using defaults."
    }
}

# ============================================================================
# LOGGING
# ============================================================================

$script:LogFile = $null
$script:TranscriptFile = $null
$script:Summary = @{
    StartTime = $null
    EndTime = $null
    ModulesChecked = 0
    ModulesUpdated = 0
    ModulesFailed = @()
    VersionsPruned = 0
    PrunesFailed = @()
    ExcludedModules = @()
}

function Initialize-Logging {
    [CmdletBinding()]
    param([string]$BasePath)

    # Ensure log directory exists
    if (-not (Test-Path $BasePath)) {
        New-Item -Path $BasePath -ItemType Directory -Force | Out-Null
    }

    $timestamp = Get-Date -Format 'yyyy-MM-dd_HHmmss'
    $script:LogFile = Join-Path $BasePath "maintenance_$timestamp.log"
    $script:TranscriptFile = Join-Path $BasePath "transcript_$timestamp.log"

    # Start transcript for full verbose capture
    Start-Transcript -Path $script:TranscriptFile -Force | Out-Null

    Write-Log "======================================================"
    Write-Log "PSModuleMaintenance started"
    Write-Log "======================================================"
    Write-Log "PowerShell Version: $($PSVersionTable.PSVersion)"
    Write-Log "Running as: $([System.Security.Principal.WindowsIdentity]::GetCurrent().Name)"
    Write-Log "Log file: $script:LogFile"
}

function Write-Log {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, Position = 0)]
        [string]$Message,

        [Parameter()]
        [ValidateSet('INFO', 'WARN', 'ERROR', 'SUCCESS')]
        [string]$Level = 'INFO'
    )

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $logEntry = "[$timestamp] [$Level] $Message"

    # Write to log file
    if ($script:LogFile) {
        Add-Content -Path $script:LogFile -Value $logEntry
    }

    # Write to console with colors
    switch ($Level) {
        'WARN'    { Write-Host $logEntry -ForegroundColor Yellow }
        'ERROR'   { Write-Host $logEntry -ForegroundColor Red }
        'SUCCESS' { Write-Host $logEntry -ForegroundColor Green }
        default   { Write-Host $logEntry }
    }
}

function Save-Summary {
    [CmdletBinding()]
    param([string]$BasePath)

    $script:Summary.EndTime = Get-Date -Format 'o'
    
    $summaryFile = Join-Path $BasePath "summary_$(Get-Date -Format 'yyyy-MM-dd_HHmmss').json"
    $script:Summary | ConvertTo-Json -Depth 3 | Set-Content -Path $summaryFile

    Write-Log "Summary saved to: $summaryFile"
}

function Remove-OldLogs {
    [CmdletBinding()]
    param(
        [string]$BasePath,
        [int]$RetentionDays
    )

    if (-not (Test-Path $BasePath)) {
        return
    }

    $cutoffDate = (Get-Date).AddDays(-$RetentionDays)
    $oldFiles = Get-ChildItem -Path $BasePath -File | Where-Object { $_.LastWriteTime -lt $cutoffDate }

    if ($oldFiles) {
        Write-Log "Removing $($oldFiles.Count) log files older than $RetentionDays days"
        $oldFiles | Remove-Item -Force -ErrorAction SilentlyContinue
    }
}

# ============================================================================
# TOAST NOTIFICATIONS
# ============================================================================

function Send-ToastNotification {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Summary,

        [switch]$SkippedUpdates,

        [switch]$SkippedPruning
    )

    try {
        # Build the notification message from summary data
        $parts = @()

        if (-not $SkippedUpdates) {
            $updateText = "Updated $($Summary.ModulesUpdated) modules"
            if ($Summary.ModulesFailed.Count -gt 0) {
                $updateText += ", $($Summary.ModulesFailed.Count) failed"
            }
            $parts += $updateText
        }

        if (-not $SkippedPruning) {
            $pruneText = "Pruned $($Summary.VersionsPruned) versions"
            if ($Summary.PrunesFailed.Count -gt 0) {
                $pruneText += ", $($Summary.PrunesFailed.Count) failed"
            }
            $parts += $pruneText
        }

        $hasFailures = ($Summary.ModulesFailed.Count -gt 0) -or ($Summary.PrunesFailed.Count -gt 0)

        $message = ($parts -join '. ') + '.'
        if (-not $hasFailures) {
            $message += ' No errors.'
        }

        # Use Windows PowerShell (5.1) for native WinRT toast support — always present on Windows 10/11
        $xmlMessage = [System.Security.SecurityElement]::Escape($message)

        $toastScript = @"
`$ErrorActionPreference = 'Stop'
try {
    [Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime] | Out-Null
    [Windows.Data.Xml.Dom.XmlDocument, Windows.Data.Xml.Dom.XmlDocument, ContentType = WindowsRuntime] | Out-Null

    `$toastXml = @'
<toast>
    <visual>
        <binding template="ToastGeneric">
            <text>PSModuleMaintenance</text>
            <text>$xmlMessage</text>
        </binding>
    </visual>
    <audio src="ms-winsoundevent:Notification.Default"/>
</toast>
'@

    `$xmlDoc = New-Object Windows.Data.Xml.Dom.XmlDocument
    `$xmlDoc.LoadXml(`$toastXml)
    `$toast = New-Object Windows.UI.Notifications.ToastNotification(`$xmlDoc)
    `$notifier = [Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier('Windows.SystemToast.SecurityAndMaintenance')
    `$notifier.Show(`$toast)
}
catch {
    exit 1
}
"@

        $encoded = [Convert]::ToBase64String([System.Text.Encoding]::Unicode.GetBytes($toastScript))
        powershell.exe -NoProfile -NonInteractive -WindowStyle Hidden -EncodedCommand $encoded

        if ($LASTEXITCODE -eq 0) {
            Write-Log "Toast notification sent: $message"
        }
        else {
            Write-Log "Toast notification subprocess exited with code $LASTEXITCODE" -Level WARN
        }
    }
    catch {
        Write-Log "Toast notification failed: $_" -Level WARN
    }
}

# ============================================================================
# MODULE OPERATIONS
# ============================================================================

function Update-AllModules {
    [CmdletBinding(SupportsShouldProcess)]
    param()

    Write-Log "Starting module updates..."

    # Get installed modules (newest version of each)
    $installed = Get-PSResource |
        Where-Object { $_.Name -notin $script:Config.ExcludedModules } |
        Group-Object Name |
        ForEach-Object {
            $newest = $_.Group | Sort-Object Version -Descending | Select-Object -First 1
            [PSCustomObject]@{ Name = $newest.Name; Version = $newest.Version }
        }

    $script:Summary.ModulesChecked = $installed.Count
    $script:Summary.ExcludedModules = @($script:Config.ExcludedModules)

    Write-Log "Found $($installed.Count) installed modules (excluding: $($script:Config.ExcludedModules -join ', '))"
    Write-Log "Checking PSGallery for available updates..."

    # Query PSGallery for latest versions (bulk request)
    try {
        $gallery = Find-PSResource -Name $installed.Name -Repository PSGallery -ErrorAction SilentlyContinue
    }
    catch {
        Write-Log "Failed to query PSGallery: $_" -Level ERROR
        return
    }

    # Find modules that need updates
    $needsUpdate = @()
    foreach ($mod in $installed) {
        $galleryVersion = ($gallery | Where-Object Name -eq $mod.Name | Sort-Object Version -Descending | Select-Object -First 1).Version
        # Normalize versions to 4-part form so 6.1907.1 and 6.1907.1.0 compare as equal
        if ($galleryVersion -and $mod.Version) {
            $normalizedInstalled = [version]::new($mod.Version.Major, $mod.Version.Minor, [Math]::Max($mod.Version.Build, 0), [Math]::Max($mod.Version.Revision, 0))
            $normalizedGallery = [version]::new($galleryVersion.Major, $galleryVersion.Minor, [Math]::Max($galleryVersion.Build, 0), [Math]::Max($galleryVersion.Revision, 0))
        }
        if ($galleryVersion -and $normalizedGallery -gt $normalizedInstalled) {
            $needsUpdate += [PSCustomObject]@{
                Name = $mod.Name
                InstalledVersion = $mod.Version
                GalleryVersion = $galleryVersion
            }
        }
    }

    if ($needsUpdate.Count -eq 0) {
        Write-Log "All modules are up to date"
        return
    }

    Write-Log "Found $($needsUpdate.Count) modules with available updates"

    # Update only modules that need it
    foreach ($module in $needsUpdate) {
        try {
            if ($PSCmdlet.ShouldProcess("$($module.Name) $($module.InstalledVersion) -> $($module.GalleryVersion)", "Update module")) {
                Write-Log "Updating: $($module.Name) ($($module.InstalledVersion) -> $($module.GalleryVersion))"

                $updateParams = @{
                    Name            = $module.Name
                    AcceptLicense   = $true
                    TrustRepository = $script:Config.TrustPSGallery
                    ErrorAction     = 'Stop'
                }

                Update-PSResource @updateParams
                $script:Summary.ModulesUpdated++
                Write-Log "Updated: $($module.Name)" -Level SUCCESS
            }
        }
        catch {
            Write-Log "Failed to update $($module.Name): $_" -Level ERROR
            $script:Summary.ModulesFailed += @{
                Module = $module.Name
                Error  = $_.Exception.Message
            }
        }
    }

    Write-Log "Module updates complete. Updated: $($script:Summary.ModulesUpdated), Failed: $($script:Summary.ModulesFailed.Count)"
}

function Remove-OldModuleVersions {
    [CmdletBinding(SupportsShouldProcess)]
    param()

    Write-Log "Starting old version cleanup..."

    $allModules = Get-PSResource | Where-Object { $_.Name -notin $script:Config.ExcludedModules }

    # Check if modules are in OneDrive-synced folder
    $oneDrivePaths = @($env:OneDrive, $env:OneDriveCommercial, $env:OneDriveConsumer) | Where-Object { $_ }
    $sampleModule = $allModules | Select-Object -First 1
    if ($sampleModule) {
        $modulePath = Split-Path $sampleModule.InstalledLocation -Parent
        $isOneDrive = $oneDrivePaths | Where-Object { $modulePath -like "$_*" }
        if ($isOneDrive) {
            Write-Log "Modules are in OneDrive-synced folder. Some deletions may require confirmation in OneDrive popup." -Level WARN
        }
    }

    $grouped = $allModules | Group-Object Name | Where-Object { $_.Count -gt 1 }

    Write-Log "Found $($grouped.Count) modules with multiple versions"

    foreach ($group in $grouped) {
        $oldVersions = $group.Group | Sort-Object Version -Descending | Select-Object -Skip 1

        foreach ($oldVersion in $oldVersions) {
            if ($PSCmdlet.ShouldProcess("$($oldVersion.Name) v$($oldVersion.Version)", "Remove old version")) {
                Write-Log "Removing: $($oldVersion.Name) v$($oldVersion.Version)"

                try {
                    Uninstall-PSResource -Name $oldVersion.Name -Version $oldVersion.Version -SkipDependencyCheck -ErrorAction Stop
                    $script:Summary.VersionsPruned++
                    Write-Log "Removed: $($oldVersion.Name) v$($oldVersion.Version)" -Level SUCCESS
                }
                catch {
                    Write-Log "Failed to remove $($oldVersion.Name) v$($oldVersion.Version): $_" -Level ERROR
                    $script:Summary.PrunesFailed += @{
                        Module  = $oldVersion.Name
                        Version = $oldVersion.Version.ToString()
                        Error   = $_.Exception.Message
                    }
                }
            }
        }
    }

    Write-Log "Version cleanup complete. Removed: $($script:Summary.VersionsPruned), Failed: $($script:Summary.PrunesFailed.Count)"
}

# ============================================================================
# MAIN EXECUTION
# ============================================================================

try {
    $script:Summary.StartTime = Get-Date -Format 'o'

    # Load configuration
    Import-MaintenanceConfig -Path $ConfigPath

    # Initialize logging
    Initialize-Logging -BasePath $LogPath

    # Log configuration
    Write-Log "Configuration loaded:"
    Write-Log "  - Excluded modules: $($script:Config.ExcludedModules.Count)"
    Write-Log "  - Log retention: $($script:Config.LogRetentionDays) days"
    Write-Log "  - Trust PSGallery: $($script:Config.TrustPSGallery)"
    Write-Log "  - Notification mode: $($script:Config.NotificationMode)"

    # Clean up old logs
    Remove-OldLogs -BasePath $LogPath -RetentionDays $script:Config.LogRetentionDays

    # Perform operations
    if (-not $PruneOnly) {
        Update-AllModules
    }

    if (-not $UpdateOnly) {
        Remove-OldModuleVersions
    }

    Write-Log "======================================================"
    Write-Log "PSModuleMaintenance completed successfully" -Level SUCCESS
    Write-Log "======================================================"
}
catch {
    Write-Log "Critical error: $_" -Level ERROR
    throw
}
finally {
    # Save summary
    Save-Summary -BasePath $LogPath

    # Send toast notification based on config
    $notifyMode = $script:Config.NotificationMode
    $hasFailures = ($script:Summary.ModulesFailed.Count -gt 0) -or ($script:Summary.PrunesFailed.Count -gt 0)

    if ($notifyMode -eq 'Always' -or ($notifyMode -eq 'OnFailure' -and $hasFailures)) {
        Send-ToastNotification -Summary $script:Summary -SkippedUpdates:$PruneOnly -SkippedPruning:$UpdateOnly
    }

    # Stop transcript
    Stop-Transcript -ErrorAction SilentlyContinue | Out-Null
}
