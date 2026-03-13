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

.PARAMETER MigrateOnly
    Only migrate modules from OneDrive to AllUsers scope, skip updates and pruning.

.PARAMETER WhatIf
    Show what would be done without making changes.

.EXAMPLE
    .\Invoke-PSModuleMaintenance.ps1

.EXAMPLE
    .\Invoke-PSModuleMaintenance.ps1 -MigrateOnly

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
    [switch]$PruneOnly,

    [Parameter()]
    [switch]$MigrateOnly
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
    MigrateFromOneDrive = $true
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
            if ($null -ne $jsonConfig.MigrateFromOneDrive) {
                $script:Config.MigrateFromOneDrive = $jsonConfig.MigrateFromOneDrive
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
    ModulesMigrated = 0
    MigrationFailed = @()
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

        if ($Summary.ModulesMigrated -gt 0) {
            $parts += "Migrated $($Summary.ModulesMigrated) modules from OneDrive"
        }

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

function Test-OneDrivePath {
    <#
    .SYNOPSIS
        Tests whether a given path is inside a OneDrive-synced folder.
    #>
    [CmdletBinding()]
    param([string]$Path)

    $oneDrivePaths = @($env:OneDrive, $env:OneDriveCommercial, $env:OneDriveConsumer) | Where-Object { $_ }
    foreach ($odPath in $oneDrivePaths) {
        if ($Path -like "$odPath*") { return $true }
    }
    return $false
}

function Move-ModulesToAllUsers {
    <#
    .SYNOPSIS
        Copies modules from OneDrive-synced CurrentUser path to AllUsers scope.
        Does not delete source copies — the prune pass handles that.
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param()

    if (-not $script:Config.MigrateFromOneDrive) {
        Write-Log "Module migration is disabled in configuration"
        return
    }

    $currentUserModulePath = Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'PowerShell\Modules'

    if (-not (Test-OneDrivePath $currentUserModulePath)) {
        Write-Log "CurrentUser module path is not in OneDrive — no migration needed"
        return
    }

    if (-not (Test-Path $currentUserModulePath)) {
        Write-Log "CurrentUser module path does not exist: $currentUserModulePath"
        return
    }

    $allUsersPath = Join-Path $env:ProgramFiles 'PowerShell\Modules'
    Write-Log "Migrating modules from OneDrive to AllUsers scope"
    Write-Log "  Source: $currentUserModulePath"
    Write-Log "  Destination: $allUsersPath"

    $moduleFolders = Get-ChildItem -Path $currentUserModulePath -Directory -ErrorAction SilentlyContinue
    if (-not $moduleFolders) {
        Write-Log "No modules found in CurrentUser path"
        return
    }

    foreach ($moduleFolder in $moduleFolders) {
        $versionFolders = Get-ChildItem -Path $moduleFolder.FullName -Directory -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -match '^\d+(\.\d+){1,3}$' }

        # Some modules don't use version subfolders — check for a manifest directly
        if (-not $versionFolders) {
            $manifest = Get-ChildItem -Path $moduleFolder.FullName -Filter '*.psd1' -ErrorAction SilentlyContinue | Select-Object -First 1
            if (-not $manifest) { continue }

            # Treat the module folder itself as the item to copy
            $destPath = Join-Path $allUsersPath $moduleFolder.Name
            if (Test-Path $destPath) {
                Write-Verbose "Already exists in AllUsers (no version folder): $($moduleFolder.Name)"
                continue
            }

            if ($PSCmdlet.ShouldProcess($moduleFolder.Name, "Copy module to AllUsers")) {
                try {
                    Copy-Item -Path $moduleFolder.FullName -Destination $destPath -Recurse -Force -ErrorAction Stop
                    $script:Summary.ModulesMigrated++
                    Write-Log "Migrated: $($moduleFolder.Name)" -Level SUCCESS
                }
                catch {
                    Write-Log "Failed to migrate $($moduleFolder.Name): $_" -Level ERROR
                    $script:Summary.MigrationFailed += @{
                        Module = $moduleFolder.Name
                        Error  = $_.Exception.Message
                    }
                }
            }
            continue
        }

        foreach ($versionFolder in $versionFolders) {
            $destPath = Join-Path $allUsersPath "$($moduleFolder.Name)\$($versionFolder.Name)"

            if (Test-Path $destPath) {
                Write-Verbose "Already exists in AllUsers: $($moduleFolder.Name) v$($versionFolder.Name)"
                continue
            }

            if ($PSCmdlet.ShouldProcess("$($moduleFolder.Name) v$($versionFolder.Name)", "Copy module to AllUsers")) {
                try {
                    $destParent = Join-Path $allUsersPath $moduleFolder.Name
                    if (-not (Test-Path $destParent)) {
                        New-Item -Path $destParent -ItemType Directory -Force | Out-Null
                    }

                    Copy-Item -Path $versionFolder.FullName -Destination $destPath -Recurse -Force -ErrorAction Stop
                    $script:Summary.ModulesMigrated++
                    Write-Log "Migrated: $($moduleFolder.Name) v$($versionFolder.Name)" -Level SUCCESS
                }
                catch {
                    Write-Log "Failed to migrate $($moduleFolder.Name) v$($versionFolder.Name): $_" -Level ERROR
                    $script:Summary.MigrationFailed += @{
                        Module  = $moduleFolder.Name
                        Version = $versionFolder.Name
                        Error   = $_.Exception.Message
                    }
                }
            }
        }
    }

    Write-Log "Migration complete. Copied: $($script:Summary.ModulesMigrated), Failed: $($script:Summary.MigrationFailed.Count)"
    if ($script:Summary.ModulesMigrated -gt 0) {
        Write-Log "OneDrive copies will be cleaned up during the prune pass"
    }
}

function Remove-LockedModuleFolder {
    <#
    .SYNOPSIS
        Force-removes a module folder that OneDrive has locked or converted to cloud placeholders.
        Escalation: normal Remove-Item → strip cloud attributes + rd /s /q → MoveFileEx reboot deletion.
        Only operates on folders inside a known PSModulePath with a version-number name.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$FolderPath
    )

    if (-not (Test-Path $FolderPath)) {
        return $true
    }

    # Safety: folder name must look like a version number (e.g. 1.11.0)
    $folderName = Split-Path $FolderPath -Leaf
    if ($folderName -notmatch '^\d+(\.\d+){1,3}$') {
        Write-Log "Refusing force-removal: '$folderName' is not a version folder" -Level ERROR
        return $false
    }

    # Safety: path must be inside one of the known PSModulePath directories
    $modulePaths = $env:PSModulePath -split [System.IO.Path]::PathSeparator
    $isInsideModulePath = $false
    foreach ($mp in $modulePaths) {
        if ($mp -and $FolderPath -like "$mp*") {
            $isInsideModulePath = $true
            break
        }
    }
    if (-not $isInsideModulePath) {
        Write-Log "Refusing force-removal: path is not inside any PSModulePath directory" -Level ERROR
        return $false
    }

    Write-Log "Force-removing locked folder: $FolderPath" -Level WARN

    # Release any .NET assembly locks the current session may hold
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()

    # Strip read-only, hidden, and system attributes from all files so they can be deleted
    Get-ChildItem -Path $FolderPath -Recurse -Force -ErrorAction SilentlyContinue | ForEach-Object {
        try {
            [System.IO.File]::SetAttributes($_.FullName, [System.IO.FileAttributes]::Normal)
        }
        catch {
            # Directories, reparse points, or already-deleted files — ignore
        }
    }

    # First attempt: remove the whole tree at once
    try {
        Remove-Item -Path $FolderPath -Recurse -Force -ErrorAction Stop
        Write-Log "Force-removed folder: $FolderPath" -Level SUCCESS
        return $true
    }
    catch {
        Write-Log "Bulk Remove-Item failed: $_" -Level WARN
    }

    # Second attempt: strip OneDrive cloud-file attributes (Pinned/Unpinned/Offline),
    # then use cmd.exe rd /s /q which handles reparse points differently than Remove-Item
    $reparsePoints = Get-ChildItem -Path $FolderPath -Recurse -Force -ErrorAction SilentlyContinue |
        Where-Object { $_.Attributes -band [System.IO.FileAttributes]::ReparsePoint }
    if ($reparsePoints) {
        Write-Log "Found $($reparsePoints.Count) OneDrive cloud placeholder(s) — stripping cloud attributes" -Level WARN
        foreach ($rp in $reparsePoints) {
            & cmd.exe /c "attrib -P -U -O `"$($rp.FullName)`"" 2>&1 | Out-Null
        }
    }
    & cmd.exe /c "rd /s /q `"$FolderPath`"" 2>&1 | Out-Null
    if (-not (Test-Path $FolderPath)) {
        Write-Log "Force-removed folder (rd /s /q after cloud attribute strip): $FolderPath" -Level SUCCESS
        return $true
    }
    Write-Log "rd /s /q could not fully remove folder — trying file-by-file + reboot fallback" -Level WARN

    # Third attempt: delete what we can file-by-file, schedule the rest for reboot
    Get-ChildItem -Path $FolderPath -Recurse -Force -File -ErrorAction SilentlyContinue |
        Remove-Item -Force -ErrorAction SilentlyContinue

    # Check if that was enough
    if (-not (Test-Path $FolderPath) -or
        @(Get-ChildItem -Path $FolderPath -Recurse -Force -ErrorAction SilentlyContinue).Count -eq 0) {
        Remove-Item -Path $FolderPath -Force -Recurse -ErrorAction SilentlyContinue
        Write-Log "Force-removed folder (file-by-file): $FolderPath" -Level SUCCESS
        return $true
    }

    # Fourth attempt: schedule remaining files for deletion on next reboot via kernel32 MoveFileEx
    if (-not ('PSModuleMaintenance.FileUtils' -as [type])) {
        Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
namespace PSModuleMaintenance {
    public class FileUtils {
        [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        public static extern bool MoveFileEx(string lpExistingFileName, string lpNewFileName, int dwFlags);
        public const int MOVEFILE_DELAY_UNTIL_REBOOT = 0x4;
    }
}
"@
    }

    $scheduledCount = 0
    $remainingFiles = Get-ChildItem -Path $FolderPath -Recurse -Force -File -ErrorAction SilentlyContinue
    foreach ($file in $remainingFiles) {
        $scheduled = [PSModuleMaintenance.FileUtils]::MoveFileEx(
            $file.FullName, $null, [PSModuleMaintenance.FileUtils]::MOVEFILE_DELAY_UNTIL_REBOOT)
        if ($scheduled) {
            $scheduledCount++
        }
        else {
            $win32Error = [System.Runtime.InteropServices.Marshal]::GetLastWin32Error()
            Write-Log "MoveFileEx failed for $($file.FullName) — Win32 error $win32Error" -Level WARN
        }
    }

    # Schedule directories for reboot deletion (deepest first so children go before parents)
    $remainingDirs = Get-ChildItem -Path $FolderPath -Recurse -Force -Directory -ErrorAction SilentlyContinue |
        Sort-Object { $_.FullName.Length } -Descending
    foreach ($dir in $remainingDirs) {
        [PSModuleMaintenance.FileUtils]::MoveFileEx(
            $dir.FullName, $null, [PSModuleMaintenance.FileUtils]::MOVEFILE_DELAY_UNTIL_REBOOT) | Out-Null
    }
    # Schedule the root version folder itself
    [PSModuleMaintenance.FileUtils]::MoveFileEx(
        $FolderPath, $null, [PSModuleMaintenance.FileUtils]::MOVEFILE_DELAY_UNTIL_REBOOT) | Out-Null

    if ($scheduledCount -gt 0) {
        Write-Log "Scheduled $scheduledCount locked file(s) for deletion on next reboot: $FolderPath" -Level WARN
        return $true
    }
    else {
        Write-Log "Could not force-remove or schedule folder for reboot deletion: $FolderPath" -Level ERROR
        return $false
    }
}

function Update-AllModules {
    [CmdletBinding(SupportsShouldProcess)]
    param()

    Write-Log "Starting module updates..."

    # Detect OneDrive on module path — use AllUsers scope to avoid file locking
    $currentUserModulePath = Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'PowerShell\Modules'
    $useAllUsersScope = Test-OneDrivePath $currentUserModulePath
    if ($useAllUsersScope) {
        Write-Log "OneDrive detected on module path — updates will target AllUsers scope"
    }

    # Get installed modules (newest version of each)
    # When OneDrive is detected, modules live in AllUsers scope — query that explicitly
    $getParams = if ($useAllUsersScope) { @{ Scope = 'AllUsers' } } else { @{} }
    $installed = Get-PSResource @getParams |
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
                if ($useAllUsersScope) {
                    $updateParams.Scope = 'AllUsers'
                }

                Update-PSResource @updateParams
                $script:Summary.ModulesUpdated++
                Write-Log "Updated: $($module.Name)" -Level SUCCESS
            }
        }
        catch {
            $errorMsg = $_.Exception.Message

            # Detect OneDrive-locked folders: "Cannot remove package path <path>"
            # Loop because meta-packages like Az can hit multiple locked sub-module folders
            $maxAttempts = 20
            $attempt = 0
            $updated = $false

            while ($errorMsg -match 'Cannot remove package path\s+(.+?)\.?\s*(The previous|$)') {
                $attempt++
                if ($attempt -gt $maxAttempts) {
                    Write-Log "Reached max force-removal attempts ($maxAttempts) for $($module.Name)" -Level ERROR
                    break
                }

                $lockedPath = $Matches[1].TrimEnd('. ')
                if (-not (Test-Path $lockedPath)) { break }

                Write-Log "Locked folder detected ($attempt): $lockedPath — force-removing" -Level WARN
                if (-not (Remove-LockedModuleFolder -FolderPath $lockedPath)) { break }

                try {
                    Update-PSResource @updateParams
                    $script:Summary.ModulesUpdated++
                    Write-Log "Updated (after $attempt force-removal(s)): $($module.Name)" -Level SUCCESS
                    $updated = $true
                    break
                }
                catch {
                    $errorMsg = $_.Exception.Message
                }
            }

            if (-not $updated) {
                Write-Log "Failed to update $($module.Name): $errorMsg" -Level ERROR
                $script:Summary.ModulesFailed += @{
                    Module = $module.Name
                    Error  = $errorMsg
                }
            }
        }
    }

    Write-Log "Module updates complete. Updated: $($script:Summary.ModulesUpdated), Failed: $($script:Summary.ModulesFailed.Count)"
}

function Remove-OldModuleVersions {
    [CmdletBinding(SupportsShouldProcess)]
    param()

    Write-Log "Starting old version cleanup..."

    # Detect OneDrive on module path
    $currentUserModulePath = Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'PowerShell\Modules'
    $isOneDrive = Test-OneDrivePath $currentUserModulePath

    # When OneDrive is detected, modules live in AllUsers scope — query that explicitly
    $getParams = if ($isOneDrive) { @{ Scope = 'AllUsers' } } else { @{} }
    $allModules = Get-PSResource @getParams | Where-Object { $_.Name -notin $script:Config.ExcludedModules }

    # --- Pass 1: Remove old versions of AllUsers modules (keep latest) ---
    if ($isOneDrive) {
        # Only prune modules NOT in OneDrive (AllUsers modules)
        $nonOneDriveModules = $allModules | Where-Object { -not ($_.InstalledLocation -like "$currentUserModulePath*") }
    }
    else {
        $nonOneDriveModules = $allModules
    }

    # Filter out built-in modules that ship with PowerShell (e.g. PackageManagement) —
    # PSResourceGet cannot uninstall these and always errors
    $psHomePath = $PSHOME
    $nonOneDriveModules = $nonOneDriveModules | Where-Object {
        -not ($_.InstalledLocation -like "$psHomePath*")
    }

    $grouped = $nonOneDriveModules | Group-Object Name | Where-Object { $_.Count -gt 1 }
    Write-Log "Found $($grouped.Count) modules with multiple versions"

    foreach ($group in $grouped) {
        $oldVersions = $group.Group | Sort-Object Version -Descending | Select-Object -Skip 1

        foreach ($oldVersion in $oldVersions) {
            if ($PSCmdlet.ShouldProcess("$($oldVersion.Name) v$($oldVersion.Version)", "Remove old version")) {
                Write-Log "Removing: $($oldVersion.Name) v$($oldVersion.Version)"

                try {
                    $uninstallParams = @{
                        Name                = $oldVersion.Name
                        Version             = $oldVersion.Version
                        SkipDependencyCheck = $true
                        ErrorAction         = 'Stop'
                    }
                    if ($isOneDrive) { $uninstallParams['Scope'] = 'AllUsers' }
                    Uninstall-PSResource @uninstallParams
                    $script:Summary.VersionsPruned++
                    Write-Log "Removed: $($oldVersion.Name) v$($oldVersion.Version)" -Level SUCCESS
                }
                catch {
                    $errorMsg = $_.Exception.Message

                    # "does not exist" = built-in module or phantom metadata — skip silently
                    if ($errorMsg -match 'does not exist') {
                        Write-Log "Skipping $($oldVersion.Name) v$($oldVersion.Version): not managed by PSResourceGet" -Level WARN
                        continue
                    }

                    # If access denied / cannot delete, try force-removing the folder directly
                    $folderPath = $oldVersion.InstalledLocation
                    if ($folderPath -and (Test-Path $folderPath) -and
                        ($errorMsg -match 'Access.*denied|could not be deleted|Cannot remove')) {
                        Write-Log "Lock detected — attempting force-removal of $folderPath" -Level WARN
                        if (Remove-LockedModuleFolder -FolderPath $folderPath) {
                            $script:Summary.VersionsPruned++
                            Write-Log "Force-removed: $($oldVersion.Name) v$($oldVersion.Version)" -Level SUCCESS
                            continue
                        }
                    }

                    Write-Log "Failed to remove $($oldVersion.Name) v$($oldVersion.Version): $errorMsg" -Level ERROR
                    $script:Summary.PrunesFailed += @{
                        Module  = $oldVersion.Name
                        Version = $oldVersion.Version.ToString()
                        Error   = $errorMsg
                    }
                }
            }
        }
    }

    # --- Pass 2: Remove ALL migrated copies from OneDrive path ---
    # Uses direct folder deletion instead of Uninstall-PSResource to avoid
    # accidentally removing AllUsers copies (Uninstall-PSResource without -Scope
    # can target any scope where the module is found).
    if ($isOneDrive -and (Test-Path $currentUserModulePath)) {
        $odModuleFolders = Get-ChildItem -Path $currentUserModulePath -Directory -ErrorAction SilentlyContinue

        if ($odModuleFolders.Count -gt 0) {
            Write-Log "Cleaning up $($odModuleFolders.Count) module folder(s) from OneDrive path"

            foreach ($moduleFolder in $odModuleFolders) {
                $versionFolders = Get-ChildItem -Path $moduleFolder.FullName -Directory -ErrorAction SilentlyContinue |
                    Where-Object { $_.Name -match '^\d+(\.\d+){1,3}$' }

                foreach ($versionFolder in $versionFolders) {
                    $moduleName = $moduleFolder.Name
                    $versionName = $versionFolder.Name

                    if ($PSCmdlet.ShouldProcess("$moduleName v$versionName (OneDrive copy)", "Remove migrated copy")) {
                        Write-Log "Removing OneDrive copy: $moduleName v$versionName"

                        try {
                            Remove-Item -Path $versionFolder.FullName -Recurse -Force -ErrorAction Stop
                            $script:Summary.VersionsPruned++
                            Write-Log "Removed OneDrive copy: $moduleName v$versionName" -Level SUCCESS
                        }
                        catch {
                            Write-Log "OneDrive lock on $moduleName — attempting force-removal" -Level WARN
                            if (Remove-LockedModuleFolder -FolderPath $versionFolder.FullName) {
                                $script:Summary.VersionsPruned++
                                Write-Log "Force-removed OneDrive copy: $moduleName v$versionName" -Level SUCCESS
                            }
                            else {
                                Write-Log "Failed to remove OneDrive copy $moduleName v${versionName}: $($_.Exception.Message)" -Level WARN
                                $script:Summary.PrunesFailed += @{
                                    Module  = $moduleName
                                    Version = $versionName
                                    Error   = $_.Exception.Message
                                }
                            }
                        }
                    }
                }

                # Clean up empty module folder after all versions removed
                if ((Test-Path $moduleFolder.FullName) -and
                    @(Get-ChildItem -Path $moduleFolder.FullName -Force -ErrorAction SilentlyContinue).Count -eq 0) {
                    Remove-Item -Path $moduleFolder.FullName -Force -ErrorAction SilentlyContinue
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
    Write-Log "  - Migrate from OneDrive: $($script:Config.MigrateFromOneDrive)"

    # Clean up old logs
    Remove-OldLogs -BasePath $LogPath -RetentionDays $script:Config.LogRetentionDays

    # Perform operations
    if ($MigrateOnly) {
        Move-ModulesToAllUsers
    }
    else {
        if (-not $PruneOnly) {
            Move-ModulesToAllUsers
            Update-AllModules
        }

        if (-not $UpdateOnly) {
            Remove-OldModuleVersions
        }
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
    $hasFailures = ($script:Summary.ModulesFailed.Count -gt 0) -or ($script:Summary.PrunesFailed.Count -gt 0) -or ($script:Summary.MigrationFailed.Count -gt 0)

    if ($notifyMode -eq 'Always' -or ($notifyMode -eq 'OnFailure' -and $hasFailures)) {
        Send-ToastNotification -Summary $script:Summary -SkippedUpdates:($PruneOnly -or $MigrateOnly) -SkippedPruning:($UpdateOnly -or $MigrateOnly)
    }

    # Stop transcript
    Stop-Transcript -ErrorAction SilentlyContinue | Out-Null
}
