<#
.SYNOPSIS
    Migrates PowerShell modules from OneDrive-synced CurrentUser path to AllUsers scope.

.DESCRIPTION
    When Windows Known Folder Move redirects Documents to OneDrive, PowerShell's CurrentUser
    module path ends up inside OneDrive. This causes sync conflicts, cloud placeholders, and
    file locks that break module updates and pruning.

    This script detects the OneDrive redirect, copies all CurrentUser modules to AllUsers scope
    ($env:ProgramFiles\PowerShell\Modules), then cleans up the OneDrive copies using a
    four-stage escalation for stubborn locked/cloud-placeholder files.

    This is a one-time migration. After running it, the weekly maintenance script
    (Invoke-PSModuleMaintenance.ps1) automatically detects OneDrive and targets AllUsers scope
    for all future updates and pruning.

.PARAMETER LogPath
    Base path for the log file. Defaults to $env:ProgramData\PSModuleMaintenance\Logs.

.PARAMETER SkipCleanup
    Copy modules to AllUsers but do not delete the OneDrive copies.

.PARAMETER WhatIf
    Show what would be done without making changes.

.EXAMPLE
    .\Invoke-OneDriveMigration.ps1 -WhatIf

.EXAMPLE
    .\Invoke-OneDriveMigration.ps1

.EXAMPLE
    .\Invoke-OneDriveMigration.ps1 -SkipCleanup

.NOTES
    Author: Haakon Wibe
    Requires: PowerShell 7+, Administrator privileges (AllUsers scope writes to Program Files)
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter()]
    [string]$LogPath = "$env:ProgramData\PSModuleMaintenance\Logs",

    [Parameter()]
    [switch]$SkipCleanup
)

#Requires -Version 7.0
#Requires -RunAsAdministrator

# ============================================================================
# LOGGING
# ============================================================================

$script:LogFile = $null

function Write-Log {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, Position = 0)]
        [string]$Message,

        [ValidateSet('INFO', 'WARN', 'ERROR', 'SUCCESS')]
        [string]$Level = 'INFO'
    )

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $entry = "[$timestamp] [$Level] $Message"

    switch ($Level) {
        'ERROR'   { Write-Host $entry -ForegroundColor Red }
        'WARN'    { Write-Host $entry -ForegroundColor Yellow }
        'SUCCESS' { Write-Host $entry -ForegroundColor Green }
        default   { Write-Host $entry }
    }

    if ($script:LogFile) {
        $entry | Out-File -FilePath $script:LogFile -Append -Encoding utf8
    }
}

# ============================================================================
# ONEDRIVE UTILITIES
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

function Remove-LockedModuleFolder {
    <#
    .SYNOPSIS
        Force-removes a module folder that OneDrive has locked or converted to cloud placeholders.
        Escalation: normal Remove-Item -> strip cloud attributes + rd /s /q -> MoveFileEx reboot deletion.
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
            # Directories, reparse points, or already-deleted files -- ignore
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

# ============================================================================
# MAIN EXECUTION
# ============================================================================

$ErrorActionPreference = 'Stop'

# Initialize log file
if (-not (Test-Path $LogPath)) {
    New-Item -Path $LogPath -ItemType Directory -Force | Out-Null
}
$timestamp = Get-Date -Format 'yyyy-MM-dd_HHmmss'
$script:LogFile = Join-Path $LogPath "migration_$timestamp.log"

Write-Log "======================================================"
Write-Log "OneDrive Module Migration started"
Write-Log "======================================================"
Write-Log "PowerShell Version: $($PSVersionTable.PSVersion)"
Write-Log "Running as: $([Security.Principal.WindowsIdentity]::GetCurrent().Name)"
Write-Log "Log file: $script:LogFile"

# Detect OneDrive
$currentUserModulePath = Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'PowerShell\Modules'

if (-not (Test-OneDrivePath $currentUserModulePath)) {
    Write-Log "CurrentUser module path is not in OneDrive — no migration needed"
    Write-Log "  Path: $currentUserModulePath"
    exit 0
}

if (-not (Test-Path $currentUserModulePath)) {
    Write-Log "CurrentUser module path does not exist: $currentUserModulePath"
    exit 0
}

$allUsersPath = Join-Path $env:ProgramFiles 'PowerShell\Modules'
Write-Log "OneDrive detected on CurrentUser module path"
Write-Log "  Source: $currentUserModulePath"
Write-Log "  Destination: $allUsersPath"

# --- Phase 1: Copy modules to AllUsers scope ---
Write-Log "Starting module copy..."

$moduleFolders = Get-ChildItem -Path $currentUserModulePath -Directory -ErrorAction SilentlyContinue
if (-not $moduleFolders) {
    Write-Log "No modules found in CurrentUser path"
    exit 0
}

$copiedCount = 0
$copyFailures = @()

foreach ($moduleFolder in $moduleFolders) {
    $versionFolders = Get-ChildItem -Path $moduleFolder.FullName -Directory -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -match '^\d+(\.\d+){1,3}$' }

    # Some modules don't use version subfolders -- check for a manifest directly
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
                $copiedCount++
                Write-Log "Copied: $($moduleFolder.Name)" -Level SUCCESS
            }
            catch {
                Write-Log "Failed to copy $($moduleFolder.Name): $_" -Level ERROR
                $copyFailures += "$($moduleFolder.Name): $($_.Exception.Message)"
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
                $copiedCount++
                Write-Log "Copied: $($moduleFolder.Name) v$($versionFolder.Name)" -Level SUCCESS
            }
            catch {
                Write-Log "Failed to copy $($moduleFolder.Name) v$($versionFolder.Name): $_" -Level ERROR
                $copyFailures += "$($moduleFolder.Name) v$($versionFolder.Name): $($_.Exception.Message)"
            }
        }
    }
}

Write-Log "Copy phase complete. Copied: $copiedCount, Failed: $($copyFailures.Count)"

# --- Phase 2: Clean up OneDrive copies ---
if ($SkipCleanup) {
    Write-Log "Cleanup skipped (-SkipCleanup). OneDrive copies remain in place."
}
elseif ($copiedCount -eq 0 -and $copyFailures.Count -eq 0) {
    Write-Log "All modules already exist in AllUsers — checking for leftover OneDrive copies"

    # Still run cleanup even if nothing was copied (previous partial run may have left copies)
    $odModuleFolders = Get-ChildItem -Path $currentUserModulePath -Directory -ErrorAction SilentlyContinue
    $hasVersionFolders = $false
    foreach ($mf in $odModuleFolders) {
        $vf = Get-ChildItem -Path $mf.FullName -Directory -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -match '^\d+(\.\d+){1,3}$' }
        if ($vf) { $hasVersionFolders = $true; break }
    }

    if (-not $hasVersionFolders) {
        Write-Log "No version folders found in OneDrive path — nothing to clean up"
    }
}

if (-not $SkipCleanup) {
    $odModuleFolders = Get-ChildItem -Path $currentUserModulePath -Directory -ErrorAction SilentlyContinue

    if ($odModuleFolders.Count -gt 0) {
        Write-Log "Cleaning up OneDrive copies from: $currentUserModulePath"

        $removedCount = 0
        $removeFailures = @()
        $rebootScheduled = 0

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
                        $removedCount++
                        Write-Log "Removed: $moduleName v$versionName" -Level SUCCESS
                    }
                    catch {
                        Write-Log "OneDrive lock on $moduleName — attempting force-removal" -Level WARN
                        if (Remove-LockedModuleFolder -FolderPath $versionFolder.FullName) {
                            $removedCount++
                            Write-Log "Force-removed: $moduleName v$versionName" -Level SUCCESS
                        }
                        else {
                            Write-Log "Failed to remove OneDrive copy $moduleName v${versionName}: $($_.Exception.Message)" -Level WARN
                            $removeFailures += "$moduleName v$versionName"
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

        Write-Log "Cleanup complete. Removed: $removedCount, Failed: $($removeFailures.Count)"
    }
}

# --- Summary ---
Write-Log "======================================================"
Write-Log "OneDrive Module Migration completed" -Level SUCCESS
Write-Log "======================================================"
Write-Log "  Modules copied to AllUsers: $copiedCount"
if (-not $SkipCleanup) {
    Write-Log "  OneDrive copies removed: $removedCount"
    if ($removeFailures.Count -gt 0) {
        Write-Log "  Failed to remove: $($removeFailures.Count) (may be scheduled for reboot deletion)" -Level WARN
    }
}
if ($copyFailures.Count -gt 0) {
    Write-Log "  Failed to copy: $($copyFailures.Count)" -Level ERROR
    foreach ($f in $copyFailures) {
        Write-Log "    $f" -Level ERROR
    }
}
