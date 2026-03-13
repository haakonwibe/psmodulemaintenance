# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

PSModuleMaintenance is a Windows-based automation tool that keeps PowerShell modules up to date. It uses Microsoft.PowerShell.PSResourceGet to update modules and prune old versions on a weekly schedule.

## Requirements

- Windows 10/11 or Windows Server 2019+
- PowerShell 7.0+
- Microsoft.PowerShell.PSResourceGet module

## Running the Scripts

```powershell
# Full maintenance (migrate + update + prune)
.\Invoke-PSModuleMaintenance.ps1

# Update only
.\Invoke-PSModuleMaintenance.ps1 -UpdateOnly

# Prune only
.\Invoke-PSModuleMaintenance.ps1 -PruneOnly

# Migrate modules out of OneDrive only
.\Invoke-PSModuleMaintenance.ps1 -MigrateOnly

# Dry run
.\Invoke-PSModuleMaintenance.ps1 -WhatIf

# Install scheduled task (requires Administrator)
.\Install-ModuleMaintenance.ps1

# Uninstall scheduled task
.\Install-ModuleMaintenance.ps1 -Uninstall
```

## Architecture

**Invoke-PSModuleMaintenance.ps1** - Main maintenance script with six sections:
1. Configuration - Loads `config.json`, merges with defaults
2. Logging - Initializes log files, transcript, and summary tracking
3. Toast Notifications - `Send-ToastNotification` function (shells out to PS 5.1 for WinRT support)
4. OneDrive Migration - `Test-OneDrivePath`, `Move-ModulesToAllUsers`, `Remove-LockedModuleFolder` functions
5. Module Operations - `Update-AllModules` and `Remove-OldModuleVersions` functions
6. Main Execution - Orchestrates the workflow with try/finally for cleanup

**Install-ModuleMaintenance.ps1** - Creates Windows Scheduled Task that runs the maintenance script weekly as the current user with elevated privileges.

**config.json** - Runtime configuration validated against `config.schema.json`:
- `ExcludedModules`: Array of module names to skip
- `LogRetentionDays`: How long to keep logs (default 180)
- `TrustPSGallery`: Trust repository during updates (default true)
- `NotificationMode`: Toast notifications — `Always` (default), `OnFailure`, or `Never`
- `MigrateFromOneDrive`: Migrate modules from OneDrive to AllUsers scope (default true)

## Key Implementation Details

- **Update optimization**: Bulk queries PSGallery via `Find-PSResource` to check for available updates, then only calls `Update-PSResource` for modules that actually need updating (avoids 150+ individual network calls)
- **OneDrive migration**: When Documents is redirected to OneDrive via Known Folder Move, the script detects this by checking `[Environment]::GetFolderPath('MyDocuments')` against OneDrive environment variables, then:
  - Copies modules to `$env:ProgramFiles\PowerShell\Modules` (AllUsers scope)
  - Updates with `-Scope AllUsers` so new versions install outside OneDrive
  - Prunes all OneDrive copies in a dedicated cleanup pass
  - Uses force-removal (`Remove-LockedModuleFolder`) with GC collection, attribute stripping, and two-pass deletion for OneDrive-locked files
  - Retry loop handles meta-packages (e.g., Az) that hit multiple locked sub-module folders
  - When OneDrive is NOT detected, all behavior is identical to pre-migration versions
- Modules are grouped by name; only the latest version is kept during pruning
- Logs go to `$env:ProgramData\PSModuleMaintenance\Logs` with three file types: structured log, full transcript, and JSON summary
- The scheduled task uses `Interactive` logon type with "run with highest privileges" (runs hidden, requires user to be logged in but `StartWhenAvailable` catches up if missed)
- Both scripts use `[CmdletBinding(SupportsShouldProcess)]` for `-WhatIf` support
