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
# Full maintenance (update + prune)
.\Invoke-PSModuleMaintenance.ps1

# Update only
.\Invoke-PSModuleMaintenance.ps1 -UpdateOnly

# Prune only
.\Invoke-PSModuleMaintenance.ps1 -PruneOnly

# Dry run
.\Invoke-PSModuleMaintenance.ps1 -WhatIf

# One-time OneDrive migration (requires Administrator)
.\Invoke-OneDriveMigration.ps1

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
4. OneDrive Utilities - `Test-OneDrivePath` and `Remove-LockedModuleFolder` functions (used for scope detection and locked-file fallback during updates/pruning)
5. Module Operations - version helpers (`ConvertTo-NormalizedVersion`, `Get-ModuleVersionKey`, `ConvertFrom-PinnedVersionString`, `Get-PinnedVersion`, `Test-IsPinnedVersion`) plus `Invoke-ModuleUpdate`, `Set-PinnedModuleVersions`, `Update-AllModules`, and `Remove-OldModuleVersions`
6. Main Execution - Orchestrates the workflow with try/finally for cleanup and incremental summary saves

**Invoke-OneDriveMigration.ps1** - Standalone one-time script that migrates modules from OneDrive-synced CurrentUser path to AllUsers scope, then cleans up OneDrive copies. Self-contained with its own copies of `Test-OneDrivePath`, `Remove-LockedModuleFolder`, and `Write-Log`. Requires Administrator.

**Install-ModuleMaintenance.ps1** - Creates Windows Scheduled Task that runs the maintenance script weekly as the current user with elevated privileges.

**config.json** - Runtime configuration validated against `config.schema.json`:
- `ExcludedModules`: Array of module names to skip
- `PinnedModules`: Object mapping module name to an exact version to hold (default `{}`)
- `LogRetentionDays`: How long to keep logs (default 180)
- `TrustPSGallery`: Trust repository during updates (default true)
- `NotificationMode`: Toast notifications — `Always` (default), `OnFailure`, or `Never`
- `ModuleUpdateTimeoutSeconds`: Per-module update timeout in seconds (default 600)

## Key Implementation Details

- **Version pinning**: `PinnedModules` holds a module at one version. `Set-PinnedModuleVersions` runs first in the update phase and installs the pinned version via `Install-PSResource` when it's missing (a bare version string is a *required* version in PSResourceGet, not a minimum — see `about_PSResourceGet`, "NuGet version ranges"). Pinned modules stay in the `Find-PSResource` bulk query (same single call either way) so the log can report `<module> is pinned to <pin> — holding back <gallery>` and record it in `Summary.PinsHoldingBack`; they are filtered out of the update loop immediately after. In `Remove-OldModuleVersions` a pinned module keeps the pin instead of the newest version, so newer versions get pruned. **Safety invariant**: if the pinned version isn't installed, that module is pruned not at all — otherwise a bad pin would delete every version. Versions are compared via `Get-ModuleVersionKey` (4-part normalization + prerelease label), so `2.19`, `2.19.0`, and `2.19.0.0` are the same pin. Exclusion beats pinning: a module in both lists is dropped from `PinnedModules` at config load with a warning
- **Config warnings before logging exists**: `Import-MaintenanceConfig` runs before `Initialize-Logging`, so pin validation problems are queued in `$script:ConfigWarnings` and written to the log as WARN once logging is up
- **Per-module timeout**: Each `Update-PSResource` call runs in an isolated PowerShell runspace via `Invoke-ModuleUpdate`. If a module exceeds `ModuleUpdateTimeoutSeconds` (default 600s), the runspace is stopped and the script moves to the next module — prevents one slow/hung module (e.g. Microsoft.Graph) from consuming all scheduled task time
- **Incremental summary saves**: `Save-Summary` is called after each phase (migration, updates, pruning), overwriting the same file. If the process is killed mid-run, the last completed phase's results are on disk
- **Progress suppression**: `$ProgressPreference = 'SilentlyContinue'` is set at script start — progress bars add significant overhead in non-interactive/scheduled task mode
- **Update optimization**: Bulk queries PSGallery via `Find-PSResource` to check for available updates, then only calls `Update-PSResource` for modules that actually need updating (avoids 150+ individual network calls)
- **PSResourceGet scope pitfalls**: `Get-PSResource` without `-Scope` defaults to CurrentUser only — after OneDrive migration, modules live in AllUsers so `-Scope AllUsers` is required. `Uninstall-PSResource` without `-Scope` can target ANY scope, potentially removing AllUsers copies instead of OneDrive copies. `InstalledLocation` on PSResource objects can return the modules root directory (e.g. `C:\Program Files\PowerShell\Modules`) instead of the version folder — the pruning fallback extracts the correct path from the PSResourceGet error message or constructs it from the module name + version
- **OneDrive handling**: When Documents is redirected to OneDrive via Known Folder Move, the main script detects this by checking `[Environment]::GetFolderPath('MyDocuments')` against OneDrive environment variables, then automatically targets AllUsers scope for updates and pruning. If modules are found in the OneDrive CurrentUser path, the prune phase logs a warning directing the user to run `Invoke-OneDriveMigration.ps1` — it does not delete them. The one-time migration itself is handled by `Invoke-OneDriveMigration.ps1` (separate script). `Remove-LockedModuleFolder` uses four-stage escalation for stubborn OneDrive files:
    1. Normal `Remove-Item -Recurse -Force`
    2. Strip OneDrive cloud-file attributes (`attrib -P -U -O`) + `cmd.exe rd /s /q` (handles reparse points/cloud placeholders)
    3. File-by-file deletion of whatever is deletable
    4. `kernel32.dll MoveFileEx` with `MOVEFILE_DELAY_UNTIL_REBOOT` — schedules remaining files for kernel-level deletion on next reboot
  - OneDrive cloud placeholders (reparse points) are the main blocker — files appear locally but are cloud-only stubs that standard file APIs reject with "Access denied"
  - Built-in modules (e.g., PackageManagement) that ship under `$PSHOME` are automatically skipped during pruning — PSResourceGet cannot uninstall these
  - When OneDrive is NOT detected, all behavior is identical to pre-migration versions
- Modules are grouped by name; only the latest version is kept during pruning (or the pinned version, for pinned modules)
- **`GroupInfo.Count` pitfall**: `$grouped = Group-Object ... | Where-Object {...}` returns a bare `GroupInfo` when exactly one group survives, and `$grouped.Count` then resolves to *that group's* item count instead of the number of groups. `Remove-OldModuleVersions` wraps the pipeline in `@()` for this reason — without it the log reported "Found 2 modules with multiple versions" for a single 2-version module
- Logs go to `$env:ProgramData\PSModuleMaintenance\Logs` with three file types: structured log, full transcript, and JSON summary
- **CMTrace-safe log wording**: CMTrace colors un-typed log lines red/yellow by scanning the text for `error`/`fail`/`warn` substrings, ignoring the `[INFO]` tag. Keep INFO/SUCCESS summary lines neutral so they don't appear as false errors — use `Unsuccessful:` not `Failed:`, and `No issues.` not `No errors.`. Only genuine `-Level WARN`/`-Level ERROR` lines should contain those words
- The scheduled task uses `Interactive` logon type with "run with highest privileges" (runs hidden, requires user to be logged in but `StartWhenAvailable` catches up if missed) with a 4-hour execution time limit
- All three scripts use `[CmdletBinding(SupportsShouldProcess)]` for `-WhatIf` support. Because `-WhatIf` sets `$WhatIfPreference` for the whole script scope, it cascades into ShouldProcess-aware cmdlets including the bookkeeping file writes. `Write-Log`'s `Out-File`/`Add-Content`, `Save-Summary`'s `Set-Content`, and the log-directory `New-Item` in both `Initialize-Logging` and `Invoke-OneDriveMigration.ps1` are called with `-WhatIf:$false` so logging and the JSON summary still run (and don't spam `What if: Performing the operation "Output to File"...`) during a dry run — only the real actions (copy/remove/update/prune) are gated by `-WhatIf`. The log directory is easy to miss because `%ProgramData%\PSModuleMaintenance\Logs` already exists in normal use; the gap only shows up on a dry run against a *fresh* `-LogPath`, where the missing directory silently suppresses both the log and the summary. `Start-Transcript` is deliberately left gated — the transcript is skipped during dry runs
