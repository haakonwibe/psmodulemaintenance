# PSModuleMaintenance

Automated PowerShell module maintenance for Windows. Updates all PSResourceGet-managed modules and prunes old versions on a weekly schedule with comprehensive logging.

## Features

- 🔄 **Automatic Updates** — Updates all installed PowerShell modules via PSResourceGet
- 🧹 **Version Pruning** — Removes old module versions, keeping only the latest
- ☁️ **OneDrive Migration** — Standalone script to migrate modules out of OneDrive-synced folders to AllUsers scope
- 📋 **Comprehensive Logging** — Structured logs with transcripts and JSON summaries
- ⚙️ **Configurable Exclusions** — Skip specific modules via config file
- ⏰ **Scheduled Execution** — Runs weekly via Windows Task Scheduler
- 🔔 **Toast Notifications** — Optional Windows toast notifications after each run
- 🛡️ **Per-Module Timeout** — Each module update runs in an isolated runspace with a configurable timeout, preventing one slow module from blocking the entire run

## Requirements

- Windows 10/11 or Windows Server 2019+
- PowerShell 7.0 or later
- [Microsoft.PowerShell.PSResourceGet](https://www.powershellgallery.com/packages/Microsoft.PowerShell.PSResourceGet) module

## Quick Start

### 1. Clone or Download

```powershell
git clone https://github.com/haakonwibe/PSModuleMaintenance.git
cd PSModuleMaintenance
```

### 2. Configure (Optional)

Edit `config.json` to exclude specific modules:

```json
{
  "ExcludedModules": [
    "Az.Accounts",
    "SomeModuleIPinToSpecificVersion"
  ],
  "LogRetentionDays": 180,
  "TrustPSGallery": true,
  "NotificationMode": "Always",
  "ModuleUpdateTimeoutSeconds": 600
}
```

### 3. OneDrive Check

If your device uses **OneDrive Known Folder Move** (common on enterprise/Intune-managed devices), your PowerShell modules are synced to OneDrive, which causes file locks and sync conflicts. Run the migration script first to move them out:

```powershell
# Check if you're affected (dry run)
.\Invoke-OneDriveMigration.ps1 -WhatIf

# If it finds modules, run the migration (requires Administrator)
.\Invoke-OneDriveMigration.ps1
```

If OneDrive is not detected, the script exits immediately. See [OneDrive Migration](#onedrive-migration) for details.

### 4. Install Scheduled Task

Run as Administrator:

```powershell
.\Install-ModuleMaintenance.ps1
```

This creates a weekly task running Sundays at 3:00 AM.

#### Custom Schedule

```powershell
.\Install-ModuleMaintenance.ps1 -DayOfWeek Saturday -Time "04:30"
```

## Manual Usage

Run the maintenance script directly:

```powershell
# Full maintenance (update + prune)
.\Invoke-PSModuleMaintenance.ps1

# Update only
.\Invoke-PSModuleMaintenance.ps1 -UpdateOnly

# Prune only
.\Invoke-PSModuleMaintenance.ps1 -PruneOnly

# Dry run - see what would happen
.\Invoke-PSModuleMaintenance.ps1 -WhatIf

# Verbose output
.\Invoke-PSModuleMaintenance.ps1 -Verbose

# Migrate modules out of OneDrive (one-time, see OneDrive Migration section below)
.\Invoke-OneDriveMigration.ps1
```

## Configuration

| Setting | Type | Default | Description |
|---------|------|---------|-------------|
| `ExcludedModules` | string[] | `[]` | Module names to skip during updates and pruning |
| `LogRetentionDays` | int | `180` | Days to keep log files before auto-cleanup |
| `TrustPSGallery` | bool | `true` | Trust PSGallery during updates (avoids prompts) |
| `NotificationMode` | string | `"Always"` | Toast notifications: `"Always"`, `"OnFailure"`, or `"Never"` |
| `ModuleUpdateTimeoutSeconds` | int | `600` | Max seconds per module update before timing out and moving to the next |

## Notifications

PSModuleMaintenance can show a Windows toast notification after each run. Set `NotificationMode` in `config.json`:

- **`"Always"`** — Notification after every run with a summary of updates and pruning (default)
- **`"OnFailure"`** — Notification only when modules fail to update or versions fail to prune
- **`"Never"`** — No notifications

The toast uses the built-in Windows "Security and Maintenance" notification channel — no additional setup required.

## Logs

Logs are written to `%ProgramData%\PSModuleMaintenance\Logs\`:

```
C:\ProgramData\PSModuleMaintenance\Logs\
├── maintenance_2024-01-15_030000.log    # Structured log
├── transcript_2024-01-15_030000.log     # Full verbose transcript
└── summary_2024-01-15_030512.json       # Machine-readable summary
```

### Summary JSON Structure

```json
{
  "StartTime": "2024-01-15T03:00:00.0000000+01:00",
  "EndTime": "2024-01-15T03:05:12.0000000+01:00",
  "ModulesChecked": 79,
  "ModulesUpdated": 12,
  "ModulesFailed": [],
  "VersionsPruned": 45,
  "PrunesFailed": [],
  "ExcludedModules": ["Az.Accounts"]
}
```

## Uninstall

Remove the scheduled task:

```powershell
.\Install-ModuleMaintenance.ps1 -Uninstall
```

Optionally remove logs:

```powershell
Remove-Item "$env:ProgramData\PSModuleMaintenance" -Recurse -Force
```

## OneDrive Migration

### The Problem

PowerShell 7 installs CurrentUser-scope modules to `$HOME\Documents\PowerShell\Modules`. On enterprise devices with **Known Folder Move** enabled, the Documents folder is redirected to OneDrive. This causes OneDrive to sync module files, leading to:

- **File locks** during sync that block `Update-PSResource` and `Uninstall-PSResource` with "Cannot remove package path" and "Access denied" errors
- **Cloud placeholders** (reparse points) — OneDrive replaces local files with cloud-only stubs that standard file APIs cannot delete
- **Deletion confirmation popups** when pruning old module versions
- **Inability to exclude the folder** from sync on managed devices (organizational policy)

**How do I know if this affects me?** Run `.\Invoke-OneDriveMigration.ps1 -WhatIf` — if it says "CurrentUser module path is not in OneDrive", you're not affected and can ignore this section entirely. If it lists modules to copy, you're affected. This is common on enterprise devices managed by Intune/SCCM with Known Folder Move policies.

### The Four Horsemen

Solving this required fighting four systems at once, each with undocumented edge cases that only revealed themselves when the previous layer was fixed:

1. **PSResourceGet's scope model** — `Get-PSResource` without `-Scope` defaults to CurrentUser only (finds nothing after migration). `Uninstall-PSResource` without `-Scope` targets *any* scope (deletes the wrong copies). `InstalledLocation` returns inconsistent paths. Each API call needed different scope handling.

2. **OneDrive Known Folder Move** — Silently redirects a system path that PowerShell depends on. No API to detect it directly — you have to infer it by comparing `[Environment]::GetFolderPath('MyDocuments')` against OneDrive environment variables.

3. **OneDrive cloud placeholders** — Files that *look* normal to `Get-ChildItem` but are actually NTFS reparse points with no local data. They return "Access denied" on delete, but `handle.exe` shows no locks and ACLs show FullControl. The fix: strip cloud attributes with `attrib -P -U -O`, then use `cmd.exe rd /s /q` which handles reparse points where PowerShell's `Remove-Item` cannot.

4. **The cascading reveal** — Each fix exposed the next bug. Fix the migration → scope mismatch deletes AllUsers modules. Fix the scope → `Get-PSResource` returns 1 module instead of 160. Fix that → `InstalledLocation` points to wrong path. Fix *that* → "Access denied" on cloud placeholders. No single system was "wrong" — the bugs only existed at the intersections.

### The Solution

`Invoke-OneDriveMigration.ps1` is a standalone one-time migration script that:

1. **Detects** if your CurrentUser module path is inside a OneDrive-synced folder
2. **Copies** all modules to AllUsers scope (`$env:ProgramFiles\PowerShell\Modules`) — safely, without deleting the originals
3. **Cleans up** the old OneDrive copies with four-stage force-removal for cloud placeholders and locked files (see [Troubleshooting](#onedrive-file-lock--cloud-placeholder-errors))

After migration, the weekly maintenance script (`Invoke-PSModuleMaintenance.ps1`) automatically detects OneDrive on the module path and targets AllUsers scope for all future updates and pruning — no configuration needed.

The migration is **idempotent and gradual** — modules that already exist at the destination are skipped, and OneDrive copies that can't be removed are either force-deleted or scheduled for reboot deletion.

If you skip migration, the weekly maintenance script will still work (it detects OneDrive and targets AllUsers scope automatically), but any modules left in the OneDrive path will trigger a warning in the logs: `Found N module(s) in OneDrive path — run Invoke-OneDriveMigration.ps1 to migrate them`.

### Usage

```powershell
# Dry-run first to see what would happen
.\Invoke-OneDriveMigration.ps1 -WhatIf

# Run the migration (requires Administrator)
.\Invoke-OneDriveMigration.ps1

# Copy modules to AllUsers but keep OneDrive copies in place
.\Invoke-OneDriveMigration.ps1 -SkipCleanup
```

## How It Works

1. **Load Configuration** — Reads `config.json` for exclusions and settings
2. **Initialize Logging** — Creates timestamped log files and starts transcript
3. **Clean Old Logs** — Removes logs older than retention period
4. **Update Modules** — Bulk checks PSGallery for available updates, then updates each module in an isolated runspace with a per-module timeout (targets AllUsers scope when OneDrive is detected)
5. **Prune Versions** — Groups modules by name, keeps newest, removes the rest (skips built-in modules like PackageManagement). When OneDrive is detected and modules are found in the CurrentUser path, logs a warning to run `Invoke-OneDriveMigration.ps1`
6. **Save Summary** — Writes JSON summary after each phase (incremental saves protect against process termination)
7. **Toast Notification** — Shows a Windows toast with the run summary (if enabled via `NotificationMode`)

## Troubleshooting

### Task doesn't run

Check Task Scheduler history. Common issues:
- Script path changed after installation
- User password changed (re-run installer)
- Network not available at scheduled time

### Modules fail to update

Check the log files for specific errors. Common causes:
- Module removed from PSGallery
- Dependency conflicts
- Network/proxy issues

### Module update timed out

Large meta-modules like `Microsoft.Graph` (40+ sub-modules) can exceed the default 10-minute timeout. Increase it in `config.json`:

```json
{
  "ModuleUpdateTimeoutSeconds": 1200
}
```

The log will show which module timed out and how far through the update list it got (e.g. "Updating module 1/40: Microsoft.Graph").

### Permission errors

The maintenance script requires Administrator privileges when modules are in AllUsers scope (after OneDrive migration). The scheduled task is configured to run with highest privileges. For manual runs, use an elevated PowerShell prompt. `Invoke-OneDriveMigration.ps1` requires Administrator (enforced via `#Requires -RunAsAdministrator`).

### OneDrive file lock / cloud placeholder errors

OneDrive can block file deletion in two ways: **sync locks** during active syncing, and **cloud placeholders** (reparse points) where OneDrive replaces local files with cloud-only stubs. Both cause "Access denied" errors. The script handles this with a four-stage escalation:

1. **Normal deletion** — `Remove-Item -Recurse -Force`
2. **Cloud attribute strip + `rd /s /q`** — Strips OneDrive cloud-file attributes (`attrib -P -U -O`) to convert placeholders back to normal files, then uses `cmd.exe rd /s /q` which handles reparse points differently than PowerShell's `Remove-Item`
3. **File-by-file deletion** — Deletes individual files, skipping those still locked
4. **Reboot-scheduled deletion** — Uses `kernel32.dll MoveFileEx` with `MOVEFILE_DELAY_UNTIL_REBOOT` to schedule remaining locked files for deletion by the Windows kernel on next reboot (before any user-mode process starts)

Most OneDrive cleanup completes at stage 2. Stage 4 is the nuclear option for truly stubborn files.

### OneDrive "large number of files deleted" warning

OneDrive may warn about mass deletions when cleaning up migrated module copies. This is expected — the modules have already been copied to AllUsers scope. Click "Delete" to allow OneDrive to sync the removal.

## Contributing

Issues and PRs welcome! Please include log output when reporting bugs.

## License

MIT License - See [LICENSE](LICENSE) for details.

## Author

**Haakon Wibe**  
- Blog: [alttabtowork.com](https://alttabtowork.com)  
- Twitter: [@HaakonWibe](https://twitter.com/HaakonWibe)