# PSModuleMaintenance

Automated PowerShell module maintenance for Windows. Updates all PSResourceGet-managed modules and prunes old versions on a weekly schedule with comprehensive logging.

## Features

- 🔄 **Automatic Updates** — Updates all installed PowerShell modules via PSResourceGet
- 🧹 **Version Pruning** — Removes old module versions, keeping only the latest
- ☁️ **OneDrive Migration** — Automatically migrates modules out of OneDrive-synced folders to AllUsers scope
- 📋 **Comprehensive Logging** — Structured logs with transcripts and JSON summaries
- ⚙️ **Configurable Exclusions** — Skip specific modules via config file
- ⏰ **Scheduled Execution** — Runs weekly via Windows Task Scheduler
- 🔔 **Toast Notifications** — Optional Windows toast notifications after each run

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
  "MigrateFromOneDrive": false
}
```

### 3. Install Scheduled Task

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
# Full maintenance (migrate + update + prune)
.\Invoke-PSModuleMaintenance.ps1

# Update only
.\Invoke-PSModuleMaintenance.ps1 -UpdateOnly

# Prune only
.\Invoke-PSModuleMaintenance.ps1 -PruneOnly

# Migrate modules out of OneDrive only (no updates or pruning)
.\Invoke-PSModuleMaintenance.ps1 -MigrateOnly

# Dry run - see what would happen
.\Invoke-PSModuleMaintenance.ps1 -WhatIf

# Verbose output
.\Invoke-PSModuleMaintenance.ps1 -Verbose
```

## Configuration

| Setting | Type | Default | Description |
|---------|------|---------|-------------|
| `ExcludedModules` | string[] | `[]` | Module names to skip during updates and pruning |
| `LogRetentionDays` | int | `180` | Days to keep log files before auto-cleanup |
| `TrustPSGallery` | bool | `true` | Trust PSGallery during updates (avoids prompts) |
| `NotificationMode` | string | `"Always"` | Toast notifications: `"Always"`, `"OnFailure"`, or `"Never"` |
| `MigrateFromOneDrive` | bool | `false` | Migrate modules from OneDrive to AllUsers scope (see below) |

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
  "ModulesMigrated": 0,
  "MigrationFailed": [],
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

### The Four Horsemen

Solving this required fighting four systems at once, each with undocumented edge cases that only revealed themselves when the previous layer was fixed:

1. **PSResourceGet's scope model** — `Get-PSResource` without `-Scope` defaults to CurrentUser only (finds nothing after migration). `Uninstall-PSResource` without `-Scope` targets *any* scope (deletes the wrong copies). `InstalledLocation` returns inconsistent paths. Each API call needed different scope handling.

2. **OneDrive Known Folder Move** — Silently redirects a system path that PowerShell depends on. No API to detect it directly — you have to infer it by comparing `[Environment]::GetFolderPath('MyDocuments')` against OneDrive environment variables.

3. **OneDrive cloud placeholders** — Files that *look* normal to `Get-ChildItem` but are actually NTFS reparse points with no local data. They return "Access denied" on delete, but `handle.exe` shows no locks and ACLs show FullControl. The fix: strip cloud attributes with `attrib -P -U -O`, then use `cmd.exe rd /s /q` which handles reparse points where PowerShell's `Remove-Item` cannot.

4. **The cascading reveal** — Each fix exposed the next bug. Fix the migration → scope mismatch deletes AllUsers modules. Fix the scope → `Get-PSResource` returns 1 module instead of 160. Fix that → `InstalledLocation` points to wrong path. Fix *that* → "Access denied" on cloud placeholders. No single system was "wrong" — the bugs only existed at the intersections.

### The Solution

When `MigrateFromOneDrive` is enabled, the script automatically:

1. **Detects** if your CurrentUser module path is inside a OneDrive-synced folder
2. **Copies** all modules to AllUsers scope (`$env:ProgramFiles\PowerShell\Modules`) — safely, without deleting the originals
3. **Updates** modules with `-Scope AllUsers` so new versions install outside OneDrive
4. **Cleans up** the old OneDrive copies during the prune pass, with four-stage force-removal for cloud placeholders and locked files (see [Troubleshooting](#onedrive-file-lock--cloud-placeholder-errors))

The migration is **idempotent and gradual** — modules that already exist at the destination are skipped, and OneDrive copies that can't be removed are either force-deleted or scheduled for reboot deletion.

### Usage

```powershell
# Dry-run first to see what would happen
.\Invoke-PSModuleMaintenance.ps1 -MigrateFromOneDrive -WhatIf

# One-time migration + full maintenance (no config change needed)
.\Invoke-PSModuleMaintenance.ps1 -MigrateFromOneDrive

# Migrate only, skip updates and pruning
.\Invoke-PSModuleMaintenance.ps1 -MigrateOnly -MigrateFromOneDrive
```

To enable migration permanently, set `"MigrateFromOneDrive": true` in `config.json`. The `-MigrateFromOneDrive` switch overrides the config for a single run. When disabled (the default), or when the module path is not in OneDrive, the script behaves exactly as before.

## How It Works

1. **Load Configuration** — Reads `config.json` for exclusions and settings
2. **Initialize Logging** — Creates timestamped log files and starts transcript
3. **Clean Old Logs** — Removes logs older than retention period
4. **Migrate Modules** — If OneDrive is detected on the module path, copies modules to AllUsers scope
5. **Update Modules** — Bulk checks PSGallery for available updates, then only updates modules that have newer versions (targets AllUsers scope when OneDrive is detected)
6. **Prune Versions** — Groups modules by name, keeps newest, removes the rest (skips built-in modules like PackageManagement). When OneDrive is detected, also removes all migrated copies from the OneDrive path
7. **Save Summary** — Writes JSON summary for monitoring/alerting integration
8. **Toast Notification** — Shows a Windows toast with the run summary (if enabled via `NotificationMode`)

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

### Permission errors

The script requires Administrator privileges when using OneDrive migration (AllUsers scope). The scheduled task is configured to run with highest privileges. For manual runs, use an elevated PowerShell prompt.

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