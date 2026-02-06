# PSModuleMaintenance

Automated PowerShell module maintenance for Windows. Updates all PSResourceGet-managed modules and prunes old versions on a weekly schedule with comprehensive logging.

## Features

- 🔄 **Automatic Updates** — Updates all installed PowerShell modules via PSResourceGet
- 🧹 **Version Pruning** — Removes old module versions, keeping only the latest
- 📋 **Comprehensive Logging** — Structured logs with transcripts and JSON summaries
- ⚙️ **Configurable Exclusions** — Skip specific modules via config file
- ⏰ **Scheduled Execution** — Runs weekly via Windows Task Scheduler
- 🔔 **Toast Notifications** — Optional Windows toast notifications after each run
- 🔒 **CurrentUser Scope** — Runs as your account for OneDrive-synced module paths

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
  "TrustPSGallery": true
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
```

## Configuration

| Setting | Type | Default | Description |
|---------|------|---------|-------------|
| `ExcludedModules` | string[] | `[]` | Module names to skip during updates and pruning |
| `LogRetentionDays` | int | `180` | Days to keep log files before auto-cleanup |
| `TrustPSGallery` | bool | `true` | Trust PSGallery during updates (avoids prompts) |
| `NotificationMode` | string | `"Always"` | Toast notifications: `"Always"`, `"OnFailure"`, or `"Never"` |

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

## How It Works

1. **Load Configuration** — Reads `config.json` for exclusions and settings
2. **Initialize Logging** — Creates timestamped log files and starts transcript
3. **Clean Old Logs** — Removes logs older than retention period
4. **Update Modules** — Bulk checks PSGallery for available updates, then only updates modules that have newer versions
5. **Prune Versions** — Groups modules by name, keeps newest, removes the rest (detects OneDrive and warns about potential popups)
6. **Save Summary** — Writes JSON summary for monitoring/alerting integration
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

### Permission errors

Ensure the script runs as your user account (not SYSTEM) since modules are installed in CurrentUser scope.

### OneDrive "large number of files deleted" warning

OneDrive may warn about mass deletions when pruning old module versions. This is expected behavior. To prevent this warning, exclude your PowerShell modules folder from OneDrive sync:

1. Open OneDrive Settings → Account → Choose folders
2. Deselect `Documents\PowerShell\Modules` (PowerShell 7+) or `Documents\WindowsPowerShell\Modules` (Windows PowerShell)

**Note:** On managed devices with Known Folder Move enabled, you may not be able to deselect subfolders of Documents. This is an organizational policy limitation — contact your IT admin or accept the OneDrive warnings when they appear.

### "Access denied" errors during pruning

OneDrive can lock files during sync, causing "Access denied" errors when removing old module versions. The script detects OneDrive-synced folders and warns you upfront. When this happens, OneDrive may show a confirmation popup — click "Delete" to allow the removal. Solutions:

1. **Confirm in OneDrive popup** — When prompted, click "Delete" to allow the removal
2. **Pause OneDrive sync** before running maintenance (right-click OneDrive tray icon → Pause syncing)
3. **Exclude the Modules folder** from OneDrive sync (see above) — this is the permanent fix

## Contributing

Issues and PRs welcome! Please include log output when reporting bugs.

## License

MIT License - See [LICENSE](LICENSE) for details.

## Author

**Haakon Wibe**  
- Blog: [alttabtowork.com](https://alttabtowork.com)  
- Twitter: [@HaakonWibe](https://twitter.com/HaakonWibe)