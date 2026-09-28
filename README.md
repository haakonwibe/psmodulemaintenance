# PSModuleMaintenance

Automated PowerShell module maintenance for Windows. Updates all PSResourceGet-managed modules and prunes old versions on a weekly schedule with comprehensive logging.

## Features

- 🔄 **Automatic Updates** — Updates all installed PowerShell modules via PSResourceGet
- 🧹 **Version Pruning** — Removes old module versions, keeping only the latest
- ☁️ **OneDrive Migration** — Standalone script to migrate modules out of OneDrive-synced folders to AllUsers scope
- 📋 **Comprehensive Logging** — Structured logs with transcripts and JSON summaries
- ⚙️ **Configurable Exclusions** — Skip specific modules via config file
- 📌 **Version Pinning** — Hold specific modules at a chosen version instead of updating them
- 🪜 **Side-by-Side Version Lines** — Keep an older major version next to the newest one, each updated within its own line
- ⏰ **Scheduled Execution** — Runs weekly via Windows Task Scheduler
- 🔔 **Toast Notifications** — Optional Windows toast notifications after each run
- 🛡️ **Per-Module Timeout** — Each module update runs in an isolated runspace with a configurable timeout, preventing one slow module from blocking the entire run
- 🔁 **Retry on Network Faults** — A dropped connection or gateway hiccup is retried (up to three attempts) instead of costing the module a week
- 📡 **Healthchecks.io Monitoring** — Optional dead-man's switch that alerts when a scheduled run never happens, not just when one fails

## Requirements

- Windows 10/11 or Windows Server 2019+
- PowerShell 7.0 or later
- [Microsoft.PowerShell.PSResourceGet](https://www.powershellgallery.com/packages/Microsoft.PowerShell.PSResourceGet) module — included with PowerShell 7.4 and later; on 7.0–7.3 install it from the gallery

## Quick Start

### 1. Clone or Download

```powershell
git clone https://github.com/haakonwibe/PSModuleMaintenance.git
cd PSModuleMaintenance
```

### 2. Configure (Optional)

Without a `config.json` the script runs on its built-in defaults. To change a setting,
copy the template and edit the copy:

```powershell
Copy-Item .\config.example.json .\config.json
```

For example, to exclude or pin specific modules:

```json
{
  "ExcludedModules": [
    "SomeModuleIManageMyself"
  ],
  "PinnedModules": {
    "Az.Accounts": "2.19.0"
  },
  "KeepVersions": {
    "Pester": ["5"]
  },
  "LogRetentionDays": 180,
  "TrustPSGallery": true,
  "NotificationMode": "Always",
  "ModuleUpdateTimeoutSeconds": 600,
  "Healthchecks": {
    "Enabled": false,
    "SecretName": "PSModuleMaintenance-Healthchecks",
    "SecretVault": "SecretStore",
    "TimeoutSeconds": 10
  }
}
```

`config.json` is ignored by git, so your settings stay on your machine and a `git pull`
never touches them.

> **Upgrading from a version where `config.json` was part of the repository:** copy your
> `config.json` somewhere safe before you pull, and copy it back afterwards. Git removes the
> file from the folder when it stops being tracked.

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
| `PinnedModules` | object | `{}` | Module name → exact version to hold (see [Version Pinning](#version-pinning)) |
| `KeepVersions` | object | `{}` | Module name → older version lines to keep and update next to the newest (see [Keeping an Older Version Line](#keeping-an-older-version-line)) |
| `LogRetentionDays` | int | `180` | Days to keep log files before auto-cleanup |
| `TrustPSGallery` | bool | `true` | Trust PSGallery during updates (avoids prompts) |
| `NotificationMode` | string | `"Always"` | Toast notifications: `"Always"`, `"OnFailure"`, or `"Never"` |
| `ModuleUpdateTimeoutSeconds` | int | `600` | Max seconds per module update before timing out and moving to the next |
| `Healthchecks` | object | see below | Healthchecks.io monitoring (see [Monitoring](#monitoring)) |

### Healthchecks sub-settings

| Setting | Type | Default | Description |
|---------|------|---------|-------------|
| `Enabled` | bool | `false` | Send start/success/fail pings |
| `SecretName` | string | `"PSModuleMaintenance-Healthchecks"` | Vault secret holding the full ping URL |
| `SecretVault` | string | `"SecretStore"` | SecretManagement vault to read from |
| `TimeoutSeconds` | int | `10` | HTTP timeout per ping attempt (retried twice) |

**The ping URL is never stored in `config.json`** — it is a bearer secret, and a plain-text settings file is too easily copied, synced or shared.

### If `config.json` cannot be read

A `config.json` that is there but cannot be read, because of a typo or a half-saved edit,
stops the run before it touches anything. **No module is updated and no version is pruned.**

Going on would mean running on the built-in defaults, under which nothing is excluded,
pinned or kept. Every module would be updated and every old version removed, including the
ones the file was written to protect. Old logs are left alone too, since how long to keep
them is set in the same file.

```
[ERROR] Could not read the config file: <reason>
[ERROR] Nothing was updated or pruned. Without the config it is not known which modules are excluded, pinned or kept
[WARN] PSModuleMaintenance completed with 1 unsuccessful operation(s) (config: 1, lookups: 0, updates: 0, pins: 0, prunes: 0)
```

The run is reported as unsuccessful everywhere it can be heard:

- **Toast:** "The config file could not be read. Nothing was updated or pruned."
- **Healthchecks:** whether monitoring is on is one of the things the file would have said.
  The script looks for the ping URL under the default secret name anyway and, if it is
  there, sends a fail ping with the reason.

Fix the file and run the task again. To check a file after editing it:

```powershell
Test-Json -Path .\config.json -SchemaFile .\config.schema.json
```

A `config.json` that does not exist is a different matter and not a problem: the script
then runs on the built-in defaults, as it always has.

### If a single entry is not understood

An entry in `PinnedModules` or `KeepVersions` that cannot be understood, such as a pin of
`"2.*"` or a line written as `"5.x"`, says that something was wanted for that module but
not what. **That module is left alone for the run: not updated and not pruned**, exactly as
if it were excluded. Every other module is maintained as usual.

Ignoring the entry would treat the module like any other, and remove the very versions the
entry was most likely there to keep.

```
[ERROR] Contoso.Tools is left alone in this run, not updated and not pruned. A config entry for it is not understood (KeepVersions: '5.x' is not a version prefix such as 5 or 5.7)
[WARN] PSModuleMaintenance completed with 1 unsuccessful operation(s) (config: 1, lookups: 0, updates: 0, pins: 0, prunes: 0)
```

| Case | What happens |
|---|---|
| One selector in a list is wrong, the others are fine | The whole entry counts as not understood |
| The module has a good pin and a bad `KeepVersions` entry, or the other way round | Both are set aside, the module is left alone |
| A version or selector without quotes | Not understood. JSON reads `5.10` as the number 5.1 |
| An empty list of selectors | Not understood |
| The module is also in `ExcludedModules` | A warning only. It is left alone anyway, and the run counts as successful |
| A whole setting has the wrong shape, such as a list where `KeepVersions` expects an object | Counts as a file that cannot be read, see above. It is not known which modules were meant |

The run is reported as unsuccessful in the closing log line, the toast and the Healthchecks
ping until the entry is put right. However many modules are affected, it counts as one
failure.

## Version Pinning

Pinning holds a module at one specific version. Use it when a newer release breaks something and you need that module to stay put while everything else keeps updating.

```json
{
  "PinnedModules": {
    "Az.Accounts": "2.19.0",
    "Pester": "5.7.1"
  }
}
```

For each pinned module the script:

1. **Skips the normal update** — the module is never updated to the latest gallery version
2. **Installs the pinned version if it's missing** — including downgrading, since PSResourceGet installs versions side-by-side
3. **Reports what the pin is holding back** — pinned modules stay in the gallery lookup (same single bulk call either way), so the log tells you which release you're declining
4. **Prunes everything else** — unlike unpinned modules where the newest version is kept, pruning keeps the *pinned* version and removes all others, newer ones included

A run with `"PSWriteColor": "1.0.2"` pinned while 1.0.3 is installed logs:

```
[INFO] Enforcing 1 pinned module version(s)...
[INFO] Installing pinned version of PSWriteColor: 1.0.2 (installed: 1.0.3.0)
[SUCCESS] Installed pinned version: PSWriteColor 1.0.2 (took 3s)
[INFO] PSWriteColor is pinned to 1.0.2 — holding back 1.0.3
...
[INFO] Removing: PSWriteColor v1.0.3
[SUCCESS] Removed: PSWriteColor v1.0.3
```

### Rules and edge cases

- **The version must be exact.** `"2.19.0"` and `"2.19"` are the same pin (versions are padded to four parts, so `2.19` = `2.19.0.0`). Ranges and wildcards are not supported — `"2.*"` or `"latest"` is not understood, and the module is [left alone](#if-a-single-entry-is-not-understood) until the pin is put right.
- **Prerelease pins work**: `"6.0.0-beta1"`. The label must match exactly.
- **Pinning never installs a module you don't already have.** A pin for a module that isn't installed logs a note and does nothing.
- **If the pinned version can't be installed** (wrong version number, gallery unreachable), pruning leaves *all* installed versions of that module in place rather than deleting the ones you have. You'll see a `WARN` in the log and a `PinsFailed` entry in the summary.
- **`ExcludedModules` wins over a pin.** Listing a module in both is contradictory — exclusion means "never touch this", so the pin is dropped with a warning. Use exclusion when you want a module left entirely alone, including its old versions; use a pin when you want one specific version kept and the rest cleaned up.
- Pins are enforced during the update phase, so `-PruneOnly` skips enforcement but still prunes pin-aware.

## Keeping an Older Version Line

Some modules make a breaking change between major versions, and you need both for a while:
the old one for what has not been moved over yet, the new one for everything else. Pester
5 and 6 are a typical pair.

`KeepVersions` keeps an older version line installed next to the newest version. The
module is updated as normal, and the kept line is updated too, within itself.

```json
{
  "KeepVersions": {
    "Pester": ["5"]
  }
}
```

| You want | Use |
|----------|-----|
| A module left entirely alone | `ExcludedModules` |
| One version, and nothing newer | `PinnedModules` |
| The newest version **and** an older line, both kept current | `KeepVersions` |

For each module listed the script:

1. **Updates the module as normal** — the newest version moves on, 6.0.0 to 6.1.0
2. **Updates each kept line within itself** — a newer release inside the line is installed next to it, 5.5.0 to 5.6.1. A line never moves into the next one
3. **Keeps the newest version of each line when pruning** — and removes the older ones in that line, as it does everywhere else

```
[INFO] Checking kept version lines of 1 module(s) for updates...
[INFO] Updating kept line 5 of Pester: 5.5.0 -> 5.6.1
[SUCCESS] Updated kept line 5 of Pester to 5.6.1 (took 3s)
[INFO] Kept line updates complete. Updated: 1, Unsuccessful: 0, Not checked: 0
...
[INFO] Keeping Pester v5.6.1 (KeepVersions: 5)
[INFO] Removing: Pester v5.5.0
[SUCCESS] Removed: Pester v5.5.0
```

### What a line is

A selector is the start of a version number. It names every version that begins with it.

| Selector | Covers | Updated to |
|----------|--------|------------|
| `"5"` | 5.0.0, 5.7.1, 5.12.3 | The newest 5.x.y |
| `"5.7"` | 5.7.0, 5.7.4 | The newest 5.7.x |
| `"5.7.1"` | 5.7.1 and its four-part builds | The newest 5.7.1.x |
| `"5.7.1.0"` | That one version | Never updated, only kept |

Numbers are compared one by one, so `"1"` does not cover 12.4.0.

### Rules and edge cases

- **Write selectors in quotes.** JSON reads an unquoted `5.10` as the number 5.1, which is a different line. Anything that is not a quoted version prefix is not understood, and the module is then [left alone](#if-a-single-entry-is-not-understood) until the entry is put right.
- **A line is only maintained once you have a version of it.** If nothing installed matches a selector, the log warns and nothing else happens: the script never brings a line onto a machine that does not have it, and the run still counts as successful. Install a version of the line yourself once, and it is kept current from then on.
- **Only stable releases are installed.** A prerelease you installed yourself is kept like any other version, but a line update never picks one.
- **Two selectors can overlap.** With `["5", "5.7"]` the 5 line moves to the newest 5.x.y and the 5.7 line to the newest 5.7.x, and both results stay.
- **A selector for the newest line changes nothing.** `"6"` while 6 is the newest major is already covered by the normal update.
- **`ExcludedModules` wins.** A module listed in both is left alone and its `KeepVersions` entry is ignored with a warning.
- **Pinning and keeping combine.** The pin holds the module's main version, and the kept lines are maintained beside it. If the pinned version lies inside a kept line, the pin decides and that line is not updated.
- **A line that could not be updated is reported like any other update**, in the log, the toast and the Healthchecks ping. So is a line lookup that PSGallery did not answer.
- Lines are updated during the update phase, so `-PruneOnly` skips that but still keeps them when pruning.

## Notifications

PSModuleMaintenance can show a Windows toast notification after each run. Set `NotificationMode` in `config.json`:

- **`"Always"`** — Notification after every run with a summary of updates and pruning (default)
- **`"OnFailure"`** — Notification only when modules fail to update or versions fail to prune
- **`"Never"`** — No notifications

The toast uses the built-in Windows "Security and Maintenance" notification channel — no additional setup required.

## Monitoring

Toast notifications only fire when the script runs. A run that never *starts* — a broken
`pwsh.exe` path, a disabled task, a machine that never wakes — is completely silent.
Healthchecks.io monitoring closes that gap with a dead-man's switch: the alert fires on the
**absence** of a ping.

### Setup

Store the full ping URL in the SecretManagement vault. It is a bearer secret — anyone
holding it can send fake success pings and suppress your real alerts — so it stays out of
`config.json`:

```powershell
Set-Secret -Name PSModuleMaintenance-Healthchecks `
           -Secret 'https://hc-ping.com/your-uuid-here' -Vault SecretStore
```

Or let the installer do it:

```powershell
.\Install-ModuleMaintenance.ps1 -HealthchecksUrl "https://hc-ping.com/your-uuid-here"
```

Then set `Healthchecks.Enabled` to `true` in `config.json`.

The vault must be readable non-interactively, or the scheduled task will block on
`Get-Secret`. The installer warns if it is not. To configure it yourself:

```powershell
Set-SecretStoreConfiguration -Authentication None -Interaction None
```

Storing the **full URL** rather than just the UUID means self-hosted Healthchecks instances
work with no extra configuration.

### Configuring the check

Use a **Period** schedule, not cron:

| Setting | Value |
|---------|-------|
| Schedule | Period, 7 days |
| Grace | 1 day |

A cron schedule matching the task (`0 3 * * 0`) looks correct but produces false alarms.
The task uses `StartWhenAvailable`, so a sleeping machine catches up hours late. A
period-based check tolerates that while still reporting a missed week within a day.

### What gets sent

- A start ping when the run begins, so Healthchecks can track duration and spot a run that
  started but never finished.
- A success or fail ping at the end, with the run summary as the body:

```
PSModuleMaintenance - Fail
Host: DESKTOP-01   Mode: Full   Duration: 4m 21s
Checked: 150  Updated: 12  Pruned: 11
Pins: 0 enforced, 0 satisfied, 0 holding back
Issues: 1
  - prune Contoso.Tools 1.4.0: Access to the path 'Contoso.Tools.dll' is denied.
Log: C:\ProgramData\PSModuleMaintenance\Logs\maintenance_2024-01-15_030000.log
```

The body includes the machine name so one account can cover several machines, but
deliberately omits the user and domain — the body is embedded in alert emails, chat
messages and webhook payloads.

Any non-empty `ModulesFailed`, `PrunesFailed`, or `PinsFailed` sends a fail ping, as does a
`config.json` that could not be read, and so
does a non-zero `ModulesUnchecked` — a run that could not reach PSGallery has not checked
anything, so it must not reset the timer. It is the same rule as
`NotificationMode: OnFailure`, so the two channels never disagree.

### Behaviour notes

- **`-WhatIf` sends no pings at all.** A success ping from a dry run would reset the
  dead-man timer and hide a genuinely missed scheduled run.
- **A ping failure never fails the run.** Network problems are logged as a warning and the
  maintenance continues.
- **A broken vault is not silent.** If the secret cannot be read, the script logs a warning
  and carries on — and because no ping is sent, Healthchecks alerts you. Broken monitoring
  surfaces as a notification rather than as silence.
- A manual `-UpdateOnly` or `-PruneOnly` run also pings, with `Mode:` naming it in the body.

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
  "ModulesUnchecked": 0,
  "GalleryFault": null,
  "ModulesUpdated": 12,
  "ModulesFailed": [],
  "VersionsPruned": 45,
  "PrunesFailed": [],
  "ExcludedModules": ["Az.Accounts"],
  "PinnedModules": { "Pester": "5.7.1" },
  "PinsSatisfied": 1,
  "PinsEnforced": 0,
  "PinsFailed": [],
  "PinsHoldingBack": [
    { "Module": "Pester", "Pinned": "5.7.1", "Available": "6.0.1" }
  ],
  "ConfigFault": null,
  "ProtectedModules": [],
  "KeepVersions": { "PSReadLine": ["2"] },
  "KeepVersionsMatched": [
    { "Module": "PSReadLine", "Selector": "2", "Version": "2.3.6" }
  ],
  "KeepVersionsUnmatched": [],
  "KeepLinesUpdated": [
    { "Module": "PSReadLine", "Line": "2", "From": "2.3.5", "To": "2.3.6" }
  ],
  "KeepLinesUnchecked": []
}
```

`PinsSatisfied` counts modules already on their pinned version, `PinsEnforced` counts pins this run had to install, and `PinsHoldingBack` lists the newer releases each pin is declining — useful for periodically reviewing whether a pin is still needed.

`ModulesUnchecked` counts installed modules that PSGallery gave no answer for, with the reason in `GalleryFault`. Any non-zero value is reported as a single failure. `ModulesChecked` excludes them, so an unreachable gallery reads `"ModulesChecked": 0` rather than looking like a full check that found nothing. A module that simply is not on PSGallery counts as checked.

`ConfigFault` holds the reason when `config.json` could not be read, and is `null` otherwise. When it is set, every count in the summary is zero, because the run stopped before touching anything.

`ProtectedModules` lists the modules that were left alone because a config entry about them was not understood, each as `{ "Module": "...", "Problem": "..." }`. Anything in it makes the run count as unsuccessful.

`KeepVersionsMatched` lists the version each kept line resolved to, and `KeepLinesUpdated` the lines that were moved forward this run. `KeepVersionsUnmatched` holds selectors that matched nothing installed; that is a warning in the log, not a failure. `KeepLinesUnchecked` holds lines PSGallery gave no answer for, which does count, together with `ModulesUnchecked`, as a single failure.

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
2. **Copies** the modules PSResourceGet installed to AllUsers scope (`$env:ProgramFiles\PowerShell\Modules`) — safely, without deleting the originals
3. **Cleans up** the old OneDrive copies with four-stage force-removal for cloud placeholders and locked files (see [Troubleshooting](#onedrive-file-lock--cloud-placeholder-errors))

After migration, the weekly maintenance script (`Invoke-PSModuleMaintenance.ps1`) automatically detects OneDrive on the module path and targets AllUsers scope for all future updates and pruning — no configuration needed.

The migration is **idempotent and gradual** — modules that already exist at the destination are skipped, and OneDrive copies that can't be removed are either force-deleted or scheduled for reboot deletion.

**An original is only removed once its copy is in place.** If a version could not be copied, its OneDrive copy is kept, the log says so, and the run ends with a warning instead of a success. Run the migration again once the cause is out of the way:

```
[ERROR] Failed to copy Contoso.Tools v1.0.0: <reason>
[WARN] Keeping the OneDrive copy of Contoso.Tools v1.0.0: it is not in AllUsers
[WARN] OneDrive Module Migration completed with 1 module version(s) not migrated
```

If you skip migration, the weekly maintenance script will still work (it detects OneDrive and targets AllUsers scope automatically), but any modules left in the OneDrive path will trigger a warning in the logs: `Found N module(s) in OneDrive path — run Invoke-OneDriveMigration.ps1 to migrate them`.

#### Modules that belong to another program

Some programs install a PowerShell module of their own into `Documents\PowerShell\Modules`.
PowerToys does this with `Microsoft.PowerToys.Configure`, and writes it again on every
PowerToys update. Such a module is not installed through PSResourceGet, so this tool can
neither update nor prune it, and moving it would only leave an orphaned copy behind.

Both scripts recognise these modules by the absence of `PSGetModuleInfo.xml`, the marker
PSResourceGet writes into everything it installs, and leave them where they are. There is
nothing to configure. The weekly log notes them without raising a warning:

```
[INFO] Leaving 1 module(s) in OneDrive path alone, installed by another program: Microsoft.PowerToys.Configure
```

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

1. **Load Configuration** — Reads `config.json` for exclusions and settings, or uses the built-in defaults if there is none
2. **Initialize Logging** — Creates timestamped log files and starts transcript
3. **Resolve Monitoring Secret** — Reads the Healthchecks ping URL from the vault and sends a start ping. Done before any module work, because the script prunes `SecretManagement` itself
4. **Self-Check** — Warns if the scheduled task launches a hard-coded interpreter path that a PowerShell reinstall would break
5. **Clean Old Logs** — Removes logs older than retention period
6. **Update Modules** — Bulk checks PSGallery for available updates, then updates each module in an isolated runspace with a per-module timeout (targets AllUsers scope when OneDrive is detected). Transient network faults are retried; timeouts are not. If PSGallery cannot be reached at all, the run is reported as unsuccessful instead of as up to date. Version lines listed in `KeepVersions` are updated within themselves first
7. **Prune Versions** — Groups modules by name, keeps newest (or the pin) and the newest version of each kept line, removes the rest (skips built-in modules like PackageManagement). When OneDrive is detected and modules installed by PSResourceGet are found in the CurrentUser path, logs a warning to run `Invoke-OneDriveMigration.ps1`. Modules another program put there are left alone
8. **Save Summary** — Writes JSON summary after each phase (incremental saves protect against process termination)
9. **Toast Notification** — Shows a Windows toast with the run summary (if enabled via `NotificationMode`)
10. **Healthchecks Ping** — Sends a success or fail ping with the run summary as the body (if enabled)

## Troubleshooting

### Task doesn't run

Check Task Scheduler history. Common issues:
- Script path changed after installation
- User password changed (re-run installer)
- Network not available at scheduled time
- **PowerShell was reinstalled to a different directory** (see below)

#### PowerShell reinstalled, task silently stops running

This is the nastiest failure mode, because it leaves *no evidence*. The task fails to
start, so the script never runs — meaning no log file, no summary, no toast. The only
symptom is a week with no log at all.

It happens when the task was registered with a hard-coded interpreter path and PowerShell
later moves. MSI, the Microsoft Store package and pwshup ZIP installs all live in different
directories, so switching between them relocates `pwsh.exe`.

Current versions register the task with a **bare** `pwsh.exe`, which Task Scheduler
re-resolves from PATH at every run. A task created by an older installer keeps its baked-in
path until you re-register it:

```powershell
# Run as Administrator
.\Install-ModuleMaintenance.ps1
```

You do not need to guess whether you are affected — each run checks and logs a warning:

```
[WARN] Scheduled task 'PSModuleMaintenance' has a hard-coded interpreter path
       (C:\Program Files\PowerShell\7\pwsh.exe). It works today but stops working if
       PowerShell is reinstalled elsewhere — re-run Install-ModuleMaintenance.ps1 as
       Administrator to switch it to PATH resolution
```

To check the registration by hand:

```powershell
(Get-ScheduledTask -TaskName PSModuleMaintenance).Actions.Execute
# "pwsh.exe"                              -> resilient
# "C:\Program Files\PowerShell\7\pwsh.exe" -> fragile, re-run the installer
```

[Monitoring](#monitoring) is the backstop: even a task that never starts trips the
Healthchecks dead-man's switch within a day.

### Modules fail to update

Check the log files for specific errors. Common causes:
- Module removed from PSGallery
- Dependency conflicts
- Network/proxy issues

#### Network faults

A passing network fault — a dropped TLS handshake, DNS not ready after wake, a gateway
502/503/504 — is retried automatically. Each module gets up to three attempts, waiting 5
seconds and then 15 seconds in between:

```
[WARN] Network fault on Microsoft.WinGet.Client (attempt 1 of 3): SSL connection could not be established. Retrying in 5s
[SUCCESS] Updated: Microsoft.WinGet.Client (took 9s)
```

A module that succeeds on a later attempt counts as a normal success. Only a module that
fails all three attempts is reported as unsuccessful, and the run then ends with:

```
[WARN] PSModuleMaintenance completed with 1 unsuccessful operation(s) (config: 0, lookups: 0, updates: 1, pins: 0, prunes: 0)
```

This matters most on laptops: the task uses `StartWhenAvailable`, so a missed 03:00 run
fires the moment the machine wakes, often before the network has settled.

#### PSGallery unreachable

The check for available updates is retried on the same schedule. If PSGallery still cannot
be reached, the run says so and is reported as unsuccessful — it does **not** claim that
everything is up to date:

```
[WARN] PSGallery gave no answer for 150 module(s) (attempt 1 of 3): No such host is known. Retrying in 5s
[WARN] PSGallery gave no answer for 150 module(s) (attempt 2 of 3): No such host is known. Retrying in 15s
[ERROR] Could not reach PSGallery, so none of the 150 installed modules were checked for updates: ...
[WARN] PSModuleMaintenance completed with 1 unsuccessful operation(s) (config: 0, lookups: 1, updates: 0, pins: 0, prunes: 0)
```

An outage counts as one unsuccessful operation, not one per module. The number of modules
it affected is in the ERROR line and in `ModulesUnchecked` in the summary.

Pruning still runs, since it needs no network. If only some lookups go unanswered, the rest
are updated as normal and the log names the modules that were skipped.

A module that is installed but not published on PSGallery is not a fault and is never
reported this way.

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

## Tests

```powershell
.\tests\Invoke-Tests.ps1            # a few seconds
.\tests\Invoke-Tests.ps1 -Detailed  # show every check
```

The tests need no network and no elevation. They install nothing, send no notification and
no ping, and only write to a temp folder that they remove again.

They cover the parts a normal run almost never reaches: what happens when the network
drops, when PSGallery cannot be reached, and when a module belongs to another program.
They run the real functions, lifted out of the scripts through the PowerShell parser, and
stand in only for the calls that would touch the network or the machine. Two of them run
a script as a whole, in a process where those calls are replaced and checked before the
script starts.

| File | Covers |
|------|--------|
| `Test-Retry.ps1` | Which errors count as a network fault, and how an update is retried |
| `Test-GalleryLookup.ps1` | The PSGallery lookup, and how an incomplete one is reported in the log, the toast and the ping |
| `Test-ModuleOwnership.ps1` | Telling modules installed by PSResourceGet from another program's, and finding a version on disk |
| `Test-Config.ps1` | That `config.example.json` is valid, matches the built-in defaults, and that a missing `config.json` is fine. Loading `KeepVersions`, and telling apart a file that cannot be read, a setting of the wrong shape and a single entry that is not understood |
| `Test-ConfigFault.ps1` | The whole script is run against made-up modules. With a config file that cannot be read it must touch no module, keep its logs, show a toast and send a fail ping. With one entry that is not understood it must leave that module alone and maintain the others |
| `Test-KeepVersions.ps1` | Version lines: what a selector covers, what is kept and what is pruned, which lines are updated and to what |
| `Test-Migration.ps1` | The OneDrive migration, run as a whole against a made-up folder tree: what is copied, what is left in place, what is cleaned up, what happens when a copy does not succeed, and what the log says |

## Contributing

Issues and PRs welcome! Please include log output when reporting bugs, and run the tests
before opening a PR.

## License

MIT License - See [LICENSE](LICENSE) for details.

## Author

**Haakon Wibe**  
- Blog: [alttabtowork.com](https://alttabtowork.com)  
- Twitter: [@HaakonWibe](https://twitter.com/HaakonWibe)