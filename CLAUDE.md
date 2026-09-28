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
2. Logging - Initializes log files, transcript, and summary tracking. `Get-SummaryFailureCount` is the single definition of "the run had failures"
3. Toast Notifications - `Send-ToastNotification` function (shells out to PS 5.1 for WinRT support)
4. Healthchecks Monitoring - `Get-HealthchecksUrl`, `Format-HealthchecksBody`, and `Send-HealthchecksPing` (dead-man's-switch pings to Healthchecks.io)
5. Self-Checks - `Test-ScheduledTaskHealth` (warns when the scheduled task launches a hard-coded interpreter path)
6. OneDrive Utilities - `Test-OneDrivePath` and `Remove-LockedModuleFolder` functions (used for scope detection and locked-file fallback during updates/pruning)
7. Module Operations - version helpers (`ConvertTo-NormalizedVersion`, `Get-ModuleVersionKey`, `ConvertFrom-PinnedVersionString`, `Get-PinnedVersion`, `Test-IsPinnedVersion`) plus `Invoke-ModuleUpdate`, `Get-TransientNetworkFault`, `Invoke-ModuleUpdateWithRetry`, `Set-PinnedModuleVersions`, `Find-GalleryModules`, `Get-GalleryFaultText`, `Update-AllModules`, and `Remove-OldModuleVersions`
8. Main Execution - Orchestrates the workflow with try/finally for cleanup and incremental summary saves

**Invoke-OneDriveMigration.ps1** - Standalone one-time script that migrates modules from OneDrive-synced CurrentUser path to AllUsers scope, then cleans up OneDrive copies. Self-contained with its own copies of `Test-OneDrivePath`, `Remove-LockedModuleFolder`, and `Write-Log`. Requires Administrator.

**Install-ModuleMaintenance.ps1** - Creates Windows Scheduled Task that runs the maintenance script weekly as the current user with elevated privileges.

**config.json** - Runtime configuration validated against `config.schema.json`:
- `ExcludedModules`: Array of module names to skip
- `PinnedModules`: Object mapping module name to an exact version to hold (default `{}`)
- `LogRetentionDays`: How long to keep logs (default 180)
- `TrustPSGallery`: Trust repository during updates (default true)
- `NotificationMode`: Toast notifications — `Always` (default), `OnFailure`, or `Never`
- `ModuleUpdateTimeoutSeconds`: Per-module update timeout in seconds (default 600)
- `Healthchecks`: Object — `Enabled` (default false), `SecretName`, `SecretVault`, `TimeoutSeconds`. The ping URL itself is **never** stored here; it lives in the SecretManagement vault

## Key Implementation Details

- **Version pinning**: `PinnedModules` holds a module at one version. `Set-PinnedModuleVersions` runs first in the update phase and installs the pinned version via `Install-PSResource` when it's missing (a bare version string is a *required* version in PSResourceGet, not a minimum — see `about_PSResourceGet`, "NuGet version ranges"). Pinned modules stay in the `Find-PSResource` bulk query (same single call either way) so the log can report `<module> is pinned to <pin> — holding back <gallery>` and record it in `Summary.PinsHoldingBack`; they are filtered out of the update loop immediately after. In `Remove-OldModuleVersions` a pinned module keeps the pin instead of the newest version, so newer versions get pruned. **Safety invariant**: if the pinned version isn't installed, that module is pruned not at all — otherwise a bad pin would delete every version. Versions are compared via `Get-ModuleVersionKey` (4-part normalization + prerelease label), so `2.19`, `2.19.0`, and `2.19.0.0` are the same pin. Exclusion beats pinning: a module in both lists is dropped from `PinnedModules` at config load with a warning
- **Config warnings before logging exists**: `Import-MaintenanceConfig` runs before `Initialize-Logging`, so pin validation problems are queued in `$script:ConfigWarnings` and written to the log as WARN once logging is up
- **Per-module timeout**: Each `Update-PSResource` call runs in an isolated PowerShell runspace via `Invoke-ModuleUpdate`. If a module exceeds `ModuleUpdateTimeoutSeconds` (default 600s), the script gives up on it and moves to the next — prevents one slow/hung module (e.g. Microsoft.Graph) from consuming all scheduled task time. **The obvious implementation of this does not work**, and both traps were measured, not guessed:
    - **`$ps.Stop()` blocks** until non-cooperative native work finishes. PSResourceGet's network and file I/O do not cooperate, so a timed-out module still ran to completion — a controlled repro took 20s against a 5s timeout. Use `BeginStop` with a short grace (`$stopGraceSeconds`, 5s), then abandon the runspace to the GC rather than `Dispose`-ing it, since disposing a still-winding-down runspace can block just as badly
    - **`AsyncWaitHandle.WaitOne` cannot be trusted as the completion signal.** It was observed returning `$true` while the pipeline was still running, after which `EndInvoke` silently absorbed the rest of the runtime and the module was logged as SUCCESS. On 2026-09-20 Microsoft.Graph ran 1261s against a 600s timeout and reported success; every other module that run took 1-7s. The fix polls in 500ms slices against a `Stopwatch` and treats only a terminal `$ps.InvocationStateInfo.State` (`Completed`/`Failed`/`Stopped`) as done, so a spurious signal costs one extra lap instead of dropping out of the loop
    - **The runspace does not inherit `$ProgressPreference`.** A fresh session state means the script-level `SilentlyContinue` does not apply inside, so progress records pile up in `$ps.Streams.Progress` (20,000 in a repro) for output nothing renders. `Invoke-ModuleUpdate`'s script block sets it again explicitly
    - Terminating errors inside the script block surface from `EndInvoke` wrapped in a `MethodInvocationException`, so the `$ps.HadErrors` branch below it is only reached for non-terminating ones. The OneDrive locked-folder retry in `Update-AllModules` matches on the message text, which still contains the original `Cannot remove package path` string through the wrapper
- **Retry on transient network faults**: `Update-AllModules` and `Set-PinnedModuleVersions` call `Invoke-ModuleUpdateWithRetry`, never `Invoke-ModuleUpdate` directly. It makes up to three attempts, waiting 5s then 15s (`-RetryDelaySeconds`, not exposed in `config.json`), and logs each retry as WARN. A module that succeeds on a later attempt counts as a plain success: nothing lands in `Summary.ModulesFailed`, so the toast and the ping stay green. It exists because of the 2026-09-27 run, where both pending updates failed 2s apart with `The SSL connection could not be established`. The task had fired via `StartWhenAvailable` seconds after the machine left Modern Standby and joined a phone hotspot; the gallery was fine, but with no retry each module waited a week
    - **`Get-TransientNetworkFault` matches on message text because nothing else survives.** A failure inside the runspace arrives as `MethodInvocationException` -> `ActionPreferenceStopException` with no `InnerException` below that — PSResourceGet has already flattened the `HttpRequestException`/`SocketException` chain into one string. Measured against a dead proxy and an unresolvable host. Keep the pattern list narrow: a false positive costs 20s per module, and `404`/"not found" must never match
    - **Timeouts are not retried.** The module already spent its whole `ModuleUpdateTimeoutSeconds`, so retrying would multiply the runtime the timeout exists to bound. The wrapper rethrows the original `ErrorRecord`, so the callers' `catch [System.TimeoutException]` and the locked-folder message match both still work through it
    - **Worst case is bounded but not free**: a module that fails all three attempts costs about 20s of waiting on top of the failures themselves, per module. With the network fully down the update loop never gets that far, because `Find-GalleryModules` gives up first and no module is queued. Pins are different: they are enforced *before* the gallery lookup, so a missing pin does go through all three attempts
- **An unreachable gallery is a failure, not "up to date"**: `Find-PSResource` reports a network fault as a *non-terminating* error per module, so under `-ErrorAction SilentlyContinue` an unreachable PSGallery simply returns nothing. The run used to log `All modules are up to date` and send a **Success** ping, resetting the dead-man timer on a run that had checked nothing. Measured behind a dead proxy: 164 modules took 5m34s to fail silently. `Find-GalleryModules` now owns the lookup:
    - **`SilentlyContinue` has to stay**, because "not on the gallery" is an everyday answer for a module installed from elsewhere. The two cases are told apart by `FullyQualifiedErrorId`: `PackageNotFound,*` means answered, anything else (`HttpRequestCallFailure` for a network fault) means not answered. **Error ids survive here even though they do not survive the isolated runspace** — do not "simplify" this to the message matching `Get-TransientNetworkFault` uses
    - **A probe runs before the bulk query**: one lookup of `Microsoft.PowerShell.PSResourceGet` with `-ErrorAction Stop`. If it fails the bulk query is not sent. That is what took the unreachable case from 5m34s to 32s, and what stops a hanging network from costing one timeout per installed module
    - Unanswered lookups are retried on the same 5s/15s schedule, asking only about modules still without an answer. Whatever is left is counted in `Summary.ModulesUnchecked` (an int, with the reason in `Summary.GalleryFault`) and subtracted from `ModulesChecked`. `Get-SummaryFailureCount` includes it, which is what turns the toast, the ping and the closing line red
    - A partial failure does not stop the run: modules that did get an answer are updated as normal, the ERROR line names the ones that did not, and "All modules are up to date" becomes "No updates found for the N modules that could be checked"
    - The module name in a `PackageNotFound` error exists only in the message text. If a future PSResourceGet rewords it, those modules stay pending and get counted as unchecked — but only in a run that already has unanswered lookups, so it errs toward reporting too much, never too little
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
- **The closing log line reports the outcome, it does not assume it.** It is `[SUCCESS] PSModuleMaintenance completed successfully` only when `Get-SummaryFailureCount` is 0; otherwise it is a WARN: `completed with N unsuccessful operation(s) (lookups: w, updates: x, pins: y, prunes: z)`. It used to be unconditional, which on 2026-09-27 printed "completed successfully" directly above `Healthchecks ping sent: Fail`
- The scheduled task uses `Interactive` logon type with "run with highest privileges" (runs hidden, requires user to be logged in but `StartWhenAvailable` catches up if missed) with a 4-hour execution time limit
- **The task registers `pwsh.exe` as a bare name, never a resolved path.** Task Scheduler re-resolves a bare executable name from PATH at every run (verified empirically before the change). A hard-coded path is what broke the 2026-09-13 run: PowerShell moved from the MSI location to the Store package, `C:\Program Files\PowerShell\7\pwsh.exe` stopped existing, and the task failed to start — producing no log, no toast and no summary, because nothing in the script ever ran. `Install-ModuleMaintenance.ps1` still calls `Get-Command pwsh.exe` but only to validate PATH resolution and print the current location. `Test-ScheduledTaskHealth` warns on every run if it finds a task whose `Execute` contains a directory separator, since a task registered by an older version of the installer keeps its baked-in path until re-registered
- **Healthchecks monitoring**: a dead-man's switch for the failure mode every other signal misses — a run that never *starts*. Toast notifications, `LastTaskResult` and the log files are all emitted by the script, so a script that never runs is silent (this actually happened on 2026-09-13, when the task's hard-coded `pwsh.exe` path went missing during a PowerShell reinstall). Key constraints:
    - **The ping URL is a bearer secret and `config.json` is tracked in git**, so it lives in the SecretManagement vault (`Get-Secret -Vault SecretStore`), never in config. The *full URL* is stored, not just the UUID, so self-hosted instances need no extra config. It is held in `$script:HealthchecksUrl` and never written to the log, transcript or summary
    - **`Get-HealthchecksUrl` runs before any update/prune work** — this script updates and prunes `SecretManagement`/`SecretStore` themselves, so the secret must be read while the module is still guaranteed loadable
    - **Soft-import, never `#Requires`**: a broken vault logs a WARN and maintenance continues. That is the fail-safe — no ping means the check goes overdue and Healthchecks alerts, so broken monitoring surfaces as a notification rather than as silence
    - **`-WhatIf` suppresses pings, which is the opposite of the logging convention below.** Bookkeeping writes pass `-WhatIf:$false` so dry runs still log; a ping must go the other way, because a success ping from a dry run would reset the dead-man timer and hide a genuinely missed run. `Invoke-RestMethod` has no ShouldProcess, so `Send-HealthchecksPing` checks `$WhatIfPreference` explicitly
    - The fail rule reuses the `$hasFailures` value already computed in `finally` for `NotificationMode: OnFailure`, which in turn comes from `Get-SummaryFailureCount`. The toast, the ping and the closing log line all go through that one function, so they can never disagree about what counts as a failure
    - Configure the check with a **Period** schedule (7 days, 1 day grace), *not* cron. The task uses `StartWhenAvailable`, so a sleeping machine catches up hours late — the 2026-09-20 run fired at 06:29 instead of 03:00 — and cron with a tight grace would false-alarm every time
- `Invoke-PSModuleMaintenance.ps1` and `Invoke-OneDriveMigration.ps1` use `[CmdletBinding(SupportsShouldProcess)]` for `-WhatIf` support (`Install-ModuleMaintenance.ps1` does not — it is plain `[CmdletBinding()]`). Because `-WhatIf` sets `$WhatIfPreference` for the whole script scope, it cascades into ShouldProcess-aware cmdlets including the bookkeeping file writes. `Write-Log`'s `Out-File`/`Add-Content`, `Save-Summary`'s `Set-Content`, and the log-directory `New-Item` in both `Initialize-Logging` and `Invoke-OneDriveMigration.ps1` are called with `-WhatIf:$false` so logging and the JSON summary still run (and don't spam `What if: Performing the operation "Output to File"...`) during a dry run — only the real actions (copy/remove/update/prune) are gated by `-WhatIf`. The log directory is easy to miss because `%ProgramData%\PSModuleMaintenance\Logs` already exists in normal use; the gap only shows up on a dry run against a *fresh* `-LogPath`, where the missing directory silently suppresses both the log and the summary. `Start-Transcript` is deliberately left gated — the transcript is skipped during dry runs
