# Healthchecks.io Monitoring Design

**Date:** 2026-09-20
**Status:** Approved

## Summary

Add Healthchecks.io dead-man's-switch monitoring to PSModuleMaintenance so a scheduled
run that never happens raises an alert instead of passing unnoticed.

## Motivation

A weekly run did not happen. No log file was ever created for it, meaning the script never
started: the scheduled task action hard-codes the path to `pwsh.exe`, and that path had
stopped existing after PowerShell was reinstalled to a different directory. The task
started working again once a later reinstall happened to restore the directory, and the
following run completed normally.

Nothing surfaced the miss:

- `LastTaskResult` reports only the most recent run and read `0x0` afterwards.
- The Task Scheduler operational log is disabled by default.
- Toast notifications only fire when the script runs, so a script that never runs is silent.

Every existing signal is emitted *by the script*. A run that never starts produces no
signal at all. A dead-man's switch inverts this: the alert fires on the **absence** of a
ping, which is exactly the failure mode that went unnoticed.

## Configuration

New non-secret block in `config.json`, schema updated to match (`additionalProperties`
stays `false`):

```json
"Healthchecks": {
  "Enabled": true,
  "SecretName": "PSModuleMaintenance-Healthchecks",
  "SecretVault": "SecretStore",
  "TimeoutSeconds": 10
}
```

Defaults merge in `Import-MaintenanceConfig` like every other key. `Enabled` defaults to
`$false` in `$script:Config`, so a config file without the block is unaffected —
monitoring is opt-in and never becomes a surprise dependency.

## Secret Handling

`config.json` is tracked in git and pushed to GitHub, so the ping URL cannot live there.
A Healthchecks ping URL is a bearer secret: anyone holding it can send fake success pings
and suppress real alerts.

The URL is stored in the SecretManagement vault instead:

```powershell
Set-Secret -Name PSModuleMaintenance-Healthchecks `
           -Secret 'https://hc-ping.com/<uuid>' -Vault SecretStore
```

This machine's SecretStore is already automation-ready (`Scope: CurrentUser`,
`Authentication: None`, `Interaction: None`), so `Get-Secret` returns without a password
prompt from the non-interactive task. The vault is user-scoped and the task runs as the
same user with elevation, so the profile and DPAPI keys match.

**The full ping URL is the secret, not just the UUID.** Self-hosted and hosted
Healthchecks instances then need no extra configuration.

### `Get-HealthchecksUrl`

Called from Main Execution immediately after `Initialize-Logging` and **before any update
or prune work**. This script updates and prunes `SecretManagement` and `SecretStore`
themselves — they live in `C:\Program Files\PowerShell\Modules` — so the secret is read
before the script is in a position to prune its own dependency.

Behaviour:

- Returns `$null` immediately when `Enabled` is false, so the import never happens when unused.
- Soft-imports `Microsoft.PowerShell.SecretManagement` in `try`/`catch` rather than via
  `#Requires`, so a missing vault module cannot stop maintenance from running.
- Passes `-Vault` explicitly rather than trusting the default-vault flag.
- Validates the value is an `http(s)` URL and trims a trailing `/`.
- On any failure logs `-Level WARN` with actionable text and returns `$null`.

The result is held in `$script:HealthchecksUrl` and is never written to the log, the
transcript, or the summary JSON.

**Fail-safe property:** a broken vault means no ping, and no ping means Healthchecks
alerts. The monitoring cannot fail silently, because its failure mode is the alarm.

## `Send-HealthchecksPing`

One function in a new region after Toast Notifications:

```powershell
Send-HealthchecksPing -Event Start|Success|Fail [-Body <string>]
```

- URL is `$script:HealthchecksUrl` plus `/start`, `''`, or `/fail`.
- No-ops instantly when `$script:HealthchecksUrl` is `$null`, so call sites stay unconditional.
- `Invoke-RestMethod -Method Post`, `-TimeoutSec` from config, with PowerShell 7's
  `-MaximumRetryCount 2 -RetryIntervalSec 5` so a transient DNS blip at 03:00 does not
  manufacture a false alarm.
- Body sent as UTF-8 `text/plain`, defensively truncated at 10 KB.
- Wrapped in `try`/`catch` that logs `-Level WARN` and **never rethrows**. Monitoring must
  not be able to break maintenance.

### `-WhatIf` gating is deliberately the opposite of the logging convention

The existing convention gives bookkeeping writes `-WhatIf:$false` so logs and the summary
still work during a dry run. Pings go the other way. A `-WhatIf` run that fired a success
ping would reset the dead-man timer and mask a genuinely missed scheduled run.

`Send-HealthchecksPing` therefore checks `$WhatIfPreference` directly and skips, logging
`Healthchecks ping skipped (WhatIf): Success`. `Invoke-RestMethod` has no `ShouldProcess`
of its own, so this must be explicit rather than inherited — the same reasoning that left
`Start-Transcript` gated. Dry runs are completely invisible to Healthchecks.

## Call Sites

Two touch points in Main Execution.

In the `try`, after `Get-HealthchecksUrl` and before `Remove-OldLogs`:

```powershell
Send-HealthchecksPing -Event Start
```

The existing `catch` sets `$script:FatalError = $_` before rethrowing so `finally` can
distinguish a crash from a clean finish. In `finally`, the decision reuses the
`$hasFailures` expression already computed there for the toast, so the two notification
channels can never disagree about what "failed" means:

```powershell
$pingEvent = if ($script:FatalError -or $hasFailures) { 'Fail' } else { 'Success' }
Send-HealthchecksPing -Event $pingEvent -Body (Format-HealthchecksBody)
```

Any non-empty `ModulesFailed`, `PrunesFailed`, or `PinsFailed` sends `/fail`. A prune that
fails on a locked file goes red under this rule.

If the script dies before `Get-HealthchecksUrl` — a malformed `config.json`, say — no ping
is sent at all and the absence alerts. Correct by construction.

## Ping Body

```
PSModuleMaintenance — Fail
Host: DESKTOP-01   Mode: Full   Duration: 4m 21s
Checked: 150  Updated: 12  Pruned: 11
Pins: 0 enforced, 0 satisfied, 0 holding back
Issues: 1 prune failure
  - Contoso.Tools 1.4.0: Access to the path 'Contoso.Tools.dll' is denied.
Log: C:\ProgramData\PSModuleMaintenance\Logs\maintenance_2024-01-15_030000.log
```

The `Log:` line points the alert email straight at the file to open.

**Hostname is included; username and domain are not.** Ping bodies are not publicly
readable — reading them needs a dashboard login or a Management API key, and public status
badges expose only up/down state. The real exposure vectors are notification fan-out (the
body is embedded in alert emails, Slack messages, and webhook payloads) and third-party
retention on hosted healthchecks.io. The machine name is worth that exposure once the
script runs on more than one box; the domain-qualified username is not, and adds nothing
diagnostically.

Error strings may still carry paths such as `C:\Program Files\PowerShell\Modules\...`,
which are stock Windows locations that disclose nothing.

A manual `-UpdateOnly` or `-PruneOnly` run still pings and resets the dead-man timer, with
`Mode:` naming it in the body. This could in principle mask a missed scheduled run. If that
becomes a problem, a `-NoPing` switch is a one-liner.

## Healthchecks Check Configuration

Use a **Period** schedule, not cron.

The obvious choice is cron `0 3 * * 0` matching the weekly task. The run logs argue against
it: a machine that is asleep at the scheduled time runs the task when it wakes, and
`StartWhenAvailable` has been seen catching up hours late. Under cron with a tight grace
period, every sleeping machine produces a false alarm.

| Setting | Value |
| --- | --- |
| Schedule | Period, 7 days |
| Grace | 1 day |

This tolerates arbitrary catch-up timing while still reporting a missed week within a day.
The missed run that prompted this design would have alerted the day after.

## Installation

`Install-ModuleMaintenance.ps1` gains an optional `-HealthchecksUrl` parameter that calls
`Set-Secret` during install, making a new machine a single command. It also inspects
`Get-SecretStoreConfiguration` and warns when `Authentication` is not `None` — on a fresh
box the vault defaults to password-protected, and `Get-Secret` would block indefinitely in
a non-interactive scheduled task.

## Verification

1. `-WhatIf` run — log shows `ping skipped (WhatIf)`, Healthchecks records nothing.
2. Secret pointed at a throwaway check, `-PruneOnly` run — start and success pings land.
3. Bogus pin added to force `PinsFailed` — confirms `/fail`.
4. Secret renamed — confirms WARN, maintenance still completes, no ping sent.
5. Search the log, transcript, and summary JSON for the UUID — must return nothing.

## File Changes

1. **`config.json`** — add the `Healthchecks` block
2. **`config.schema.json`** — add the `Healthchecks` object with its properties and defaults
3. **`Invoke-PSModuleMaintenance.ps1`**
   - `$script:Config`: add `Healthchecks` defaults (`Enabled = $false`)
   - `Import-MaintenanceConfig`: merge the block
   - New Healthchecks region: `Get-HealthchecksUrl`, `Format-HealthchecksBody`, `Send-HealthchecksPing`
   - Main Execution: resolve URL, `Start` ping, `$script:FatalError` in `catch`, final ping in `finally`
4. **`Install-ModuleMaintenance.ps1`** — `-HealthchecksUrl` parameter and vault readiness warning
5. **`README.md`** — Healthchecks setup section and config table entry
6. **`CLAUDE.md`** — record the `-WhatIf` gating exception and the secret-before-prune ordering

## Out of Scope

- Fixing the per-module timeout defect found during design: `ModuleUpdateTimeoutSeconds` is
  600, yet a log shows a module taking more than twice that on the plain success path.
  Tracked separately; it is why the grace period cannot assume a bounded run.
- Monitoring the scheduled task's own existence or its hard-coded `pwsh.exe` path.
- Modules left in the OneDrive module path.
