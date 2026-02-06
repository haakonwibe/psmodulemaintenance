# Toast Notifications Design

**Date:** 2026-02-06
**Status:** Implemented

## Summary

Add native Windows toast notifications to PSModuleMaintenance so users get a popup summary after each maintenance run without checking logs manually.

## Configuration

New `NotificationMode` option in `config.json` with three values:

- **`"Always"`** — Toast after every run (default)
- **`"OnFailure"`** — Toast only when `ModulesFailed > 0` or `PrunesFailed > 0`
- **`"Never"`** — No toasts

```json
{
  "ExcludedModules": [],
  "LogRetentionDays": 180,
  "TrustPSGallery": true,
  "NotificationMode": "Never"
}
```

## Toast Content

The toast uses the `ToastGeneric` template:

```xml
<toast>
  <visual>
    <binding template="ToastGeneric">
      <text>PSModuleMaintenance</text>
      <text>Updated 148 modules. Pruned 53 versions. No errors.</text>
    </binding>
  </visual>
</toast>
```

**Message logic:**
- Always shows modules updated and versions pruned counts
- Appends failure counts only when there are failures
- If run was `UpdateOnly` or `PruneOnly`, only mentions the relevant operation

Examples:
- Clean run: `Updated 148 modules. Pruned 53 versions. No errors.`
- With failures: `Updated 148 modules, 4 failed. Pruned 53 versions, 47 failed.`
- Update only: `Updated 148 modules. No errors.`

## Implementation

### Approach

Shells out to **Windows PowerShell 5.1** (`powershell.exe`) which has native WinRT type loading via `ContentType = WindowsRuntime`. PowerShell 7 (.NET 5+) removed built-in WinRT support, making direct access unreliable. The PS 5.1 subprocess approach is zero-dependency (always present on Windows 10/11), proven reliable (same pattern used in Registry Configuration Engine), and keeps the main script running entirely in PS 7.

Key implementation details:
- Uses `| Out-Null` for WinRT type loading (not `[void]`)
- Uses `New-Object` for WinRT objects (not `::new()`)
- Uses `-EncodedCommand` to pass the script (avoids escaping issues)
- Adds `-WindowStyle Hidden` to suppress the PS 5.1 console flash
- XML-escapes the message via `[System.Security.SecurityElement]::Escape()`

### AppId

Uses the built-in Windows AppId `Windows.SystemToast.SecurityAndMaintenance`. This is a system-provided notification channel that requires no shortcut registration or setup — it works on any Windows 10/11 machine out of the box.

### New Function: `Send-ToastNotification`

Located in `Invoke-PSModuleMaintenance.ps1` in its own Toast Notifications section.

**Parameters:**
- `$Summary` — reuses the existing summary hashtable (no duplication)
- `-SkippedUpdates` — omits update counts from the message (set when `-PruneOnly`)
- `-SkippedPruning` — omits prune counts from the message (set when `-UpdateOnly`)

**Error handling:** If the toast fails (notification center disabled, Focus Assist blocking, subprocess exit code non-zero), it logs a warning via `Write-Log` and moves on. A toast failure never causes the maintenance run to fail.

### Call Site

Called in the `finally` block, right after `Save-Summary`:

```
if NotificationMode == "Always"  -> send toast
if NotificationMode == "OnFailure" and (failures > 0)  -> send toast
if NotificationMode == "Never"  -> skip
```

## File Changes

### Modified Files

1. **`config.json`** — Add `"NotificationMode": "Never"`
2. **`config.schema.json`** — Add `NotificationMode` property with enum `["Always", "OnFailure", "Never"]` and default
3. **`Invoke-PSModuleMaintenance.ps1`** — Three changes:
   - `Import-MaintenanceConfig`: add `NotificationMode` to defaults merge
   - New `Send-ToastNotification` function in the Logging section
   - Call it after `Save-Summary` in the main execution block
4. **`README.md`** — Add `NotificationMode` to the configuration table and a brief "Notifications" section

### No New Files or Dependencies

- Self-contained within existing scripts
- No new modules required
- No changes to update/prune logic
- No changes to existing logging behavior
