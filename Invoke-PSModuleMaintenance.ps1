<#
.SYNOPSIS
    Automated PowerShell module maintenance - updates modules and prunes old versions.

.DESCRIPTION
    This script performs automated maintenance of PowerShell modules installed via PSResourceGet:
    - Updates all installed modules to their latest versions
    - Removes old versions while keeping the latest
    - Logs all operations with configurable retention
    - Supports module exclusions via config file

.PARAMETER ConfigPath
    Path to the JSON configuration file. Defaults to script directory's config.json.

.PARAMETER LogPath
    Base path for logs. Defaults to $env:ProgramData\PSModuleMaintenance\Logs

.PARAMETER UpdateOnly
    Only perform updates, skip pruning old versions.

.PARAMETER PruneOnly
    Only prune old versions, skip updates.

.PARAMETER WhatIf
    Show what would be done without making changes.

.EXAMPLE
    .\Invoke-PSModuleMaintenance.ps1

.EXAMPLE
    .\Invoke-PSModuleMaintenance.ps1 -PruneOnly -WhatIf

.NOTES
    Author: Haakon Wibe
    Requires: PowerShell 7+, Microsoft.PowerShell.PSResourceGet
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter()]
    [string]$ConfigPath,

    [Parameter()]
    [string]$LogPath = "$env:ProgramData\PSModuleMaintenance\Logs",

    [Parameter()]
    [switch]$UpdateOnly,

    [Parameter()]
    [switch]$PruneOnly
)

#Requires -Version 7.0
#Requires -Modules Microsoft.PowerShell.PSResourceGet

# Suppress progress bars — they add significant overhead in non-interactive/scheduled task mode
$ProgressPreference = 'SilentlyContinue'

# ============================================================================
# CONFIGURATION
# ============================================================================

$script:Config = @{
    ExcludedModules = @()
    PinnedModules = @{}
    KeepVersions = @{}
    LogRetentionDays = 180
    TrustPSGallery = $true
    NotificationMode = 'Always'
    ModuleUpdateTimeoutSeconds = 600
    Healthchecks = @{
        Enabled        = $false
        SecretName     = 'PSModuleMaintenance-Healthchecks'
        SecretVault    = 'SecretStore'
        TimeoutSeconds = 10
    }
}

# Config is loaded before logging starts, so problems found there are queued and
# written to the log once Initialize-Logging has run
$script:ConfigWarnings = @()

# Why the config file could not be read, if it could not. When this is set the run
# touches no module at all: without the config it is not known which modules are
# excluded, pinned or kept, and the built-in defaults would update and prune them all
$script:ConfigFault = $null

# Healthchecks ping URL, resolved once at startup from the SecretManagement vault.
# Held in memory for the run and deliberately never written to the log, the transcript
# or the summary JSON — the ping URL is a bearer secret
$script:HealthchecksUrl = $null

# Set by the main catch block before it rethrows, so the finally block can tell a
# crash apart from a clean finish when deciding which ping to send
$script:FatalError = $null

function ConvertTo-PinnedModuleTable {
    <#
    .SYNOPSIS
        Converts the PinnedModules JSON object into a hashtable of parsed pin objects.
        Entries with an unparseable version are dropped with a warning.
    #>
    [CmdletBinding()]
    param($PinnedConfig)

    $table = @{}
    foreach ($entry in $PinnedConfig.PSObject.Properties) {
        $parsed = ConvertFrom-PinnedVersionString $entry.Value
        if (-not $parsed) {
            $script:ConfigWarnings += "Ignoring pin for '$($entry.Name)': '$($entry.Value)' is not a valid version"
            continue
        }
        $table[$entry.Name] = $parsed
    }
    return $table
}

function ConvertTo-KeepVersionTable {
    <#
    .SYNOPSIS
        Converts the KeepVersions JSON object into a hashtable of module name to parsed
        selectors. Anything unusable is dropped with a warning.

    .NOTES
        The table has to be a plain @{}, which looks names up without regard to case. The
        name in the config and the name of the installed module need not be cased alike.
    #>
    [CmdletBinding()]
    param($KeepConfig)

    $table = @{}

    if ($KeepConfig -isnot [PSCustomObject]) {
        $script:ConfigWarnings += 'Ignoring KeepVersions: it has to be an object that maps a module name to a list of version prefixes'
    }
    else {
        foreach ($entry in $KeepConfig.PSObject.Properties) {
            $selectors = @()
            $seen = @()

            # @() also accepts a single value written without the list brackets
            foreach ($value in @($entry.Value)) {
                if ($null -eq $value) {
                    continue
                }
                if ($value -isnot [string]) {
                    $script:ConfigWarnings += "Ignoring a KeepVersions selector for '$($entry.Name)': it has to be a string in quotes, such as `"5`""
                    continue
                }

                $selector = ConvertFrom-KeepVersionSelector -Value $value
                if (-not $selector) {
                    $script:ConfigWarnings += "Ignoring KeepVersions selector for '$($entry.Name)': '$value' is not a version prefix such as 5 or 5.7"
                    continue
                }

                if ($selector.Key -notin $seen) {
                    $seen += $selector.Key
                    $selectors += $selector
                }
            }

            if ($selectors.Count -eq 0) {
                $script:ConfigWarnings += "Ignoring KeepVersions for '$($entry.Name)': the list holds no usable selector"
                continue
            }

            $table[$entry.Name] = @($selectors)
        }
    }

    return $table
}

function Import-MaintenanceConfig {
    [CmdletBinding()]
    param([string]$Path)

    if ([string]::IsNullOrEmpty($Path)) {
        $Path = Join-Path $PSScriptRoot 'config.json'
    }

    if (Test-Path $Path) {
        try {
            $jsonConfig = Get-Content $Path -Raw | ConvertFrom-Json
            
            if ($jsonConfig.ExcludedModules) {
                $script:Config.ExcludedModules = @($jsonConfig.ExcludedModules)
            }
            if ($jsonConfig.PinnedModules) {
                $script:Config.PinnedModules = ConvertTo-PinnedModuleTable $jsonConfig.PinnedModules
            }
            if ($jsonConfig.KeepVersions) {
                $script:Config.KeepVersions = ConvertTo-KeepVersionTable $jsonConfig.KeepVersions
            }
            if ($null -ne $jsonConfig.LogRetentionDays) {
                $script:Config.LogRetentionDays = $jsonConfig.LogRetentionDays
            }
            if ($null -ne $jsonConfig.TrustPSGallery) {
                $script:Config.TrustPSGallery = $jsonConfig.TrustPSGallery
            }
            if ($null -ne $jsonConfig.NotificationMode) {
                $script:Config.NotificationMode = $jsonConfig.NotificationMode
            }
            if ($null -ne $jsonConfig.ModuleUpdateTimeoutSeconds) {
                $script:Config.ModuleUpdateTimeoutSeconds = $jsonConfig.ModuleUpdateTimeoutSeconds
            }
            if ($null -ne $jsonConfig.Healthchecks) {
                # Merge key by key so a config that sets only Enabled keeps the other defaults
                foreach ($hcKey in @('Enabled', 'SecretName', 'SecretVault', 'TimeoutSeconds')) {
                    $hcValue = $jsonConfig.Healthchecks.$hcKey
                    if ($null -ne $hcValue) {
                        $script:Config.Healthchecks[$hcKey] = $hcValue
                    }
                }
            }

            # A module that is both excluded and pinned is contradictory — exclusion means
            # "never touch it", so it wins and the pin is dropped
            foreach ($name in @($script:Config.PinnedModules.Keys)) {
                if ($name -in $script:Config.ExcludedModules) {
                    $script:ConfigWarnings += "'$name' is both excluded and pinned — exclusion takes precedence, the pin is ignored"
                    $script:Config.PinnedModules.Remove($name)
                }
            }

            # The same goes for a module that is both excluded and has versions to keep
            foreach ($name in @($script:Config.KeepVersions.Keys)) {
                if ($name -in $script:Config.ExcludedModules) {
                    $script:ConfigWarnings += "'$name' is both excluded and listed in KeepVersions - exclusion takes precedence, the KeepVersions entry is ignored"
                    $script:Config.KeepVersions.Remove($name)
                }
            }

            Write-Verbose "Loaded configuration from: $Path"
        }
        catch {
            Write-Warning "Failed to parse config file: $_"

            # Write-Warning only reaches the console, and nobody watches the console of a
            # scheduled task. The main block reads this, writes it to the log and stops
            # before any module is touched
            $script:ConfigFault = $_.Exception.Message
        }
    }
    else {
        Write-Verbose "No config file found at $Path. Using defaults."
    }
}

# ============================================================================
# LOGGING
# ============================================================================

$script:LogFile = $null
$script:TranscriptFile = $null
$script:Summary = @{
    StartTime = $null
    EndTime = $null
    ModulesChecked = 0
    ModulesUnchecked = 0
    GalleryFault = $null
    ModulesUpdated = 0
    ModulesFailed = @()
    VersionsPruned = 0
    PrunesFailed = @()
    ExcludedModules = @()
    PinnedModules = @{}
    PinsSatisfied = 0
    PinsEnforced = 0
    PinsFailed = @()
    PinsHoldingBack = @()
    ConfigFault = $null
    KeepVersions = @{}
    KeepVersionsMatched = @()
    KeepVersionsUnmatched = @()
    KeepLinesUpdated = @()
    KeepLinesUnchecked = @()
}

function Initialize-Logging {
    [CmdletBinding()]
    param([string]$BasePath)

    # Ensure log directory exists. -WhatIf:$false for the same reason as the writes in
    # Write-Log and Save-Summary: without the directory, a dry run against a log path that
    # doesn't exist yet silently produces no log file and no summary.
    if (-not (Test-Path $BasePath)) {
        New-Item -Path $BasePath -ItemType Directory -Force -WhatIf:$false | Out-Null
    }

    $timestamp = Get-Date -Format 'yyyy-MM-dd_HHmmss'
    $script:LogFile = Join-Path $BasePath "maintenance_$timestamp.log"
    $script:TranscriptFile = Join-Path $BasePath "transcript_$timestamp.log"

    # Start transcript for full verbose capture
    Start-Transcript -Path $script:TranscriptFile -Force | Out-Null

    Write-Log "======================================================"
    Write-Log "PSModuleMaintenance started"
    Write-Log "======================================================"
    Write-Log "PowerShell Version: $($PSVersionTable.PSVersion)"
    Write-Log "Running as: $([System.Security.Principal.WindowsIdentity]::GetCurrent().Name)"
    Write-Log "Log file: $script:LogFile"
}

function Write-Log {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, Position = 0)]
        [string]$Message,

        [Parameter()]
        [ValidateSet('INFO', 'WARN', 'ERROR', 'SUCCESS')]
        [string]$Level = 'INFO'
    )

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $logEntry = "[$timestamp] [$Level] $Message"

    # Write to log file
    if ($script:LogFile) {
        Add-Content -Path $script:LogFile -Value $logEntry -WhatIf:$false
    }

    # Write to console with colors
    switch ($Level) {
        'WARN'    { Write-Host $logEntry -ForegroundColor Yellow }
        'ERROR'   { Write-Host $logEntry -ForegroundColor Red }
        'SUCCESS' { Write-Host $logEntry -ForegroundColor Green }
        default   { Write-Host $logEntry }
    }
}

function Save-Summary {
    [CmdletBinding()]
    param([string]$BasePath)

    $script:Summary.EndTime = Get-Date -Format 'o'

    # Reuse the same file path so incremental saves overwrite rather than create duplicates
    if (-not $script:SummaryFile) {
        $script:SummaryFile = Join-Path $BasePath "summary_$(Get-Date -Format 'yyyy-MM-dd_HHmmss').json"
    }

    $script:Summary | ConvertTo-Json -Depth 3 | Set-Content -Path $script:SummaryFile -WhatIf:$false

    Write-Log "Summary saved to: $script:SummaryFile"
}

function Get-SummaryFailures {
    <#
    .SYNOPSIS
        What did not succeed this run, by kind and in total. This is the one definition of
        "the run had failures", shared by the closing log line, the toast and the
        Healthchecks ping so the three cannot disagree.

    .NOTES
        The gallery lookup counts once, however many modules it left unchecked. An outage
        is one thing going wrong, and counting it per module turned a single unreachable
        gallery into one unsuccessful operation for every installed module. How many
        modules were affected is still in Summary.ModulesUnchecked and in the log.

        It does have to count, though. Leaving it out is how an unreachable gallery used
        to pass as a clean run.

        A kept version line that could not be looked up is the same kind of thing, and
        shares that one count. A kept line whose update did not succeed is an update like
        any other and sits in ModulesFailed. A KeepVersions selector that matches nothing
        installed is deliberately not counted at all: it only gets a line in the log.

        A config file that could not be read counts as one. Nothing else can have gone
        wrong in such a run, because it stops before any module is touched.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Summary
    )

    $config = 0
    if ($Summary.ConfigFault) {
        $config = 1
    }

    $lookups = 0
    if (($Summary.ModulesUnchecked -gt 0) -or ($Summary.KeepLinesUnchecked.Count -gt 0)) {
        $lookups = 1
    }
    $updates = $Summary.ModulesFailed.Count
    $pins = $Summary.PinsFailed.Count
    $prunes = $Summary.PrunesFailed.Count

    return [PSCustomObject]@{
        Config  = $config
        Lookups = $lookups
        Updates = $updates
        Pins    = $pins
        Prunes  = $prunes
        Total   = $config + $lookups + $updates + $pins + $prunes
    }
}

function Remove-OldLogs {
    [CmdletBinding()]
    param(
        [string]$BasePath,
        [int]$RetentionDays
    )

    if (-not (Test-Path $BasePath)) {
        return
    }

    $cutoffDate = (Get-Date).AddDays(-$RetentionDays)
    $oldFiles = Get-ChildItem -Path $BasePath -File | Where-Object { $_.LastWriteTime -lt $cutoffDate }

    if ($oldFiles) {
        Write-Log "Removing $($oldFiles.Count) log files older than $RetentionDays days"
        $oldFiles | Remove-Item -Force -ErrorAction SilentlyContinue
    }
}

# ============================================================================
# TOAST NOTIFICATIONS
# ============================================================================

function Send-ToastNotification {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Summary,

        [switch]$SkippedUpdates,

        [switch]$SkippedPruning
    )

    try {
        # Build the notification message from summary data
        $parts = @()

        if (-not $SkippedUpdates) {
            $updateText = "Updated $($Summary.ModulesUpdated) modules"
            if ($Summary.ModulesFailed.Count -gt 0) {
                $updateText += ", $($Summary.ModulesFailed.Count) unsuccessful"
            }
            if ($Summary.ModulesUnchecked -gt 0) {
                $updateText += ", $($Summary.ModulesUnchecked) not checked"
            }
            if ($Summary.KeepLinesUnchecked.Count -gt 0) {
                $updateText += ", $($Summary.KeepLinesUnchecked.Count) kept line(s) not checked"
            }
            $parts += $updateText
        }

        # Only mention pins when something actually happened to one
        if (-not $SkippedUpdates -and ($Summary.PinsEnforced -gt 0 -or $Summary.PinsFailed.Count -gt 0)) {
            $pinText = "Pinned $($Summary.PinsEnforced) modules"
            if ($Summary.PinsFailed.Count -gt 0) {
                $pinText += ", $($Summary.PinsFailed.Count) unsuccessful"
            }
            $parts += $pinText
        }

        if (-not $SkippedPruning) {
            $pruneText = "Pruned $($Summary.VersionsPruned) versions"
            if ($Summary.PrunesFailed.Count -gt 0) {
                $pruneText += ", $($Summary.PrunesFailed.Count) unsuccessful"
            }
            $parts += $pruneText
        }

        $hasFailures = (Get-SummaryFailures -Summary $Summary).Total -gt 0

        $message = ($parts -join '. ') + '.'
        if (-not $hasFailures) {
            $message += ' No issues.'
        }

        if ($Summary.ConfigFault) {
            # Counts of 0 would read like a quiet week. Say what actually happened
            $message = 'The config file could not be read. Nothing was updated or pruned.'
        }

        # Use Windows PowerShell (5.1) for native WinRT toast support — always present on Windows 10/11
        $xmlMessage = [System.Security.SecurityElement]::Escape($message)

        $toastScript = @"
`$ErrorActionPreference = 'Stop'
try {
    [Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime] | Out-Null
    [Windows.Data.Xml.Dom.XmlDocument, Windows.Data.Xml.Dom.XmlDocument, ContentType = WindowsRuntime] | Out-Null

    `$toastXml = @'
<toast>
    <visual>
        <binding template="ToastGeneric">
            <text>PSModuleMaintenance</text>
            <text>$xmlMessage</text>
        </binding>
    </visual>
    <audio src="ms-winsoundevent:Notification.Default"/>
</toast>
'@

    `$xmlDoc = New-Object Windows.Data.Xml.Dom.XmlDocument
    `$xmlDoc.LoadXml(`$toastXml)
    `$toast = New-Object Windows.UI.Notifications.ToastNotification(`$xmlDoc)
    `$notifier = [Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier('Windows.SystemToast.SecurityAndMaintenance')
    `$notifier.Show(`$toast)
}
catch {
    exit 1
}
"@

        $encoded = [Convert]::ToBase64String([System.Text.Encoding]::Unicode.GetBytes($toastScript))
        powershell.exe -NoProfile -NonInteractive -WindowStyle Hidden -EncodedCommand $encoded

        if ($LASTEXITCODE -eq 0) {
            Write-Log "Toast notification sent: $message"
        }
        else {
            Write-Log "Toast notification subprocess exited with code $LASTEXITCODE" -Level WARN
        }
    }
    catch {
        Write-Log "Toast notification failed: $_" -Level WARN
    }
}

# ============================================================================
# HEALTHCHECKS MONITORING
# ============================================================================

function Get-HealthchecksUrl {
    <#
    .SYNOPSIS
        Resolves the Healthchecks ping URL from the SecretManagement vault.

        Must be called before any update or prune work. This script updates and prunes
        SecretManagement and SecretStore themselves — they live under
        C:\Program Files\PowerShell\Modules — so the secret has to be read while the
        module is still guaranteed loadable.

        Returns $null on any problem instead of throwing. That is the point: no ping
        means the check goes overdue and Healthchecks alerts, so broken monitoring
        surfaces as a notification rather than as silence.

    .PARAMETER WithoutConfig
        For a run whose config file could not be read. Whether monitoring is switched on
        is one of the things that file would have said, so this looks for the secret
        under the name and vault that are in force, which are the built-in defaults, and
        uses it if it is there. If it is not, that is not a problem worth a warning:
        monitoring may simply never have been set up.
    #>
    [CmdletBinding()]
    param(
        [switch]$WithoutConfig
    )

    $hc = $script:Config.Healthchecks

    if ((-not $hc.Enabled) -and (-not $WithoutConfig)) {
        Write-Log "Healthchecks monitoring is off"
        return $null
    }

    $secretName = $hc.SecretName
    $secretVault = $hc.SecretVault

    if ($WithoutConfig) {
        $found = $null
        try {
            Import-Module Microsoft.PowerShell.SecretManagement -ErrorAction Stop

            $lookupParams = @{
                Name        = $secretName
                AsPlainText = $true
                ErrorAction = 'Stop'
            }
            if ($secretVault) { $lookupParams['Vault'] = $secretVault }

            $found = Get-Secret @lookupParams
        }
        catch {
            $found = $null
        }

        $usable = $null
        if (-not [string]::IsNullOrWhiteSpace($found)) {
            $candidate = "$found".Trim().TrimEnd('/')
            if ($candidate -match '^https?://') {
                $usable = $candidate
            }
        }

        if ($usable) {
            Write-Log "Healthchecks secret found under its default name '$secretName' - this run will be reported"
        }
        else {
            Write-Log "No usable Healthchecks secret under the default name '$secretName' - this run will not be reported there"
        }
        return $usable
    }

    try {
        Import-Module Microsoft.PowerShell.SecretManagement -ErrorAction Stop

        $secretParams = @{
            Name        = $secretName
            AsPlainText = $true
            ErrorAction = 'Stop'
        }
        # Name the vault explicitly rather than relying on which one is flagged default
        if ($secretVault) { $secretParams['Vault'] = $secretVault }

        $url = Get-Secret @secretParams
    }
    catch {
        $reason = $_.Exception.Message
        Write-Log "Healthchecks ping URL unavailable ($reason) - runs will not be monitored. Store it with: Set-Secret -Name '$secretName' -Vault '$secretVault'" -Level WARN
        return $null
    }

    if ([string]::IsNullOrWhiteSpace($url)) {
        Write-Log "Healthchecks secret '$secretName' is empty - runs will not be monitored" -Level WARN
        return $null
    }

    $url = $url.Trim().TrimEnd('/')

    if ($url -notmatch '^https?://') {
        Write-Log "Healthchecks secret '$secretName' is not an http(s) URL - runs will not be monitored" -Level WARN
        return $null
    }

    Write-Log "Healthchecks monitoring is on"
    return $url
}

function Format-HealthchecksBody {
    <#
    .SYNOPSIS
        Builds the plain-text ping body posted alongside the success/fail ping.

        Includes the machine name, because one Healthchecks account may cover several
        machines. Deliberately omits the user and domain: the body is embedded in alert
        emails, chat messages and webhook payloads, and the account name adds nothing
        diagnostically.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        $Summary,

        [string]$Mode = 'Full',

        [switch]$IsFailure
    )

    $status = if ($IsFailure) { 'Fail' } else { 'Success' }

    $duration = 'unknown'
    if ($Summary.StartTime -and $Summary.EndTime) {
        $span = [datetime]$Summary.EndTime - [datetime]$Summary.StartTime
        # Floor, not [int] — a PowerShell [int] cast ROUNDS, so a 51s run would report
        # "1m 51s" (TotalMinutes 0.85 rounds to 1). TotalMinutes rather than .Minutes so
        # a run past the hour reads "75m 30s" instead of wrapping to "15m 30s"
        $totalMinutes = [int][math]::Floor($span.TotalMinutes)
        $duration = '{0}m {1}s' -f $totalMinutes, $span.Seconds
    }

    $checked = $Summary.ModulesChecked
    $updated = $Summary.ModulesUpdated
    $pruned = $Summary.VersionsPruned
    $pinsEnforced = $Summary.PinsEnforced
    $pinsSatisfied = $Summary.PinsSatisfied
    $pinsHolding = @($Summary.PinsHoldingBack).Count
    $machine = $env:COMPUTERNAME

    $lines = @(
        "PSModuleMaintenance - $status"
        "Host: $machine   Mode: $Mode   Duration: $duration"
        "Checked: $checked  Updated: $updated  Pruned: $pruned"
        "Pins: $pinsEnforced enforced, $pinsSatisfied satisfied, $pinsHolding holding back"
    )

    $issues = @()
    if ($Summary.ConfigFault) {
        $issues += "  - config: could not be read, so nothing was updated or pruned: $($Summary.ConfigFault)"
    }
    if ($Summary.ModulesUnchecked -gt 0) {
        $issues += "  - lookup: $($Summary.ModulesUnchecked) module(s) not checked: $($Summary.GalleryFault)"
    }
    $uncheckedLines = @($Summary.KeepLinesUnchecked | Where-Object { $_ })
    if ($uncheckedLines.Count -gt 0) {
        $issues += "  - lookup: $($uncheckedLines.Count) kept line(s) not checked: $($uncheckedLines[0].Fault)"
    }
    foreach ($item in @($Summary.ModulesFailed)) {
        $what = $item.Module
        if ($item.Line) {
            # A kept version line, so say which one
            $what = "$($item.Module) (kept line $($item.Line))"
        }
        $issues += "  - update ${what}: $($item.Error)"
    }
    foreach ($item in @($Summary.PrunesFailed)) {
        $issues += "  - prune $($item.Module) $($item.Version): $($item.Error)"
    }
    foreach ($item in @($Summary.PinsFailed)) {
        $issues += "  - pin $($item.Module): $($item.Error)"
    }

    if ($issues.Count -gt 0) {
        $lines += "Issues: $($issues.Count)"
        $lines += $issues
    }
    else {
        $lines += "Issues: none"
    }

    if ($script:LogFile) {
        $lines += "Log: $script:LogFile"
    }

    $body = $lines -join "`n"

    # Healthchecks caps the stored body; keep it well under and leave a marker if trimmed
    if ($body.Length -gt 10240) {
        $body = $body.Substring(0, 10240) + "`n[truncated]"
    }

    return $body
}

function Send-HealthchecksPing {
    <#
    .SYNOPSIS
        Sends a start, success or fail ping. Never throws — a monitoring problem must
        not be able to break a maintenance run.

        Pings are suppressed under -WhatIf, which is the opposite of how logging is
        handled in this script. Write-Log and Save-Summary pass -WhatIf:$false so dry
        runs still produce a log and a summary. A ping must go the other way: a success
        ping from a dry run would reset the dead-man timer and hide a scheduled run that
        never happened. Invoke-RestMethod has no ShouldProcess of its own, so the check
        has to be made explicitly here rather than inherited.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateSet('Start', 'Success', 'Fail')]
        [string]$PingType,

        [string]$Body
    )

    if (-not $script:HealthchecksUrl) { return }

    if ($WhatIfPreference) {
        Write-Log "Healthchecks ping skipped (WhatIf): $PingType"
        return
    }

    $suffix = switch ($PingType) {
        'Start' { '/start' }
        'Fail'  { '/fail' }
        default { '' }
    }

    try {
        $requestParams = @{
            Uri               = "$script:HealthchecksUrl$suffix"
            Method            = 'Post'
            TimeoutSec        = $script:Config.Healthchecks.TimeoutSeconds
            MaximumRetryCount = 2
            RetryIntervalSec  = 5
            ErrorAction       = 'Stop'
        }

        if ($Body) {
            $requestParams['Body'] = $Body
            $requestParams['ContentType'] = 'text/plain; charset=utf-8'
        }

        Invoke-RestMethod @requestParams | Out-Null
        Write-Log "Healthchecks ping sent: $PingType"
    }
    catch {
        $reason = $_.Exception.Message
        Write-Log "Healthchecks ping ($PingType) failed: $reason" -Level WARN
    }
}

# ============================================================================
# SELF-CHECKS
# ============================================================================

function Test-ScheduledTaskHealth {
    <#
    .SYNOPSIS
        Warns when a scheduled task pointing at this script launches a hard-coded
        interpreter path instead of resolving pwsh.exe from PATH.

        A baked-in path stops working the moment PowerShell is reinstalled to a
        different directory, and it fails *before* anything in this script runs — no
        log, no toast, no trace. A run that currently works is the only opportunity to
        warn about it in advance, which is why the check lives here rather than in the
        installer.

        Read-only and best-effort: it never throws and never modifies the task.
    #>
    [CmdletBinding()]
    param()

    try {
        $thisScript = $PSCommandPath
        if (-not $thisScript) { return }

        $tasks = @(Get-ScheduledTask -ErrorAction Stop | Where-Object {
            $argString = @($_.Actions.Arguments) -join ' '
            $argString -and $argString.Contains($thisScript, [StringComparison]::OrdinalIgnoreCase)
        })

        if ($tasks.Count -eq 0) {
            Write-Log "No scheduled task references this script — running ad hoc"
            return
        }

        foreach ($task in $tasks) {
            $taskName = $task.TaskName
            foreach ($action in @($task.Actions)) {
                $exe = $action.Execute
                if (-not $exe) { continue }

                # A bare name has no directory separator, so Task Scheduler re-resolves
                # it from PATH on every run — which is what we want
                if ($exe -notmatch '[\\/]') {
                    Write-Log "Scheduled task '$taskName' launches '$exe' from PATH — survives a PowerShell reinstall"
                    continue
                }

                if (Test-Path $exe) {
                    Write-Log "Scheduled task '$taskName' has a hard-coded interpreter path ($exe). It works today but stops working if PowerShell is reinstalled elsewhere — re-run Install-ModuleMaintenance.ps1 as Administrator to switch it to PATH resolution" -Level WARN
                }
                else {
                    Write-Log "Scheduled task '$taskName' points at an interpreter that no longer exists ($exe) — re-run Install-ModuleMaintenance.ps1 as Administrator" -Level WARN
                }
            }
        }
    }
    catch {
        # Enumerating tasks can be denied or slow depending on machine policy; this is a
        # diagnostic nicety, never a reason to disrupt maintenance
        Write-Verbose "Could not inspect scheduled task registration: $_"
    }
}

# ============================================================================
# MODULE OPERATIONS
# ============================================================================

function ConvertTo-NormalizedVersion {
    <#
    .SYNOPSIS
        Pads a version to four parts so 6.1907.1 and 6.1907.1.0 compare as equal.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][version]$Version)

    return [version]::new(
        $Version.Major, $Version.Minor,
        [Math]::Max($Version.Build, 0), [Math]::Max($Version.Revision, 0))
}

function Get-ModuleVersionKey {
    <#
    .SYNOPSIS
        Builds a comparable string for an installed version, including its prerelease label.
        PSResourceGet exposes the numeric version and prerelease label as separate properties.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][version]$Version,
        [string]$Prerelease
    )

    $key = (ConvertTo-NormalizedVersion $Version).ToString()
    if ($Prerelease) { $key += "-$Prerelease" }
    return $key
}

function ConvertFrom-PinnedVersionString {
    <#
    .SYNOPSIS
        Parses a pinned version such as '2.19.0' or '2.0.0-beta1'.
        Returns $null if the string is not a valid version.
    #>
    [CmdletBinding()]
    param([string]$Value)

    if ([string]::IsNullOrWhiteSpace($Value)) { return $null }

    $numeric, $prerelease = $Value.Trim() -split '-', 2
    $parsed = $null
    if (-not [version]::TryParse($numeric, [ref]$parsed)) { return $null }

    return [PSCustomObject]@{
        Version    = $parsed
        Prerelease = $prerelease
        Requested  = $Value.Trim()
        Key        = Get-ModuleVersionKey -Version $parsed -Prerelease $prerelease
    }
}

function Get-PinnedVersion {
    <#
    .SYNOPSIS
        Returns the parsed pin for a module name, or $null if it is not pinned.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Name)

    if ($script:Config.PinnedModules.ContainsKey($Name)) {
        return $script:Config.PinnedModules[$Name]
    }
    return $null
}

function Test-IsPinnedVersion {
    <#
    .SYNOPSIS
        Tests whether an installed PSResource matches the given pin.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Resource,
        [Parameter(Mandatory)]$Pin
    )

    return (Get-ModuleVersionKey -Version $Resource.Version -Prerelease $Resource.Prerelease) -eq $Pin.Key
}

function ConvertFrom-KeepVersionSelector {
    <#
    .SYNOPSIS
        Parses a KeepVersions selector such as '5', '5.7' or '5.7.1'.
        Returns $null if the value is not a version prefix.

    .NOTES
        A selector names a version line, so it is a prefix of one to four numbers, not a
        full version. [version] is no help here: it cannot parse '5', and it accepts
        things a selector must not be, such as ' 5.7 ' with the spaces or '05.7'.

        Only a string is accepted. JSON turns an unquoted 5.10 into the number 5.1, which
        would silently select the wrong line.
    #>
    [CmdletBinding()]
    param($Value)

    $selector = $null

    if ($Value -is [string]) {
        $text = $Value.Trim()
        if ($text -match '^[0-9]+(\.[0-9]+){0,3}$') {
            $parts = @()
            $valid = $true
            foreach ($piece in $text.Split('.')) {
                $number = 0
                if ([int]::TryParse($piece, [ref]$number)) {
                    $parts += $number
                }
                else {
                    $valid = $false
                }
            }

            if ($valid) {
                $selector = [PSCustomObject]@{
                    Requested = $text
                    Parts     = $parts
                    Key       = ($parts -join '.')
                }
            }
        }
    }

    return $selector
}

function Test-KeepVersionMatch {
    <#
    .SYNOPSIS
        Tests whether a version belongs to the line a selector names. The comparison is
        number by number, so '1' does not match 12.4.0 and '5.7' does not match 5.70.1.
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param(
        [Parameter(Mandatory)][version]$Version,
        [Parameter(Mandatory)]$Selector
    )

    $normalized = ConvertTo-NormalizedVersion $Version
    $actual = @($normalized.Major, $normalized.Minor, $normalized.Build, $normalized.Revision)

    $isMatch = $true
    for ($i = 0; $i -lt $Selector.Parts.Count; $i++) {
        if ($actual[$i] -ne $Selector.Parts[$i]) {
            $isMatch = $false
        }
    }
    return $isMatch
}

function Get-KeepVersionSelectors {
    <#
    .SYNOPSIS
        Returns the parsed KeepVersions selectors for a module name, or an empty list.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Name)

    $selectors = @()
    if ($script:Config.KeepVersions.ContainsKey($Name)) {
        $selectors = @($script:Config.KeepVersions[$Name])
    }
    return $selectors
}

function Get-KeepVersionRange {
    <#
    .SYNOPSIS
        Builds the NuGet version range that covers every given selector, for one gallery
        lookup per module. '5' gives '[5.0.0.0, 6.0.0.0)'.

    .NOTES
        The range is what Find-PSResource needs to return every version in a line. A
        wildcard must not be used for this: '5.*' was seen returning versions from the
        line above.

        Several selectors share one range that spans them all. The answer is then
        filtered per selector with Test-KeepVersionMatch, so what lies between two lines
        is simply ignored.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param([Parameter(Mandatory)][array]$Selectors)

    $lowest = $null
    $highest = $null

    foreach ($selector in $Selectors) {
        # Copies, so the selector itself is not changed
        $from = @($selector.Parts)
        $to = @($selector.Parts)
        $to[$to.Count - 1] = $to[$to.Count - 1] + 1

        while ($from.Count -lt 4) { $from += 0 }
        while ($to.Count -lt 4) { $to += 0 }

        $lower = [version]::new($from[0], $from[1], $from[2], $from[3])
        $upper = [version]::new($to[0], $to[1], $to[2], $to[3])

        if (($null -eq $lowest) -or ($lower -lt $lowest)) { $lowest = $lower }
        if (($null -eq $highest) -or ($upper -gt $highest)) { $highest = $upper }
    }

    return "[$lowest, $highest)"
}

function Get-ModulePrunePlan {
    <#
    .SYNOPSIS
        Decides which installed versions of one module stay and which go.

    .DESCRIPTION
        What stays is the newest version, or the pinned one if the module is pinned, plus
        the newest installed version of every line named in KeepVersions.

        Nothing is removed when the module is pinned and the pinned version is not
        installed. Removing the rest could leave the module with no version at all.

    .NOTES
        Pure on purpose: no logging and no script state, so the rules can be tested
        without standing in for PSResourceGet.

        A version that stays is never removed, even if it is listed twice. The same
        version can be installed in two scopes, and the older logic treated the second
        copy as an old version.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        [array]$Resources,

        $Pin,

        [AllowNull()]
        [AllowEmptyCollection()]
        [array]$Selectors
    )

    $sorted = @($Resources | Sort-Object -Property @{ Expression = { ConvertTo-NormalizedVersion $_.Version } } -Descending)
    $selectorList = @($Selectors | Where-Object { $_ })

    $keepKeys = @()
    $pinMissing = $false
    $base = $null

    if ($Pin) {
        $base = $sorted | Where-Object { Test-IsPinnedVersion -Resource $_ -Pin $Pin } | Select-Object -First 1
        if (-not $base) {
            $pinMissing = $true
        }
    }
    else {
        $base = $sorted | Select-Object -First 1
    }

    if ($base) {
        $keepKeys += Get-ModuleVersionKey -Version $base.Version -Prerelease $base.Prerelease
    }

    $matched = @()
    $unmatched = @()
    foreach ($selector in $selectorList) {
        $hit = $sorted | Where-Object { Test-KeepVersionMatch -Version $_.Version -Selector $selector } |
            Select-Object -First 1

        if (-not $hit) {
            $unmatched += $selector.Requested
            continue
        }

        $keepKeys += Get-ModuleVersionKey -Version $hit.Version -Prerelease $hit.Prerelease

        $display = $hit.Version.ToString()
        if ($hit.Prerelease) {
            $display += "-$($hit.Prerelease)"
        }
        $matched += [PSCustomObject]@{
            Selector = $selector.Requested
            Version  = $display
            Resource = $hit
        }
    }

    $keep = $sorted
    $remove = @()
    if (-not $pinMissing) {
        $keep = @($sorted | Where-Object {
            (Get-ModuleVersionKey -Version $_.Version -Prerelease $_.Prerelease) -in $keepKeys
        })
        $remove = @($sorted | Where-Object {
            (Get-ModuleVersionKey -Version $_.Version -Prerelease $_.Prerelease) -notin $keepKeys
        })
    }

    return [PSCustomObject]@{
        Keep       = @($keep)
        Remove     = @($remove)
        Matched    = @($matched)
        Unmatched  = @($unmatched)
        PinMissing = $pinMissing
    }
}

function Get-KeptLinePlan {
    <#
    .SYNOPSIS
        Decides, for each kept line of one module, whether it needs a gallery lookup, can
        be compared straight away, or is left alone.

    .DESCRIPTION
        Returns one object per selector with an Action:

        Skip     nothing to do. Reason says why: NotInstalled, Exact, PinInLine or
                 MainUpdate.
        Compare  the newest release on the gallery lies in this line, so it is the
                 target and no further lookup is needed.
        Lookup   the gallery has to be asked for the versions in this line.

    .NOTES
        "Covered by the normal update" has to be judged from the gallery, not from what
        is installed. With only an old line installed, that line looks like the newest
        one, and skipping it would leave it a release behind.

        The normal update drops pinned modules, so it never covers a line of one.

        Pure on purpose, like Get-ModulePrunePlan.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        [array]$Resources,

        [AllowNull()]
        [AllowEmptyCollection()]
        [array]$Selectors,

        $Pin,

        # Newest release of the module on the gallery, from the main lookup
        [AllowNull()]
        [version]$GalleryNewest
    )

    $sorted = @($Resources | Sort-Object -Property @{ Expression = { ConvertTo-NormalizedVersion $_.Version } } -Descending)
    $newestInstalled = $sorted | Select-Object -First 1

    $plan = @()
    foreach ($selector in @($Selectors | Where-Object { $_ })) {
        $installed = $sorted | Where-Object { Test-KeepVersionMatch -Version $_.Version -Selector $selector } |
            Select-Object -First 1

        $galleryInLine = $false
        if ($GalleryNewest) {
            $galleryInLine = Test-KeepVersionMatch -Version $GalleryNewest -Selector $selector
        }

        $newestInLine = $false
        if ($newestInstalled) {
            $newestInLine = Test-KeepVersionMatch -Version $newestInstalled.Version -Selector $selector
        }

        $pinInLine = $false
        if ($Pin) {
            $pinInLine = Test-KeepVersionMatch -Version $Pin.Version -Selector $selector
        }

        $action = 'Lookup'
        $reason = $null
        $target = $null

        if (-not $installed) {
            $action = 'Skip'
            $reason = 'NotInstalled'
        }
        elseif ($selector.Parts.Count -ge 4) {
            $action = 'Skip'
            $reason = 'Exact'
        }
        elseif ($pinInLine) {
            $action = 'Skip'
            $reason = 'PinInLine'
        }
        elseif ($galleryInLine -and $newestInLine -and (-not $Pin)) {
            $action = 'Skip'
            $reason = 'MainUpdate'
        }
        elseif ($galleryInLine) {
            $action = 'Compare'
            $target = $GalleryNewest
        }

        $plan += [PSCustomObject]@{
            Selector  = $selector
            Installed = $installed
            Action    = $action
            Reason    = $reason
            Target    = $target
        }
    }

    return $plan
}

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

function Test-ManagedModuleFolder {
    <#
    .SYNOPSIS
        Tests whether a module folder was installed by PSResourceGet (or PowerShellGet), as
        opposed to being put there by another program's installer.

    .NOTES
        The marker is PSGetModuleInfo.xml, which both write into every version folder they
        install. Without it Get-PSResource does not list the module, so this script can
        neither update nor prune it, and has no business reporting or moving it.

        Microsoft.PowerToys.Configure is the case that prompted this. PowerToys installs it
        to Documents\PowerShell\Modules by itself and re-creates it on every update, so the
        "run Invoke-OneDriveMigration.ps1" warning came back week after week and could
        never be cleared by doing what it said.
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param(
        [Parameter(Mandatory)]
        [string]$ModuleFolder
    )

    # Depth 1 reaches the version folders, and also covers a module with no version
    # folder. -Force because the marker is a hidden file
    $marker = Get-ChildItem -LiteralPath $ModuleFolder -Filter 'PSGetModuleInfo.xml' -File `
        -Recurse -Depth 1 -Force -ErrorAction SilentlyContinue | Select-Object -First 1

    return [bool]$marker
}

function Resolve-ModuleVersionFolder {
    <#
    .SYNOPSIS
        Returns the folder a module version actually occupies under a modules root, or
        $null if it is not there.

    .NOTES
        Looks on disk instead of trusting InstalledLocation, which is written at install
        time and never corrected. A module migrated out of OneDrive still reports its old
        OneDrive path there while living in Program Files.

        Versions are compared, not folder names, because the folder can be 6.1907.1.0
        where PSResourceGet reports 6.1907.1.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory)]
        [string]$ModulesRoot,

        [Parameter(Mandatory)]
        [string]$Name,

        [Parameter(Mandatory)]
        [version]$Version
    )

    $folder = $null
    $moduleFolder = Join-Path $ModulesRoot $Name

    if (Test-Path -LiteralPath $moduleFolder) {
        $wanted = ConvertTo-NormalizedVersion $Version

        $match = Get-ChildItem -LiteralPath $moduleFolder -Directory -Force -ErrorAction SilentlyContinue |
            Where-Object {
                $parsed = $null
                ([version]::TryParse($_.Name, [ref]$parsed)) -and
                    ((ConvertTo-NormalizedVersion $parsed) -eq $wanted)
            } | Select-Object -First 1

        if ($match) {
            $folder = $match.FullName
        }
    }

    return $folder
}

function Remove-LockedModuleFolder {
    <#
    .SYNOPSIS
        Force-removes a module folder that OneDrive has locked or converted to cloud placeholders.
        Escalation: normal Remove-Item → strip cloud attributes + rd /s /q → MoveFileEx reboot deletion.
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
            # Directories, reparse points, or already-deleted files — ignore
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

function Invoke-ModuleUpdate {
    <#
    .SYNOPSIS
        Runs Update-PSResource in an isolated runspace with a timeout — or Install-PSResource
        when -Version is given, which is how pinned modules are put back on their version.
        Prevents a single slow/hung module from consuming all scheduled task time.

    .NOTES
        The deadline is enforced with a Stopwatch and a check of the pipeline's own
        InvocationStateInfo, NOT by trusting AsyncWaitHandle.WaitOne's return value.

        Two separate faults made the original WaitOne + Stop() approach ineffective:

        1. Stop() BLOCKS until non-cooperative native work finishes. PSResourceGet's
           network and file I/O do not cooperate, so a timed-out module still consumed
           its full runtime — measured at 20s for a 5s timeout in a controlled repro.
        2. WaitOne was observed returning $true while the pipeline was still running,
           after which EndInvoke silently absorbed the remaining time and the module was
           reported as a SUCCESS. In one run a large meta-module took more than twice the
           600s timeout and logged SUCCESS, while every other module took a few seconds.

        Hence: poll in short slices against a Stopwatch, treat only a terminal pipeline
        state as completion, and request the stop asynchronously so a stubborn native
        call cannot re-block the main thread.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Name,

        [string]$Scope,

        [bool]$TrustRepository = $true,

        [int]$TimeoutSeconds = 600,

        [string]$Version,

        [switch]$Prerelease
    )

    # How long to give an asynchronous stop before abandoning the runspace. Deliberately
    # short: a cooperative pipeline stops in well under a second, and one blocked in a
    # native call will not stop no matter how long we wait. So this checks whether the
    # stop took effect promptly — it is not a wait for the work to finish.
    $stopGraceSeconds = 5

    $ps = [powershell]::Create()
    $abandoned = $false

    try {
        $ps.AddScript({
            param($n, $s, $t, $v, $pre)

            # The isolated runspace gets a fresh session state, so the parent script's
            # $ProgressPreference = 'SilentlyContinue' does NOT carry over. Without this
            # the progress records pile up in $ps.Streams.Progress for the whole update
            # (measured at 20,000 records in a controlled repro) for no benefit, since
            # nothing renders them in a scheduled task.
            $ProgressPreference = 'SilentlyContinue'

            $params = @{
                Name            = $n
                AcceptLicense   = $true
                TrustRepository = $t
                ErrorAction     = 'Stop'
            }
            if ($s) { $params['Scope'] = $s }

            if ($v) {
                # PSResourceGet treats a bare version string as a required (exact) version
                # rather than a minimum, so this installs precisely $v side-by-side with
                # whatever else is present. See about_PSResourceGet, "NuGet version ranges".
                $params['Version'] = $v
                if ($pre) { $params['Prerelease'] = $true }
                Install-PSResource @params
            }
            else {
                Update-PSResource @params
            }
        }).AddArgument($Name).AddArgument($Scope).AddArgument($TrustRepository).
            AddArgument($Version).AddArgument($Prerelease.IsPresent) | Out-Null

        $handle = $ps.BeginInvoke()

        # A terminal state is the only trustworthy signal that the pipeline is done.
        # WaitOne's return value is not: see the timeout notes in the comment block above.
        $terminalStates = @('Completed', 'Failed', 'Stopped')
        $timer = [System.Diagnostics.Stopwatch]::StartNew()
        $deadline = [TimeSpan]::FromSeconds($TimeoutSeconds)

        while (($timer.Elapsed -lt $deadline) -and
               ($ps.InvocationStateInfo.State -notin $terminalStates)) {
            # Short slices rather than one long wait, so a spurious signal just costs
            # another lap instead of dropping us out of the loop early
            $handle.AsyncWaitHandle.WaitOne(500) | Out-Null
        }

        if ($ps.InvocationStateInfo.State -notin $terminalStates) {
            $elapsed = [math]::Round($timer.Elapsed.TotalSeconds)

            # BeginStop, not Stop: the synchronous Stop() blocks until the underlying
            # native work finishes, which is exactly what defeats the timeout
            $stopHandle = $ps.BeginStop($null, $null)
            $stoppedInTime = $stopHandle.AsyncWaitHandle.WaitOne(
                [TimeSpan]::FromSeconds($stopGraceSeconds))

            # If it ignored the stop request, leave the runspace to the garbage collector.
            # Disposing one whose pipeline is still winding down can block just as badly.
            $abandoned = -not $stoppedInTime

            throw [System.TimeoutException]::new(
                "Operation timed out after ${elapsed}s (limit ${TimeoutSeconds}s) for module '$Name'")
        }

        $ps.EndInvoke($handle) | Out-Null

        if ($ps.HadErrors) {
            throw $ps.Streams.Error[0].Exception
        }
    }
    finally {
        if (-not $abandoned) {
            $ps.Dispose()
        }
    }
}

function Get-TransientNetworkFault {
    <#
    .SYNOPSIS
        Returns the part of an error message that marks it as a passing network fault
        (dropped TLS handshake, DNS not ready yet, gateway 5xx), or $null when the failure
        is about the module itself and trying again would not help.

    .NOTES
        Matches on message text because there is nothing else left to match on. A failure
        inside the isolated runspace reaches the caller as MethodInvocationException ->
        ActionPreferenceStopException with no InnerException below that: PSResourceGet has
        already flattened the HttpRequestException/SocketException chain into one string.
        Measured against a dead proxy and an unresolvable host, not assumed.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory)]
        [AllowEmptyString()]
        [string]$Message
    )

    $patterns = @(
        'SSL connection could not be established'
        'Unable to (read|write) data (from|to) the transport connection'
        'forcibly closed by the remote host'
        'connection was (closed|aborted)'
        'response ended prematurely'
        'No such host is known'
        'remote name could not be resolved'
        'connected party did not properly respond'
        'actively refused'
        'unreachable (network|host)'
        'HttpClient\.Timeout'
        'error occurred while sending the request'
        'status code does not indicate success: (429|502|503|504)'
    )

    $fault = $null
    if ($Message -match ($patterns -join '|')) {
        $fault = $Matches[0]
    }
    return $fault
}

function Invoke-ModuleUpdateWithRetry {
    <#
    .SYNOPSIS
        Invoke-ModuleUpdate, tried again when the failure is a passing network fault.
        Takes the same parameters and throws the same errors, so callers keep their
        existing catch blocks.

    .NOTES
        In one run every pending update failed within seconds with "The SSL connection
        could not be established". The task had fired right after the machine woke up,
        before the network had settled. The same call succeeded later that day, but with
        no retry each module had to wait a week for the next run.

        Not retried:
        - Timeouts. The module already used its whole ModuleUpdateTimeoutSeconds, and
          trying again would multiply the runtime the timeout exists to bound.
        - Anything that is not network-shaped, including the locked-folder failures that
          Update-AllModules recovers from by itself.

        Each attempt gets a fresh runspace from Invoke-ModuleUpdate, so nothing carries
        over from the one that failed.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Name,

        [string]$Scope,

        [bool]$TrustRepository = $true,

        [int]$TimeoutSeconds = 600,

        [string]$Version,

        [switch]$Prerelease,

        # Seconds to wait before each retry. One entry per retry, so two entries means
        # three attempts in total
        [int[]]$RetryDelaySeconds = @(5, 15)
    )

    $updateParams = [hashtable]::new($PSBoundParameters)
    $updateParams.Remove('RetryDelaySeconds')

    $maxAttempts = $RetryDelaySeconds.Count + 1

    for ($attempt = 1; $attempt -le $maxAttempts; $attempt++) {
        $failure = $null
        try {
            Invoke-ModuleUpdate @updateParams
        }
        catch {
            $failure = $_
        }

        if (-not $failure) {
            return
        }

        $fault = $null
        if ($failure.Exception -isnot [System.TimeoutException]) {
            $fault = Get-TransientNetworkFault -Message $failure.Exception.Message
        }

        if ((-not $fault) -or ($attempt -ge $maxAttempts)) {
            throw $failure
        }

        $delay = $RetryDelaySeconds[$attempt - 1]
        Write-Log "Network fault on $Name (attempt $attempt of $maxAttempts): $fault. Retrying in ${delay}s" -Level WARN
        Start-Sleep -Seconds $delay
    }
}

function Set-PinnedModuleVersions {
    <#
    .SYNOPSIS
        Ensures every pinned module has its pinned version installed. Other versions are
        left alone here — Remove-OldModuleVersions prunes everything except the pin.
        Pinning never installs a module that isn't already present.
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        [array]$InstalledResources,

        [string]$Scope
    )

    $pinnedNames = @($script:Config.PinnedModules.Keys)
    if ($pinnedNames.Count -eq 0) { return }

    Write-Log "Enforcing $($pinnedNames.Count) pinned module version(s)..."
    $timeout = $script:Config.ModuleUpdateTimeoutSeconds

    foreach ($name in $pinnedNames) {
        $pin = $script:Config.PinnedModules[$name]
        $versions = @($InstalledResources | Where-Object { $_.Name -eq $name })

        if ($versions.Count -eq 0) {
            Write-Log "Pinned module $name is not installed — nothing to enforce (pinning does not install new modules)"
            continue
        }

        if ($versions | Where-Object { Test-IsPinnedVersion -Resource $_ -Pin $pin }) {
            Write-Log "$name is at its pinned version $($pin.Requested)"
            $script:Summary.PinsSatisfied++
            continue
        }

        $installedList = ($versions | Sort-Object Version -Descending | ForEach-Object {
            Get-ModuleVersionKey -Version $_.Version -Prerelease $_.Prerelease
        }) -join ', '

        try {
            if ($PSCmdlet.ShouldProcess("$name -> $($pin.Requested)", "Install pinned version")) {
                Write-Log "Installing pinned version of ${name}: $($pin.Requested) (installed: $installedList)"

                $pinTimer = [System.Diagnostics.Stopwatch]::StartNew()

                Invoke-ModuleUpdateWithRetry -Name $name -Scope $Scope `
                    -TrustRepository $script:Config.TrustPSGallery -TimeoutSeconds $timeout `
                    -Version $pin.Requested -Prerelease:([bool]$pin.Prerelease)

                $script:Summary.PinsEnforced++
                Write-Log "Installed pinned version: $name $($pin.Requested) (took $([math]::Round($pinTimer.Elapsed.TotalSeconds))s)" -Level SUCCESS
            }
        }
        catch [System.TimeoutException] {
            Write-Log "Timed out installing pinned $name $($pin.Requested) after ${timeout}s — skipping" -Level ERROR
            $script:Summary.PinsFailed += @{
                Module  = $name
                Version = $pin.Requested
                Error   = $_.Exception.Message
            }
        }
        catch {
            Write-Log "Failed to install pinned $name $($pin.Requested): $($_.Exception.Message)" -Level ERROR
            $script:Summary.PinsFailed += @{
                Module  = $name
                Version = $pin.Requested
                Error   = $_.Exception.Message
            }
        }
    }

    Write-Log "Pin enforcement complete. Already pinned: $($script:Summary.PinsSatisfied), Installed: $($script:Summary.PinsEnforced), Unsuccessful: $($script:Summary.PinsFailed.Count)"
}

function Find-GalleryModules {
    <#
    .SYNOPSIS
        The bulk PSGallery lookup for Update-AllModules, able to tell "no update
        available" apart from "could not ask".

    .NOTES
        Find-PSResource reports a network fault as a NON-terminating error per module, so
        under -ErrorAction SilentlyContinue an unreachable gallery just returns nothing.
        The run then logged "All modules are up to date" and sent a Success ping, which
        reset the dead-man timer on a run that had checked nothing. Measured behind a dead
        proxy: about two seconds per installed module, several minutes in all, without a
        word in the log.

        SilentlyContinue has to stay, because "not on the gallery" is an everyday answer
        for a module installed from somewhere else. The two cases are told apart by error
        id, which survives here even though it does not survive the isolated runspace:
            PackageNotFound    answered, the module is simply not on PSGallery
            anything else      not answered (HttpRequestCallFailure for a network fault)

        One module is looked up before the bulk query, as a probe. If the gallery gives no
        answer the bulk query is not sent, which keeps a dead network from costing one
        failed request per installed module. The probe is classified like any other
        lookup, so "not found" counts as an answer: it proves the gallery is reachable,
        which is all the probe is for.

        Unanswered lookups are retried on the same schedule as updates, and only the
        modules still without an answer are asked about again.

        With -Version the question is "which versions exist in this range" and the answer
        is one resource per version. A range with nothing in it returns nothing and raises
        NO error, not even PackageNotFound, so an empty answer is an answer. Only an error
        that is not PackageNotFound means the gallery did not reply.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string[]]$Name,

        # A NuGet version range such as '[5.0.0.0, 6.0.0.0)'. Never a wildcard: '5.*' was
        # seen returning versions from the line above
        [string]$Version,

        # Seconds to wait before each retry. One entry per retry, so two entries means
        # three attempts in total
        [int[]]$RetryDelaySeconds = @(5, 15)
    )

    # Only used to see whether the gallery answers. Nothing depends on this module staying
    # published, and it does not have to be installed from the gallery: on PowerShell 7.4
    # and later it is the copy bundled under $PSHOME, which this script does not manage
    $probeName = 'Microsoft.PowerShell.PSResourceGet'

    $found = @()
    $pending = @($Name)
    $detail = $null
    $maxAttempts = $RetryDelaySeconds.Count + 1

    for ($attempt = 1; $attempt -le $maxAttempts; $attempt++) {
        $unanswered = @()
        $probeErrors = @()
        $lookupErrors = @()

        try {
            Find-PSResource -Name $probeName -Repository PSGallery `
                -ErrorAction SilentlyContinue -ErrorVariable probeErrors | Out-Null
            $unanswered = @($probeErrors | Where-Object { $_.FullyQualifiedErrorId -notlike 'PackageNotFound,*' })

            if ($unanswered.Count -eq 0) {
                if ($Version) {
                    $batch = @(Find-PSResource -Name $pending -Version $Version -Repository PSGallery `
                        -ErrorAction SilentlyContinue -ErrorVariable lookupErrors)
                }
                else {
                    $batch = @(Find-PSResource -Name $pending -Repository PSGallery `
                        -ErrorAction SilentlyContinue -ErrorVariable lookupErrors)
                }
                $found += $batch

                $notFoundErrors = @($lookupErrors | Where-Object { $_.FullyQualifiedErrorId -like 'PackageNotFound,*' })
                $unanswered = @($lookupErrors | Where-Object { $_.FullyQualifiedErrorId -notlike 'PackageNotFound,*' })

                # The name is only available inside the message. If a future PSResourceGet
                # words it differently, these modules stay in $pending and are counted as
                # unchecked, which errs on the side of reporting too much
                $notOnGallery = @($notFoundErrors | ForEach-Object {
                    if ($_.Exception.Message -match "'([^']+)'") { $Matches[1] }
                })
                $pending = @($pending | Where-Object { ($_ -notin $batch.Name) -and ($_ -notin $notOnGallery) })
            }
        }
        catch {
            # The probe or the bulk query threw outright. Either way nothing still
            # pending got an answer this attempt
            $unanswered = @($_)
        }

        if ($unanswered.Count -eq 0) {
            $pending = @()
            $detail = $null
            break
        }

        $detail = $unanswered[0].Exception.Message

        if ($attempt -lt $maxAttempts) {
            $delay = $RetryDelaySeconds[$attempt - 1]
            $pendingCount = $pending.Count
            $reason = Get-GalleryFaultText -Message $detail
            Write-Log "PSGallery gave no answer for $pendingCount module(s) (attempt $attempt of $maxAttempts): $reason. Retrying in ${delay}s" -Level WARN
            Start-Sleep -Seconds $delay
        }
    }

    $fault = $null
    if ($detail) {
        $fault = Get-GalleryFaultText -Message $detail
    }

    return [PSCustomObject]@{
        Resources = $found
        Unchecked = @($pending)
        Fault     = $fault
        Detail    = $detail
    }
}

function Get-GalleryFaultText {
    <#
    .SYNOPSIS
        A lookup error cut down to something that fits a log line, a toast or a ping body.
        The full text runs to several hundred characters because it carries the request URL.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory)]
        [string]$Message
    )

    $text = Get-TransientNetworkFault -Message $Message
    if (-not $text) {
        $text = $Message
        if ($text.Length -gt 120) {
            $text = $text.Substring(0, 120) + '...'
        }
    }
    return $text
}

function Update-KeptVersionLines {
    <#
    .SYNOPSIS
        Updates every kept version line within itself. A newer release inside a line is
        installed next to what is already there, and pruning then removes the older one.

    .NOTES
        A line is only maintained once a version of it is installed. This never brings a
        line onto a machine that does not have it.

        Has to run before the normal update loop, not after it. Update-AllModules returns
        early when no module needs updating, which is what happens most weeks.

        The gallery is asked once per module, with a range spanning its lines. Once it
        fails to answer, the remaining lines are recorded as not checked instead of being
        asked about one by one.

        An exact version is installed into a folder of its own, so the locked-folder
        recovery that the normal update needs has nothing to do here.
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        [array]$InstalledResources,

        # What the main lookup returned: the newest release of each module
        [AllowEmptyCollection()]
        [array]$GalleryResources = @(),

        # Modules the main lookup got no answer for. Those are already counted
        [AllowEmptyCollection()]
        [string[]]$Unchecked = @(),

        [string]$Scope
    )

    $moduleNames = @($script:Config.KeepVersions.Keys)
    if ($moduleNames.Count -eq 0) {
        return
    }

    Write-Log "Checking kept version lines of $($moduleNames.Count) module(s) for updates..."

    $timeout = $script:Config.ModuleUpdateTimeoutSeconds
    $updatedCount = 0
    $unsuccessfulCount = 0
    $uncheckedCount = 0
    $galleryAnswers = $true
    $galleryFault = $null

    foreach ($configName in $moduleNames) {
        $resources = @($InstalledResources | Where-Object { $_.Name -eq $configName })
        if ($resources.Count -eq 0) {
            Write-Log "Kept lines of $configName skipped: the module is not installed"
            continue
        }

        # From here on the name as installed, which may be cased differently
        $name = $resources[0].Name

        if ($name -in $Unchecked) {
            Write-Log "Kept lines of $name not checked: PSGallery gave no answer for this module" -Level WARN
            continue
        }

        $galleryNewest = ($GalleryResources | Where-Object { $_.Name -eq $name } |
            Sort-Object Version -Descending | Select-Object -First 1).Version
        if (-not $galleryNewest) {
            Write-Log "Kept lines of $name skipped: the module is not on PSGallery"
            continue
        }

        $selectors = @(Get-KeepVersionSelectors -Name $configName)
        $pin = Get-PinnedVersion $configName
        $plan = @(Get-KeptLinePlan -Resources $resources -Selectors $selectors -Pin $pin -GalleryNewest $galleryNewest)

        foreach ($entry in @($plan | Where-Object { $_.Action -eq 'Skip' })) {
            $line = $entry.Selector.Requested
            switch ($entry.Reason) {
                'NotInstalled' { Write-Log "Kept line $line of $name has no installed version - nothing to update" }
                'Exact'        { Write-Log "Kept line $line of $name names one exact version - nothing to update" }
                'PinInLine'    { Write-Log "Kept line $line of $name holds the pinned version - the pin decides, no line update" }
                'MainUpdate'   { Write-Log "Kept line $line of $name holds the newest release - covered by the normal update" }
            }
        }

        # One lookup for all the lines of this module that need one
        $candidates = @()
        $lookupAnswered = $true
        $toLookUp = @($plan | Where-Object { $_.Action -eq 'Lookup' })

        if ($toLookUp.Count -gt 0) {
            $lineList = ($toLookUp | ForEach-Object { $_.Selector.Requested }) -join ', '

            if ($galleryAnswers) {
                $range = Get-KeepVersionRange -Selectors @($toLookUp | ForEach-Object { $_.Selector })
                $lineLookup = Find-GalleryModules -Name $name -Version $range

                if ($lineLookup.Unchecked.Count -gt 0) {
                    $galleryAnswers = $false
                    $galleryFault = $lineLookup.Fault
                    $lookupAnswered = $false
                    Write-Log "Could not check kept line(s) $lineList of $name for updates: $($lineLookup.Detail)" -Level ERROR
                }
                else {
                    # A line update never installs a prerelease
                    $candidates = @($lineLookup.Resources | Where-Object { -not $_.Prerelease })
                }
            }
            else {
                $lookupAnswered = $false
                Write-Log "Kept line(s) $lineList of $name not checked: PSGallery stopped answering earlier in this run" -Level WARN
            }

            if (-not $lookupAnswered) {
                foreach ($entry in $toLookUp) {
                    $script:Summary.KeepLinesUnchecked += @{
                        Module = $name
                        Line   = $entry.Selector.Requested
                        Fault  = $galleryFault
                    }
                    $uncheckedCount++
                }
            }
        }

        # Work out the target of each line and whether it is ahead of what is installed
        $pending = @()
        foreach ($entry in @($plan | Where-Object { $_.Action -in 'Compare', 'Lookup' })) {
            if (($entry.Action -eq 'Lookup') -and (-not $lookupAnswered)) {
                continue
            }

            $line = $entry.Selector.Requested
            $target = $entry.Target

            if ($entry.Action -eq 'Lookup') {
                $selector = $entry.Selector
                $target = ($candidates |
                    Where-Object { Test-KeepVersionMatch -Version $_.Version -Selector $selector } |
                    Sort-Object -Property @{ Expression = { ConvertTo-NormalizedVersion $_.Version } } -Descending |
                    Select-Object -First 1).Version

                if (-not $target) {
                    Write-Log "Kept line $line of $name has no release on PSGallery - nothing to update"
                    continue
                }
            }

            $installedText = $entry.Installed.Version.ToString()
            if ($entry.Installed.Prerelease) {
                $installedText += "-$($entry.Installed.Prerelease)"
            }

            # By number only, as the normal update compares. A prerelease with the same
            # number as the stable release shares its folder and is not replaced
            $normalizedTarget = ConvertTo-NormalizedVersion $target
            $normalizedInstalled = ConvertTo-NormalizedVersion $entry.Installed.Version

            if ($normalizedTarget -gt $normalizedInstalled) {
                $pending += [PSCustomObject]@{
                    Line   = $line
                    From   = $installedText
                    Target = $target.ToString()
                }
            }
            elseif ($normalizedTarget -eq $normalizedInstalled) {
                Write-Log "Kept line $line of $name is up to date at $installedText"
            }
            else {
                Write-Log "Kept line $line of $name is at $installedText, ahead of $target on PSGallery - left as it is"
            }
        }

        # Two lines can share a target. It is installed once
        foreach ($group in @($pending | Group-Object Target)) {
            $targetText = $group.Name
            $lines = ($group.Group | ForEach-Object { $_.Line }) -join ', '
            $from = $group.Group[0].From

            try {
                if ($PSCmdlet.ShouldProcess("$name line $lines`: $from -> $targetText", 'Update kept version line')) {
                    Write-Log "Updating kept line $lines of ${name}: $from -> $targetText"

                    $lineTimer = [System.Diagnostics.Stopwatch]::StartNew()

                    Invoke-ModuleUpdateWithRetry -Name $name -Scope $Scope `
                        -TrustRepository $script:Config.TrustPSGallery -TimeoutSeconds $timeout `
                        -Version $targetText

                    $script:Summary.ModulesUpdated++
                    $updatedCount++
                    foreach ($item in $group.Group) {
                        $script:Summary.KeepLinesUpdated += @{
                            Module = $name
                            Line   = $item.Line
                            From   = $item.From
                            To     = $targetText
                        }
                    }
                    Write-Log "Updated kept line $lines of $name to $targetText (took $([math]::Round($lineTimer.Elapsed.TotalSeconds))s)" -Level SUCCESS
                }
            }
            catch [System.TimeoutException] {
                Write-Log "Timed out updating kept line $lines of $name to $targetText after ${timeout}s - skipping" -Level ERROR
                $script:Summary.ModulesFailed += @{
                    Module  = $name
                    Version = $targetText
                    Line    = $lines
                    Error   = $_.Exception.Message
                }
                $unsuccessfulCount++
            }
            catch {
                Write-Log "Failed to update kept line $lines of $name to ${targetText}: $($_.Exception.Message)" -Level ERROR
                $script:Summary.ModulesFailed += @{
                    Module  = $name
                    Version = $targetText
                    Line    = $lines
                    Error   = $_.Exception.Message
                }
                $unsuccessfulCount++
            }
        }
    }

    Write-Log "Kept line updates complete. Updated: $updatedCount, Unsuccessful: $unsuccessfulCount, Not checked: $uncheckedCount"
}

function Update-AllModules {
    [CmdletBinding(SupportsShouldProcess)]
    param()

    Write-Log "Starting module updates..."

    # Detect OneDrive on module path — use AllUsers scope to avoid file locking
    $currentUserModulePath = Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'PowerShell\Modules'
    $useAllUsersScope = Test-OneDrivePath $currentUserModulePath
    if ($useAllUsersScope) {
        Write-Log "OneDrive detected on module path — updates will target AllUsers scope"
    }

    # Get installed modules (newest version of each)
    # When OneDrive is detected, modules live in AllUsers scope — query that explicitly
    $getParams = if ($useAllUsersScope) { @{ Scope = 'AllUsers' } } else { @{} }
    $allResources = @(Get-PSResource @getParams |
        Where-Object { $_.Name -notin $script:Config.ExcludedModules })

    $installed = @($allResources |
        Group-Object Name |
        ForEach-Object {
            $newest = $_.Group | Sort-Object Version -Descending | Select-Object -First 1
            [PSCustomObject]@{ Name = $newest.Name; Version = $newest.Version }
        })

    $script:Summary.ModulesChecked = $installed.Count
    $script:Summary.ExcludedModules = @($script:Config.ExcludedModules)

    Write-Log "Found $($installed.Count) installed modules (excluding: $($script:Config.ExcludedModules -join ', '))"

    # Pinned modules are held at a specific version — enforce those before updating anything
    $scope = if ($useAllUsersScope) { 'AllUsers' } else { $null }
    Set-PinnedModuleVersions -InstalledResources $allResources -Scope $scope

    if ($installed.Count -eq 0) {
        Write-Log "No modules left to check for updates"
        return
    }

    Write-Log "Checking PSGallery for available updates..."

    # Query PSGallery for latest versions (bulk request). Pinned modules stay in this query
    # even though they will not be updated — it is the same single call either way, and it
    # lets the log report which release a pin is holding back.
    $lookup = Find-GalleryModules -Name $installed.Name
    $gallery = $lookup.Resources
    $uncheckedCount = $lookup.Unchecked.Count

    if ($uncheckedCount -gt 0) {
        $installedCount = $installed.Count
        $lookupDetail = $lookup.Detail

        $script:Summary.ModulesUnchecked = $uncheckedCount
        $script:Summary.GalleryFault = $lookup.Fault
        $script:Summary.ModulesChecked = $installedCount - $uncheckedCount

        if ($uncheckedCount -ge $installedCount) {
            Write-Log "Could not reach PSGallery, so none of the $installedCount installed modules were checked for updates: $lookupDetail" -Level ERROR
            return
        }

        # Name the modules, but not all of them if the list is long
        $shown = @($lookup.Unchecked | Select-Object -First 10)
        $nameList = $shown -join ', '
        if ($uncheckedCount -gt $shown.Count) {
            $nameList += ", and $($uncheckedCount - $shown.Count) more"
        }
        Write-Log "Could not check $uncheckedCount of $installedCount modules for updates ($nameList): $lookupDetail" -Level ERROR
    }

    # Report what each pin is holding back, then drop pinned modules from the update flow
    foreach ($mod in @($installed | Where-Object { Get-PinnedVersion $_.Name })) {
        $pin = Get-PinnedVersion $mod.Name
        $galleryVersion = ($gallery | Where-Object Name -eq $mod.Name |
            Sort-Object Version -Descending | Select-Object -First 1).Version
        if (-not $galleryVersion) { continue }

        # Find-PSResource returns stable releases here, so a stable build of the same number
        # is newer than a prerelease pin (2.0.0 beats 2.0.0-beta1)
        $normalizedGallery = ConvertTo-NormalizedVersion $galleryVersion
        $normalizedPin = ConvertTo-NormalizedVersion $pin.Version
        if ($normalizedGallery -gt $normalizedPin -or ($normalizedGallery -eq $normalizedPin -and $pin.Prerelease)) {
            Write-Log "$($mod.Name) is pinned to $($pin.Requested) — holding back $galleryVersion"
            $script:Summary.PinsHoldingBack += @{
                Module    = $mod.Name
                Pinned    = $pin.Requested
                Available = $galleryVersion.ToString()
            }
        }
    }

    $installed = @($installed | Where-Object { -not (Get-PinnedVersion $_.Name) })

    # Kept version lines are updated here, before the normal updates. Further down this
    # function returns as soon as no module needs updating, which is most weeks
    Update-KeptVersionLines -InstalledResources $allResources -GalleryResources @($gallery) `
        -Unchecked @($lookup.Unchecked) -Scope $scope

    # Find modules that need updates
    $needsUpdate = @()
    foreach ($mod in $installed) {
        $galleryVersion = ($gallery | Where-Object Name -eq $mod.Name | Sort-Object Version -Descending | Select-Object -First 1).Version
        # Normalize versions to 4-part form so 6.1907.1 and 6.1907.1.0 compare as equal
        if ($galleryVersion -and $mod.Version) {
            $normalizedInstalled = ConvertTo-NormalizedVersion $mod.Version
            $normalizedGallery = ConvertTo-NormalizedVersion $galleryVersion
        }
        if ($galleryVersion -and $normalizedGallery -gt $normalizedInstalled) {
            $needsUpdate += [PSCustomObject]@{
                Name = $mod.Name
                InstalledVersion = $mod.Version
                GalleryVersion = $galleryVersion
            }
        }
    }

    if ($needsUpdate.Count -eq 0) {
        if ($uncheckedCount -gt 0) {
            # Not "all": some modules never got an answer from the gallery
            $checkedCount = $script:Summary.ModulesChecked
            Write-Log "No updates found for the $checkedCount modules that could be checked"
        }
        else {
            Write-Log "All modules are up to date"
        }
        return
    }

    Write-Log "Found $($needsUpdate.Count) modules with available updates"

    $timeout = $script:Config.ModuleUpdateTimeoutSeconds

    # Update only modules that need it
    $moduleIndex = 0
    foreach ($module in $needsUpdate) {
        $moduleIndex++

        try {
            if ($PSCmdlet.ShouldProcess("$($module.Name) $($module.InstalledVersion) -> $($module.GalleryVersion)", "Update module")) {
                Write-Log "Updating module $moduleIndex/$($needsUpdate.Count): $($module.Name) ($($module.InstalledVersion) -> $($module.GalleryVersion))"

                $moduleTimer = [System.Diagnostics.Stopwatch]::StartNew()

                Invoke-ModuleUpdateWithRetry -Name $module.Name -Scope $scope `
                    -TrustRepository $script:Config.TrustPSGallery -TimeoutSeconds $timeout

                $script:Summary.ModulesUpdated++
                Write-Log "Updated: $($module.Name) (took $([math]::Round($moduleTimer.Elapsed.TotalSeconds))s)" -Level SUCCESS
            }
        }
        catch [System.TimeoutException] {
            Write-Log "Timed out updating $($module.Name) after ${timeout}s — skipping" -Level ERROR
            $script:Summary.ModulesFailed += @{
                Module = $module.Name
                Error  = $_.Exception.Message
            }
        }
        catch {
            $errorMsg = $_.Exception.Message

            # Detect OneDrive-locked folders: "Cannot remove package path <path>"
            # Loop because meta-packages like Az can hit multiple locked sub-module folders
            $maxAttempts = 20
            $attempt = 0
            $updated = $false

            while ($errorMsg -match 'Cannot remove package path\s+(.+?)\.?\s*(The previous|$)') {
                $attempt++
                if ($attempt -gt $maxAttempts) {
                    Write-Log "Reached max force-removal attempts ($maxAttempts) for $($module.Name)" -Level ERROR
                    break
                }

                $lockedPath = $Matches[1].TrimEnd('. ')
                if (-not (Test-Path $lockedPath)) { break }

                Write-Log "Locked folder detected ($attempt): $lockedPath — force-removing" -Level WARN
                if (-not (Remove-LockedModuleFolder -FolderPath $lockedPath)) { break }

                try {
                    Invoke-ModuleUpdateWithRetry -Name $module.Name -Scope $scope `
                        -TrustRepository $script:Config.TrustPSGallery -TimeoutSeconds $timeout
                    $script:Summary.ModulesUpdated++
                    Write-Log "Updated (after $attempt force-removal(s)): $($module.Name) (took $([math]::Round($moduleTimer.Elapsed.TotalSeconds))s)" -Level SUCCESS
                    $updated = $true
                    break
                }
                catch [System.TimeoutException] {
                    Write-Log "Timed out updating $($module.Name) after ${timeout}s — skipping" -Level ERROR
                    $script:Summary.ModulesFailed += @{
                        Module = $module.Name
                        Error  = $_.Exception.Message
                    }
                    $updated = $true  # Prevent double-logging below
                    break
                }
                catch {
                    $errorMsg = $_.Exception.Message
                }
            }

            if (-not $updated) {
                Write-Log "Failed to update $($module.Name): $errorMsg" -Level ERROR
                $script:Summary.ModulesFailed += @{
                    Module = $module.Name
                    Error  = $errorMsg
                }
            }
        }
    }

    Write-Log "Module updates complete. Updated: $($script:Summary.ModulesUpdated), Unsuccessful: $($script:Summary.ModulesFailed.Count)"
}

function Confirm-KeptModuleVersions {
    <#
    .SYNOPSIS
        Logs which versions KeepVersions protects from pruning, and warns about every
        selector that matches nothing installed.

    .NOTES
        This needs a pass of its own. The prune loop only looks at modules with more than
        one installed version, so a module with a single version and a selector that
        matches nothing would never be looked at.

        It belongs to the prune phase and is handed the list the prune loop works on.
        That list is read after the updates, so a line that was just updated shows its
        new version here.

        A selector that matches nothing is a WARN and nothing more. It installs nothing
        and does not make the run count as unsuccessful.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        [array]$InstalledResources
    )

    $moduleNames = @($script:Config.KeepVersions.Keys)
    if ($moduleNames.Count -eq 0) {
        return
    }

    Write-Log "Checking keep-version rules for $($moduleNames.Count) module(s)..."

    foreach ($configName in $moduleNames) {
        $selectors = @(Get-KeepVersionSelectors -Name $configName)
        $resources = @($InstalledResources | Where-Object { $_.Name -eq $configName })

        if ($resources.Count -eq 0) {
            $lineList = ($selectors | ForEach-Object { "'$($_.Requested)'" }) -join ', '
            Write-Log "KeepVersions: $configName is not installed, so nothing is kept or updated for $lineList" -Level WARN
            foreach ($selector in $selectors) {
                $script:Summary.KeepVersionsUnmatched += @{
                    Module   = $configName
                    Selector = $selector.Requested
                }
            }
            continue
        }

        # The name as installed, which may be cased differently from the config
        $name = $resources[0].Name
        $pin = Get-PinnedVersion $configName
        $plan = Get-ModulePrunePlan -Resources $resources -Pin $pin -Selectors $selectors

        # Two selectors can land on the same version. One line in the log names both
        foreach ($group in @($plan.Matched | Group-Object Version)) {
            $lines = ($group.Group | ForEach-Object { $_.Selector }) -join ', '
            Write-Log "Keeping $name v$($group.Name) (KeepVersions: $lines)"

            foreach ($item in $group.Group) {
                $script:Summary.KeepVersionsMatched += @{
                    Module   = $name
                    Selector = $item.Selector
                    Version  = $item.Version
                }
            }
        }

        foreach ($requested in $plan.Unmatched) {
            Write-Log "KeepVersions: no installed version of $name matches '$requested'. Nothing is kept or updated for it - a line is only maintained once a version of it is installed" -Level WARN
            $script:Summary.KeepVersionsUnmatched += @{
                Module   = $name
                Selector = $requested
            }
        }
    }
}

function Remove-OldModuleVersions {
    [CmdletBinding(SupportsShouldProcess)]
    param()

    Write-Log "Starting old version cleanup..."

    # Detect OneDrive on module path
    $currentUserModulePath = Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'PowerShell\Modules'
    $isOneDrive = Test-OneDrivePath $currentUserModulePath

    # When OneDrive is detected, modules live in AllUsers scope — query that explicitly
    $getParams = if ($isOneDrive) { @{ Scope = 'AllUsers' } } else { @{} }
    $allModules = Get-PSResource @getParams | Where-Object { $_.Name -notin $script:Config.ExcludedModules }

    # --- Pass 1: Remove old versions of AllUsers modules (keep latest) ---
    $allUsersModulePath = Join-Path $env:ProgramFiles 'PowerShell\Modules'

    if ($isOneDrive) {
        # Only prune what is physically in the AllUsers path. This used to filter on
        # InstalledLocation, but a module migrated out of OneDrive still reports its old
        # OneDrive path there. That hid every migrated module from pruning for good, so
        # old versions stayed while the log said "Found 0 modules with multiple versions"
        $nonOneDriveModules = $allModules | Where-Object {
            Resolve-ModuleVersionFolder -ModulesRoot $allUsersModulePath -Name $_.Name -Version $_.Version
        }
    }
    else {
        $nonOneDriveModules = $allModules
    }

    # Filter out built-in modules that ship with PowerShell (e.g. PackageManagement) —
    # PSResourceGet cannot uninstall these and always errors
    $psHomePath = $PSHOME
    $nonOneDriveModules = $nonOneDriveModules | Where-Object {
        -not ($_.InstalledLocation -like "$psHomePath*")
    }

    # Say what KeepVersions protects before anything is removed. It sees the same list as
    # the loop below
    Confirm-KeptModuleVersions -InstalledResources @($nonOneDriveModules)

    # @() matters: a single surviving group is a bare GroupInfo, whose own .Count member is
    # the number of versions in that group, not the number of groups
    $grouped = @($nonOneDriveModules | Group-Object Name | Where-Object { $_.Count -gt 1 })
    Write-Log "Found $($grouped.Count) modules with multiple versions"

    foreach ($group in $grouped) {
        # What stays is the newest version, or the pin, plus the newest version of every
        # kept line. Everything else goes, newer versions than a pin included
        $pin = Get-PinnedVersion $group.Name
        $selectors = @(Get-KeepVersionSelectors -Name $group.Name)
        $plan = Get-ModulePrunePlan -Resources @($group.Group) -Pin $pin -Selectors $selectors

        if ($plan.PinMissing) {
            # If the pin isn't installed, prune nothing: removing the rest would leave the
            # module with no version at all.
            Write-Log "$($group.Name) is pinned to $($pin.Requested) but that version is not installed — leaving all $($group.Count) installed version(s) in place" -Level WARN
            continue
        }

        $oldVersions = $plan.Remove

        foreach ($oldVersion in $oldVersions) {
            if ($PSCmdlet.ShouldProcess("$($oldVersion.Name) v$($oldVersion.Version)", "Remove old version")) {
                Write-Log "Removing: $($oldVersion.Name) v$($oldVersion.Version)"

                try {
                    $uninstallParams = @{
                        Name                = $oldVersion.Name
                        Version             = $oldVersion.Version
                        SkipDependencyCheck = $true
                        ErrorAction         = 'Stop'
                    }
                    if ($isOneDrive) { $uninstallParams['Scope'] = 'AllUsers' }
                    Uninstall-PSResource @uninstallParams
                    $script:Summary.VersionsPruned++
                    Write-Log "Removed: $($oldVersion.Name) v$($oldVersion.Version)" -Level SUCCESS
                }
                catch {
                    $errorMsg = $_.Exception.Message

                    # "does not exist" = built-in module or phantom metadata — skip silently
                    if ($errorMsg -match 'does not exist') {
                        Write-Log "Skipping $($oldVersion.Name) v$($oldVersion.Version): not managed by PSResourceGet" -Level WARN
                        continue
                    }

                    # If access denied / cannot delete, try force-removing the folder directly.
                    # With OneDrive in play the folder is looked up on disk, because
                    # InstalledLocation may still name the OneDrive path it was migrated from
                    $folderPath = $null
                    if ($isOneDrive) {
                        $folderPath = Resolve-ModuleVersionFolder -ModulesRoot $allUsersModulePath `
                            -Name $oldVersion.Name -Version $oldVersion.Version
                    }
                    $foundOnDisk = [bool]$folderPath
                    if (-not $foundOnDisk) {
                        $folderPath = $oldVersion.InstalledLocation
                    }

                    # PSResourceGet sometimes returns the modules root or module base instead
                    # of the version folder — detect and correct this
                    $versionString = $oldVersion.Version.ToString()
                    if ((-not $foundOnDisk) -and $folderPath -and -not $folderPath.EndsWith($versionString)) {
                        # Try to extract the correct path from the error message
                        if ($errorMsg -match "Parent directory '([^']+)'") {
                            $folderPath = $Matches[1]
                        }
                        else {
                            # Construct it: InstalledLocation may be the modules root or module base
                            $leaf = Split-Path $folderPath -Leaf
                            if ($leaf -eq $oldVersion.Name) {
                                $folderPath = Join-Path $folderPath $versionString
                            }
                            else {
                                $folderPath = Join-Path $folderPath $oldVersion.Name $versionString
                            }
                        }
                    }
                    if ($folderPath -and (Test-Path $folderPath) -and
                        ($errorMsg -match 'Access.*denied|could not be deleted|Cannot remove')) {
                        Write-Log "Lock detected — attempting force-removal of $folderPath" -Level WARN
                        if (Remove-LockedModuleFolder -FolderPath $folderPath) {
                            $script:Summary.VersionsPruned++
                            Write-Log "Force-removed: $($oldVersion.Name) v$($oldVersion.Version)" -Level SUCCESS
                            continue
                        }
                    }

                    Write-Log "Failed to remove $($oldVersion.Name) v$($oldVersion.Version): $errorMsg" -Level ERROR
                    $script:Summary.PrunesFailed += @{
                        Module  = $oldVersion.Name
                        Version = $oldVersion.Version.ToString()
                        Error   = $errorMsg
                    }
                }
            }
        }
    }

    # --- Warn about modules in OneDrive path ---
    # After migration, modules should live in AllUsers only. If new modules appear
    # in the OneDrive CurrentUser path, warn the user instead of silently deleting them.
    if ($isOneDrive -and (Test-Path $currentUserModulePath)) {
        $odModuleFolders = @(Get-ChildItem -Path $currentUserModulePath -Directory -ErrorAction SilentlyContinue |
            Where-Object {
                # Only count folders that contain version subfolders (real modules)
                Get-ChildItem -Path $_.FullName -Directory -ErrorAction SilentlyContinue |
                    Where-Object { $_.Name -match '^\d+(\.\d+){1,3}$' }
            })

        # Only a module PSResourceGet installed can be migrated and then kept up to date.
        # One that another program put here belongs to that program and stays put
        $toMigrate = @($odModuleFolders | Where-Object { Test-ManagedModuleFolder -ModuleFolder $_.FullName })
        $leftAlone = @($odModuleFolders | Where-Object { $_.Name -notin $toMigrate.Name })

        if ($toMigrate.Count -gt 0) {
            $moduleNames = ($toMigrate | Select-Object -ExpandProperty Name) -join ', '
            Write-Log "Found $($toMigrate.Count) module(s) in OneDrive path: $moduleNames — run Invoke-OneDriveMigration.ps1 to migrate them" -Level WARN
        }

        if ($leftAlone.Count -gt 0) {
            $leftAloneNames = ($leftAlone | Select-Object -ExpandProperty Name) -join ', '
            Write-Log "Leaving $($leftAlone.Count) module(s) in OneDrive path alone, installed by another program: $leftAloneNames"
        }
    }

    Write-Log "Version cleanup complete. Removed: $($script:Summary.VersionsPruned), Unsuccessful: $($script:Summary.PrunesFailed.Count)"
}

# ============================================================================
# MAIN EXECUTION
# ============================================================================

try {
    $script:Summary.StartTime = Get-Date -Format 'o'

    # Load configuration
    Import-MaintenanceConfig -Path $ConfigPath

    # Initialize logging
    Initialize-Logging -BasePath $LogPath

    # Surface anything the config parser flagged before the log file existed
    foreach ($configWarning in $script:ConfigWarnings) {
        Write-Log $configWarning -Level WARN
    }

    if ($script:ConfigFault) {
        # A config file that is there but cannot be read. Going on would mean running on
        # the built-in defaults, under which nothing is excluded, pinned or kept: every
        # module would be updated and every old version pruned, including the ones the
        # file was written to protect. So this run touches no module, and no log either,
        # since how long logs are kept is in that file too
        $script:Summary.ConfigFault = $script:ConfigFault
        Write-Log "Could not read the config file: $($script:ConfigFault)" -Level ERROR
        Write-Log "Nothing was updated or pruned. Without the config it is not known which modules are excluded, pinned or kept" -Level ERROR

        # Whether monitoring is on is in that file as well. Look for the secret anyway,
        # so that this is heard about today and not when the check goes overdue
        $script:HealthchecksUrl = Get-HealthchecksUrl -WithoutConfig
    }
    else {
        # Record the effective pins in the summary (name -> version) so they show up even for -PruneOnly
        foreach ($pinName in $script:Config.PinnedModules.Keys) {
            $script:Summary.PinnedModules[$pinName] = $script:Config.PinnedModules[$pinName].Requested
        }

        # The same for the kept version lines (name -> list of selectors as written)
        foreach ($keepName in $script:Config.KeepVersions.Keys) {
            $script:Summary.KeepVersions[$keepName] = @($script:Config.KeepVersions[$keepName] |
                ForEach-Object { $_.Requested })
        }

        # Log configuration
        Write-Log "Configuration loaded:"
        Write-Log "  - Excluded modules: $($script:Config.ExcludedModules.Count)"
        Write-Log "  - Pinned modules: $($script:Config.PinnedModules.Count)"
        Write-Log "  - Modules with kept versions: $($script:Config.KeepVersions.Count)"
        Write-Log "  - Log retention: $($script:Config.LogRetentionDays) days"
        Write-Log "  - Trust PSGallery: $($script:Config.TrustPSGallery)"
        Write-Log "  - Notification mode: $($script:Config.NotificationMode)"
        Write-Log "  - Module update timeout: $($script:Config.ModuleUpdateTimeoutSeconds)s"
        $hcEnabled = $script:Config.Healthchecks.Enabled
        Write-Log "  - Healthchecks monitoring: $hcEnabled"

        # Resolve the ping URL before any update or prune work — this script prunes
        # SecretManagement itself, so the secret is read while the module is still loadable
        $script:HealthchecksUrl = Get-HealthchecksUrl
        Send-HealthchecksPing -PingType Start

        # Flag a fragile task registration while the script can still be heard
        Test-ScheduledTaskHealth

        # Clean up old logs
        Remove-OldLogs -BasePath $LogPath -RetentionDays $script:Config.LogRetentionDays

        # Perform operations — save summary after each phase so state is preserved if the process is killed
        if (-not $PruneOnly) {
            Update-AllModules
            Save-Summary -BasePath $LogPath
        }

        if (-not $UpdateOnly) {
            Remove-OldModuleVersions
            Save-Summary -BasePath $LogPath
        }
    }

    # The closing line has to agree with the toast and the ping. It used to claim success
    # unconditionally, which could put "completed successfully" directly above a Fail
    # ping for a run whose updates had been unsuccessful
    $failures = Get-SummaryFailures -Summary $script:Summary

    Write-Log "======================================================"
    if ($failures.Total -gt 0) {
        $failureCount = $failures.Total
        $configFailed = $failures.Config
        $lookupsFailed = $failures.Lookups
        $updatesFailed = $failures.Updates
        $pinsFailed = $failures.Pins
        $prunesFailed = $failures.Prunes
        Write-Log "PSModuleMaintenance completed with $failureCount unsuccessful operation(s) (config: $configFailed, lookups: $lookupsFailed, updates: $updatesFailed, pins: $pinsFailed, prunes: $prunesFailed)" -Level WARN
    }
    else {
        Write-Log "PSModuleMaintenance completed successfully" -Level SUCCESS
    }
    Write-Log "======================================================"
}
catch {
    # Recorded so the finally block can distinguish a crash from a clean finish
    $script:FatalError = $_
    Write-Log "Critical error: $_" -Level ERROR
    throw
}
finally {
    # Save summary
    Save-Summary -BasePath $LogPath

    # Send toast notification based on config
    $notifyMode = $script:Config.NotificationMode
    $hasFailures = (Get-SummaryFailures -Summary $script:Summary).Total -gt 0

    if ($notifyMode -eq 'Always' -or ($notifyMode -eq 'OnFailure' -and $hasFailures)) {
        Send-ToastNotification -Summary $script:Summary -SkippedUpdates:$PruneOnly -SkippedPruning:$UpdateOnly
    }

    # Healthchecks closing ping. Reuses the same $hasFailures rule the toast uses, so the
    # two notification channels can never disagree about what counts as a failure
    $runMode = if ($UpdateOnly) { 'UpdateOnly' } elseif ($PruneOnly) { 'PruneOnly' } else { 'Full' }
    $isFailure = ($null -ne $script:FatalError) -or $hasFailures
    $pingBody = Format-HealthchecksBody -Summary $script:Summary -Mode $runMode -IsFailure:$isFailure
    $pingType = if ($isFailure) { 'Fail' } else { 'Success' }
    Send-HealthchecksPing -PingType $pingType -Body $pingBody

    # Stop transcript
    Stop-Transcript -ErrorAction SilentlyContinue | Out-Null
}
