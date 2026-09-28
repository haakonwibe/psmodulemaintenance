<#
.SYNOPSIS
    Tests for a config file that cannot be read: the run must touch no module, and must
    say so where it will be heard.

.DESCRIPTION
    The first part covers Get-HealthchecksUrl -WithoutConfig, lifted out of the script.

    The second part runs the whole script, because what is being tested is what the main
    block does and does not call. It runs in a process of its own in which everything
    that could change the machine or reach the network is replaced by a stand-in that
    only writes down that it was called: the PSResourceGet cmdlets, Get-Secret,
    Invoke-RestMethod and the toast. That process checks its stand-ins before it starts
    the script, and stops if one is missing.
#>
param([switch]$Quiet)

. (Join-Path $PSScriptRoot 'TestHelpers.ps1')

# =====================================================================================
# Part 1: looking for the secret without a config
# =====================================================================================

. ([scriptblock]::Create((Get-ScriptFunctionText -Path $script:MainScript -Name 'Get-HealthchecksUrl')))

$script:LogLines = @()
function Write-Log {
    param([Parameter(Position = 0)][string]$Message, [string]$Level = 'INFO')
    $script:LogLines += "[$Level] $Message"
}

# Stand-ins for the vault. Defined as functions, which are found before cmdlets
$script:SecretAnswer = { 'https://hc.invalid/ping/made-up' }
$script:SecretCalls = @()
function Import-Module {
    # Only the vault module is kept from loading, anything else goes to the real cmdlet
    if ("$args" -like '*Microsoft.PowerShell.SecretManagement*') { return }
    Microsoft.PowerShell.Core\Import-Module @args
}
function Get-Secret {
    [CmdletBinding()]
    param([string]$Name, [string]$Vault, [switch]$AsPlainText)
    $script:SecretCalls += "$Vault/$Name"
    & $script:SecretAnswer
}

function Reset-Lookup {
    param([bool]$Enabled = $false)
    $script:Config = @{
        Healthchecks = @{
            Enabled = $Enabled; SecretName = 'PSModuleMaintenance-Healthchecks'
            SecretVault = 'SecretStore'; TimeoutSeconds = 10
        }
    }
    $script:LogLines = @()
    $script:SecretCalls = @()
}

Write-Section 'Get-HealthchecksUrl -WithoutConfig'

Reset-Lookup
$script:SecretAnswer = { ' https://hc.invalid/ping/made-up/ ' }
$url = Get-HealthchecksUrl -WithoutConfig
Assert-That ($url -eq 'https://hc.invalid/ping/made-up') 'a secret under the default name is used, tidied up'
Assert-That (($script:SecretCalls -join ',') -eq 'SecretStore/PSModuleMaintenance-Healthchecks') 'it is looked for under the default name and vault'
Assert-That ($script:LogLines -contains "[INFO] Healthchecks secret found under its default name 'PSModuleMaintenance-Healthchecks' - this run will be reported") 'the log says the run will be reported'
Assert-That (@($script:LogLines | Where-Object { $_ -like '*hc.invalid*' }).Count -eq 0) 'the URL itself is never written to the log'

Reset-Lookup
$script:SecretAnswer = { throw 'The secret PSModuleMaintenance-Healthchecks was not found.' }
$url = Get-HealthchecksUrl -WithoutConfig
Assert-That ($null -eq $url) 'no secret: nothing to use'
Assert-That (@($script:LogLines | Where-Object { ($_ -like '`[WARN`]*') -or ($_ -like '`[ERROR`]*') }).Count -eq 0) 'and that is not reported as a problem, monitoring may never have been set up'
Assert-That ($script:LogLines -contains "[INFO] No usable Healthchecks secret under the default name 'PSModuleMaintenance-Healthchecks' - this run will not be reported there") 'the log says the run will not be reported'

Reset-Lookup
$script:SecretAnswer = { 'not a web address' }
Assert-That ($null -eq (Get-HealthchecksUrl -WithoutConfig)) 'a secret that is not a web address is not used'

Reset-Lookup
$script:SecretAnswer = { '   ' }
Assert-That ($null -eq (Get-HealthchecksUrl -WithoutConfig)) 'nor is an empty one'

Reset-Lookup -Enabled $false
$script:SecretAnswer = { 'https://hc.invalid/ping/made-up' }
$url = Get-HealthchecksUrl
Assert-That (($null -eq $url) -and ($script:SecretCalls.Count -eq 0)) 'with a config that says monitoring is off, the vault is not even opened'
Assert-That ($script:LogLines -contains '[INFO] Healthchecks monitoring is off') 'as before'

Reset-Lookup -Enabled $true
$url = Get-HealthchecksUrl
Assert-That ($url -eq 'https://hc.invalid/ping/made-up') 'with a config that says monitoring is on, the secret is read as before'
Assert-That ($script:LogLines -contains '[INFO] Healthchecks monitoring is on') 'and logged as before'

# =====================================================================================
# Part 2: the whole script
# =====================================================================================

$root = Join-Path ([System.IO.Path]::GetTempPath()) ('PSModuleMaintenance-tests-' + [guid]::NewGuid().ToString('N'))
$runnerPath = Join-Path $root 'runner.ps1'

# Runs in the child process. Everything that could change the machine or reach the
# network is a stand-in here, and each one writes down that it was called
$runner = @'
param([string]$ScriptPath, [string]$ConfigPath, [string]$LogPath, [string]$CallsPath, [string]$Mode)

# Get-PSResource is not a cmdlet. It is an alias the module gives Get-InstalledPSResource,
# and an alias is found before a function, so a stand-in named Get-PSResource is passed
# over as soon as the module is loaded. The module is therefore loaded first, here, and
# its aliases are taken away before the stand-ins are defined. The script's own
# "#Requires -Modules" then finds the module loaded and does not bring them back
Microsoft.PowerShell.Core\Import-Module Microsoft.PowerShell.PSResourceGet -ErrorAction Stop
foreach ($alias in @(Get-Alias | Where-Object { $_.Source -eq 'Microsoft.PowerShell.PSResourceGet' })) {
    Remove-Alias -Name $alias.Name -Scope Global -Force
}

function global:Add-Call { param([string]$Text) Add-Content -LiteralPath $CallsPath -Value $Text }

function global:powershell.exe { Add-Call 'toast'; $global:LASTEXITCODE = 0 }

# Only the vault module is kept from loading. Everything else goes to the real cmdlet:
# "#Requires -Modules" finds Import-Module by name, so it ends up here as well
function global:Import-Module {
    if ("$args" -like '*Microsoft.PowerShell.SecretManagement*') { return }
    Microsoft.PowerShell.Core\Import-Module @args
}

function global:Get-Secret {
    [CmdletBinding()] param([string]$Name, [string]$Vault, [switch]$AsPlainText)
    Add-Call "secret $Vault/$Name"
    'https://hc.invalid/ping/made-up'
}
function global:Invoke-RestMethod {
    [CmdletBinding()]
    param($Uri, $Method, $TimeoutSec, $MaximumRetryCount, $RetryIntervalSec, $Body, $ContentType)
    Add-Call "ping $Uri"
    if ($Body) { Add-Call ('body ' + (($Body -split "`r?`n") -join ' | ')) }
}
function global:Get-PSResource { [CmdletBinding()] param($Name, $Scope) Add-Call 'Get-PSResource' }
function global:Get-InstalledPSResource { [CmdletBinding()] param($Name, $Scope) Add-Call 'Get-PSResource' }
function global:Find-PSResource { [CmdletBinding()] param($Name, $Repository, $Version) Add-Call 'Find-PSResource' }
function global:Update-PSResource { [CmdletBinding()] param($Name, $Scope) Add-Call 'Update-PSResource' }
function global:Install-PSResource { [CmdletBinding()] param($Name, $Scope, $Version) Add-Call 'Install-PSResource' }
function global:Uninstall-PSResource { [CmdletBinding()] param($Name, $Scope, $Version) Add-Call 'Uninstall-PSResource' }
function global:Get-ScheduledTask { [CmdletBinding()] param($TaskName) }

# Checked with the module loaded, which is the state the script will run in
$mustBeStandIns = 'Get-PSResource', 'Get-InstalledPSResource', 'Find-PSResource', 'Update-PSResource',
                  'Install-PSResource', 'Uninstall-PSResource', 'Get-Secret', 'Invoke-RestMethod',
                  'powershell.exe'
foreach ($name in $mustBeStandIns) {
    if ((Get-Command $name).CommandType -ne 'Function') {
        Add-Call "NOT A STAND-IN: $name"
        exit 3
    }
}

$arguments = @{ ConfigPath = $ConfigPath; LogPath = $LogPath }
if ($Mode) { $arguments[$Mode] = $true }
& $ScriptPath @arguments
'@

# Runs the script once and returns what it logged, what it called and what it summed up
function Invoke-WholeScript {
    param([string]$ConfigText, [string]$Mode, [switch]$WithOldLog)

    $folder = Join-Path $root ([guid]::NewGuid().ToString('N'))
    $logFolder = Join-Path $folder 'Logs'
    New-Item -Path $logFolder -ItemType Directory -Force | Out-Null
    $configPath = Join-Path $folder 'config.json'
    $callsPath = Join-Path $folder 'calls.txt'
    Set-Content -LiteralPath $configPath -Value $ConfigText
    Set-Content -LiteralPath $callsPath -Value 'started'

    $oldLog = Join-Path $logFolder 'maintenance_2020-01-05_030000.log'
    if ($WithOldLog) {
        Set-Content -LiteralPath $oldLog -Value 'an old log'
        (Get-Item -LiteralPath $oldLog).LastWriteTime = (Get-Date).AddDays(-400)
    }

    $all = @('-NoProfile', '-NonInteractive', '-File', $runnerPath, '-ScriptPath', $script:MainScript,
             '-ConfigPath', $configPath, '-LogPath', $logFolder, '-CallsPath', $callsPath)
    if ($Mode) { $all += @('-Mode', $Mode) }
    & pwsh @all *> $null
    $exitCode = $LASTEXITCODE

    $logFile = Get-ChildItem -LiteralPath $logFolder -Filter 'maintenance_*.log' |
        Where-Object { $_.FullName -ne $oldLog } | Select-Object -First 1
    $summaryFile = Get-ChildItem -LiteralPath $logFolder -Filter 'summary_*.json' | Select-Object -First 1

    $logLines = @()
    if ($logFile) {
        $logLines = @(Get-Content -LiteralPath $logFile.FullName | ForEach-Object { $_ -replace '^\[[\d\- :]+\] ', '' })
    }
    $summaryText = ''
    if ($summaryFile) {
        $summaryText = Get-Content -LiteralPath $summaryFile.FullName -Raw
    }

    return [PSCustomObject]@{
        ExitCode    = $exitCode
        Log         = $logLines
        Calls       = @(Get-Content -LiteralPath $callsPath)
        SummaryText = $summaryText
        OldLogKept  = (Test-Path -LiteralPath $oldLog)
    }
}

$touching = 'Get-PSResource', 'Find-PSResource', 'Update-PSResource', 'Install-PSResource', 'Uninstall-PSResource'
$brokenConfig = '{ "ExcludedModules": ["Contoso.Tools"], "KeepVersions": { "Fabrikam.Core": ["5"] '
$goodConfig = '{ "NotificationMode": "Never", "LogRetentionDays": 180, "Healthchecks": { "Enabled": false } }'

try {
    New-Item -Path $root -ItemType Directory -Force | Out-Null
    Set-Content -LiteralPath $runnerPath -Value $runner

    # --- A config that can be read, for comparison -----------------------------------
    Write-Section 'The whole script, with a config that can be read'

    $run = Invoke-WholeScript -ConfigText $goodConfig -WithOldLog
    Assert-That ($run.ExitCode -ne 3) 'the stand-ins were in place'
    Assert-That ($run.Log -contains '[INFO] Found 0 installed modules (excluding: )') 'the script sees the stand-in, not the modules on this machine'
    Assert-That ($run.Log -contains '[INFO] Starting module updates...') 'the update phase is entered'
    Assert-That ($run.Log -contains '[INFO] Starting old version cleanup...') 'the prune phase is entered'
    Assert-That ($run.Calls -contains 'Get-PSResource') 'installed modules are looked at'
    Assert-That ($run.Log -contains '[SUCCESS] PSModuleMaintenance completed successfully') 'the run ends as a success'
    Assert-That (-not $run.OldLogKept) 'a log older than the retention is removed'
    Assert-That (@($run.Calls | Where-Object { $_ -like 'ping*' -or $_ -eq 'toast' }).Count -eq 0) 'nothing is sent, as the config says'

    # --- A config that cannot be read ------------------------------------------------
    Write-Section 'The whole script, with a config that cannot be read'

    $run = Invoke-WholeScript -ConfigText $brokenConfig -WithOldLog
    Assert-That ($run.ExitCode -ne 3) 'the stand-ins were in place'

    Assert-That (@($run.Calls | Where-Object { $_ -in $touching }).Count -eq 0) "no module is looked at, updated or removed (calls: $(($run.Calls | Where-Object { $_ -in $touching }) -join ', '))"
    Assert-That ($run.Log -notcontains '[INFO] Starting module updates...') 'the update phase is not entered'
    Assert-That ($run.Log -notcontains '[INFO] Starting old version cleanup...') 'the prune phase is not entered'
    Assert-That ($run.OldLogKept) 'no old log is removed either, how long to keep them is in that file too'

    Assert-That (@($run.Log | Where-Object { $_ -like '`[ERROR`] Could not read the config file: *' }).Count -eq 1) 'the log says the file could not be read, with the reason'
    Assert-That ($run.Log -contains '[ERROR] Nothing was updated or pruned. Without the config it is not known which modules are excluded, pinned or kept') 'and what that meant for this run'
    Assert-That ($run.Log -contains '[WARN] PSModuleMaintenance completed with 1 unsuccessful operation(s) (config: 1, lookups: 0, updates: 0, pins: 0, prunes: 0)') 'the closing line counts it'
    Assert-That (@($run.Log | Where-Object { $_ -like '*completed successfully*' }).Count -eq 0) 'and does not claim success'
    Assert-That ($run.Log -notcontains '[INFO] Configuration loaded:') 'no configuration is reported as loaded'

    Assert-That ($run.Calls -contains 'toast') 'a toast is shown'
    Assert-That ($run.Log -contains '[INFO] Toast notification sent: The config file could not be read. Nothing was updated or pruned.') 'saying what happened'

    Assert-That ($run.Calls -contains 'secret SecretStore/PSModuleMaintenance-Healthchecks') 'the secret is looked for under its default name'
    Assert-That ($run.Calls -contains 'ping https://hc.invalid/ping/made-up/fail') 'a fail ping is sent'
    Assert-That (@($run.Calls | Where-Object { $_ -like 'ping *' }).Count -eq 1) 'and no other ping, not even a start ping'
    Assert-That (@($run.Calls | Where-Object { $_ -like 'body *config: could not be read, so nothing was updated or pruned*' }).Count -eq 1) 'its body says why'
    Assert-That ($run.Log -contains '[INFO] Healthchecks ping sent: Fail') 'the log records the ping'

    Assert-That (@($run.Log | Where-Object { $_ -like '*hc.invalid*' }).Count -eq 0) 'the ping URL is not in the log'
    Assert-That ($run.SummaryText -notlike '*hc.invalid*') 'nor in the summary'
    $summary = $run.SummaryText | ConvertFrom-Json
    Assert-That (-not [string]::IsNullOrWhiteSpace($summary.ConfigFault)) 'the summary records the fault'
    Assert-That (($summary.ModulesUpdated -eq 0) -and ($summary.VersionsPruned -eq 0)) 'and that nothing was updated or pruned'

    # --- The same, whatever the run was asked to do ----------------------------------
    Write-Section 'A config that cannot be read, in the other modes'

    foreach ($mode in 'PruneOnly', 'UpdateOnly') {
        $run = Invoke-WholeScript -ConfigText $brokenConfig -Mode $mode
        Assert-That (@($run.Calls | Where-Object { $_ -in $touching }).Count -eq 0) "-${mode}: no module is touched"
        Assert-That (@($run.Log | Where-Object { $_ -like '`[ERROR`] Could not read the config file: *' }).Count -eq 1) "-${mode}: and the log says why"
    }

    $run = Invoke-WholeScript -ConfigText $brokenConfig -Mode 'WhatIf'
    Assert-That (@($run.Calls | Where-Object { $_ -in $touching }).Count -eq 0) '-WhatIf: no module is touched'
    Assert-That (@($run.Calls | Where-Object { $_ -like 'ping *' }).Count -eq 0) '-WhatIf: no ping is sent, a dry run never pings'
    Assert-That ($run.Log -contains '[INFO] Healthchecks ping skipped (WhatIf): Fail') '-WhatIf: and the log says it was skipped'

    # --- A file that is simply not there ---------------------------------------------
    Write-Section 'No config file at all'

    $folder = Join-Path $root 'no-config'
    $logFolder = Join-Path $folder 'Logs'
    New-Item -Path $logFolder -ItemType Directory -Force | Out-Null
    $callsPath = Join-Path $folder 'calls.txt'
    Set-Content -LiteralPath $callsPath -Value 'started'
    & pwsh -NoProfile -NonInteractive -File $runnerPath -ScriptPath $script:MainScript `
        -ConfigPath (Join-Path $folder 'config.json') -LogPath $logFolder -CallsPath $callsPath -Mode 'WhatIf' *> $null
    $noConfigLog = @(Get-ChildItem -LiteralPath $logFolder -Filter 'maintenance_*.log' | Get-Content |
        ForEach-Object { $_ -replace '^\[[\d\- :]+\] ', '' })
    Assert-That ($noConfigLog -contains '[INFO] Starting module updates...') 'the run goes ahead on the built-in defaults, as documented'
    Assert-That (@($noConfigLog | Where-Object { $_ -like '*Could not read the config file*' }).Count -eq 0) 'a missing file is not an unreadable one'
}
finally {
    if (Test-Path -LiteralPath $root) {
        Remove-Item -LiteralPath $root -Recurse -Force
    }
}

Complete-Tests
