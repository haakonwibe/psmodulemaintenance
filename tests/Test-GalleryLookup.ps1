<#
.SYNOPSIS
    Tests for the PSGallery lookup and for how an incomplete lookup is reported.

.DESCRIPTION
    Covers Find-GalleryModules, Get-GalleryFaultText, Get-SummaryFailures, and the parts
    of Format-HealthchecksBody and Send-ToastNotification that report a lookup problem.

    Find-PSResource is replaced by a stand-in that raises errors with the ids the real
    cmdlet uses. powershell.exe is replaced too, so no notification is ever shown.
#>
param([switch]$Quiet)

. (Join-Path $PSScriptRoot 'TestHelpers.ps1')

$wanted = 'Get-SummaryFailures', 'Get-TransientNetworkFault', 'Get-GalleryFaultText',
          'Find-GalleryModules', 'Format-HealthchecksBody', 'Send-ToastNotification'
foreach ($name in $wanted) {
    . ([scriptblock]::Create((Get-ScriptFunctionText -Path $script:MainScript -Name $name)))
}

# The ping body names the machine. Fixed here so the test reads the same everywhere
$env:COMPUTERNAME = 'TEST-HOST'

# --- Stand-ins -----------------------------------------------------------------------
$script:LogLines = @()
function Write-Log {
    param([Parameter(Position = 0)][string]$Message, [string]$Level = 'INFO')
    $script:LogLines += "[$Level] $Message"
}

$probe = 'Microsoft.PowerShell.PSResourceGet'
$script:FindCalls = @()
$script:Outcome = { 'ok' }

# The outcome script block is given the module name and the call number, and answers
# ok, notfound, reworded, network or throw
function Find-PSResource {
    [CmdletBinding()]
    param([string[]]$Name, [string]$Repository)

    $script:FindCalls += , @($Name)
    $callNumber = $script:FindCalls.Count

    foreach ($n in $Name) {
        switch (& $script:Outcome $n $callNumber) {
            'ok' {
                [PSCustomObject]@{ Name = $n; Version = [version]'9.9.9' }
            }
            'notfound' {
                Write-Error -ErrorId 'PackageNotFound' -Category ObjectNotFound `
                    -Message "Package with name '$n' could not be found in repository 'PSGallery'."
            }
            'reworded' {
                Write-Error -ErrorId 'PackageNotFound' -Category ObjectNotFound `
                    -Message "Nothing called $n exists in PSGallery."
            }
            'network' {
                Write-Error -ErrorId 'HttpRequestCallFailure' `
                    -Message "'No such host is known. (proxy.invalid:8080)' Request sent: 'https://www.powershellgallery.com/api/v2/FindPackagesById()?id=%27$n%27'"
            }
            'throw' {
                throw "Repository 'PSGallery' is not registered."
            }
        }
    }
}

function Reset-Test {
    $script:FindCalls = @()
    $script:LogLines = @()
}

function Get-BulkCalls {
    , @($script:FindCalls | Where-Object { -not (($_.Count -eq 1) -and ($_[0] -eq $probe)) })
}

function Get-Warnings {
    , @($script:LogLines | Where-Object { $_ -like '`[WARN`]*' })
}

$names = 'Pester', 'Az.Accounts', 'Private.Module', 'Microsoft.Graph', 'PSReadLine'

# --- Healthy gallery -----------------------------------------------------------------
Write-Section 'Find-GalleryModules: healthy gallery'
Reset-Test
$script:Outcome = { param($n) if ($n -eq 'Private.Module') { 'notfound' } else { 'ok' } }
$r = Find-GalleryModules -Name $names -RetryDelaySeconds 0, 0
Assert-That ($r.Resources.Count -eq 4) "4 of 5 found (got $($r.Resources.Count))"
Assert-That ($r.Unchecked.Count -eq 0) 'a module that is not on the gallery is not unchecked'
Assert-That (($null -eq $r.Fault) -and ($null -eq $r.Detail)) 'no fault reported'
Assert-That ($script:FindCalls.Count -eq 2) "2 calls: probe and bulk (got $($script:FindCalls.Count))"
Assert-That ($script:LogLines.Count -eq 0) 'nothing logged'

# --- Gallery unreachable -------------------------------------------------------------
Write-Section 'Find-GalleryModules: gallery unreachable'
Reset-Test
$script:Outcome = { 'network' }
$r = Find-GalleryModules -Name $names -RetryDelaySeconds 0, 0
$warnings = Get-Warnings
Assert-That ($r.Resources.Count -eq 0) 'nothing found'
Assert-That ($r.Unchecked.Count -eq 5) "all 5 reported unchecked (got $($r.Unchecked.Count))"
Assert-That ($r.Fault -eq 'No such host is known') "short fault text: '$($r.Fault)'"
Assert-That ($r.Detail -like '*Request sent*') 'full detail kept for the log'
Assert-That ($script:FindCalls.Count -eq 3) "3 probes (got $($script:FindCalls.Count))"
Assert-That ((Get-BulkCalls).Count -eq 0) 'bulk query never sent while the probe goes unanswered'
Assert-That ($warnings.Count -eq 2) '2 WARN lines, one per retry'
Assert-That ($warnings[0] -eq '[WARN] PSGallery gave no answer for 5 module(s) (attempt 1 of 3): No such host is known. Retrying in 0s') 'WARN line reads as documented'

# --- Network comes up on the second attempt ------------------------------------------
Write-Section 'Find-GalleryModules: network comes up on the second attempt'
Reset-Test
$script:Outcome = {
    param($n, $call)
    if ($call -eq 1) { 'network' } elseif ($n -eq 'Private.Module') { 'notfound' } else { 'ok' }
}
$r = Find-GalleryModules -Name $names -RetryDelaySeconds 0, 0
Assert-That ($r.Resources.Count -eq 4) "4 found after the retry (got $($r.Resources.Count))"
Assert-That ($r.Unchecked.Count -eq 0) 'nothing unchecked'
Assert-That ($null -eq $r.Fault) 'fault cleared once the lookup succeeded'
Assert-That ((Get-Warnings).Count -eq 1) '1 WARN line'

# --- Two lookups dropped, then answered ----------------------------------------------
Write-Section 'Find-GalleryModules: two lookups dropped, then answered'
Reset-Test
$script:Outcome = {
    param($n, $call)
    if ($n -eq 'Private.Module') { 'notfound' }
    elseif (($call -le 2) -and ($n -in 'Az.Accounts', 'Microsoft.Graph')) { 'network' }
    else { 'ok' }
}
$r = Find-GalleryModules -Name $names -RetryDelaySeconds 0, 0
$bulk = Get-BulkCalls
Assert-That ($r.Resources.Count -eq 4) "4 found in total (got $($r.Resources.Count))"
Assert-That (@($r.Resources.Name | Sort-Object -Unique).Count -eq 4) 'no module returned twice'
Assert-That ($r.Unchecked.Count -eq 0) 'nothing unchecked'
Assert-That ($bulk.Count -eq 2) "2 bulk queries (got $($bulk.Count))"
Assert-That (($bulk[1] -join ',') -eq 'Az.Accounts,Microsoft.Graph') "the retry asks only about the dropped ones: $($bulk[1] -join ',')"

# --- One lookup never answered -------------------------------------------------------
Write-Section 'Find-GalleryModules: one lookup never answered'
Reset-Test
$script:Outcome = {
    param($n)
    if ($n -eq 'Private.Module') { 'notfound' } elseif ($n -eq 'Microsoft.Graph') { 'network' } else { 'ok' }
}
$r = Find-GalleryModules -Name $names -RetryDelaySeconds 0, 0
$bulk = Get-BulkCalls
Assert-That ($r.Resources.Count -eq 3) "3 found (got $($r.Resources.Count))"
Assert-That (($r.Unchecked -join ',') -eq 'Microsoft.Graph') "only the unanswered module is unchecked: $($r.Unchecked -join ',')"
Assert-That ($bulk.Count -eq 3) "3 bulk queries (got $($bulk.Count))"
Assert-That (($bulk[2] -join ',') -eq 'Microsoft.Graph') 'the last retry asks about that module alone'

# --- Terminating error ---------------------------------------------------------------
Write-Section 'Find-GalleryModules: terminating error'
Reset-Test
$script:Outcome = { 'throw' }
$threw = $false
try { $r = Find-GalleryModules -Name $names -RetryDelaySeconds 0 } catch { $threw = $true }
Assert-That (-not $threw) 'does not throw'
Assert-That ($r.Unchecked.Count -eq 5) 'all reported unchecked'
Assert-That ($r.Fault -like '*not registered*') "a fault that is not a network fault is passed through: '$($r.Fault)'"

# --- Not-found wording changes -------------------------------------------------------
Write-Section 'Find-GalleryModules: reworded not-found message'
Reset-Test
$script:Outcome = { param($n) if ($n -eq 'Private.Module') { 'reworded' } else { 'ok' } }
$r = Find-GalleryModules -Name $names -RetryDelaySeconds 0, 0
Assert-That ($r.Unchecked.Count -eq 0) 'still not unchecked: classified by error id, not by wording'
Assert-That ($script:FindCalls.Count -eq 2) 'no retry'

# --- Single module -------------------------------------------------------------------
Write-Section 'Find-GalleryModules: single name'
Reset-Test
$script:Outcome = { 'ok' }
$r = Find-GalleryModules -Name 'Pester' -RetryDelaySeconds 0, 0
Assert-That (($r.Resources.Count -eq 1) -and ($r.Unchecked.Count -eq 0)) 'one name in, one resource out'

Reset-Test
$script:Outcome = { 'network' }
$r = Find-GalleryModules -Name 'Pester' -RetryDelaySeconds 0
Assert-That ($r.Unchecked.Count -eq 1) 'one unchecked name still counts as 1'

# --- The probe module is not on the gallery ------------------------------------------
# The probe is there to see whether the gallery answers. "Not found" is an answer
Write-Section 'Find-GalleryModules: probe module not on the gallery'
Reset-Test
$script:Outcome = { param($n) if ($n -in $probe, 'Private.Module') { 'notfound' } else { 'ok' } }
$r = Find-GalleryModules -Name $names -RetryDelaySeconds 0, 0
Assert-That ($r.Resources.Count -eq 4) "the bulk query still runs, 4 found (got $($r.Resources.Count))"
Assert-That ($r.Unchecked.Count -eq 0) '"not found" for the probe is not an outage'
Assert-That ($script:FindCalls.Count -eq 2) "no retry: probe and bulk (got $($script:FindCalls.Count))"
Assert-That ($script:LogLines.Count -eq 0) 'nothing logged'

Reset-Test
$script:Outcome = { param($n) if ($n -eq $probe) { 'reworded' } else { 'ok' } }
$r = Find-GalleryModules -Name $names -RetryDelaySeconds 0, 0
Assert-That (($r.Unchecked.Count -eq 0) -and ($r.Resources.Count -eq 5)) 'same when the not-found wording changes'

Reset-Test
$script:Outcome = { 'ok' }
$r = Find-GalleryModules -Name ($names + $probe) -RetryDelaySeconds 0, 0
Assert-That (($r.Resources.Count -eq 6) -and ($r.Unchecked.Count -eq 0)) 'the probe module in the installed list comes back once, as a normal result'

# --- Get-GalleryFaultText ------------------------------------------------------------
Write-Section 'Get-GalleryFaultText'
$long = 'x' * 400
Assert-That ((Get-GalleryFaultText -Message $long).Length -eq 123) 'unrecognised text is cut to 120 characters and an ellipsis'
Assert-That ((Get-GalleryFaultText -Message 'short text') -eq 'short text') 'short text is left alone'

# --- Reporting -----------------------------------------------------------------------
Write-Section 'Reporting'

function New-Summary {
    @{
        StartTime = '2026-01-04T03:00:00+01:00'; EndTime = '2026-01-04T03:00:30+01:00'
        ModulesChecked = 0; ModulesUnchecked = 0; GalleryFault = $null; ModulesUpdated = 0
        ModulesFailed = @(); VersionsPruned = 0; PrunesFailed = @(); ExcludedModules = @()
        PinnedModules = @{}; PinsSatisfied = 0; PinsEnforced = 0; PinsFailed = @()
        PinsHoldingBack = @()
    }
}

$outage = New-Summary
$outage.ModulesUnchecked = 150
$outage.GalleryFault = 'No such host is known'

$f = Get-SummaryFailures -Summary $outage
Assert-That (($f.Total -eq 1) -and ($f.Lookups -eq 1)) "an outage that left 150 modules unchecked is one failure (total $($f.Total))"
Assert-That ($outage.ModulesUnchecked -eq 150) 'the module count itself stays in the summary'

$one = New-Summary
$one.ModulesUnchecked = 1
Assert-That ((Get-SummaryFailures -Summary $one).Total -eq 1) 'a single unanswered lookup is also one failure'

$f = Get-SummaryFailures -Summary (New-Summary)
Assert-That (($f.Total -eq 0) -and ($f.Lookups -eq 0)) 'a clean summary counts 0'

$mixed = New-Summary
$mixed.ModulesUnchecked = 3
$mixed.ModulesFailed += @{ Module = 'A'; Error = 'x' }
$mixed.ModulesFailed += @{ Module = 'B'; Error = 'y' }
$mixed.PrunesFailed += @{ Module = 'C'; Version = '1.0'; Error = 'z' }
$mixed.PinsFailed += @{ Module = 'D'; Version = '2.0'; Error = 'q' }
$f = Get-SummaryFailures -Summary $mixed
Assert-That (($f.Lookups -eq 1) -and ($f.Updates -eq 2) -and ($f.Prunes -eq 1) -and ($f.Pins -eq 1)) "breakdown: lookups $($f.Lookups), updates $($f.Updates), pins $($f.Pins), prunes $($f.Prunes)"
Assert-That ($f.Total -eq 5) "the total is the sum of the breakdown: $($f.Total)"

# A failure is a hashtable with several keys. One of them must count as 1, not as its
# number of keys
$single = New-Summary
$single.PinsFailed += @{ Module = 'D'; Version = '2.0'; Error = 'q' }
Assert-That ((Get-SummaryFailures -Summary $single).Total -eq 1) 'one failure with three keys counts as 1'

$older = @{ ModulesFailed = @(); PrunesFailed = @(); PinsFailed = @() }
Assert-That ((Get-SummaryFailures -Summary $older).Total -eq 0) 'a summary without the newer keys does not break the count'

$body = Format-HealthchecksBody -Summary $outage -Mode 'Full' -IsFailure
Assert-That ($body -like '*PSModuleMaintenance - Fail*') 'ping body: status is Fail'
Assert-That ($body -like '*Host: TEST-HOST *') 'ping body: names the machine'
Assert-That ($body -like '*Checked: 0 *') 'ping body: Checked is 0, not the number installed'
Assert-That ($body -like '*Issues: 1*') 'ping body: one issue'
Assert-That ($body -like '*lookup: 150 module(s) not checked: No such host is known*') 'ping body: names the lookup problem'

$cleanBody = Format-HealthchecksBody -Summary (New-Summary) -Mode 'Full'
Assert-That ($cleanBody -like '*Issues: none*') 'ping body: a clean run is unchanged'

# The toast is sent by starting Windows PowerShell. This stand-in takes its place
function powershell.exe { $global:LASTEXITCODE = 0 }

$script:LogLines = @()
Send-ToastNotification -Summary $outage
$toastLine = $script:LogLines | Where-Object { $_ -like '*Toast notification sent*' }
Assert-That ($toastLine -like '*Updated 0 modules, 150 not checked.*') "toast: $toastLine"
Assert-That ($toastLine -notlike '*No issues*') 'toast: does not say "No issues"'

$script:LogLines = @()
Send-ToastNotification -Summary (New-Summary)
$toastLine = $script:LogLines | Where-Object { $_ -like '*Toast notification sent*' }
Assert-That ($toastLine -like '*Updated 0 modules. Pruned 0 versions. No issues.*') "toast, clean run unchanged: $toastLine"

Complete-Tests
