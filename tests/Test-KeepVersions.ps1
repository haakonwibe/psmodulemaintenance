<#
.SYNOPSIS
    Tests for KeepVersions: keeping an older version line next to the newest version,
    and updating that line within itself.

.DESCRIPTION
    Covers the selector parser and matcher, Get-KeepVersionRange, the two pure planning
    functions Get-ModulePrunePlan and Get-KeptLinePlan, the two passes that log and act
    on them, and the wiring into Remove-OldModuleVersions and Update-AllModules.

    Everything that would touch the gallery or the machine is replaced by a stand-in:
    Get-PSResource, Uninstall-PSResource, Find-GalleryModules and
    Invoke-ModuleUpdateWithRetry. Nothing is installed or removed.
#>
param([switch]$Quiet)

. (Join-Path $PSScriptRoot 'TestHelpers.ps1')

$wanted = 'ConvertTo-NormalizedVersion', 'Get-ModuleVersionKey', 'ConvertFrom-PinnedVersionString',
          'Get-PinnedVersion', 'Test-IsPinnedVersion', 'ConvertFrom-KeepVersionSelector',
          'Test-KeepVersionMatch', 'Get-KeepVersionSelectors', 'Get-KeepVersionRange',
          'Get-ModulePrunePlan', 'Get-KeptLinePlan', 'Confirm-KeptModuleVersions',
          'Update-KeptVersionLines', 'Set-PinnedModuleVersions', 'Remove-OldModuleVersions',
          'Update-AllModules'
foreach ($name in $wanted) {
    . ([scriptblock]::Create((Get-ScriptFunctionText -Path $script:MainScript -Name $name)))
}

# --- Stand-ins and helpers -----------------------------------------------------------
$script:LogLines = @()
function Write-Log {
    param([Parameter(Position = 0)][string]$Message, [string]$Level = 'INFO')
    $script:LogLines += "[$Level] $Message"
}

function New-Resource {
    param([string]$Name = 'Contoso.Tools', [string]$Version, [string]$Prerelease)
    [PSCustomObject]@{
        Name              = $Name
        Version           = [version]$Version
        Prerelease        = $Prerelease
        InstalledLocation = "C:\Modules\$Name\$Version"
    }
}

function New-Selectors {
    param([string[]]$Text)
    @($Text | ForEach-Object { ConvertFrom-KeepVersionSelector -Value $_ })
}

function Get-VersionList {
    param($Resources)
    (@($Resources) | ForEach-Object {
        $text = $_.Version.ToString()
        if ($_.Prerelease) { $text += "-$($_.Prerelease)" }
        $text
    }) -join ', '
}

function Reset-State {
    param([hashtable]$Keep = @{}, [hashtable]$Pins = @{}, [string[]]$Excluded = @())

    $keepTable = @{}
    foreach ($key in $Keep.Keys) {
        $keepTable[$key] = @(New-Selectors $Keep[$key])
    }
    $pinTable = @{}
    foreach ($key in $Pins.Keys) {
        $pinTable[$key] = ConvertFrom-PinnedVersionString $Pins[$key]
    }

    $script:Config = @{
        ExcludedModules            = $Excluded
        PinnedModules              = $pinTable
        KeepVersions               = $keepTable
        TrustPSGallery             = $true
        ModuleUpdateTimeoutSeconds = 600
    }
    $script:Summary = @{
        ModulesChecked = 0; ModulesUnchecked = 0; GalleryFault = $null; ModulesUpdated = 0
        ModulesFailed = @(); VersionsPruned = 0; PrunesFailed = @(); ExcludedModules = @()
        PinnedModules = @{}; PinsSatisfied = 0; PinsEnforced = 0; PinsFailed = @()
        PinsHoldingBack = @()
        KeepVersions = @{}; KeepVersionsMatched = @(); KeepVersionsUnmatched = @()
        KeepLinesUpdated = @(); KeepLinesUnchecked = @()
    }
    $script:LogLines = @()
    $script:LookupCalls = @()
    $script:InstallCalls = @()
    $script:UninstallCalls = @()
    $script:Installed = @()
    $script:MainAnswer = { param($Names) New-Answer }
    $script:LineAnswer = { param($Name, $Range) New-Answer }
    $script:InstallBehaviour = { }
}

function New-Answer {
    param($Resources = @(), [string[]]$Unchecked = @(), [string]$Fault, [string]$Detail)
    [PSCustomObject]@{ Resources = @($Resources); Unchecked = @($Unchecked); Fault = $Fault; Detail = $Detail }
}

function Find-GalleryModules {
    param([string[]]$Name, [string]$Version, [int[]]$RetryDelaySeconds)

    $script:LookupCalls += , @{ Name = $Name; Version = $Version }
    if ($Version) {
        return (& $script:LineAnswer $Name[0] $Version)
    }
    return (& $script:MainAnswer $Name)
}

function Invoke-ModuleUpdateWithRetry {
    [CmdletBinding()]
    param(
        [string]$Name, [string]$Scope, [bool]$TrustRepository = $true,
        [int]$TimeoutSeconds = 600, [string]$Version, [switch]$Prerelease
    )
    $script:InstallCalls += , @{ Name = $Name; Version = $Version; Scope = $Scope }
    & $script:InstallBehaviour $Name $Version
}

function Get-PSResource {
    param([string]$Scope)
    $script:Installed
}

function Uninstall-PSResource {
    [CmdletBinding()]
    param([string]$Name, $Version, [switch]$SkipDependencyCheck, [string]$Scope)
    $script:UninstallCalls += "$Name $Version"
}

function Test-OneDrivePath {
    param([string]$Path)
    $false
}

function Get-LogLines {
    param([string]$Level)
    , @($script:LogLines | Where-Object { $_ -like "``[$Level``]*" })
}

# The wiring tests call functions that remove and install. Make sure the stand-ins are
# what they reach, not the real cmdlets
$safe = ((Get-Command Uninstall-PSResource).CommandType -eq 'Function') -and
        ((Get-Command Get-PSResource).CommandType -eq 'Function') -and
        ((Get-Command Invoke-ModuleUpdateWithRetry).CommandType -eq 'Function') -and
        ((Get-Command Find-GalleryModules).CommandType -eq 'Function')
if (-not $safe) {
    throw 'A stand-in is not in place. Stopping before anything real can be called.'
}

# --- Selector parser -----------------------------------------------------------------
Write-Section 'ConvertFrom-KeepVersionSelector'

$valid = [ordered]@{ '5' = '5'; '5.7' = '5.7'; '5.7.1' = '5.7.1'; '5.7.1.0' = '5.7.1.0'; ' 5 ' = '5'; '05' = '5' }
foreach ($case in $valid.GetEnumerator()) {
    $selector = ConvertFrom-KeepVersionSelector -Value $case.Key
    Assert-That (($null -ne $selector) -and ($selector.Key -eq $case.Value)) "'$($case.Key)' is the line $($case.Value)"
}

foreach ($text in '', '5.', '.5', '5.*', '5-rc1', '1.2.3.4.5', 'latest', '99999999999', '5,7') {
    Assert-That ($null -eq (ConvertFrom-KeepVersionSelector -Value $text)) "'$text' is rejected"
}
Assert-That ($null -eq (ConvertFrom-KeepVersionSelector -Value 5)) 'the number 5 is rejected, only text is a selector'
Assert-That ($null -eq (ConvertFrom-KeepVersionSelector -Value 5.1)) 'the number 5.1 is rejected'
Assert-That ($null -eq (ConvertFrom-KeepVersionSelector -Value $null)) 'nothing at all is rejected'

$selector = ConvertFrom-KeepVersionSelector -Value '5.7'
Assert-That (($selector.Parts.Count -eq 2) -and ($selector.Parts[0] -eq 5) -and ($selector.Parts[1] -eq 7)) 'the parts are numbers'
Assert-That ($selector.Requested -eq '5.7') 'the text as written is kept for the log'

# --- Matching ------------------------------------------------------------------------
Write-Section 'Test-KeepVersionMatch'

$cases = @(
    @{ Selector = '5';       Version = '5.7.1';   Expected = $true }
    @{ Selector = '5';       Version = '5.0';     Expected = $true }
    @{ Selector = '5';       Version = '6.0.0';   Expected = $false }
    @{ Selector = '1';       Version = '12.4.0';  Expected = $false }
    @{ Selector = '15';      Version = '1.5.0';   Expected = $false }
    @{ Selector = '5.7';     Version = '5.7.1';   Expected = $true }
    @{ Selector = '5.7';     Version = '5.70.1';  Expected = $false }
    @{ Selector = '5.7';     Version = '5.8.0';   Expected = $false }
    @{ Selector = '5.7.1';   Version = '5.7.1';   Expected = $true }
    @{ Selector = '5.7.1';   Version = '5.7.1.0'; Expected = $true }
    @{ Selector = '5.7.1';   Version = '5.7.1.4'; Expected = $true }
    @{ Selector = '5.7.1';   Version = '5.7.10';  Expected = $false }
    @{ Selector = '5.7.0';   Version = '5.7';     Expected = $true }
    @{ Selector = '5.7.1.0'; Version = '5.7.1';   Expected = $true }
    @{ Selector = '5.7.1.0'; Version = '5.7.1.1'; Expected = $false }
)
foreach ($case in $cases) {
    $selector = ConvertFrom-KeepVersionSelector -Value $case.Selector
    $actual = Test-KeepVersionMatch -Version $case.Version -Selector $selector
    $verb = if ($case.Expected) { 'matches' } else { 'does not match' }
    Assert-That ($actual -eq $case.Expected) "'$($case.Selector)' $verb $($case.Version)"
}

# --- Range ---------------------------------------------------------------------------
Write-Section 'Get-KeepVersionRange'

Assert-That ((Get-KeepVersionRange -Selectors (New-Selectors '5')) -eq '[5.0.0.0, 6.0.0.0)') "'5' is [5.0.0.0, 6.0.0.0)"
Assert-That ((Get-KeepVersionRange -Selectors (New-Selectors '5.7')) -eq '[5.7.0.0, 5.8.0.0)') "'5.7' is [5.7.0.0, 5.8.0.0)"
Assert-That ((Get-KeepVersionRange -Selectors (New-Selectors '5.7.1')) -eq '[5.7.1.0, 5.7.2.0)') "'5.7.1' is [5.7.1.0, 5.7.2.0)"
Assert-That ((Get-KeepVersionRange -Selectors (New-Selectors '3', '5.7')) -eq '[3.0.0.0, 5.8.0.0)') 'two lines share one range that spans both'
Assert-That ((Get-KeepVersionRange -Selectors (New-Selectors '5.7', '3')) -eq '[3.0.0.0, 5.8.0.0)') 'whatever order they come in'

$selectors = New-Selectors '5', '5.7'
$null = Get-KeepVersionRange -Selectors $selectors
Assert-That ((($selectors[0].Parts -join '.') -eq '5') -and (($selectors[1].Parts -join '.') -eq '5.7')) 'building the range leaves the selectors as they were'

# --- What stays and what goes --------------------------------------------------------
Write-Section 'Get-ModulePrunePlan'

$four = @((New-Resource -Version '5.6.0'), (New-Resource -Version '6.1.0'), (New-Resource -Version '5.7.0'), (New-Resource -Version '5.7.1'))

$plan = Get-ModulePrunePlan -Resources $four -Selectors @()
Assert-That ((Get-VersionList $plan.Keep) -eq '6.1.0') 'no selectors: only the newest stays'
Assert-That ((Get-VersionList $plan.Remove) -eq '5.7.1, 5.7.0, 5.6.0') 'and the rest goes'

$plan = Get-ModulePrunePlan -Resources $four -Selectors $null
Assert-That ((Get-VersionList $plan.Keep) -eq '6.1.0') 'no selector list at all behaves the same'

$plan = Get-ModulePrunePlan -Resources $four -Selectors (New-Selectors '5')
Assert-That ((Get-VersionList $plan.Keep) -eq '6.1.0, 5.7.1') "'5': the newest version and the newest 5.x stay"
Assert-That ((Get-VersionList $plan.Remove) -eq '5.7.0, 5.6.0') 'older versions in the kept line go'
Assert-That (($plan.Matched.Count -eq 1) -and ($plan.Matched[0].Version -eq '5.7.1')) 'the match is reported'

$plan = Get-ModulePrunePlan -Resources $four -Selectors (New-Selectors '5.6', '5.7')
Assert-That ((Get-VersionList $plan.Keep) -eq '6.1.0, 5.7.1, 5.6.0') "'5.6' and '5.7': one version of each line stays"
Assert-That ((Get-VersionList $plan.Remove) -eq '5.7.0') 'only 5.7.0 goes'

$plan = Get-ModulePrunePlan -Resources $four -Selectors (New-Selectors '6')
Assert-That ((Get-VersionList $plan.Keep) -eq '6.1.0') "'6' names the newest line: nothing extra stays"
Assert-That (($plan.Matched.Count -eq 1) -and ($plan.Unmatched.Count -eq 0)) 'and it still counts as matched'

$plan = Get-ModulePrunePlan -Resources $four -Selectors (New-Selectors '5', '5.7')
Assert-That ((Get-VersionList $plan.Keep) -eq '6.1.0, 5.7.1') "'5' and '5.7' land on the same version, which stays once"
Assert-That ($plan.Matched.Count -eq 2) 'both selectors are reported as matched'

$plan = Get-ModulePrunePlan -Resources $four -Selectors (New-Selectors '4')
Assert-That (($plan.Unmatched -join ',') -eq '4') "'4' matches nothing and is reported"
Assert-That ((Get-VersionList $plan.Remove) -eq '5.7.1, 5.7.0, 5.6.0') 'pruning goes ahead as if it were not there'

$plan = Get-ModulePrunePlan -Resources @((New-Resource -Version '6.1.0')) -Selectors (New-Selectors '5')
Assert-That (($plan.Remove.Count -eq 0) -and (($plan.Unmatched -join ',') -eq '5')) 'a single installed version: nothing goes, the selector is reported'

$plan = Get-ModulePrunePlan -Resources @() -Selectors (New-Selectors '5')
Assert-That (($plan.Keep.Count -eq 0) -and ($plan.Remove.Count -eq 0) -and (($plan.Unmatched -join ',') -eq '5')) 'nothing installed: nothing to decide'

$plan = Get-ModulePrunePlan -Resources @((New-Resource -Version '12.4.0'), (New-Resource -Version '13.0.0')) -Selectors (New-Selectors '1')
Assert-That (((Get-VersionList $plan.Remove) -eq '12.4.0') -and (($plan.Unmatched -join ',') -eq '1')) "'1' does not protect 12.4.0"

$pin = ConvertFrom-PinnedVersionString '6.0.0'
$pinned = @((New-Resource -Version '6.1.0'), (New-Resource -Version '6.0.0'), (New-Resource -Version '5.7.1'), (New-Resource -Version '5.7.0'))
$plan = Get-ModulePrunePlan -Resources $pinned -Pin $pin -Selectors (New-Selectors '5')
Assert-That ((Get-VersionList $plan.Keep) -eq '6.0.0, 5.7.1') 'pinned and kept: the pin and the newest of the line stay'
Assert-That ((Get-VersionList $plan.Remove) -eq '6.1.0, 5.7.0') 'a version newer than the pin still goes'

$plan = Get-ModulePrunePlan -Resources $pinned -Pin $pin -Selectors @()
Assert-That ((Get-VersionList $plan.Keep) -eq '6.0.0') 'pinned without selectors: only the pin stays, as before'

$plan = Get-ModulePrunePlan -Resources $pinned -Pin (ConvertFrom-PinnedVersionString '7.0.0') -Selectors (New-Selectors '5')
Assert-That ($plan.PinMissing) 'a pin that is not installed is flagged'
Assert-That (($plan.Remove.Count -eq 0) -and ($plan.Keep.Count -eq 4)) 'and then nothing at all goes'

$plan = Get-ModulePrunePlan -Resources $pinned -Pin (ConvertFrom-PinnedVersionString '5.7.0') -Selectors (New-Selectors '5')
Assert-That ((Get-VersionList $plan.Keep) -eq '5.7.1, 5.7.0') 'a pin inside the kept line: the pin and the newest of the line stay'
Assert-That ((Get-VersionList $plan.Remove) -eq '6.1.0, 6.0.0') 'everything outside goes'

$withPrerelease = @((New-Resource -Version '6.1.0'), (New-Resource -Version '5.8.0' -Prerelease 'beta1'), (New-Resource -Version '5.7.1'))
$plan = Get-ModulePrunePlan -Resources $withPrerelease -Selectors (New-Selectors '5')
Assert-That ((Get-VersionList $plan.Keep) -eq '6.1.0, 5.8.0-beta1') 'a prerelease that is the newest of its line is the one that stays'
Assert-That ($plan.Matched[0].Version -eq '5.8.0-beta1') 'and is shown with its label'

$mixed = @((New-Resource -Version '6.1.0'), (New-Resource -Version '5.7.1.0'), (New-Resource -Version '5.7'))
$plan = Get-ModulePrunePlan -Resources $mixed -Selectors (New-Selectors '5.7.1')
Assert-That ((Get-VersionList $plan.Keep) -eq '6.1.0, 5.7.1.0') 'three and four part versions compare by their numbers'

$twice = @((New-Resource -Version '6.1.0'), (New-Resource -Version '6.1.0'), (New-Resource -Version '6.0.0'))
$plan = Get-ModulePrunePlan -Resources $twice -Selectors @()
Assert-That (($plan.Keep.Count -eq 2) -and ((Get-VersionList $plan.Remove) -eq '6.0.0')) 'a version that stays is never removed, even when it is listed twice'

# --- Which lines need a lookup -------------------------------------------------------
Write-Section 'Get-KeptLinePlan'

$both = @((New-Resource -Version '6.1.0'), (New-Resource -Version '5.7.1'))

$entry = @(Get-KeptLinePlan -Resources $both -Selectors (New-Selectors '4') -GalleryNewest '6.2.0')[0]
Assert-That (($entry.Action -eq 'Skip') -and ($entry.Reason -eq 'NotInstalled')) 'a line with nothing installed is left alone'

$entry = @(Get-KeptLinePlan -Resources $both -Selectors (New-Selectors '5.7.1.0') -GalleryNewest '6.2.0')[0]
Assert-That (($entry.Action -eq 'Skip') -and ($entry.Reason -eq 'Exact')) 'four numbers name one exact version, which is never updated'

$entry = @(Get-KeptLinePlan -Resources $both -Selectors (New-Selectors '5') -Pin (ConvertFrom-PinnedVersionString '5.7.1') -GalleryNewest '6.2.0')[0]
Assert-That (($entry.Action -eq 'Skip') -and ($entry.Reason -eq 'PinInLine')) 'a pin inside the line decides'

$entry = @(Get-KeptLinePlan -Resources $both -Selectors (New-Selectors '6') -GalleryNewest '6.2.0')[0]
Assert-That (($entry.Action -eq 'Skip') -and ($entry.Reason -eq 'MainUpdate')) 'the line holding the newest release is left to the normal update'

$entry = @(Get-KeptLinePlan -Resources $both -Selectors (New-Selectors '5') -GalleryNewest '6.2.0')[0]
Assert-That ($entry.Action -eq 'Lookup') 'an older line has to be looked up'
Assert-That ($entry.Installed.Version -eq [version]'5.7.1') 'and knows what is installed in it'

# With only the old line installed it looks like the newest line. Judged by that alone
# it would be skipped and stay a release behind
$entry = @(Get-KeptLinePlan -Resources @((New-Resource -Version '5.7.1')) -Selectors (New-Selectors '5') -GalleryNewest '6.2.0')[0]
Assert-That ($entry.Action -eq 'Lookup') 'only the old line installed: still looked up'

# A prerelease of the next line is installed. The normal update compares against that,
# so it will never update line 6
$ahead = @((New-Resource -Version '7.0.0' -Prerelease 'preview1'), (New-Resource -Version '6.1.0'))
$entry = @(Get-KeptLinePlan -Resources $ahead -Selectors (New-Selectors '6') -GalleryNewest '6.2.0')[0]
Assert-That (($entry.Action -eq 'Compare') -and ($entry.Target -eq [version]'6.2.0')) 'newest release in the line but something newer installed: compared directly'

# The normal update drops pinned modules, so it cannot cover a line of one
$entry = @(Get-KeptLinePlan -Resources $both -Selectors (New-Selectors '6') -Pin (ConvertFrom-PinnedVersionString '5.7.1') -GalleryNewest '6.2.0')[0]
Assert-That (($entry.Action -eq 'Compare') -and ($entry.Target -eq [version]'6.2.0')) 'pinned module, line holds the newest release: compared directly'

$plan = @(Get-KeptLinePlan -Resources $both -Selectors (New-Selectors '5', '6', '4') -GalleryNewest '6.2.0')
Assert-That ((($plan | ForEach-Object { $_.Action }) -join ',') -eq 'Lookup,Skip,Skip') 'one entry per selector, in order'

Assert-That (@(Get-KeptLinePlan -Resources $both -Selectors @() -GalleryNewest '6.2.0').Count -eq 0) 'no selectors, no entries'

# --- The pass that says what is kept -------------------------------------------------
Write-Section 'Confirm-KeptModuleVersions'

Reset-State -Keep @{ 'Contoso.Tools' = @('5') }
Confirm-KeptModuleVersions -InstalledResources @((New-Resource -Version '5.7.1'))
Assert-That ($script:LogLines -contains '[INFO] Keeping Contoso.Tools v5.7.1 (KeepVersions: 5)') 'a module with a single version is looked at too'
Assert-That ($script:Summary.KeepVersionsMatched.Count -eq 1) 'and recorded in the summary'

Reset-State -Keep @{ 'Contoso.Tools' = @('5', '5.7', '4') }
Confirm-KeptModuleVersions -InstalledResources $four
Assert-That ($script:LogLines -contains '[INFO] Keeping Contoso.Tools v5.7.1 (KeepVersions: 5, 5.7)') 'two selectors on one version share a line in the log'
Assert-That ($script:LogLines -contains "[WARN] KeepVersions: no installed version of Contoso.Tools matches '4'. Nothing is kept or updated for it - a line is only maintained once a version of it is installed") 'a selector that matches nothing gets its warning, word for word'
Assert-That (($script:Summary.KeepVersionsMatched.Count -eq 2) -and ($script:Summary.KeepVersionsUnmatched.Count -eq 1)) 'the summary holds two matches and one miss'

Reset-State -Keep @{ 'Contoso.Tools' = @('5', '4') }
Confirm-KeptModuleVersions -InstalledResources @((New-Resource -Name 'Fabrikam.Core' -Version '1.0.0'))
Assert-That ($script:LogLines -contains "[WARN] KeepVersions: Contoso.Tools is not installed, so nothing is kept or updated for '5', '4'") 'a module that is not installed gets its warning, word for word'
Assert-That ($script:Summary.KeepVersionsUnmatched.Count -eq 2) 'with one entry per selector in the summary'

Reset-State -Keep @{ 'contoso.tools' = @('5') }
Confirm-KeptModuleVersions -InstalledResources $four
Assert-That ($script:LogLines -contains '[INFO] Keeping Contoso.Tools v5.7.1 (KeepVersions: 5)') 'the name in the config may be cased differently, the log uses the installed name'

Reset-State
Confirm-KeptModuleVersions -InstalledResources $four
Assert-That ($script:LogLines.Count -eq 0) 'without KeepVersions nothing is logged'

# --- The pass that updates the lines -------------------------------------------------
Write-Section 'Update-KeptVersionLines'

$galleryNewest = @((New-Resource -Version '6.2.0'))
$installedBoth = @((New-Resource -Version '6.1.0'), (New-Resource -Version '5.7.0'))
$inLineFive = { param($Name, $Range) New-Answer -Resources @((New-Resource -Name $Name -Version '5.6.0'), (New-Resource -Name $Name -Version '5.7.1'), (New-Resource -Name $Name -Version '5.7.0')) }

Reset-State -Keep @{ 'Contoso.Tools' = @('5') }
$script:LineAnswer = $inLineFive
Update-KeptVersionLines -InstalledResources $installedBoth -GalleryResources $galleryNewest -Scope 'AllUsers'
Assert-That (($script:InstallCalls.Count -eq 1) -and ($script:InstallCalls[0].Version -eq '5.7.1')) 'a line that is behind: the newest release in it is installed, as an exact version'
Assert-That ($script:InstallCalls[0].Scope -eq 'AllUsers') 'into the scope it was given'
Assert-That ($script:LookupCalls[0].Version -eq '[5.0.0.0, 6.0.0.0)') 'the gallery was asked with a range, not a wildcard'
Assert-That ($script:LogLines -contains '[INFO] Updating kept line 5 of Contoso.Tools: 5.7.0 -> 5.7.1') 'the update is announced'
Assert-That (@($script:LogLines | Where-Object { $_ -like '`[SUCCESS`] Updated kept line 5 of Contoso.Tools to 5.7.1 (took *s)' }).Count -eq 1) 'and confirmed'
Assert-That ($script:LogLines -contains '[INFO] Kept line updates complete. Updated: 1, Unsuccessful: 0, Not checked: 0') 'the closing line counts it'
Assert-That ($script:Summary.ModulesUpdated -eq 1) 'it counts as an update'
$record = $script:Summary.KeepLinesUpdated[0]
Assert-That (($record.Module -eq 'Contoso.Tools') -and ($record.Line -eq '5') -and ($record.From -eq '5.7.0') -and ($record.To -eq '5.7.1')) 'the summary says from what to what'

Reset-State -Keep @{ 'Contoso.Tools' = @('5') }
$script:LineAnswer = $inLineFive
Update-KeptVersionLines -InstalledResources @((New-Resource -Version '6.1.0'), (New-Resource -Version '5.7.1')) -GalleryResources $galleryNewest
Assert-That ($script:InstallCalls.Count -eq 0) 'a line that is current: nothing installed'
Assert-That ($script:LogLines -contains '[INFO] Kept line 5 of Contoso.Tools is up to date at 5.7.1') 'and the log says so'

Reset-State -Keep @{ 'Contoso.Tools' = @('5') }
$script:LineAnswer = $inLineFive
Update-KeptVersionLines -InstalledResources @((New-Resource -Version '6.1.0'), (New-Resource -Version '5.8.0')) -GalleryResources $galleryNewest
Assert-That ($script:InstallCalls.Count -eq 0) 'a line that is ahead of the gallery: nothing installed'
Assert-That ($script:LogLines -contains '[INFO] Kept line 5 of Contoso.Tools is at 5.8.0, ahead of 5.7.1 on PSGallery - left as it is') 'and the log says so'

Reset-State -Keep @{ 'Contoso.Tools' = @('5') }
Update-KeptVersionLines -InstalledResources $installedBoth -GalleryResources $galleryNewest
Assert-That ($script:InstallCalls.Count -eq 0) 'a line that no longer exists on the gallery: nothing installed'
Assert-That ($script:LogLines -contains '[INFO] Kept line 5 of Contoso.Tools has no release on PSGallery - nothing to update') 'and the log says so'
Assert-That ($script:Summary.KeepLinesUnchecked.Count -eq 0) 'an empty answer is an answer, not a lookup that went wrong'

Reset-State -Keep @{ 'Contoso.Tools' = @('5') }
$script:LineAnswer = { param($Name, $Range) New-Answer -Resources @((New-Resource -Name $Name -Version '5.9.0' -Prerelease 'rc1'), (New-Resource -Name $Name -Version '5.7.1')) }
Update-KeptVersionLines -InstalledResources $installedBoth -GalleryResources $galleryNewest
Assert-That (($script:InstallCalls.Count -eq 1) -and ($script:InstallCalls[0].Version -eq '5.7.1')) 'a prerelease is never the target'

Reset-State -Keep @{ 'Contoso.Tools' = @('5', '5.7') }
$script:LineAnswer = { param($Name, $Range) New-Answer -Resources @((New-Resource -Name $Name -Version '5.9.1'), (New-Resource -Name $Name -Version '5.7.1')) }
Update-KeptVersionLines -InstalledResources $installedBoth -GalleryResources $galleryNewest
Assert-That ((($script:InstallCalls | ForEach-Object { $_.Version } | Sort-Object) -join ',') -eq '5.7.1,5.9.1') 'overlapping lines each get their own target'
Assert-That ($script:LookupCalls.Count -eq 1) 'from a single lookup'
Assert-That ($script:LookupCalls[0].Version -eq '[5.0.0.0, 6.0.0.0)') 'whose range spans both'

Reset-State -Keep @{ 'Contoso.Tools' = @('5', '5.9') }
$script:LineAnswer = { param($Name, $Range) New-Answer -Resources @((New-Resource -Name $Name -Version '5.9.1')) }
Update-KeptVersionLines -InstalledResources @((New-Resource -Version '6.1.0'), (New-Resource -Version '5.9.0')) -GalleryResources $galleryNewest
Assert-That ($script:InstallCalls.Count -eq 1) 'two lines with the same target: installed once'
Assert-That ($script:LogLines -contains '[INFO] Updating kept line 5, 5.9 of Contoso.Tools: 5.9.0 -> 5.9.1') 'one line in the log names both'
Assert-That ($script:Summary.KeepLinesUpdated.Count -eq 2) 'the summary records both'

# The gallery stops answering
Reset-State -Keep @{ 'Contoso.Tools' = @('5'); 'Fabrikam.Core' = @('2') }
$script:LineAnswer = { param($Name, $Range) New-Answer -Unchecked @($Name) -Fault 'No such host is known' -Detail 'No such host is known. (proxy.invalid:8080)' }
$twoModules = @((New-Resource -Version '6.1.0'), (New-Resource -Version '5.7.0'), (New-Resource -Name 'Fabrikam.Core' -Version '3.0.0'), (New-Resource -Name 'Fabrikam.Core' -Version '2.1.0'))
$twoNewest = @((New-Resource -Version '6.2.0'), (New-Resource -Name 'Fabrikam.Core' -Version '3.0.0'))
Update-KeptVersionLines -InstalledResources $twoModules -GalleryResources $twoNewest
Assert-That ($script:LookupCalls.Count -eq 1) 'once the gallery gives no answer it is not asked again'
Assert-That ($script:InstallCalls.Count -eq 0) 'nothing installed'
Assert-That ((Get-LogLines 'ERROR').Count -eq 1) 'the first line that could not be checked is an ERROR'
Assert-That (@((Get-LogLines 'ERROR') | Where-Object { $_ -like '*Could not check kept line(s) * for updates: No such host is known. (proxy.invalid:8080)' }).Count -eq 1) 'with the full detail'
Assert-That (@((Get-LogLines 'WARN') | Where-Object { $_ -like '*not checked: PSGallery stopped answering earlier in this run' }).Count -eq 1) 'the next one is noted as not checked'
Assert-That ($script:Summary.KeepLinesUnchecked.Count -eq 2) 'both are in the summary'
Assert-That ($script:Summary.KeepLinesUnchecked[1].Fault -eq 'No such host is known') 'with the reason'
Assert-That ($script:LogLines -contains '[INFO] Kept line updates complete. Updated: 0, Unsuccessful: 0, Not checked: 2') 'the closing line counts them'

Reset-State -Keep @{ 'Contoso.Tools' = @('5') }
$script:LineAnswer = $inLineFive
Update-KeptVersionLines -InstalledResources $installedBoth -GalleryResources $galleryNewest -Unchecked @('contoso.tools')
Assert-That (($script:LookupCalls.Count -eq 0) -and ($script:InstallCalls.Count -eq 0)) 'a module the main lookup got no answer for is not asked about again'
Assert-That ($script:LogLines -contains '[WARN] Kept lines of Contoso.Tools not checked: PSGallery gave no answer for this module') 'and the log says so'
Assert-That ($script:Summary.KeepLinesUnchecked.Count -eq 0) 'it was counted by the main lookup already'

Reset-State -Keep @{ 'Contoso.Tools' = @('5') }
Update-KeptVersionLines -InstalledResources $installedBoth -GalleryResources @()
Assert-That ($script:LogLines -contains '[INFO] Kept lines of Contoso.Tools skipped: the module is not on PSGallery') 'a module that is not on the gallery is skipped'
Assert-That ($script:LookupCalls.Count -eq 0) 'without a lookup'

Reset-State -Keep @{ 'Contoso.Tools' = @('5') }
Update-KeptVersionLines -InstalledResources @((New-Resource -Name 'Fabrikam.Core' -Version '1.0.0')) -GalleryResources $galleryNewest
Assert-That ($script:LogLines -contains '[INFO] Kept lines of Contoso.Tools skipped: the module is not installed') 'a module that is not installed is skipped'

Reset-State -Keep @{ 'Contoso.Tools' = @('4') }
Update-KeptVersionLines -InstalledResources $installedBoth -GalleryResources $galleryNewest
Assert-That ($script:LogLines -contains '[INFO] Kept line 4 of Contoso.Tools has no installed version - nothing to update') 'a line with nothing installed is never brought in'
Assert-That (($script:LookupCalls.Count -eq 0) -and ($script:InstallCalls.Count -eq 0)) 'no lookup, no install'

Reset-State -Keep @{ 'contoso.tools' = @('5') }
$script:LineAnswer = $inLineFive
Update-KeptVersionLines -InstalledResources $installedBoth -GalleryResources $galleryNewest
Assert-That (($script:InstallCalls.Count -eq 1) -and ($script:InstallCalls[0].Name -eq 'Contoso.Tools')) 'the config name may be cased differently, the installed name is used'

# Pins
Reset-State -Keep @{ 'Contoso.Tools' = @('5') } -Pins @{ 'Contoso.Tools' = '5.7.0' }
$script:LineAnswer = $inLineFive
Update-KeptVersionLines -InstalledResources $installedBoth -GalleryResources $galleryNewest
Assert-That ($script:InstallCalls.Count -eq 0) 'a pin inside the line: the line is not updated past it'
Assert-That ($script:LogLines -contains '[INFO] Kept line 5 of Contoso.Tools holds the pinned version - the pin decides, no line update') 'and the log says so'

Reset-State -Keep @{ 'Contoso.Tools' = @('5') } -Pins @{ 'Contoso.Tools' = '6.1.0' }
$script:LineAnswer = $inLineFive
Update-KeptVersionLines -InstalledResources $installedBoth -GalleryResources $galleryNewest
Assert-That (($script:InstallCalls.Count -eq 1) -and ($script:InstallCalls[0].Version -eq '5.7.1')) 'a pin outside the line: the line updates as usual'

Reset-State -Keep @{ 'Contoso.Tools' = @('6') } -Pins @{ 'Contoso.Tools' = '5.7.0' }
Update-KeptVersionLines -InstalledResources $installedBoth -GalleryResources $galleryNewest
Assert-That (($script:InstallCalls.Count -eq 1) -and ($script:InstallCalls[0].Version -eq '6.2.0')) 'a pinned module whose kept line holds the newest release: the line update brings it in'
Assert-That ($script:LookupCalls.Count -eq 0) 'without another lookup, the main one already had the answer'

Reset-State -Keep @{ 'Contoso.Tools' = @('6') }
Update-KeptVersionLines -InstalledResources $installedBoth -GalleryResources $galleryNewest
Assert-That ($script:InstallCalls.Count -eq 0) 'the line holding the newest release is not done twice'
Assert-That ($script:LogLines -contains '[INFO] Kept line 6 of Contoso.Tools holds the newest release - covered by the normal update') 'and the log says why'

# An update that does not succeed
Reset-State -Keep @{ 'Contoso.Tools' = @('5') }
$script:LineAnswer = $inLineFive
$script:InstallBehaviour = { throw 'Access to the path is denied.' }
Update-KeptVersionLines -InstalledResources $installedBoth -GalleryResources $galleryNewest
Assert-That ($script:LogLines -contains '[ERROR] Failed to update kept line 5 of Contoso.Tools to 5.7.1: Access to the path is denied.') 'an install that throws is an ERROR'
$failure = $script:Summary.ModulesFailed[0]
Assert-That (($script:Summary.ModulesFailed.Count -eq 1) -and ($failure.Module -eq 'Contoso.Tools') -and ($failure.Line -eq '5') -and ($failure.Version -eq '5.7.1')) 'and lands with the failed updates, naming the line'
Assert-That (($script:Summary.ModulesUpdated -eq 0) -and ($script:Summary.KeepLinesUpdated.Count -eq 0)) 'nothing is counted as updated'
Assert-That ($script:LogLines -contains '[INFO] Kept line updates complete. Updated: 0, Unsuccessful: 1, Not checked: 0') 'the closing line counts it'

Reset-State -Keep @{ 'Contoso.Tools' = @('5') }
$script:LineAnswer = $inLineFive
$script:InstallBehaviour = { throw [System.TimeoutException]::new('Operation timed out after 605s (limit 600s)') }
Update-KeptVersionLines -InstalledResources $installedBoth -GalleryResources $galleryNewest
Assert-That ($script:LogLines -contains '[ERROR] Timed out updating kept line 5 of Contoso.Tools to 5.7.1 after 600s - skipping') 'a timeout has its own wording'
Assert-That ($script:Summary.ModulesFailed.Count -eq 1) 'and counts the same'

Reset-State -Keep @{ 'Contoso.Tools' = @('5') }
$script:LineAnswer = $inLineFive
Update-KeptVersionLines -InstalledResources $installedBoth -GalleryResources $galleryNewest -WhatIf
Assert-That ($script:InstallCalls.Count -eq 0) '-WhatIf installs nothing'
Assert-That ($script:LookupCalls.Count -eq 1) 'but still asks the gallery, which changes nothing'
Assert-That (($script:Summary.ModulesUpdated -eq 0) -and ($script:Summary.KeepLinesUpdated.Count -eq 0)) 'and records nothing as updated'

Reset-State
Update-KeptVersionLines -InstalledResources $installedBoth -GalleryResources $galleryNewest
Assert-That (($script:LogLines.Count -eq 0) -and ($script:LookupCalls.Count -eq 0)) 'without KeepVersions nothing happens at all'

# --- Wording that a log viewer would colour wrongly ----------------------------------
Write-Section 'Log wording'

# CMTrace colours a line by the words in it, whatever its level. Run every branch that
# logs at INFO or SUCCESS and look at what it wrote
Reset-State -Keep @{ 'Contoso.Tools' = @('5', '5.7', '4', '6', '5.7.0.0'); 'Fabrikam.Core' = @('2'); 'Northwind.Data' = @('1') }
$script:LineAnswer = $inLineFive
$many = @((New-Resource -Version '6.1.0'), (New-Resource -Version '5.7.0'), (New-Resource -Name 'Fabrikam.Core' -Version '2.0.0'))
Update-KeptVersionLines -InstalledResources $many -GalleryResources $galleryNewest
Confirm-KeptModuleVersions -InstalledResources $many
$calmLines = @($script:LogLines | Where-Object { ($_ -like '`[INFO`]*') -or ($_ -like '`[SUCCESS`]*') })
Assert-That ($calmLines.Count -ge 8) "a good number of INFO and SUCCESS lines were produced ($($calmLines.Count))"
Assert-That (@($calmLines | Where-Object { $_ -match '(?i)error|fail|warn' }).Count -eq 0) 'none of them holds a word that would colour it as a problem'

# --- Wiring into the prune phase -----------------------------------------------------
Write-Section 'Remove-OldModuleVersions'

$machine = @(
    (New-Resource -Version '6.1.0'), (New-Resource -Version '5.7.1'), (New-Resource -Version '5.6.0')
    (New-Resource -Name 'Fabrikam.Core' -Version '2.0.0'), (New-Resource -Name 'Fabrikam.Core' -Version '1.0.0')
    (New-Resource -Name 'Northwind.Data' -Version '3.0.0')
)

Reset-State -Keep @{ 'Contoso.Tools' = @('5') }
$script:Installed = $machine
Remove-OldModuleVersions
Assert-That ((($script:UninstallCalls | Sort-Object) -join '; ') -eq 'Contoso.Tools 5.6.0; Fabrikam.Core 1.0.0') "with KeepVersions: $($script:UninstallCalls -join '; ')"
Assert-That ($script:Summary.VersionsPruned -eq 2) 'two versions pruned'
Assert-That ($script:LogLines -contains '[INFO] Keeping Contoso.Tools v5.7.1 (KeepVersions: 5)') 'the kept version is named before anything is removed'
$keepingAt = [array]::IndexOf($script:LogLines, '[INFO] Keeping Contoso.Tools v5.7.1 (KeepVersions: 5)')
$removingAt = [array]::IndexOf($script:LogLines, '[INFO] Removing: Contoso.Tools v5.6.0')
Assert-That (($keepingAt -ge 0) -and ($removingAt -gt $keepingAt)) 'in that order'

Reset-State
$script:Installed = $machine
Remove-OldModuleVersions
Assert-That ((($script:UninstallCalls | Sort-Object) -join '; ') -eq 'Contoso.Tools 5.6.0; Contoso.Tools 5.7.1; Fabrikam.Core 1.0.0') 'without KeepVersions every old version goes, as before'

Reset-State -Keep @{ 'Contoso.Tools' = @('5') } -Excluded @('Contoso.Tools')
$script:Installed = $machine
Remove-OldModuleVersions
Assert-That (($script:UninstallCalls -join '; ') -eq 'Fabrikam.Core 1.0.0') 'an excluded module is not touched'

Reset-State -Keep @{ 'Contoso.Tools' = @('5') } -Pins @{ 'Contoso.Tools' = '7.0.0' }
$script:Installed = $machine
Remove-OldModuleVersions
Assert-That (($script:UninstallCalls -join '; ') -eq 'Fabrikam.Core 1.0.0') 'a pin that is not installed: nothing of that module is removed'
Assert-That (@((Get-LogLines 'WARN') | Where-Object { $_ -like 'Contoso.Tools is pinned to 7.0.0 but that version is not installed*' -or $_ -like '*Contoso.Tools is pinned to 7.0.0 but that version is not installed*' }).Count -eq 1) 'with the warning it always had'

Reset-State -Keep @{ 'Contoso.Tools' = @('5') }
$script:Installed = $machine
Remove-OldModuleVersions -WhatIf
Assert-That ($script:UninstallCalls.Count -eq 0) '-WhatIf removes nothing'
Assert-That ($script:LogLines -contains '[INFO] Keeping Contoso.Tools v5.7.1 (KeepVersions: 5)') 'but still says what would be kept'

Reset-State -Keep @{ 'Northwind.Data' = @('2') }
$script:Installed = $machine
Remove-OldModuleVersions
Assert-That (@((Get-LogLines 'WARN') | Where-Object { $_ -like "*no installed version of Northwind.Data matches '2'*" }).Count -eq 1) 'a module with one version and a selector that matches nothing is still warned about'
Assert-That ($script:Summary.PrunesFailed.Count -eq 0) 'and that is not an unsuccessful prune'

# --- Wiring into the update phase ----------------------------------------------------
Write-Section 'Update-AllModules'

$upToDate = {
    param($Names)
    New-Answer -Resources @(
        (New-Resource -Version '6.1.0'), (New-Resource -Name 'Fabrikam.Core' -Version '2.0.0')
        (New-Resource -Name 'Northwind.Data' -Version '3.0.0')
    )
}

# The normal week: no module needs updating. The line update has to run all the same
Reset-State -Keep @{ 'Contoso.Tools' = @('5') }
$script:Installed = $machine
$script:MainAnswer = $upToDate
$script:LineAnswer = { param($Name, $Range) New-Answer -Resources @((New-Resource -Name $Name -Version '5.8.0'), (New-Resource -Name $Name -Version '5.7.1')) }
Update-AllModules
Assert-That ($script:LogLines -contains '[INFO] All modules are up to date') 'no module needs the normal update'
Assert-That (($script:InstallCalls.Count -eq 1) -and ($script:InstallCalls[0].Version -eq '5.8.0')) 'and the kept line is updated all the same'
$lineAt = [array]::IndexOf($script:LogLines, '[INFO] Kept line updates complete. Updated: 1, Unsuccessful: 0, Not checked: 0')
$doneAt = [array]::IndexOf($script:LogLines, '[INFO] All modules are up to date')
Assert-That (($lineAt -ge 0) -and ($doneAt -gt $lineAt)) 'before the normal updates, not after'

# The normal update of the same module still happens
Reset-State -Keep @{ 'Contoso.Tools' = @('5') }
$script:Installed = $machine
$script:MainAnswer = { param($Names) New-Answer -Resources @((New-Resource -Version '6.2.0'), (New-Resource -Name 'Fabrikam.Core' -Version '2.0.0'), (New-Resource -Name 'Northwind.Data' -Version '3.0.0')) }
$script:LineAnswer = { param($Name, $Range) New-Answer -Resources @((New-Resource -Name $Name -Version '5.8.0')) }
Update-AllModules
$versions = @($script:InstallCalls | ForEach-Object { "$($_.Name) $($_.Version)".Trim() })
Assert-That (($versions.Count -eq 2) -and ($versions[0] -eq 'Contoso.Tools 5.8.0') -and ($versions[1] -eq 'Contoso.Tools')) "the line gets its exact version, then the module its normal update: $($versions -join '; ')"
Assert-That ($script:Summary.ModulesUpdated -eq 2) 'both count as updates'

# The gallery cannot be reached at all
Reset-State -Keep @{ 'Contoso.Tools' = @('5') }
$script:Installed = $machine
$script:MainAnswer = { param($Names) New-Answer -Unchecked $Names -Fault 'No such host is known' -Detail 'No such host is known.' }
Update-AllModules
Assert-That ($script:LookupCalls.Count -eq 1) 'gallery unreachable: the kept lines are not asked about on top of that'
Assert-That ($script:InstallCalls.Count -eq 0) 'nothing installed'
Assert-That ($script:Summary.KeepLinesUnchecked.Count -eq 0) 'and the outage is counted once, by the main lookup'

# An excluded module is left alone by the line update as well
Reset-State -Keep @{ 'Contoso.Tools' = @('5') } -Excluded @('Contoso.Tools')
$script:Installed = $machine
$script:MainAnswer = $upToDate
$script:LineAnswer = { param($Name, $Range) New-Answer -Resources @((New-Resource -Name $Name -Version '5.8.0')) }
Update-AllModules
Assert-That ($script:InstallCalls.Count -eq 0) 'an excluded module gets no line update'

Complete-Tests
