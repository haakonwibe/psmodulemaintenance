<#
.SYNOPSIS
    Tests for the configuration template and for running without a config file.

.DESCRIPTION
    config.json is not in the repository. config.example.json is, and it has to stay
    valid, neutral and equal to the defaults built into the script. The script also has
    to run when there is no config.json at all.
#>
param([switch]$Quiet)

. (Join-Path $PSScriptRoot 'TestHelpers.ps1')

$examplePath = Join-Path $script:RepoRoot 'config.example.json'
$schemaPath = Join-Path $script:RepoRoot 'config.schema.json'
$ignorePath = Join-Path $script:RepoRoot '.gitignore'

# The defaults are assigned at the top of the script, outside any function. This lifts
# that one statement out, the way Get-ScriptFunctionText lifts a function
function Get-ScriptAssignmentText {
    param([string]$Path, [string]$Variable)

    $ast = [System.Management.Automation.Language.Parser]::ParseFile($Path, [ref]$null, [ref]$null)
    $assignment = $ast.FindAll(
        { param($node) $node -is [System.Management.Automation.Language.AssignmentStatementAst] }, $false) |
        Where-Object { $_.Left.Extent.Text -eq $Variable } | Select-Object -First 1

    if (-not $assignment) {
        throw "No assignment to $Variable found in $Path"
    }
    return $assignment.Extent.Text
}

function Reset-Config {
    . ([scriptblock]::Create((Get-ScriptAssignmentText -Path $script:MainScript -Variable '$script:Config')))
    $script:ConfigWarnings = @()
    $script:ConfigFault = $null
    $script:ProtectedModules = @{}
    $script:ConfigBadEntries = @()
}

$wanted = 'ConvertTo-NormalizedVersion', 'Get-ModuleVersionKey', 'ConvertFrom-PinnedVersionString',
          'ConvertTo-PinnedModuleTable', 'ConvertFrom-KeepVersionSelector', 'ConvertTo-KeepVersionTable',
          'Import-MaintenanceConfig'
foreach ($name in $wanted) {
    . ([scriptblock]::Create((Get-ScriptFunctionText -Path $script:MainScript -Name $name)))
}

# --- The template itself -------------------------------------------------------------
Write-Section 'config.example.json'

Assert-That (Test-Path -LiteralPath $examplePath) 'the template exists'

$raw = Get-Content -LiteralPath $examplePath -Raw
$example = $null
try { $example = $raw | ConvertFrom-Json -ErrorAction Stop } catch { $example = $null }
Assert-That ($null -ne $example) 'it is valid JSON'

$valid = $false
try { $valid = Test-Json -Json $raw -SchemaFile $schemaPath -ErrorAction Stop } catch { $valid = $false }
Assert-That $valid 'it validates against config.schema.json'

# Whatever ships in a public repository must not switch anything on or single anything out
Assert-That ($example.Healthchecks.Enabled -eq $false) 'Healthchecks is off'
Assert-That (@($example.ExcludedModules).Count -eq 0) 'no module is excluded'
Assert-That (@($example.PinnedModules.PSObject.Properties).Count -eq 0) 'no module is pinned'
Assert-That (@($example.KeepVersions.PSObject.Properties).Count -eq 0) 'no version is kept'
Assert-That ($raw -notmatch 'hc-ping\.com|healthchecks\.io/ping') 'it holds no ping URL'

# --- The template against the built-in defaults --------------------------------------
Write-Section 'Template and built-in defaults'

Reset-Config
$defaults = $script:Config

Reset-Config
Import-MaintenanceConfig -Path $examplePath
$loaded = $script:Config

foreach ($key in 'LogRetentionDays', 'TrustPSGallery', 'NotificationMode', 'ModuleUpdateTimeoutSeconds') {
    Assert-That ($loaded[$key] -eq $defaults[$key]) "$key is the default ($($defaults[$key]))"
}
foreach ($key in 'Enabled', 'SecretName', 'SecretVault', 'TimeoutSeconds') {
    Assert-That ($loaded.Healthchecks[$key] -eq $defaults.Healthchecks[$key]) "Healthchecks.$key is the default ($($defaults.Healthchecks[$key]))"
}
Assert-That (@($loaded.ExcludedModules).Count -eq 0) 'ExcludedModules loads as empty'
Assert-That ($loaded.PinnedModules.Count -eq 0) 'PinnedModules loads as empty'
Assert-That ($loaded.KeepVersions.Count -eq 0) 'KeepVersions loads as empty'
Assert-That ($script:ConfigWarnings.Count -eq 0) 'loading it raises no warning'

$templateKeys = @($example.PSObject.Properties.Name | Where-Object { $_ -notin '$schema', '_comment' } | Sort-Object)
$defaultKeys = @($defaults.Keys | Sort-Object)
Assert-That (($templateKeys -join ',') -eq ($defaultKeys -join ',')) 'the template lists every setting, and no setting the script does not know'

# --- No config file at all -----------------------------------------------------------
Write-Section 'Running without config.json'

Reset-Config
$threw = $false
$missing = Join-Path ([System.IO.Path]::GetTempPath()) ('no-such-config-' + [guid]::NewGuid().ToString('N') + '.json')
try { Import-MaintenanceConfig -Path $missing } catch { $threw = $true }
Assert-That (-not $threw) 'a missing file is not an error'
Assert-That ($script:Config.Healthchecks.Enabled -eq $false) 'Healthchecks stays off'
Assert-That ($script:Config.ModuleUpdateTimeoutSeconds -eq $defaults.ModuleUpdateTimeoutSeconds) 'the defaults are in force'

# --- Loading KeepVersions ------------------------------------------------------------
Write-Section 'Loading KeepVersions'

$configFolder = Join-Path ([System.IO.Path]::GetTempPath()) ('PSModuleMaintenance-tests-' + [guid]::NewGuid().ToString('N'))
New-Item -Path $configFolder -ItemType Directory -Force | Out-Null

# Writes the text to a config file and loads it over fresh defaults
function Import-TestConfig {
    param([string]$Json)

    $path = Join-Path $configFolder 'config.json'
    Set-Content -LiteralPath $path -Value $Json
    Reset-Config
    Import-MaintenanceConfig -Path $path 3>$null
}

try {
    Import-TestConfig '{ "KeepVersions": { "Contoso.Tools": ["5", "5.7"] } }'
    $kept = @($script:Config.KeepVersions['Contoso.Tools'])
    Assert-That ($script:Config.KeepVersions.Count -eq 1) 'a valid entry loads'
    Assert-That (($kept.Count -eq 2) -and ($kept[0].Requested -eq '5') -and ($kept[1].Requested -eq '5.7')) 'both selectors are there, in order'
    Assert-That (($kept[1].Parts -join ',') -eq '5,7') 'a selector is parsed into its numbers'
    Assert-That ($script:ConfigWarnings.Count -eq 0) 'without a warning'
    Assert-That ($script:Config.KeepVersions.ContainsKey('contoso.tools')) 'the module name is looked up without regard to case'

    Import-TestConfig '{ "KeepVersions": { "Contoso.Tools": ["5", "5", "5.7"] } }'
    Assert-That (@($script:Config.KeepVersions['Contoso.Tools']).Count -eq 2) 'a selector written twice counts once'
    Assert-That ($script:ProtectedModules.Count -eq 0) 'and that is not a problem'

    Import-TestConfig '{ "KeepVersions": { "Contoso.Tools": "5" } }'
    $kept = @($script:Config.KeepVersions['Contoso.Tools'])
    Assert-That (($kept.Count -eq 1) -and ($kept[0].Requested -eq '5')) 'a single value without the list brackets is accepted'

    Import-TestConfig '{ "ExcludedModules": ["Contoso.Tools"], "KeepVersions": { "contoso.tools": ["5"], "Fabrikam.Core": ["2"] } }'
    Assert-That (-not $script:Config.KeepVersions.ContainsKey('Contoso.Tools')) 'an excluded module loses its KeepVersions entry'
    Assert-That ($script:Config.KeepVersions.ContainsKey('Fabrikam.Core')) 'other entries stay'
    Assert-That (@($script:ConfigWarnings | Where-Object { $_ -like '*both excluded and listed in KeepVersions*' }).Count -eq 1) 'the conflict is reported'

    Import-TestConfig '{ "PinnedModules": { "Contoso.Tools": "6.1.0" }, "KeepVersions": { "Contoso.Tools": ["5"] } }'
    Assert-That (($script:Config.PinnedModules.ContainsKey('Contoso.Tools')) -and ($script:Config.KeepVersions.ContainsKey('Contoso.Tools'))) 'pinning and keeping combine'
    Assert-That ($script:ConfigWarnings.Count -eq 0) 'without a warning'

    # --- An entry that is not understood ---------------------------------------------
    # It says that something was wanted for that module, but not what. The module is
    # left alone for the run instead of being treated like any other
    Write-Section 'An entry that is not understood'

    Import-TestConfig '{ "KeepVersions": { "Contoso.Tools": ["5.*"], "Fabrikam.Core": ["2"] } }'
    Assert-That ($script:ProtectedModules.ContainsKey('Contoso.Tools')) 'a selector that is not a version prefix: the module is protected'
    Assert-That (-not $script:Config.KeepVersions.ContainsKey('Contoso.Tools')) 'and its entry is set aside'
    Assert-That ($script:Config.KeepVersions.ContainsKey('Fabrikam.Core')) 'the entry next to it is not affected'
    Assert-That (-not $script:ProtectedModules.ContainsKey('Fabrikam.Core')) 'nor is that module protected'
    Assert-That ((@($script:ProtectedModules['Contoso.Tools']) -join ' ') -like "*KeepVersions: '5.`**' is not a version prefix such as 5 or 5.7*") 'the reason names the setting and the value'
    Assert-That ($null -eq $script:ConfigFault) 'the file itself still counts as readable'

    Import-TestConfig '{ "KeepVersions": { "Contoso.Tools": ["5", "5.x"] } }'
    Assert-That ($script:ProtectedModules.ContainsKey('Contoso.Tools')) 'one bad selector among good ones spoils the entry'
    Assert-That (-not $script:Config.KeepVersions.ContainsKey('Contoso.Tools')) 'the good one is not used on its own: what was meant is not known'

    # Unquoted, 5.10 reaches the script as the number 5.1 and would select the wrong line
    Import-TestConfig '{ "KeepVersions": { "Contoso.Tools": [5.10] } }'
    Assert-That ($script:ProtectedModules.ContainsKey('Contoso.Tools')) 'a number where a quoted selector belongs: the module is protected'
    Assert-That ((@($script:ProtectedModules['Contoso.Tools']) -join ' ') -like '*has to be a string in quotes*') 'and the reason says to quote it'

    Import-TestConfig '{ "KeepVersions": { "Contoso.Tools": [] } }'
    Assert-That ($script:ProtectedModules.ContainsKey('Contoso.Tools')) 'an entry with an empty list: the module is protected'

    Import-TestConfig '{ "PinnedModules": { "Contoso.Tools": "2.*", "Fabrikam.Core": "1.4.0" } }'
    Assert-That ($script:ProtectedModules.ContainsKey('Contoso.Tools')) 'a pin that is not a version: the module is protected'
    Assert-That (-not $script:Config.PinnedModules.ContainsKey('Contoso.Tools')) 'and not pinned to anything'
    Assert-That ($script:Config.PinnedModules.ContainsKey('Fabrikam.Core')) 'the pin next to it still holds'
    Assert-That ((@($script:ProtectedModules['Contoso.Tools']) -join ' ') -like '*PinnedModules: ''2.`*'' is not an exact version such as 2.19.0*') 'the reason names the setting and the value, and shows what a pin looks like'

    Import-TestConfig '{ "PinnedModules": { "Contoso.Tools": 2.10 } }'
    Assert-That ($script:ProtectedModules.ContainsKey('Contoso.Tools')) 'a pin written as a number: the module is protected, 2.10 would have arrived as 2.1'

    Import-TestConfig '{ "PinnedModules": { "Contoso.Tools": "6.1.0" }, "KeepVersions": { "Contoso.Tools": ["5.x"] } }'
    Assert-That ($script:ProtectedModules.ContainsKey('Contoso.Tools')) 'a good pin and a bad KeepVersions entry: the module is protected'
    Assert-That (-not $script:Config.PinnedModules.ContainsKey('Contoso.Tools')) 'and the good pin is set aside with it, the module is left alone altogether'

    Import-TestConfig '{ "PinnedModules": { "Contoso.Tools": "latest" }, "KeepVersions": { "Contoso.Tools": ["5"] } }'
    Assert-That ($script:ProtectedModules.ContainsKey('Contoso.Tools')) 'a bad pin and a good KeepVersions entry: the module is protected'
    Assert-That (-not $script:Config.KeepVersions.ContainsKey('Contoso.Tools')) 'and the good entry is set aside with it, no line of it is updated'

    Import-TestConfig '{ "PinnedModules": { "Contoso.Tools": "latest" }, "KeepVersions": { "Contoso.Tools": ["5.x"] } }'
    Assert-That (@($script:ProtectedModules['Contoso.Tools']).Count -eq 2) 'two bad entries for one module: both reasons are kept'
    Assert-That ($script:ProtectedModules.Count -eq 1) 'and the module is listed once'

    Import-TestConfig '{ "ExcludedModules": ["Contoso.Tools"], "KeepVersions": { "Contoso.Tools": ["5.x"] } }'
    Assert-That (-not $script:ProtectedModules.ContainsKey('Contoso.Tools')) 'a bad entry for an excluded module: no protection is needed, it is left alone anyway'
    Assert-That (@($script:ConfigWarnings | Where-Object { $_ -like "'Contoso.Tools' is excluded, so its KeepVersions entry is ignored*" }).Count -eq 1) 'a warning says so'

    Import-TestConfig '{ "KeepVersions": { "contoso.tools": ["5.x"] } }'
    Assert-That ($script:ProtectedModules.ContainsKey('Contoso.Tools')) 'the protected name is looked up without regard to case'

    # The script loads its config once, but what a load finds must not depend on that
    $again = Join-Path $configFolder 'config.json'
    Set-Content -LiteralPath $again -Value '{ "KeepVersions": { "Contoso.Tools": ["5"] } }'
    Import-MaintenanceConfig -Path $again 3>$null
    Assert-That ($script:ProtectedModules.Count -eq 0) 'a second load starts afresh: the entry was put right, the module is no longer protected'
    Assert-That ($script:Config.KeepVersions.ContainsKey('Contoso.Tools')) 'and its entry is in force'

    # --- A setting of the wrong shape ------------------------------------------------
    # One entry gone wrong names its module. A whole setting gone wrong names none, so
    # it is not known which modules to leave alone, and the file counts as unreadable
    Write-Section 'A setting of the wrong shape'

    Import-TestConfig '{ "KeepVersions": ["Contoso.Tools"], "LogRetentionDays": 30 }'
    Assert-That ($script:ConfigFault -like 'KeepVersions has to be an object*') 'KeepVersions as a list: the file counts as unreadable'

    Import-TestConfig '{ "PinnedModules": "Contoso.Tools" }'
    Assert-That ($script:ConfigFault -like 'PinnedModules has to be an object*') 'PinnedModules as text: the same'

    Import-TestConfig '{ "ExcludedModules": { "Contoso.Tools": true } }'
    Assert-That ($script:ConfigFault -like 'ExcludedModules has to be a list*') 'ExcludedModules as an object: the same'

    Import-TestConfig '{ "ExcludedModules": ["Contoso.Tools", 5] }'
    Assert-That ($script:ConfigFault -like 'ExcludedModules has to be a list*') 'ExcludedModules with a number in it: the same'

    Import-TestConfig '{ "ExcludedModules": "Contoso.Tools" }'
    Assert-That (($null -eq $script:ConfigFault) -and (@($script:Config.ExcludedModules) -contains 'Contoso.Tools')) 'a single name without the list brackets is accepted'

    Import-TestConfig '{ "ExcludedModules": [], "PinnedModules": [], "KeepVersions": [] }'
    Assert-That (($null -eq $script:ConfigFault) -and ($script:ProtectedModules.Count -eq 0)) 'empty lists all round mean "none" and are let through'

    # --- A config file that cannot be read -------------------------------------------
    Write-Section 'A config file that cannot be read'

    Import-TestConfig '{ "ExcludedModules": ["Contoso.Tools"], '
    Assert-That (-not [string]::IsNullOrWhiteSpace($script:ConfigFault)) 'the fault is recorded, with its reason, for the main block to act on'

    Import-TestConfig 'not json at all'
    Assert-That (-not [string]::IsNullOrWhiteSpace($script:ConfigFault)) 'the same for a file that is not JSON'

    Import-TestConfig '{ "LogRetentionDays": 30 }'
    Assert-That ($null -eq $script:ConfigFault) 'a file that can be read records no fault'

    Reset-Config
    $missing = Join-Path $configFolder 'there-is-no-such-file.json'
    Import-MaintenanceConfig -Path $missing
    Assert-That ($null -eq $script:ConfigFault) 'a file that does not exist is not a fault: the defaults are what was asked for'
}
finally {
    if (Test-Path -LiteralPath $configFolder) {
        Remove-Item -LiteralPath $configFolder -Recurse -Force
    }
}

# --- config.json stays out of the repository -----------------------------------------
Write-Section 'config.json is ignored'

$ignoreLines = @(Get-Content -LiteralPath $ignorePath | ForEach-Object { $_.Trim() })
Assert-That (($ignoreLines -contains '/config.json') -or ($ignoreLines -contains 'config.json')) '.gitignore lists config.json'

Complete-Tests
