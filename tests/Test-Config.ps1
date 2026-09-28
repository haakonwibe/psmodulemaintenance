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

    Import-TestConfig '{ "KeepVersions": { "Contoso.Tools": ["5", "5.*", "5"] } }'
    $kept = @($script:Config.KeepVersions['Contoso.Tools'])
    Assert-That (($kept.Count -eq 1) -and ($kept[0].Requested -eq '5')) 'a bad selector is dropped, its sibling stays, a repeat counts once'
    Assert-That (@($script:ConfigWarnings | Where-Object { $_ -like "*'5.*' is not a version prefix*" }).Count -eq 1) 'the bad selector is reported'

    # Unquoted, 5.10 reaches the script as the number 5.1 and would select the wrong line
    Import-TestConfig '{ "KeepVersions": { "Contoso.Tools": [5.10] } }'
    Assert-That ($script:Config.KeepVersions.Count -eq 0) 'a number is rejected'
    Assert-That (@($script:ConfigWarnings | Where-Object { $_ -like '*has to be a string in quotes*' }).Count -eq 1) 'and reported'
    Assert-That (@($script:ConfigWarnings | Where-Object { $_ -like '*holds no usable selector*' }).Count -eq 1) 'an entry left with nothing is reported too'

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

    Import-TestConfig '{ "KeepVersions": ["Contoso.Tools"], "LogRetentionDays": 30 }'
    Assert-That ($script:Config.KeepVersions.Count -eq 0) 'a list where an object belongs is ignored'
    Assert-That (@($script:ConfigWarnings | Where-Object { $_ -like 'Ignoring KeepVersions: it has to be an object*' }).Count -eq 1) 'and reported'
    Assert-That ($script:Config.LogRetentionDays -eq 30) 'the rest of that file still loads'
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
