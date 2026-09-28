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
          'ConvertTo-PinnedModuleTable', 'Import-MaintenanceConfig'
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

# --- config.json stays out of the repository -----------------------------------------
Write-Section 'config.json is ignored'

$ignoreLines = @(Get-Content -LiteralPath $ignorePath | ForEach-Object { $_.Trim() })
Assert-That (($ignoreLines -contains '/config.json') -or ($ignoreLines -contains 'config.json')) '.gitignore lists config.json'

Complete-Tests
