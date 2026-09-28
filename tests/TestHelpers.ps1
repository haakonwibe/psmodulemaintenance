# Shared by the test scripts, which dot-source it.
#
# The tests run the real functions. Get-ScriptFunctionText lifts a function out of one of
# the scripts through the PowerShell parser, so the script itself never has to run, and a
# test stands in only for what would touch the network or the machine.

$script:RepoRoot = Split-Path -Path $PSScriptRoot -Parent
$script:MainScript = Join-Path $script:RepoRoot 'Invoke-PSModuleMaintenance.ps1'
$script:MigrationScript = Join-Path $script:RepoRoot 'Invoke-OneDriveMigration.ps1'

$script:Pass = 0
$script:Fail = 0

function Write-Section {
    param([string]$Title)

    if (-not $Quiet) {
        Write-Host "== $Title =="
    }
}

function Assert-That {
    param(
        [bool]$Condition,
        [string]$Name
    )

    if ($Condition) {
        $script:Pass++
        if (-not $Quiet) {
            Write-Host "  [OK]   $Name"
        }
    }
    else {
        $script:Fail++
        Write-Host "  [FAIL] $Name" -ForegroundColor Red
    }
}

function Get-ScriptFunctionAst {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$Name
    )

    $parseErrors = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($Path, [ref]$null, [ref]$parseErrors)
    if ($parseErrors.Count -gt 0) {
        $first = $parseErrors[0]
        throw "$Path does not parse: $($first.Message) (line $($first.Extent.StartLineNumber))"
    }

    $definition = $ast.FindAll(
        { param($node) $node -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $true) |
        Where-Object { $_.Name -eq $Name } | Select-Object -First 1

    if (-not $definition) {
        throw "Function $Name not found in $Path"
    }
    return $definition
}

function Get-ScriptFunctionText {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$Name
    )

    return (Get-ScriptFunctionAst -Path $Path -Name $Name).Extent.Text
}

function Complete-Tests {
    Write-Host "RESULT: $script:Pass passed, $script:Fail failed"
    if ($script:Fail -gt 0) {
        exit 1
    }
    exit 0
}
