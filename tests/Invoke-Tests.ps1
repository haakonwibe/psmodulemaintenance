<#
.SYNOPSIS
    Runs the tests.

.DESCRIPTION
    The tests need no network and no elevation. They install nothing, send no
    notification and no ping, and write only to a temp folder that they remove again.

    Each test file runs in a process of its own, so the stand-ins one file defines for
    commands like Find-PSResource cannot leak into the next.

.PARAMETER Detailed
    Show every check, not only the ones that did not pass.

.EXAMPLE
    .\tests\Invoke-Tests.ps1

.EXAMPLE
    .\tests\Invoke-Tests.ps1 -Detailed
#>
[CmdletBinding()]
param(
    [switch]$Detailed
)

#Requires -Version 7.0

$testFiles = @(Get-ChildItem -Path $PSScriptRoot -Filter 'Test-*.ps1' | Sort-Object Name)
$notPassed = @()
$timer = [System.Diagnostics.Stopwatch]::StartNew()

foreach ($testFile in $testFiles) {
    Write-Host $testFile.BaseName -ForegroundColor Cyan

    $arguments = @('-NoProfile', '-NonInteractive', '-File', $testFile.FullName)
    if (-not $Detailed) {
        $arguments += '-Quiet'
    }
    & pwsh @arguments

    if ($LASTEXITCODE -ne 0) {
        $notPassed += $testFile.BaseName
    }
}

$seconds = [math]::Round($timer.Elapsed.TotalSeconds, 1)
Write-Host ''

if ($notPassed.Count -gt 0) {
    Write-Host "NOT PASSED: $($notPassed -join ', ')  ($seconds s)" -ForegroundColor Red
    exit 1
}

Write-Host "All $($testFiles.Count) test files passed  ($seconds s)" -ForegroundColor Green
exit 0
