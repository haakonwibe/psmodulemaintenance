<#
.SYNOPSIS
    Tests for Invoke-OneDriveMigration.ps1, run against a made-up folder tree.

.DESCRIPTION
    The migration script is one long script, not a set of functions, so it is run as a
    whole. It runs from a copy that differs from the real one in three lines: the demand
    for elevation is dropped, and the two folders it works on, the user's module folder
    and the AllUsers module folder, point into a temp folder. OneDrive is redirected
    through the environment.

    The copy is checked before it is run, and it carries a guard of its own that stops
    it before the first write unless both folders lie inside that temp folder. Nothing
    outside it is touched, and the temp folder is removed afterwards.

.NOTES
    Program Files cannot be redirected through the environment. Windows sets that
    variable afresh for every process it starts, so a child process sees the real
    folder whatever the parent put there. The path has to be replaced in the text.
#>
param([switch]$Quiet)

. (Join-Path $PSScriptRoot 'TestHelpers.ps1')

$root = Join-Path ([System.IO.Path]::GetTempPath()) ('PSModuleMaintenance-tests-' + [guid]::NewGuid().ToString('N'))
$fakeOneDrive = Join-Path $root 'OneDrive'
$fakeDocuments = Join-Path $fakeOneDrive 'Documents'
$fakeUserModules = Join-Path $fakeDocuments 'PowerShell\Modules'
$fakeAllUsers = Join-Path $root 'ProgramFiles\PowerShell\Modules'
$logFolder = Join-Path $root 'Logs'
$copyPath = Join-Path $root 'Invoke-OneDriveMigration.copy.ps1'

function New-FakeModule {
    param([string]$Name, [string]$Version, [switch]$Managed)

    $folder = Join-Path $fakeUserModules "$Name\$Version"
    New-Item -Path $folder -ItemType Directory -Force | Out-Null
    Set-Content -Path (Join-Path $folder "$Name.psd1") -Value "@{ ModuleVersion = '$Version' }"
    Set-Content -Path (Join-Path $folder "$Name.psm1") -Value '# made up'

    if ($Managed) {
        $marker = Join-Path $folder 'PSGetModuleInfo.xml'
        Set-Content -Path $marker -Value '<Objs />'
        (Get-Item -LiteralPath $marker).Attributes = 'Hidden'
    }
}

# Runs the copy in a process of its own, because the script ends with "exit" in places.
# Returns the lines of the log it wrote
function Invoke-MigrationCopy {
    param(
        [string[]]$Arguments = @(),

        # Where the process is told OneDrive lives. The made-up one unless stated
        [string]$OneDrive = $fakeOneDrive
    )

    if (Test-Path -LiteralPath $logFolder) {
        Remove-Item -LiteralPath $logFolder -Recurse -Force
    }

    $saved = @{
        OneDrive           = $env:OneDrive
        OneDriveCommercial = $env:OneDriveCommercial
        OneDriveConsumer   = $env:OneDriveConsumer
    }
    try {
        $env:OneDrive = $OneDrive
        $env:OneDriveCommercial = $null
        $env:OneDriveConsumer = $null

        $all = @('-NoProfile', '-NonInteractive', '-File', $copyPath, '-LogPath', $logFolder) + $Arguments
        & pwsh @all *> $null
    }
    finally {
        foreach ($key in $saved.Keys) {
            Set-Item -Path "Env:$key" -Value $saved[$key]
        }
    }

    $lines = @()
    $logFile = Get-ChildItem -LiteralPath $logFolder -Filter 'migration_*.log' -ErrorAction SilentlyContinue |
        Select-Object -First 1
    if ($logFile) {
        $lines = @(Get-Content -LiteralPath $logFile.FullName | ForEach-Object { $_ -replace '^\[[\d\- :]+\] ', '' })
    }
    return $lines
}

try {
    New-Item -Path $fakeUserModules -ItemType Directory -Force | Out-Null
    New-Item -Path $fakeAllUsers -ItemType Directory -Force | Out-Null

    # --- The copy and its guard ------------------------------------------------------
    Write-Section 'The copy that is run'

    $text = Get-Content -LiteralPath $script:MigrationScript -Raw

    # The guard sits where the AllUsers folder is set, which is before the first write
    $allUsersLine = "`$allUsersPath = Join-Path `$env:ProgramFiles 'PowerShell\Modules'"
    $guarded = @"
`$allUsersPath = '$fakeAllUsers'
if ((-not `$allUsersPath.StartsWith('$root')) -or (-not `$currentUserModulePath.StartsWith('$root'))) {
    throw 'Test guard: this copy only works on the made-up folder tree.'
}
"@
    $copy = $text -replace '(?m)^#Requires -RunAsAdministrator[ \t]*\r?$', '# (elevation is not needed for the made-up folder tree)'
    $copy = $copy.Replace("[Environment]::GetFolderPath('MyDocuments')", "'$fakeDocuments'")
    $copy = $copy.Replace($allUsersLine, $guarded)

    $isSafe = ($copy -notmatch '#Requires -RunAsAdministrator') -and
              ($copy -notmatch 'GetFolderPath') -and
              (-not $copy.Contains('Join-Path $env:ProgramFiles')) -and
              ($copy.Contains('Test guard:')) -and
              ($copy.Contains("'$fakeDocuments'")) -and
              ($copy.Contains("'$fakeAllUsers'"))
    Assert-That $isSafe 'all three lines were replaced, so the copy cannot reach a real folder'
    if (-not $isSafe) {
        throw 'The copy is not safe to run. Stopping before anything is touched.'
    }

    $parseErrors = $null
    $null = [System.Management.Automation.Language.Parser]::ParseInput($copy, [ref]$null, [ref]$parseErrors)
    Assert-That ($parseErrors.Count -eq 0) 'the copy parses'
    Set-Content -LiteralPath $copyPath -Value $copy

    # --- Outside OneDrive ------------------------------------------------------------
    Write-Section 'A module folder that is not in OneDrive'

    New-FakeModule -Name 'Contoso.Tools' -Version '1.0.0' -Managed
    $log = Invoke-MigrationCopy -OneDrive (Join-Path $root 'SomewhereElse')
    Assert-That (@($log | Where-Object { $_ -like '*is not in OneDrive*no migration needed' }).Count -eq 1) 'the script says there is nothing to migrate'
    Assert-That (Test-Path (Join-Path $fakeUserModules 'Contoso.Tools\1.0.0')) 'and leaves the module where it is'
    Assert-That (-not (Test-Path (Join-Path $fakeAllUsers 'Contoso.Tools'))) 'without copying it'

    # --- A first run -----------------------------------------------------------------
    Write-Section 'A first run'

    New-FakeModule -Name 'Contoso.Tools' -Version '1.1.0' -Managed
    New-FakeModule -Name 'Fabrikam.Agent' -Version '2.0.0'
    New-Item -Path (Join-Path $fakeUserModules 'Northwind.Data') -ItemType Directory -Force | Out-Null

    $log = Invoke-MigrationCopy -Arguments '-WhatIf'
    Assert-That (-not (Test-Path (Join-Path $fakeAllUsers 'Contoso.Tools'))) '-WhatIf copies nothing'
    Assert-That (Test-Path (Join-Path $fakeUserModules 'Contoso.Tools\1.0.0')) '-WhatIf removes nothing'
    Assert-That ($log -contains '[INFO] Leaving Fabrikam.Agent in place: not installed by PSResourceGet, so it belongs to another program') '-WhatIf still says what would be left in place'

    $log = Invoke-MigrationCopy
    Assert-That (Test-Path (Join-Path $fakeAllUsers 'Contoso.Tools\1.0.0\Contoso.Tools.psd1')) 'a module installed by PSResourceGet is copied'
    Assert-That (Test-Path (Join-Path $fakeAllUsers 'Contoso.Tools\1.1.0\PSGetModuleInfo.xml') ) 'every version of it, with its hidden marker'
    Assert-That (-not (Test-Path (Join-Path $fakeUserModules 'Contoso.Tools'))) 'and removed from the OneDrive path'
    Assert-That (-not (Test-Path (Join-Path $fakeAllUsers 'Fabrikam.Agent'))) 'a module another program installed is not copied'
    Assert-That (Test-Path (Join-Path $fakeUserModules 'Fabrikam.Agent\2.0.0\Fabrikam.Agent.psd1')) 'and not removed'
    Assert-That (-not (Test-Path (Join-Path $fakeUserModules 'Northwind.Data'))) 'an empty folder left by an earlier run is removed'

    Assert-That ($log -contains '[INFO] Copy phase complete. Copied: 2, Unsuccessful: 0') 'the log counts two copies'
    Assert-That (@($log | Where-Object { $_ -like '`[INFO`] Cleaning up OneDrive copies from:*' }).Count -eq 1) 'says it is cleaning up'
    Assert-That ($log -contains '[INFO] Cleanup complete. Removed: 2, Unsuccessful: 0') 'and counts what it removed'
    Assert-That ($log -notcontains '[INFO] Nothing to clean up in the OneDrive path') 'without also saying there was nothing to do'
    Assert-That ($log -contains '[INFO]   OneDrive copies removed: 2') 'the summary has the number'
    Assert-That ($log -contains '[INFO]   Left in place, installed by another program: Fabrikam.Agent') 'and names what was left in place'

    # --- A second run, with nothing left to do ---------------------------------------
    Write-Section 'A second run'

    $log = Invoke-MigrationCopy
    Assert-That ($log -contains '[INFO] Copy phase complete. Copied: 0, Unsuccessful: 0') 'nothing is copied'
    Assert-That ($log -contains '[INFO] Nothing to clean up in the OneDrive path') 'the log says there is nothing to clean up'
    Assert-That (@($log | Where-Object { $_ -like '*Cleaning up OneDrive copies*' -or $_ -like '*Cleanup complete*' }).Count -eq 0) 'and does not go on to say it is cleaning up'
    Assert-That ($log -contains '[INFO]   OneDrive copies removed: 0') 'the summary says 0, not an empty space'
    Assert-That (@($log | Where-Object { $_ -like '*already exist in AllUsers*' }).Count -eq 0) 'nothing is claimed about modules that were only left in place'
    Assert-That (Test-Path (Join-Path $fakeUserModules 'Fabrikam.Agent\2.0.0\Fabrikam.Agent.psd1')) 'the other program''s module is still there'

    # --- A run that was cut short earlier --------------------------------------------
    Write-Section 'Leftovers from a run that was cut short'

    # Already in AllUsers from the first run, and back in the OneDrive path
    New-FakeModule -Name 'Contoso.Tools' -Version '1.1.0' -Managed
    $log = Invoke-MigrationCopy
    Assert-That ($log -contains '[INFO] Copy phase complete. Copied: 0, Unsuccessful: 0') 'nothing needs copying'
    Assert-That ($log -contains '[INFO] Cleanup complete. Removed: 1, Unsuccessful: 0') 'the leftover copy is still cleaned up'
    Assert-That (-not (Test-Path (Join-Path $fakeUserModules 'Contoso.Tools'))) 'and gone from the OneDrive path'

    # --- Keeping the OneDrive copies -------------------------------------------------
    Write-Section '-SkipCleanup'

    New-FakeModule -Name 'Contoso.Reports' -Version '3.0.0' -Managed
    $log = Invoke-MigrationCopy -Arguments '-SkipCleanup'
    Assert-That (Test-Path (Join-Path $fakeAllUsers 'Contoso.Reports\3.0.0')) 'the module is copied'
    Assert-That (Test-Path (Join-Path $fakeUserModules 'Contoso.Reports\3.0.0')) 'and the OneDrive copy stays'
    Assert-That ($log -contains '[INFO] Cleanup skipped (-SkipCleanup). OneDrive copies remain in place.') 'the log says so'
    Assert-That (@($log | Where-Object { $_ -like '*Nothing to clean up*' -or $_ -like '*OneDrive copies removed*' }).Count -eq 0) 'and says nothing about cleaning up'

    # --- Wording ---------------------------------------------------------------------
    Write-Section 'Log wording'

    $calmLines = @($log | Where-Object { ($_ -like '`[INFO`]*') -or ($_ -like '`[SUCCESS`]*') })
    Assert-That (@($calmLines | Where-Object { $_ -match '(?i)error|fail|warn' }).Count -eq 0) 'no INFO or SUCCESS line holds a word that would colour it as a problem'
}
finally {
    if (Test-Path -LiteralPath $root) {
        Remove-Item -LiteralPath $root -Recurse -Force
    }
}

Complete-Tests
