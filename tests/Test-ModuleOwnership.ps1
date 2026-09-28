<#
.SYNOPSIS
    Tests for telling our modules from another program's, and for finding where a module
    version really is on disk.

.DESCRIPTION
    Covers Test-ManagedModuleFolder, in both scripts, and Resolve-ModuleVersionFolder.
    Works on a made-up module tree in the temp folder, which is removed afterwards.
#>
param([switch]$Quiet)

. (Join-Path $PSScriptRoot 'TestHelpers.ps1')

$root = Join-Path ([System.IO.Path]::GetTempPath()) ('PSModuleMaintenance-tests-' + [guid]::NewGuid().ToString('N'))

function New-FakeVersion {
    param(
        [string]$Module, [string]$Version,
        [switch]$Managed, [switch]$NoManifest, [switch]$Empty
    )

    $folder = Join-Path $root "$Module\$Version"
    New-Item -Path $folder -ItemType Directory -Force | Out-Null
    if ($Empty) { return }

    if (-not $NoManifest) {
        Set-Content -Path (Join-Path $folder "$Module.psd1") -Value "@{ ModuleVersion = '$Version' }"
    }
    Set-Content -Path (Join-Path $folder "$Module.dll") -Value 'x'

    if ($Managed) {
        # Hidden, the way PSResourceGet writes it
        $marker = Join-Path $folder 'PSGetModuleInfo.xml'
        Set-Content -Path $marker -Value '<Objs />'
        (Get-Item -LiteralPath $marker).Attributes = 'Hidden'
    }
}

try {
    New-FakeVersion -Module 'Gallery.Module' -Version '2.0.0' -Managed
    New-FakeVersion -Module 'Gallery.TwoVersions' -Version '1.0.0'
    New-FakeVersion -Module 'Gallery.TwoVersions' -Version '2.0.0' -Managed
    New-FakeVersion -Module 'Microsoft.PowerToys.Configure' -Version '0.90.0.0'
    New-FakeVersion -Module 'Four.Part' -Version '6.1907.1.0' -Managed
    New-FakeVersion -Module 'Three.Part' -Version '5.7.1' -Managed
    New-FakeVersion -Module 'Leftover' -Version '1.0.0' -NoManifest
    New-FakeVersion -Module 'EmptyVersion' -Version '3.0.0' -Empty
    New-Item -Path (Join-Path $root 'EmptyModule') -ItemType Directory -Force | Out-Null

    # No version folder: the files sit directly in the module folder
    $flat = Join-Path $root 'Flat.Managed'
    New-Item -Path $flat -ItemType Directory -Force | Out-Null
    Set-Content -Path (Join-Path $flat 'Flat.Managed.psd1') -Value '@{}'
    Set-Content -Path (Join-Path $flat 'PSGetModuleInfo.xml') -Value '<Objs />'
    (Get-Item -LiteralPath (Join-Path $flat 'PSGetModuleInfo.xml')).Attributes = 'Hidden'

    # A marker further down belongs to a bundled dependency, not to the module
    $deep = Join-Path $root 'Deep.Unmanaged\1.0.0\lib\inner'
    New-Item -Path $deep -ItemType Directory -Force | Out-Null
    Set-Content -Path (Join-Path $root 'Deep.Unmanaged\1.0.0\Deep.Unmanaged.psd1') -Value '@{}'
    Set-Content -Path (Join-Path $deep 'PSGetModuleInfo.xml') -Value '<Objs />'

    $expected = [ordered]@{
        'Gallery.Module'                = $true
        'Gallery.TwoVersions'           = $true
        'Microsoft.PowerToys.Configure' = $false
        'Four.Part'                     = $true
        'Leftover'                      = $false
        'EmptyVersion'                  = $false
        'EmptyModule'                   = $false
        'Flat.Managed'                  = $true
        'Deep.Unmanaged'                = $false
    }

    $sources = @(
        @{ Label = 'maintenance script'; Path = $script:MainScript }
        @{ Label = 'migration script';   Path = $script:MigrationScript }
    )

    foreach ($source in $sources) {
        Write-Section "Test-ManagedModuleFolder ($($source.Label))"
        . ([scriptblock]::Create((Get-ScriptFunctionText -Path $source.Path -Name 'Test-ManagedModuleFolder')))

        foreach ($case in $expected.GetEnumerator()) {
            $actual = Test-ManagedModuleFolder -ModuleFolder (Join-Path $root $case.Key)
            Assert-That ($actual -eq $case.Value) "$($case.Key) -> managed: $actual"
        }

        $missing = Test-ManagedModuleFolder -ModuleFolder (Join-Path $root 'Does.Not.Exist')
        Assert-That ($missing -eq $false) 'a folder that does not exist -> false, and no error'
        Assert-That ($missing -is [bool]) 'returns a real boolean'
    }

    # The migration script carries its own copy so that it can run by itself. The two
    # must not drift apart
    Write-Section 'The two copies of Test-ManagedModuleFolder'
    $inMain = (Get-ScriptFunctionText -Path $script:MainScript -Name 'Test-ManagedModuleFolder') -replace '(?s)<#.*?#>', ''
    $inMigration = (Get-ScriptFunctionText -Path $script:MigrationScript -Name 'Test-ManagedModuleFolder') -replace '(?s)<#.*?#>', ''
    Assert-That ($inMain -eq $inMigration) 'the code is identical apart from the comment block'

    Write-Section 'Resolve-ModuleVersionFolder'
    foreach ($name in 'ConvertTo-NormalizedVersion', 'Resolve-ModuleVersionFolder') {
        . ([scriptblock]::Create((Get-ScriptFunctionText -Path $script:MainScript -Name $name)))
    }

    $found = Resolve-ModuleVersionFolder -ModulesRoot $root -Name 'Gallery.Module' -Version '2.0.0'
    Assert-That ($found -eq (Join-Path $root 'Gallery.Module\2.0.0')) 'exact match'

    $found = Resolve-ModuleVersionFolder -ModulesRoot $root -Name 'Four.Part' -Version '6.1907.1'
    Assert-That ($found -eq (Join-Path $root 'Four.Part\6.1907.1.0')) 'a version reported as 6.1907.1 finds the folder 6.1907.1.0'

    $found = Resolve-ModuleVersionFolder -ModulesRoot $root -Name 'Three.Part' -Version '5.7.1.0'
    Assert-That ($found -eq (Join-Path $root 'Three.Part\5.7.1')) 'a version reported as 5.7.1.0 finds the folder 5.7.1'

    $found = Resolve-ModuleVersionFolder -ModulesRoot $root -Name 'Gallery.TwoVersions' -Version '1.0.0'
    Assert-That ($found -eq (Join-Path $root 'Gallery.TwoVersions\1.0.0')) 'picks the right one of two versions'

    Assert-That ($null -eq (Resolve-ModuleVersionFolder -ModulesRoot $root -Name 'Gallery.Module' -Version '9.9.9')) 'a version that is not on disk -> null'
    Assert-That ($null -eq (Resolve-ModuleVersionFolder -ModulesRoot $root -Name 'Not.Installed' -Version '1.0.0')) 'a module that is not on disk -> null'
    Assert-That ($null -eq (Resolve-ModuleVersionFolder -ModulesRoot (Join-Path $root 'nowhere') -Name 'X' -Version '1.0.0')) 'a root that is not on disk -> null'

    # The way the prune phase uses it: keep what is on disk, whatever the recorded
    # location claims
    $resources = @(
        [PSCustomObject]@{
            Name = 'Gallery.Module'; Version = [version]'2.0.0'
            InstalledLocation = 'C:\Users\someone\OneDrive\Documents\PowerShell\Modules\Gallery.Module\2.0.0'
        }
        [PSCustomObject]@{
            Name = 'Ghost.Module'; Version = [version]'1.0.0'
            InstalledLocation = (Join-Path $root 'Ghost.Module\1.0.0')
        }
    )
    $kept = @($resources | Where-Object {
        Resolve-ModuleVersionFolder -ModulesRoot $root -Name $_.Name -Version $_.Version
    })
    Assert-That (($kept.Count -eq 1) -and ($kept[0].Name -eq 'Gallery.Module')) 'a stale recorded location is kept, a module missing from disk is dropped'
}
finally {
    if (Test-Path -LiteralPath $root) {
        Remove-Item -LiteralPath $root -Recurse -Force
    }
}

Complete-Tests
