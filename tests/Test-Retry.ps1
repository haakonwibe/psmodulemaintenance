<#
.SYNOPSIS
    Tests for retrying a module update after a network fault.

.DESCRIPTION
    Covers Get-TransientNetworkFault and Invoke-ModuleUpdateWithRetry. Invoke-ModuleUpdate
    is replaced by a stand-in, so nothing is installed and the network is never used.
#>
param([switch]$Quiet)

. (Join-Path $PSScriptRoot 'TestHelpers.ps1')

foreach ($name in 'Get-TransientNetworkFault', 'Invoke-ModuleUpdateWithRetry') {
    . ([scriptblock]::Create((Get-ScriptFunctionText -Path $script:MainScript -Name $name)))
}

# --- Stand-ins -----------------------------------------------------------------------
$script:LogLines = @()
function Write-Log {
    param([Parameter(Position = 0)][string]$Message, [string]$Level = 'INFO')
    $script:LogLines += "[$Level] $Message"
}

$script:Calls = @()
$script:Behaviour = { }
function Invoke-ModuleUpdate {
    [CmdletBinding()]
    param(
        [string]$Name, [string]$Scope, [bool]$TrustRepository = $true,
        [int]$TimeoutSeconds = 600, [string]$Version, [switch]$Prerelease
    )
    $script:Calls += , @{ Bound = [hashtable]::new($PSBoundParameters) }
    & $script:Behaviour $script:Calls.Count
}

function Reset-Test {
    $script:Calls = @()
    $script:LogLines = @()
}

# An error the way it arrives from the isolated runspace: text inside a
# MethodInvocationException, with no typed exception chain underneath
function New-WrappedError {
    param([string]$Inner)
    $message = 'Exception calling "EndInvoke" with "1" argument(s): "' + $Inner + '"'
    return [System.Management.Automation.MethodInvocationException]::new($message)
}

# --- Recorded messages ---------------------------------------------------------------
# The first is the error that made a scheduled run skip its updates, word for word apart
# from the module name. The next two were produced on purpose, behind a proxy that
# refuses connections and behind one whose host name does not resolve.
$tlsReset = @'
The running command stopped because the preference variable "ErrorActionPreference" or common parameter is set to Stop: 'The SSL connection could not be established, see inner exception.' Request sent: 'https://www.powershellgallery.com/api/v2/FindPackagesById()?%24filter=Id+eq+%27Contoso.Tools%27+and+IsLatestVersion+eq+true&%24inlinecount=allpages&id=%27Contoso.Tools%27' Inner exception: 'Unable to read data from the transport connection: An existing connection was forcibly closed by the remote host..'
'@
$refused = "'No connection could be made because the target machine actively refused it. (127.0.0.1:9)' Request sent: 'https://www.powershellgallery.com/api/v2/FindPackagesById()'"
$noSuchHost = "'No such host is known. (proxy.invalid:8080)' Request sent: 'https://www.powershellgallery.com/api/v2/'"

# --- Get-TransientNetworkFault -------------------------------------------------------
Write-Section 'Get-TransientNetworkFault'

$wrapped = (New-WrappedError $tlsReset).Message
$fault = Get-TransientNetworkFault -Message $wrapped
Assert-That ($fault -eq 'SSL connection could not be established') "recorded TLS reset is recognised -> '$fault'"

$shouldMatch = [ordered]@{
    'connection refused'  = $refused
    'host not resolved'   = $noSuchHost
    'gateway 502'         = 'Response status code does not indicate success: 502 (Bad Gateway).'
    'gateway 503'         = 'Response status code does not indicate success: 503 (Service Unavailable).'
    'gateway 504'         = 'Response status code does not indicate success: 504 (Gateway Timeout).'
    'throttled 429'       = 'Response status code does not indicate success: 429 (Too Many Requests).'
    'http client timeout' = 'The request was canceled due to the configured HttpClient.Timeout of 100 seconds elapsing.'
    'connection aborted'  = 'An established connection was aborted by the software in your host machine.'
}
foreach ($case in $shouldMatch.GetEnumerator()) {
    Assert-That ([bool](Get-TransientNetworkFault -Message $case.Value)) "matches: $($case.Key)"
}

$shouldNotMatch = [ordered]@{
    'locked folder'        = 'Cannot remove package path C:\Program Files\PowerShell\Modules\Contoso.Tools\1.0.0. The previous version could not be removed.'
    'our own timeout'      = "Operation timed out after 605s (limit 600s) for module 'Contoso.Tools'"
    'access denied'        = "Access to the path 'Contoso.dll' is denied."
    'package not found'    = "Package 'Does.Not.Exist' could not be found in repository 'PSGallery'."
    'not found 404'        = 'Response status code does not indicate success: 404 (Not Found).'
    'server error 500'     = 'Response status code does not indicate success: 500 (Internal Server Error).'
    'needs admin'          = 'Install-PSResource requires administrator rights for AllUsers scope.'
    'module name with 503' = "Package 'Contoso.503.Tools' version 5.0.4 has a dependency conflict."
    'empty message'        = ''
}
foreach ($case in $shouldNotMatch.GetEnumerator()) {
    Assert-That ($null -eq (Get-TransientNetworkFault -Message $case.Value)) "ignores: $($case.Key)"
}

# --- Invoke-ModuleUpdateWithRetry ----------------------------------------------------
Write-Section 'Invoke-ModuleUpdateWithRetry'

Reset-Test
$script:Behaviour = { }
Invoke-ModuleUpdateWithRetry -Name 'Mod.A' -Scope 'AllUsers' -RetryDelaySeconds 0, 0
Assert-That ($script:Calls.Count -eq 1) 'success first time: 1 call'
Assert-That ($script:LogLines.Count -eq 0) 'success first time: nothing logged'

Reset-Test
$script:Behaviour = { param($n) if ($n -le 2) { throw (New-WrappedError $tlsReset) } }
$threw = $false
try { Invoke-ModuleUpdateWithRetry -Name 'Mod.B' -Scope 'AllUsers' -RetryDelaySeconds 0, 0 } catch { $threw = $true }
$warnings = @($script:LogLines | Where-Object { $_ -like '`[WARN`]*' })
Assert-That (-not $threw) 'two faults then success: does not throw'
Assert-That ($script:Calls.Count -eq 3) "two faults then success: 3 calls (got $($script:Calls.Count))"
Assert-That ($warnings.Count -eq 2) 'two faults then success: 2 WARN lines'
Assert-That ($warnings[0] -eq '[WARN] Network fault on Mod.B (attempt 1 of 3): SSL connection could not be established. Retrying in 0s') 'WARN line reads as documented'

Reset-Test
$script:Behaviour = { throw (New-WrappedError $tlsReset) }
$caught = $null
try { Invoke-ModuleUpdateWithRetry -Name 'Mod.C' -RetryDelaySeconds 0, 0 } catch { $caught = $_ }
Assert-That ($script:Calls.Count -eq 3) "persistent fault: 3 calls, then gives up (got $($script:Calls.Count))"
Assert-That ($null -ne $caught) 'persistent fault: throws'
Assert-That ($caught.Exception -is [System.Management.Automation.MethodInvocationException]) 'persistent fault: exception type preserved'
Assert-That ($caught.Exception.Message -eq (New-WrappedError $tlsReset).Message) 'persistent fault: message preserved word for word'
Assert-That (@($script:LogLines).Count -eq 2) 'persistent fault: 2 WARN lines, none for the last attempt'

# A timeout already cost the module its whole time budget. It must not be retried, and
# it must still land in the typed catch the callers use
Reset-Test
$script:Behaviour = { throw [System.TimeoutException]::new("Operation timed out after 605s (limit 600s) for module 'Mod.D'") }
$branch = 'none'
try { Invoke-ModuleUpdateWithRetry -Name 'Mod.D' -RetryDelaySeconds 0, 0 }
catch [System.TimeoutException] { $branch = 'timeout' }
catch { $branch = 'generic' }
Assert-That ($script:Calls.Count -eq 1) "timeout: 1 call, no retry (got $($script:Calls.Count))"
Assert-That ($branch -eq 'timeout') "timeout: lands in catch [System.TimeoutException] (got '$branch')"

# Update-AllModules recovers from a locked folder by matching on the message
Reset-Test
$lockedMessage = 'Cannot remove package path C:\Program Files\PowerShell\Modules\Contoso.Tools\1.0.0. The previous version could not be removed.'
$script:Behaviour = { throw (New-WrappedError $lockedMessage) }
$caught = $null
try { Invoke-ModuleUpdateWithRetry -Name 'Mod.E' -RetryDelaySeconds 0, 0 } catch { $caught = $_ }
Assert-That ($script:Calls.Count -eq 1) "locked folder: 1 call, no retry (got $($script:Calls.Count))"
Assert-That ($caught.Exception.Message -match 'Cannot remove package path\s+(.+?)\.?\s*(The previous|$)') 'locked folder: the caller''s pattern still matches'

Reset-Test
$script:Behaviour = { }
Invoke-ModuleUpdateWithRetry -Name 'Mod.F' -Scope 'AllUsers' -TrustRepository $false -TimeoutSeconds 42 `
    -Version '1.2.3-beta1' -Prerelease:$true -RetryDelaySeconds 0
$bound = $script:Calls[0].Bound
Assert-That (($bound.Name -eq 'Mod.F') -and ($bound.Scope -eq 'AllUsers')) 'forwarding: Name and Scope'
Assert-That (($bound.TrustRepository -eq $false) -and ($bound.TimeoutSeconds -eq 42)) 'forwarding: TrustRepository and TimeoutSeconds'
Assert-That (($bound.Version -eq '1.2.3-beta1') -and [bool]$bound.Prerelease) 'forwarding: Version and Prerelease'
Assert-That (-not $bound.ContainsKey('RetryDelaySeconds')) 'forwarding: RetryDelaySeconds is not passed down'

# Update-AllModules passes a null scope when OneDrive is not in play
Reset-Test
$script:Behaviour = { }
$threw = $false
try { Invoke-ModuleUpdateWithRetry -Name 'Mod.G' -Scope $null -TrustRepository $true -TimeoutSeconds 600 } catch { $threw = $true }
Assert-That ((-not $threw) -and ($script:Calls.Count -eq 1)) 'null scope: accepted'

$definition = Get-ScriptFunctionAst -Path $script:MainScript -Name 'Invoke-ModuleUpdateWithRetry'
$retryParameter = $definition.Body.ParamBlock.Parameters |
    Where-Object { $_.Name.VariablePath.UserPath -eq 'RetryDelaySeconds' }
$defaultText = $retryParameter.DefaultValue.Extent.Text
Assert-That ($defaultText -eq '@(5, 15)') "default delays are $defaultText"

Reset-Test
$script:Behaviour = { param($n) if ($n -le 1) { throw (New-WrappedError $tlsReset) } }
$timer = [System.Diagnostics.Stopwatch]::StartNew()
Invoke-ModuleUpdateWithRetry -Name 'Mod.H' -RetryDelaySeconds 2
$waited = $timer.Elapsed.TotalSeconds
Assert-That (($waited -ge 2) -and ($waited -lt 4)) "delay is honoured: waited $([math]::Round($waited, 1))s for a 2s delay"

Reset-Test
$script:Behaviour = { throw (New-WrappedError $tlsReset) }
try { Invoke-ModuleUpdateWithRetry -Name 'Mod.I' -RetryDelaySeconds @() } catch { $null = $_ }
Assert-That ($script:Calls.Count -eq 1) "empty delay list turns retrying off (got $($script:Calls.Count) call)"

Complete-Tests
