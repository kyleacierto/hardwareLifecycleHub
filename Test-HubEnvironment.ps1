#Requires -Version 7.4
<#
.SYNOPSIS
    Pre-build environment check for the lifecycle / test bed hub.

.DESCRIPTION
    Confirms this machine can run the hub's PowerShell engine:
      - PowerShell 7, full language mode, normal (non-elevated) session
      - Managed device (AD- or Entra-joined), not a personal machine
      - Documents folder not redirected to a personal OneDrive
      - Read/write on the CUI file share, and running a script from it
      - Bundling the ImportExcel module and writing an .xlsx
      - Rendering a PDF with headless Microsoft Edge
      - Reaching the SharePoint site over HTTPS (optional)

    Uses no real data. The only thing it creates on the share is a temporary
    folder named _hubtest_xxxxxxxx, which it always deletes before exiting.

.PARAMETER SharePath
    UNC path to the folder on the CUI file share where the hub will live,
    for example \\server\share\Hub. Use the UNC form, not a mapped drive letter.

.PARAMETER SiteUrl
    Optional. SharePoint site URL, used only to confirm HTTPS reachability.
    No sign-in is attempted.

.PARAMETER ImportExcelVersion
    ImportExcel version to test bundling with. Default 7.8.10.

.EXAMPLE
    pwsh -NoProfile -File .\Test-HubEnvironment.ps1 -SharePath '\\server\share\Hub' -SiteUrl 'https://tenant.sharepoint.com/sites/Example'
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$SharePath,
    [string]$SiteUrl,
    [string]$ImportExcelVersion = '7.8.10'
)

# Language mode first, using only constructs that work in Constrained Language Mode.
if ($ExecutionContext.SessionState.LanguageMode -ne 'FullLanguage') {
    Write-Host "FAIL  Language mode is $($ExecutionContext.SessionState.LanguageMode). The PowerShell engine cannot run on this machine." -ForegroundColor Red
    exit 1
}

$ErrorActionPreference = 'Stop'
$results = [System.Collections.Generic.List[object]]::new()
function Add-Result {
    param([string]$Check, [ValidateSet('PASS', 'WARN', 'FAIL', 'SKIP')][string]$Status, [string]$Detail)
    $results.Add([pscustomobject]@{ Check = $Check; Status = $Status; Detail = $Detail })
}

$runId   = [guid]::NewGuid().ToString('N').Substring(0, 8)
$tempDir = Join-Path ([IO.Path]::GetTempPath()) "hubtest_$runId"
$testDir = $null

try {
    New-Item -ItemType Directory -Path $tempDir -Force | Out-Null

    # 1. PowerShell and session
    Add-Result 'PowerShell version' 'PASS' "$($PSVersionTable.PSVersion) ($($PSVersionTable.PSEdition))"
    Add-Result 'Language mode' 'PASS' 'FullLanguage'

    try {
        $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
            [Security.Principal.WindowsBuiltInRole]::Administrator)
        if ($isAdmin) { Add-Result 'Normal (non-elevated) session' 'WARN' 'Running as administrator. Rerun in a normal window.' }
        else          { Add-Result 'Normal (non-elevated) session' 'PASS' 'Normal user session' }
    }
    catch { Add-Result 'Normal (non-elevated) session' 'WARN' "Could not determine: $($_.Exception.Message)" }

    # 2. Managed device
    try {
        $cs = Get-CimInstance -ClassName Win32_ComputerSystem
        $ds = @{}
        foreach ($line in (& dsregcmd.exe /status 2>$null)) {
            if ($line -match '^\s*(\w+)\s*:\s*(.+?)\s*$') { $ds[$Matches[1]] = $Matches[2] }
        }
        $entraJoined = $ds['AzureAdJoined'] -eq 'YES'
        $registered  = $ds['WorkplaceJoined'] -eq 'YES'
        $mdm = $null
        try {
            $mdm = Get-ChildItem -Path 'HKLM:\SOFTWARE\Microsoft\Enrollments' -ErrorAction Stop |
                Get-ItemProperty -ErrorAction SilentlyContinue |
                Where-Object { $_.ProviderID -eq 'MS DM Server' -and $_.UPN } |
                Select-Object -First 1
        }
        catch { }
        $adPart     = if ($cs.PartOfDomain) { "AD domain $($cs.Domain)" } else { "no AD domain (workgroup $($cs.Workgroup))" }
        $entraPart  = "Entra joined: $($ds['AzureAdJoined'] ?? 'unknown')"
        $tenantPart = if ($ds['TenantName']) { "; tenant: $($ds['TenantName'])" } else { '' }
        $regPart    = "; registered (workplace-joined): $($ds['WorkplaceJoined'] ?? 'unknown')"
        $mdmPart    = if ($mdm) { "; MDM enrolled: yes (@$(($mdm.UPN -split '@')[-1]))" } else { '; MDM enrolled: no' }
        $detail     = "$env:COMPUTERNAME; $adPart; $entraPart$tenantPart$regPart$mdmPart"
        if ($cs.PartOfDomain -or $entraJoined) { Add-Result 'Managed device' 'PASS' $detail }
        elseif ($mdm) { Add-Result 'Managed device' 'WARN' "$detail. Personal device enrolled in company management: CUI use depends on company approval." }
        elseif ($registered) { Add-Result 'Managed device' 'FAIL' "$detail. Registered for sign-in only, not managed: development with fake data only." }
        else { Add-Result 'Managed device' 'FAIL' "$detail. Looks unmanaged: development with fake data only, never real CUI." }
    }
    catch { Add-Result 'Managed device' 'WARN' "Could not determine: $($_.Exception.Message)" }

    # 3. Documents folder redirection
    try {
        $docs = [Environment]::GetFolderPath('MyDocuments')
        if ($docs -match '\\OneDrive\\') {
            Add-Result 'Documents folder' 'WARN' "$docs (personal OneDrive backup). Keep all hub files off this path."
        }
        else { Add-Result 'Documents folder' 'PASS' $docs }
    }
    catch { Add-Result 'Documents folder' 'WARN' "Could not determine: $($_.Exception.Message)" }

    # 4. Share path form
    if ($SharePath -like '\\*') { Add-Result 'Share path is UNC' 'PASS' $SharePath }
    else { Add-Result 'Share path is UNC' 'WARN' "$SharePath. Scheduled tasks can't see mapped drives; use the \\server\share form." }

    # 5. Share read/write
    try {
        if (-not (Test-Path -LiteralPath $SharePath)) { throw "Path not found or no access: $SharePath" }
        Get-ChildItem -LiteralPath $SharePath -Force | Select-Object -First 1 | Out-Null
        $testDir = Join-Path $SharePath "_hubtest_$runId"
        New-Item -ItemType Directory -Path $testDir | Out-Null
        $probe = 'hub environment test (no real data)'
        $txt = Join-Path $testDir 'write-test.txt'
        Set-Content -LiteralPath $txt -Value $probe
        if ((Get-Content -LiteralPath $txt -Raw).Trim() -ne $probe) { throw 'Read-back did not match what was written' }
        Add-Result 'Share read/write' 'PASS' 'Created folder, wrote and read back a test file'
    }
    catch {
        Add-Result 'Share read/write' 'FAIL' $_.Exception.Message
        $testDir = $null
    }

    # 6. Running a script stored on the share (execution policy / zone check)
    if ($testDir) {
        try {
            $ps1 = Join-Path $testDir 'exec-test.ps1'
            Set-Content -LiteralPath $ps1 -Value "'EXEC_OK'"
            $pwshExe = (Get-Command pwsh -CommandType Application -ErrorAction SilentlyContinue | Select-Object -First 1).Source
            if (-not $pwshExe) { $pwshExe = Join-Path $PSHOME 'pwsh.exe' }
            $out = (& $pwshExe -NoProfile -NonInteractive -File $ps1 2>&1 | Out-String).Trim()
            if ($out -match 'EXEC_OK') { Add-Result 'Run script from share' 'PASS' 'Script on the share ran without prompts' }
            else { Add-Result 'Run script from share' 'WARN' "Blocked: $out" }
        }
        catch { Add-Result 'Run script from share' 'WARN' "Blocked: $($_.Exception.Message)" }
    }
    else { Add-Result 'Run script from share' 'SKIP' 'Share not writable' }

    # 7. Bundle ImportExcel and write an .xlsx
    $moduleLoaded = $false
    try {
        $modDir = Join-Path $tempDir 'modules'
        New-Item -ItemType Directory -Path $modDir -Force | Out-Null
        Save-Module -Name ImportExcel -RequiredVersion $ImportExcelVersion -Path $modDir -Repository PSGallery
        Import-Module (Join-Path $modDir "ImportExcel\$ImportExcelVersion\ImportExcel.psd1") -Force
        $moduleLoaded = $true
        Add-Result 'Bundle ImportExcel' 'PASS' "Downloaded and loaded $ImportExcelVersion from a local folder"
    }
    catch {
        Add-Result 'Bundle ImportExcel' 'WARN' "PSGallery download failed ($($_.Exception.Message)). The module folder can be copied over from the dev machine instead."
    }
    if ($moduleLoaded) {
        try {
            $xlsxDir = if ($testDir) { $testDir } else { $tempDir }
            $xlsx = Join-Path $xlsxDir 'excel-test.xlsx'
            [pscustomobject]@{ Model = 'EXAMPLE-1000'; Count = 3; EndOfSupport = (Get-Date).AddYears(1).ToString('yyyy-MM-dd') } |
                Export-Excel -Path $xlsx -WorksheetName 'Test' -AutoSize
            if ((Test-Path -LiteralPath $xlsx) -and (Get-Item -LiteralPath $xlsx).Length -gt 0) {
                Add-Result 'Write .xlsx' 'PASS' "Workbook written to $(if ($testDir) { 'share' } else { 'local temp' })"
            }
            else { throw 'File missing or empty' }
        }
        catch { Add-Result 'Write .xlsx' 'FAIL' $_.Exception.Message }
    }
    else { Add-Result 'Write .xlsx' 'SKIP' 'Module not loaded' }

    # 8. Headless Edge PDF
    try {
        $edge = @(
            "${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe",
            "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe"
        ) | Where-Object { $_ -and (Test-Path -LiteralPath $_) } | Select-Object -First 1
        if (-not $edge) { throw 'msedge.exe not found' }

        $html = Join-Path $tempDir 'pdf-test.html'
        $pdf  = Join-Path $tempDir 'pdf-test.pdf'
        $prof = Join-Path $tempDir 'edge-profile'
        Set-Content -LiteralPath $html -Encoding utf8 -Value @'
<!DOCTYPE html><html><head><meta charset="utf-8"><title>Hub PDF test</title></head>
<body style="font-family:Segoe UI,Arial,sans-serif">
<p style="text-align:center;font-weight:bold">TEST BANNER</p>
<h1>Hub environment test</h1><p>Sample text only. No real data.</p>
</body></html>
'@
        $edgeArgs = @(
            '--headless', '--disable-gpu', '--no-first-run', '--no-pdf-header-footer',
            "--user-data-dir=`"$prof`"", "--print-to-pdf=`"$pdf`"", ([Uri]$html).AbsoluteUri
        )
        $p = Start-Process -FilePath $edge -ArgumentList $edgeArgs -PassThru -WindowStyle Hidden
        if (-not $p.WaitForExit(60000)) { try { $p.Kill($true) } catch {}; throw 'Edge timed out after 60 seconds' }
        $deadline = (Get-Date).AddSeconds(10)
        while (-not (Test-Path -LiteralPath $pdf) -and (Get-Date) -lt $deadline) { Start-Sleep -Milliseconds 500 }
        if (-not ((Test-Path -LiteralPath $pdf) -and (Get-Item -LiteralPath $pdf).Length -gt 0)) {
            throw 'Edge exited but no PDF was produced (possibly blocked by browser policy)'
        }
        if ($testDir) { Copy-Item -LiteralPath $pdf -Destination $testDir }
        Add-Result 'Headless Edge PDF' 'PASS' "Rendered with $edge"
    }
    catch { Add-Result 'Headless Edge PDF' 'FAIL' $_.Exception.Message }

    # 9. SharePoint reachability (no sign-in)
    if ($SiteUrl) {
        try {
            $r = Invoke-WebRequest -Uri $SiteUrl -Method Head -SkipHttpErrorCheck -TimeoutSec 20
            Add-Result 'SharePoint reachable' 'PASS' "HTTPS status $($r.StatusCode) (401/403 is normal here; sign-in is not tested)"
        }
        catch { Add-Result 'SharePoint reachable' 'FAIL' $_.Exception.Message }
    }
    else { Add-Result 'SharePoint reachable' 'SKIP' 'No -SiteUrl given' }
}
catch {
    Add-Result 'Unexpected error' 'FAIL' "$($_.Exception.Message) (line $($_.InvocationInfo.ScriptLineNumber))"
}
finally {
    # Cleanup: always remove what this script created.
    $cleanup = @()
    if ($testDir -and (Test-Path -LiteralPath $testDir)) {
        try { Remove-Item -LiteralPath $testDir -Recurse -Force; $cleanup += 'share test folder removed' }
        catch { $cleanup += "COULD NOT REMOVE $testDir; delete it by hand" }
    }
    if (Test-Path -LiteralPath $tempDir) {
        try { Remove-Item -LiteralPath $tempDir -Recurse -Force; $cleanup += 'local temp removed' }
        catch { $cleanup += "could not remove $tempDir (safe to delete later)" }
    }

    Write-Host ($results | Format-Table -AutoSize -Wrap | Out-String -Width 220)
    $pass = @($results | Where-Object Status -eq 'PASS').Count
    $warn = @($results | Where-Object Status -eq 'WARN').Count
    $fail = @($results | Where-Object Status -eq 'FAIL').Count
    $color = if ($fail) { 'Red' } elseif ($warn) { 'Yellow' } else { 'Green' }
    Write-Host "Summary: $pass PASS, $warn WARN, $fail FAIL" -ForegroundColor $color
    if ($cleanup) { Write-Host "Cleanup: $($cleanup -join '; ')" }
}

if (@($results | Where-Object Status -eq 'FAIL').Count) { exit 1 }
