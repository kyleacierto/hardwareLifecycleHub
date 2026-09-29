#Requires -Version 7.4
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$SiteUrl,
    [Parameter(Mandatory)][string]$ClientId,
    [string]$CsvPath,
    [string]$ImportFileName = 'fake-cmdb.csv',
    [switch]$WhatIfCounts
)
$ErrorActionPreference = 'Stop'
Import-Module PnP.PowerShell
Write-Host "Connecting..." -ForegroundColor Cyan
Connect-PnPOnline -Url $SiteUrl -ClientId $ClientId -Interactive

if (-not $CsvPath) {
    $rel = ([Uri]$SiteUrl).AbsolutePath.TrimEnd('/')
    $CsvPath = Join-Path $env:TEMP $ImportFileName
    Get-PnPFile -Url "$rel/LC Imports/$ImportFileName" -Path $env:TEMP -Filename $ImportFileName -AsFile -Force
}
$rows = Import-Csv -Path $CsvPath
Write-Host ("Rows read: {0}" -f $rows.Count)

$statusMap = @{}
foreach ($m in (Get-PnPListItem -List 'LC Status Map' -PageSize 500)) { $statusMap[[string]$m['Title']] = [string]$m['StatusCategory'] }
$modelBand = @{}; $knownModels = @{}
foreach ($m in (Get-PnPListItem -List 'LC Models' -PageSize 5000)) { $mk=[string]$m['Title']; $knownModels[$mk]=$true; $modelBand[$mk]=[string]$m['LifecycleBand'] }
$existingDevices = @{}
$fields = 'ID','Title','ObjectType','Base','Building','Enclave','ServiceComponent','ModelKey','Manufacturer','ModelNumber','SerialNumber','OperatingSystem','CIStatusRaw','StatusCategory','LifecycleBand'
foreach ($d in (Get-PnPListItem -List 'LC Devices' -PageSize 5000 -Fields $fields)) { $existingDevices[[string]$d['Title']] = $d }
Write-Host ("Existing devices: {0} ; known models: {1}" -f $existingDevices.Count, $knownModels.Count)

$counts = [ordered]@{ Active=0; Inventory=0; Retired=0; Unmapped=0; Unknown=0 }
$dq = New-Object System.Collections.Generic.List[object]
$toAdd = New-Object System.Collections.Generic.List[object]
$toUpdate = New-Object System.Collections.Generic.List[object]
$newModels = @{}
$seenKeys = New-Object System.Collections.Generic.HashSet[string]
$dupCount = 0
$now = (Get-Date).ToString('yyyy-MM-dd')

foreach ($r in $rows) {
    $mfg = ([string]$r.'Manufacturer').Trim()
    $serial = ([string]$r.'Serial Number').Trim()
    $rawStatus = ([string]$r.'CI Status').Trim()
    $modelNum = ([string]$r.'Model Number').Trim()
    $cat = if ($statusMap.ContainsKey($rawStatus) -and $rawStatus -ne '') { $statusMap[$rawStatus] } else { 'Unmapped' }
    if ($cat -eq 'Unmapped') { $counts.Unmapped++; $dq.Add([pscustomobject]@{ IssueType='Unmapped status'; Detail="CI Status '$rawStatus' for $mfg $modelNum $serial" }); continue }
    if ([string]::IsNullOrWhiteSpace($serial)) { $dq.Add([pscustomobject]@{ IssueType='Blank serial'; Detail="$mfg $modelNum ($rawStatus) has no serial; skipped" }); continue }
    $key = "$mfg|$serial"
    if (-not $seenKeys.Add($key)) { $dupCount++; $dq.Add([pscustomobject]@{ IssueType='Duplicate key'; Detail="$key appears more than once" }); continue }
    $counts[$cat]++
    $modelKey = "$mfg|$modelNum"
    $band = if ($knownModels.ContainsKey($modelKey)) { $modelBand[$modelKey] } else { 'Unknown' }
    if ($band -eq 'Unknown') { $counts.Unknown++; if (-not $knownModels.ContainsKey($modelKey) -and -not $newModels.ContainsKey($modelKey)) { $newModels[$modelKey]=[pscustomobject]@{ ObjectType=([string]$r.'Object Type').Trim(); Manufacturer=$mfg; ModelNumber=$modelNum } } }
    $vals = @{ Title=$key; ObjectType=([string]$r.'Object Type').Trim(); Base=([string]$r.'Base').Trim(); Building=([string]$r.'Building').Trim(); Enclave=([string]$r.'Enclave').Trim(); ServiceComponent=([string]$r.'Service Component').Trim(); ModelKey=$modelKey; Manufacturer=$mfg; ModelNumber=$modelNum; SerialNumber=$serial; OperatingSystem=([string]$r.'Operating System').Trim(); CIStatusRaw=$rawStatus; StatusCategory=$cat; LifecycleBand=$band; LastSeen=$now }
    if ($existingDevices.ContainsKey($key)) {
        $ex = $existingDevices[$key]; $changed = $false
        foreach ($f in 'ObjectType','Base','Building','Enclave','ServiceComponent','ModelKey','OperatingSystem','CIStatusRaw','StatusCategory','LifecycleBand') { if ([string]$ex[$f] -ne [string]$vals[$f]) { $changed=$true; break } }
        if ($changed) { $vals['ID']=$ex.Id; $toUpdate.Add($vals) }
    } else { $vals['FirstSeen']=$now; $toAdd.Add($vals) }
}
$retireIds = New-Object System.Collections.Generic.List[int]
foreach ($k in $existingDevices.Keys) { if (-not $seenKeys.Contains($k)) { $retireIds.Add([int]$existingDevices[$k].Id) } }

Write-Host "`n--- Counts ---" -ForegroundColor Cyan
$counts.GetEnumerator() | ForEach-Object { Write-Host ("  {0,-10} {1}" -f $_.Key, $_.Value) }
Write-Host ("  {0,-12} {1}" -f 'Blank serial', ($dq | Where-Object IssueType -eq 'Blank serial').Count)
Write-Host ("  {0,-12} {1}" -f 'Duplicates', $dupCount)
Write-Host ("`nPlanned: {0} add, {1} update, {2} retire; {3} new models; {4} data-quality issues" -f $toAdd.Count, $toUpdate.Count, $retireIds.Count, $newModels.Count, $dq.Count)
if ($WhatIfCounts) { Write-Host "`nDRY RUN: nothing written." -ForegroundColor Yellow; Disconnect-PnPOnline; return }

if ($newModels.Count -gt 0) {
    Write-Host "`nAdding $($newModels.Count) new models..." -ForegroundColor Cyan
    foreach ($mk in $newModels.Keys) { $nm=$newModels[$mk]; Add-PnPListItem -List 'LC Models' -Values @{ Title=$mk; ObjectType=$nm.ObjectType; Manufacturer=$nm.Manufacturer; ModelNumber=$nm.ModelNumber; LifecycleBand='Unknown'; DatesSource='needs lifecycle data' } | Out-Null }
}
function Write-Batch($list,$items,$verb) {
    if ($items.Count -eq 0) { return }
    Write-Host ("{0} {1} device rows..." -f $verb,$items.Count) -ForegroundColor Cyan
    $i=0; foreach ($v in $items) { $id=$v['ID']; $v.Remove('ID')|Out-Null; if ($id) { Set-PnPListItem -List $list -Identity $id -Values $v|Out-Null } else { Add-PnPListItem -List $list -Values $v|Out-Null }; $i++; if ($i % 500 -eq 0) { Write-Host ("  {0}/{1}" -f $i,$items.Count) } }
}
Write-Batch 'LC Devices' $toAdd 'Adding'
Write-Batch 'LC Devices' $toUpdate 'Updating'
if ($retireIds.Count -gt 0) { Write-Host ("Marking {0} retired..." -f $retireIds.Count) -ForegroundColor Cyan; foreach ($id in $retireIds) { Set-PnPListItem -List 'LC Devices' -Identity $id -Values @{ StatusCategory='Retired'; LastChangeType='Retired'; RetiredOn=$now }|Out-Null } }
$dqCap=200; $k=0; foreach ($issue in $dq) { if ($k -ge $dqCap) { break }; Add-PnPListItem -List 'LC Data Quality' -Values @{ Title=$issue.IssueType; IssueType=$issue.IssueType; Detail=$issue.Detail; FoundOn=$now; Resolved='No' }|Out-Null; $k++ }
Add-PnPListItem -List 'LC Import Log' -Values @{ Title="Import $now"; RunOn=$now; RowsRead=$rows.Count; ActiveRows=$counts.Active; RetiredRows=$counts.Retired; UnmappedRows=$counts.Unmapped; NewRows=$toAdd.Count; ChangedRows=$toUpdate.Count; ErrorCount=$dq.Count }|Out-Null
Write-Host "`nDone. Active should now be $($counts.Active)." -ForegroundColor Green
Disconnect-PnPOnline
