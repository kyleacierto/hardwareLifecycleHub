#Requires -Version 7.4
[CmdletBinding()]
param([Parameter(Mandatory)][string]$SiteUrl,[Parameter(Mandatory)][string]$ClientId,[string]$OutPath="$env:USERPROFILE\HubTest\dashboard-data.json")
$ErrorActionPreference='Stop'; Import-Module PnP.PowerShell
Connect-PnPOnline -Url $SiteUrl -ClientId $ClientId -Interactive
$headline=@{}; foreach($h in (Get-PnPListItem -List 'LC Sum Headline' -PageSize 200)){$headline[[string]$h['MetricKey']]=@{value=$h['MetricValue'];text=[string]$h['MetricText']}}
$counts=@{}; foreach($c in (Get-PnPListItem -List 'LC Sum Counts' -PageSize 5000)){$dim=[string]$c['Dimension']; if(-not $counts.ContainsKey($dim)){$counts[$dim]=@()}; $counts[$dim]+=[pscustomobject]@{member=[string]$c['MemberValue'];active=[int]$c['ActiveCount'];unsup=[int]$c['UnsupCount'];soon=[int]$c['SoonCount']}}
$countsOut=@{}; foreach($dim in $counts.Keys){ $rows=$counts[$dim]|Sort-Object active -Descending
  if($rows.Count -gt 8){ $top=$rows|Select-Object -First 8; $rest=$rows|Select-Object -Skip 8
    $other=[pscustomobject]@{member="Other ($($rest.Count))";active=($rest|Measure-Object active -Sum).Sum;unsup=($rest|Measure-Object unsup -Sum).Sum;soon=($rest|Measure-Object soon -Sum).Sum}
    $countsOut[$dim]=@($top)+@($other) } else { $countsOut[$dim]=$rows } }
$outlook=@(); foreach($o in (Get-PnPListItem -List 'LC Sum Outlook' -PageSize 200)){$outlook+=[pscustomobject]@{fy=[string]$o['FiscalYear'];count=[int]$o['DeviceCount'];cost=[int]$o['MSRPCost']}}
$attention=@(); foreach($a in (Get-PnPListItem -List 'LC Attention' -PageSize 200)){$attention+=[pscustomobject]@{category=[string]$a['Category'];detail=[string]$a['Detail'];rank=[int]$a['RankOrder']}}
$attention=$attention|Sort-Object rank
$settings=@{}; foreach($s in (Get-PnPListItem -List 'LC Settings' -PageSize 200)){$settings[[string]$s['Title']]=[string]$s['SettingValue']}
([ordered]@{generatedOn=(Get-Date).ToString('yyyy-MM-dd HH:mm');settings=$settings;headline=$headline;counts=$countsOut;outlook=$outlook;attention=$attention}) | ConvertTo-Json -Depth 8 | Set-Content -Path $OutPath -Encoding UTF8
Write-Host "Wrote $OutPath" -ForegroundColor Green
Disconnect-PnPOnline
