#Requires -Version 7.4
[CmdletBinding()]
param([Parameter(Mandatory)][string]$SiteUrl,[Parameter(Mandatory)][string]$ClientId)
$ErrorActionPreference='Stop'; Import-Module PnP.PowerShell
Write-Host "Connecting..." -ForegroundColor Cyan
Connect-PnPOnline -Url $SiteUrl -ClientId $ClientId -Interactive
$today=Get-Date; $now=$today.ToString('yyyy-MM-dd')
$settings=@{}; foreach($s in (Get-PnPListItem -List 'LC Settings' -PageSize 200)){$settings[[string]$s['Title']]=[string]$s['SettingValue']}
$fyStart=[int]($settings['FiscalYearStartMonth'] ?? '10')
function Get-FY([datetime]$d){ if($d.Month -ge $fyStart){$d.Year+1}else{$d.Year} }
$curFY=Get-FY $today
$bands=@(); foreach($b in (Get-PnPListItem -List 'LC Lifecycle Bands' -PageSize 200)){
 $bands+=[pscustomobject]@{ Name=[string]$b['Title']; Order=[int]($b['BandOrder'] ?? 999)
  Min= if($b['MinMonths'] -ne $null -and "$($b['MinMonths'])" -ne ''){[double]$b['MinMonths']}else{$null}
  Max= if($b['MaxMonths'] -ne $null -and "$($b['MaxMonths'])" -ne ''){[double]$b['MaxMonths']}else{$null} } }
function Band-ForMonths($m){ foreach($b in ($bands|Where-Object{$_.Name -ne 'Unknown'}|Sort-Object Order)){
 $okMin=($b.Min -eq $null)-or($m -ge $b.Min); $okMax=($b.Max -eq $null)-or($m -lt $b.Max); if($okMin -and $okMax){return $b.Name} } 'Supported' }
$modelLDoS=@{}; $modelRepl=@{}
foreach($m in (Get-PnPListItem -List 'LC Models' -PageSize 5000)){ $mk=[string]$m['Title']
 $eff=$null; foreach($x in @($m['LastDateOfSupport'],$m['PlatformEndOfSupport'])){ if($x){$dt=[datetime]$x; if($eff -eq $null -or $dt -lt $eff){$eff=$dt}} }
 if($eff){$modelLDoS[$mk]=$eff}; $modelRepl[$mk]=[string]$m['ReplacementModel'] }
$msrp=@{}; foreach($rc in (Get-PnPListItem -List 'LC Replacement Catalog' -PageSize 1000)){ $msrp[[string]$rc['Title']]= if($rc['MSRPUSD']){[double]$rc['MSRPUSD']}else{0} }
$devFields='ID','ObjectType','Base','Enclave','ServiceComponent','ModelKey','StatusCategory'
$devs=Get-PnPListItem -List 'LC Devices' -PageSize 5000 -Fields $devFields
Write-Host ("Devices read: {0}" -f $devs.Count)
$dims=@{ Enclave=@{}; Base=@{}; ServiceComponent=@{}; ObjectType=@{} }
$fyBucket=@{}; $activeTotal=0;$inventoryTotal=0;$retiredTotal=0;$supported=0;$unsupported=0;$soon=0;$unknownDevices=0; $modelActive=@{}
foreach($d in $devs){ $cat=[string]$d['StatusCategory']
 if($cat -eq 'Inventory'){$inventoryTotal++;continue}; if($cat -eq 'Retired'){$retiredTotal++;continue}; if($cat -ne 'Active'){continue}
 $activeTotal++; $mk=[string]$d['ModelKey']; if(-not $modelActive.ContainsKey($mk)){$modelActive[$mk]=0}; $modelActive[$mk]++
 if($modelLDoS.ContainsKey($mk)){ $months=($modelLDoS[$mk]-$today).TotalDays/30.44
   $band= if($months -lt 0){'Unsupported'}else{Band-ForMonths $months}
   $fy= if($months -lt 0){'Past LDoS'}else{'FY'+((Get-FY $modelLDoS[$mk])%100).ToString('00')} }
 else{ $band='Unknown'; $unknownDevices++; $fy='Unknown' }
 if($band -eq 'Unsupported'){$unsupported++}elseif($band -in @('Replace','Plan','Supported')){$supported++}
 if($band -eq 'Replace'){$soon++}
 if(-not $fyBucket.ContainsKey($fy)){$fyBucket[$fy]=@{count=0;cost=0.0}}; $fyBucket[$fy].count++
 $repl=$modelRepl[$mk]; if($repl -and $msrp.ContainsKey($repl)){$fyBucket[$fy].cost+=$msrp[$repl]}
 foreach($dim in $dims.Keys){ $val=[string]$d[$dim]; if($val -eq ''){$val='Unassigned'}
   if(-not $dims[$dim].ContainsKey($val)){$dims[$dim][$val]=@{active=0;unsup=0;soon=0}}
   $dims[$dim][$val].active++; if($band -eq 'Unsupported'){$dims[$dim][$val].unsup++}; if($band -eq 'Replace'){$dims[$dim][$val].soon++} } }
$pctSupported= if($activeTotal -gt 0){[math]::Round(100.0*$supported/$activeTotal,1)}else{0}
function Clear-List($name){ foreach($i in (Get-PnPListItem -List $name -PageSize 2000)){Remove-PnPListItem -List $name -Identity $i.Id -Force|Out-Null} }
Write-Host "Clearing old summaries..." -ForegroundColor Cyan
'LC Sum Headline','LC Sum Counts','LC Sum Outlook','LC Attention'|ForEach-Object{Clear-List $_}
$headline=@(
 @{k='ActiveDevices';v=$activeTotal;t='In Production, Active, Deployment, Action Needed'},
 @{k='PercentSupported';v=$pctSupported;t="$supported of $activeTotal before end of support"},
 @{k='UnsupportedNow';v=$unsupported;t='Past last date of support'},
 @{k='Soon12mo';v=$soon;t="Reach end of support in FY$(($curFY)%100)"},
 @{k='MissingLifecycle';v=$unknownDevices;t='Active devices on models with no lifecycle date'},
 @{k='Inventory';v=$inventoryTotal;t='Warehouse stock, not fielded'} )
foreach($h in $headline){Add-PnPListItem -List 'LC Sum Headline' -Values @{Title=$now;MetricKey=$h.k;MetricValue=$h.v;MetricText=$h.t}|Out-Null}
foreach($dim in $dims.Keys){ foreach($val in $dims[$dim].Keys){ $c=$dims[$dim][$val]
 Add-PnPListItem -List 'LC Sum Counts' -Values @{Title=$now;Dimension=$dim;MemberValue=$val;ActiveCount=$c.active;UnsupCount=$c.unsup;SoonCount=$c.soon}|Out-Null } }
foreach($fy in $fyBucket.Keys){ $b=$fyBucket[$fy]
 Add-PnPListItem -List 'LC Sum Outlook' -Values @{Title=$now;FiscalYear=$fy;DeviceCount=$b.count;MSRPCost=[int]$b.cost;BandName=''}|Out-Null }
$rank=1; $attn=New-Object System.Collections.Generic.List[object]
$unsupModels=$modelActive.GetEnumerator()|Where-Object{$modelLDoS.ContainsKey($_.Key) -and $modelLDoS[$_.Key] -lt $today}|Sort-Object Value -Descending|Select-Object -First 3
foreach($um in $unsupModels){$attn.Add(@{c='Unsupported';d="$($um.Key) — $($um.Value) active devices past end of support"})}
$attn.Add(@{c='Data';d="$($modelActive.Keys.Count) models in inventory; $unknownDevices active devices on models with no lifecycle date"})
$dqOpen=(Get-PnPListItem -List 'LC Data Quality' -PageSize 2000|Where-Object{[string]$_['Resolved'] -ne 'Yes'}).Count
$attn.Add(@{c='Data';d="$dqOpen open data-quality issues (unmapped statuses, blank serials, duplicates)"})
foreach($a in $attn){Add-PnPListItem -List 'LC Attention' -Values @{Title=$a.c;Category=$a.c;Detail=$a.d;RankOrder=$rank;LinkUrl=''}|Out-Null;$rank++}
Write-Host "`n--- Summary ---" -ForegroundColor Green
Write-Host ("Active {0} | Supported {1}% | Unsupported {2} | FY{3} due {4} | Unknown {5} | Inventory {6} | Retired {7}" -f $activeTotal,$pctSupported,$unsupported,($curFY%100),$soon,$unknownDevices,$inventoryTotal,$retiredTotal)
Write-Host ("Outlook buckets: {0} | Members: Enclave {1}, Base {2}, SvcComp {3}, ObjType {4}" -f $fyBucket.Keys.Count,$dims.Enclave.Keys.Count,$dims.Base.Keys.Count,$dims.ServiceComponent.Keys.Count,$dims.ObjectType.Keys.Count)
Disconnect-PnPOnline
