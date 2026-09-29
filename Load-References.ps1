#Requires -Version 7.4
[CmdletBinding()]
param([Parameter(Mandatory)][string]$SiteUrl,[Parameter(Mandatory)][string]$ClientId,[string]$Dir="$env:USERPROFILE\HubTest")
$ErrorActionPreference='Stop'; Import-Module PnP.PowerShell
Connect-PnPOnline -Url $SiteUrl -ClientId $ClientId -Interactive
function ParseDate($s){ $s=([string]$s).Trim(); if(-not $s){return $null}
  foreach($fmt in 'yyyy-MM-dd','dd-MMM-yyyy','yyyy.MM.dd','M/d/yyyy','MM/dd/yyyy'){
    try{ return [datetime]::ParseExact($s,$fmt,[Globalization.CultureInfo]::InvariantCulture) }catch{} }
  try{ return [datetime]$s }catch{ return $null } }
function Cap($s,$n){ $s=([string]$s) -replace '[\r\n\t]',' '; if($s.Length -gt $n){ return $s.Substring(0,$n) } else { return $s } }

Write-Host "Loading APL..." -ForegroundColor Cyan
foreach($i in (Get-PnPListItem -List 'LC Ref APL' -PageSize 2000)){ Remove-PnPListItem -List 'LC Ref APL' -Identity $i.Id -Force|Out-Null }
$apl=Import-Csv "$Dir\apl-all.csv"
foreach($r in $apl){ $exp=ParseDate $r.Expiration
  Add-PnPListItem -List 'LC Ref APL' -Values @{ Title=(Cap $r.TN 200); Vendor=(Cap $r.Vendor 200); Product=(Cap $r.Product 250); Version=(Cap $r.Version 100); Expiration= $(if($exp){$exp.ToString('yyyy-MM-dd')}else{$null}); DeviceType=(Cap $r.DeviceType 200); CoveredModels=[string]$r.CoveredModels }|Out-Null }
Write-Host ("  APL rows: {0}" -f $apl.Count)

Write-Host "Loading NIAP..." -ForegroundColor Cyan
foreach($i in (Get-PnPListItem -List 'LC Ref NIAP' -PageSize 2000)){ Remove-PnPListItem -List 'LC Ref NIAP' -Identity $i.Id -Force|Out-Null }
$niap=Import-Csv "$Dir\niap-all.csv"
foreach($r in $niap){ $cd=ParseDate $r.'Certification Date'; $amd=ParseDate $r.'Assurance Maintenance Date'
  $rawkey= if([string]$r.VID){[string]$r.VID}else{ "$($r.Vendor)|$($r.Product)" }
  Add-PnPListItem -List 'LC Ref NIAP' -Values @{ Title=(Cap $rawkey 200); Product=(Cap "$($r.Vendor) $($r.Product)" 250); CertificationDate= $(if($cd){$cd.ToString('yyyy-MM-dd')}else{$null}); AssuranceMaintenanceDate= $(if($amd){$amd.ToString('yyyy-MM-dd')}else{$null}); TechTypes=(Cap $r.'Tech Types' 250) }|Out-Null }
Write-Host ("  NIAP rows: {0}" -f $niap.Count)

Write-Host "Loading CSfC..." -ForegroundColor Cyan
foreach($i in (Get-PnPListItem -List 'LC Ref CSfC' -PageSize 2000)){ Remove-PnPListItem -List 'LC Ref CSfC' -Identity $i.Id -Force|Out-Null }
$csfc=Import-Csv "$Dir\csfc-all.csv"
foreach($r in $csfc){ $cd=ParseDate $r.CertificationDate
  Add-PnPListItem -List 'LC Ref CSfC' -Values @{ Title=(Cap $r.Category 200); Vendor=(Cap $r.Vendor 200); Model=(Cap $r.Model 250); Version=(Cap $r.Version 100); CertificationDate= $(if($cd){$cd.ToString('yyyy-MM-dd')}else{$null}); RefreshedOn=(Get-Date).ToString('yyyy-MM-dd') }|Out-Null }
Write-Host ("  CSfC rows: {0}" -f $csfc.Count)

# Searchable index (keep full model text for matching, not the capped copy)
$aplIdx=@(); foreach($r in $apl){ $exp=ParseDate $r.Expiration; if($exp){ $aplIdx+=[pscustomobject]@{ text="$($r.Product) $($r.CoveredModels)"; date=$exp; src='APL'; key=[string]$r.TN } } }
$niapIdx=@(); foreach($r in $niap){ $amd=ParseDate $r.'Assurance Maintenance Date'; if($amd){ $niapIdx+=[pscustomobject]@{ text=[string]$r.Product; date=$amd; src='NIAP'; key= $(if([string]$r.VID){[string]$r.VID}else{"$($r.Vendor)|$($r.Product)"}) } } }
$csfcIdx=@(); foreach($r in $csfc){ $cd=ParseDate $r.CertificationDate; $csfcIdx+=[pscustomobject]@{ text=[string]$r.Model; date=$cd; src='CSfC'; key=[string]$r.Category } }
$allIdx=$aplIdx+$niapIdx+$csfcIdx

Write-Host "Matching models..." -ForegroundColor Cyan
foreach($i in (Get-PnPListItem -List 'LC Model Matches' -PageSize 2000)){ Remove-PnPListItem -List 'LC Model Matches' -Identity $i.Id -Force|Out-Null }
$models=Get-PnPListItem -List 'LC Models' -PageSize 5000 -Fields 'ID','Title','ModelNumber'
$matched=0
foreach($m in $models){
  $mn=([string]$m['ModelNumber']).Trim(); if($mn.Length -lt 3){ continue }
  $rx='(^|[^A-Za-z0-9])'+[Regex]::Escape($mn)+'([^A-Za-z0-9]|$)'
  $hits=@(); foreach($e in $allIdx){ if($e.text -and [Regex]::IsMatch($e.text,$rx)){ $hits+=$e } }
  if($hits.Count -eq 0){ continue }
  $earliest=($hits|Where-Object{$_.date}|Sort-Object date|Select-Object -First 1)
  foreach($h in $hits){ Add-PnPListItem -List 'LC Model Matches' -Values @{ Title=[string]$m['Title']; RefList=$h.src; RefKey=(Cap $h.key 200); MatchMethod='auto'; Confirmed='No' }|Out-Null }
  if($earliest){ Set-PnPListItem -List 'LC Models' -Identity $m.Id -Values @{ PlatformEndOfSupport=$earliest.date.ToString('yyyy-MM-dd') }|Out-Null }
  $matched++
}
Write-Host ("Matched {0} of {1} models to a reference entry." -f $matched,$models.Count) -ForegroundColor Green
Disconnect-PnPOnline
