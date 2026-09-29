#Requires -Version 7.4
[CmdletBinding()]
param([Parameter(Mandatory)][string]$SiteUrl,[Parameter(Mandatory)][string]$ClientId,[string]$File="$env:USERPROFILE\HubTest\cisco-rfi.xlsx")
$ErrorActionPreference='Stop'; Import-Module PnP.PowerShell; Import-Module ImportExcel
Connect-PnPOnline -Url $SiteUrl -ClientId $ClientId -Interactive
$today=Get-Date
function PD($v){ if(-not $v){return $null}; try{ if($v -is [datetime]){return $v}; return [datetime]$v }catch{ return $null } }
function ToInt($v){ $n=0; if([double]::TryParse((("{0}" -f $v) -replace '[^\d.]',''),[ref]$n)){ return [int]$n } else { return $null } }

$rows = Import-Excel -Path $File -WorksheetName 'HARDWARE RFI'
Write-Host ("RFI rows: {0}" -f $rows.Count)

$models=@{}
foreach($m in (Get-PnPListItem -List 'LC Models' -PageSize 5000 -Fields 'ID','Title','Manufacturer','ModelNumber')){
  $key=("{0}|{1}" -f ([string]$m['Manufacturer']).Trim(),([string]$m['ModelNumber']).Trim())
  $models[$key.ToLower()]=$m
}
$catalog=@{}
foreach($c in (Get-PnPListItem -List 'LC Replacement Catalog' -PageSize 1000 -Fields 'ID','Title')){ $catalog[([string]$c['Title']).Trim().ToLower()]=$true }

$matched=0; $missed=New-Object System.Collections.Generic.List[string]; $catAdd=0
foreach($r in $rows){
  $mfg=([string]$r.'Manufacturer').Trim(); $mn=([string]$r.'Model Number').Trim()
  if(-not $mn){ continue }
  $eos=PD $r.'END OF SUPPORT DATE'; if(-not $eos){ $eos=PD $r.'END OF LIFE DATE' }
  $repl=([string]$r.'SUGGESTED/RECOMMENDED REPLACEMENT').Trim()
  $rmsrp=ToInt $r.'MSRP FOR REPLACEMENT'
  $key=("{0}|{1}" -f $mfg,$mn).ToLower()
  if($models.ContainsKey($key)){
    $mi=$models[$key]
    $vals=@{ DatesSource='Cisco RFI'; DatesConfirmedBy='vendor'; DatesConfirmedOn=$today.ToString('yyyy-MM-dd') }
    if($eos){ $vals['LastDateOfSupport']=$eos.ToString('yyyy-MM-dd') }
    if($repl){ $vals['ReplacementModel']=$repl }
    Set-PnPListItem -List 'LC Models' -Identity $mi.Id -Values $vals | Out-Null
    $matched++
    if($repl -and -not $catalog.ContainsKey($repl.ToLower())){
      $cv=@{ Title=$repl; TestBedValidated='No'; MSRPAsOf=$today.ToString('yyyy-MM-dd') }
      if($rmsrp -ne $null){ $cv['MSRPUSD']=$rmsrp }
      Add-PnPListItem -List 'LC Replacement Catalog' -Values $cv | Out-Null
      $catalog[$repl.ToLower()]=$true; $catAdd++
    }
  } else { $missed.Add("$mfg $mn") }
}
Write-Host ("Matched {0} RFI rows to owned models; {1} not owned/mismatched; {2} replacements added." -f $matched,$missed.Count,$catAdd) -ForegroundColor Green
if($missed.Count -gt 0){ Write-Host "Unmatched (first 15):" -ForegroundColor Yellow; $missed | Select-Object -First 15 | ForEach-Object { Write-Host "  $_" } }
Disconnect-PnPOnline
