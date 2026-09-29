#Requires -Version 7.4
[CmdletBinding()]
param([Parameter(Mandatory)][string]$SiteUrl,[Parameter(Mandatory)][string]$ClientId,[string]$Dir="$env:USERPROFILE\HubTest")
$ErrorActionPreference='Stop'; Import-Module PnP.PowerShell; Import-Module ImportExcel
Connect-PnPOnline -Url $SiteUrl -ClientId $ClientId -Interactive
$today=Get-Date

function PD($v){
  if($null -eq $v){return $null}
  if($v -is [datetime]){return $v}
  $s=("{0}" -f $v).Trim()
  if($s -eq '' -or $s -match '(?i)not announced|n/?a|varies|tbd|unknown'){return $null}
  foreach($f in 'yyyy-MM-dd','M/d/yyyy','MMM d yyyy','MMM dd yyyy','d-MMM-yyyy','yyyy'){
    try{ return [datetime]::ParseExact($s,$f,[Globalization.CultureInfo]::InvariantCulture) }catch{}
  }
  try{ return [datetime]$s }catch{ return $null }
}
function ToInt($v){ $n=0; if([double]::TryParse((("{0}" -f $v) -replace '[^\d.]',''),[ref]$n) -and $n -gt 0){ return [int]$n } else { return $null } }
# find a column value by keyword(s) from a row object
function ColBy($row,[string[]]$must,[string[]]$mustNot){
  foreach($p in $row.PSObject.Properties){
    $h=("{0}" -f $p.Name).ToUpper()
    $ok=$true; foreach($m in $must){ if($h -notlike "*$($m.ToUpper())*"){ $ok=$false; break } }
    if($ok -and $mustNot){ foreach($x in $mustNot){ if($h -like "*$($x.ToUpper())*"){ $ok=$false; break } } }
    if($ok){ return $p.Value }
  }
  return $null
}

# owned models
$models=@{}
foreach($m in (Get-PnPListItem -List 'LC Models' -PageSize 5000 -Fields 'ID','Title','Manufacturer','ModelNumber')){
  $models[("{0}|{1}" -f ([string]$m['Manufacturer']).Trim(),([string]$m['ModelNumber']).Trim()).ToLower()]=$m
}
$catalog=@{}
foreach($c in (Get-PnPListItem -List 'LC Replacement Catalog' -PageSize 2000 -Fields 'Title')){ $catalog[([string]$c['Title']).Trim().ToLower()]=$true }

$files=Get-ChildItem "$Dir\*.xlsx" | Where-Object { $_.Name -match '(?i)RFI|HARDWARE|EOL|REQUEST' -and $_.Name -notmatch '(?i)^~' }
Write-Host ("Files: {0}" -f $files.Count) -ForegroundColor Cyan
$totM=0;$totU=0;$catAdd=0
foreach($f in $files){
  $sheet=(Import-Excel -Path $f.FullName -WorksheetName (Get-ExcelSheetInfo $f.FullName)[0].Name)
  $fm=0;$fu=0
  foreach($r in $sheet){
    $mfg=("{0}" -f (ColBy $r @('MANUFACTURER'))).Trim()
    $mn=("{0}" -f (ColBy $r @('MODEL','NUMBER'))).Trim()
    if(-not $mn){ continue }
    $eos=PD (ColBy $r @('SUPPORT') @('SALE','COST'))
    if(-not $eos){ $eos=PD (ColBy $r @('END','LIFE') @('SALE','MAINTENANCE')) }
    $repl=("{0}" -f (ColBy $r @('REPLACEMENT') @('MSRP'))).Trim()
    $rmsrp=ToInt (ColBy $r @('MSRP','REPLACEMENT'))
    $key=("{0}|{1}" -f $mfg,$mn).ToLower()
    if($models.ContainsKey($key)){
      $mi=$models[$key]
      $vals=@{ DatesSource=("{0} RFI" -f $mfg); DatesConfirmedBy='vendor'; DatesConfirmedOn=$today.ToString('yyyy-MM-dd') }
      if($eos){ $vals['LastDateOfSupport']=$eos.ToString('yyyy-MM-dd') }
      if($repl -and $repl -notmatch '(?i)^(n/?a|varies|tbd)$'){ $vals['ReplacementModel']=$repl }
      Set-PnPListItem -List 'LC Models' -Identity $mi.Id -Values $vals | Out-Null
      $fm++
      if($vals['ReplacementModel'] -and -not $catalog.ContainsKey($repl.ToLower())){
        $cv=@{ Title=$repl; TestBedValidated='No'; MSRPAsOf=$today.ToString('yyyy-MM-dd') }
        if($rmsrp -ne $null){ $cv['MSRPUSD']=$rmsrp }
        Add-PnPListItem -List 'LC Replacement Catalog' -Values $cv | Out-Null
        $catalog[$repl.ToLower()]=$true; $catAdd++
      }
    } else { $fu++ }
  }
  Write-Host ("  {0,-45} matched {1,3}  unmatched {2,3}" -f $f.Name,$fm,$fu)
  $totM+=$fm; $totU+=$fu
}
Write-Host ("`nTotal matched {0}; unmatched {1}; replacements added {2}." -f $totM,$totU,$catAdd) -ForegroundColor Green
Disconnect-PnPOnline
