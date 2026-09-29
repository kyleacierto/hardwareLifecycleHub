#Requires -Version 7.4
[CmdletBinding()]
param([Parameter(Mandatory)][string]$SiteUrl,[Parameter(Mandatory)][string]$ClientId)
$ErrorActionPreference='Stop'; Import-Module PnP.PowerShell
Connect-PnPOnline -Url $SiteUrl -ClientId $ClientId -Interactive
$today=(Get-Date).ToString('yyyy-MM-dd')
$models=@{}
foreach($m in (Get-PnPListItem -List 'LC Models' -PageSize 5000 -Fields 'ID','Title')){ $models[[string]$m['Title']]=$m.Id }
$promoted=0
foreach($u in (Get-PnPListItem -List 'LC Unknowns' -PageSize 2000)){
  $eos=$u['EndOfSupport']; $repl=[string]$u['Replacement']
  if(-not $eos -and -not $repl){ continue }   # nothing filled in
  $mk=[string]$u['Title']
  if(-not $models.ContainsKey($mk)){ continue }
  $vals=@{ DatesSource='Manual research'; DatesConfirmedBy=([string]$u['ResearchedBy']); DatesConfirmedOn=$today }
  if($eos){ $vals['LastDateOfSupport']=([datetime]$eos).ToString('yyyy-MM-dd') }
  if($repl){ $vals['ReplacementModel']=$repl }
  Set-PnPListItem -List 'LC Models' -Identity $models[$mk] -Values $vals | Out-Null
  $promoted++
}
Write-Host ("Promoted {0} researched models into LC Models." -f $promoted) -ForegroundColor Green
Write-Host "Now re-run Rebuild-Summaries and Export-Dashboard to see Unknown shrink." -ForegroundColor Yellow
Disconnect-PnPOnline
