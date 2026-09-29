#Requires -Version 7.4
[CmdletBinding()]
param([Parameter(Mandatory)][string]$SiteUrl,[Parameter(Mandatory)][string]$ClientId)
$ErrorActionPreference='Stop'; Import-Module PnP.PowerShell
Connect-PnPOnline -Url $SiteUrl -ClientId $ClientId -Interactive
$today=(Get-Date).ToString('yyyy-MM-dd')

# active device count per model
$active=@{}
foreach($d in (Get-PnPListItem -List 'LC Devices' -PageSize 5000 -Fields 'ModelKey','StatusCategory')){
  if([string]$d['StatusCategory'] -eq 'Active'){ $mk=[string]$d['ModelKey']; if($mk){ if(-not $active.ContainsKey($mk)){$active[$mk]=0}; $active[$mk]++ } }
}

# ensure LC Unknowns list exists (create via site script if missing)
if(-not (Get-PnPList -Identity 'LC Unknowns' -ErrorAction SilentlyContinue)){
  New-PnPList -Title 'LC Unknowns' -Template GenericList -OnQuickLaunch:$true | Out-Null
  Add-PnPField -List 'LC Unknowns' -DisplayName 'Object type' -InternalName 'ObjectType' -Type Text -AddToDefaultView | Out-Null
  Add-PnPField -List 'LC Unknowns' -DisplayName 'Manufacturer' -InternalName 'Manufacturer' -Type Text -AddToDefaultView | Out-Null
  Add-PnPField -List 'LC Unknowns' -DisplayName 'Model number' -InternalName 'ModelNumber' -Type Text -AddToDefaultView | Out-Null
  Add-PnPField -List 'LC Unknowns' -DisplayName 'Active devices' -InternalName 'ActiveDevices' -Type Number -AddToDefaultView | Out-Null
  Add-PnPField -List 'LC Unknowns' -DisplayName 'End of support (fill in)' -InternalName 'EndOfSupport' -Type DateTime -AddToDefaultView | Out-Null
  Add-PnPField -List 'LC Unknowns' -DisplayName 'Replacement (fill in)' -InternalName 'Replacement' -Type Text -AddToDefaultView | Out-Null
  Add-PnPField -List 'LC Unknowns' -DisplayName 'Researched by' -InternalName 'ResearchedBy' -Type Text -AddToDefaultView | Out-Null
  Add-PnPField -List 'LC Unknowns' -DisplayName 'Notes' -InternalName 'ResearchNotes' -Type Note -AddToDefaultView | Out-Null
}

# clear and repopulate
foreach($i in (Get-PnPListItem -List 'LC Unknowns' -PageSize 2000)){ Remove-PnPListItem -List 'LC Unknowns' -Identity $i.Id -Force | Out-Null }

$models=Get-PnPListItem -List 'LC Models' -PageSize 5000 -Fields 'Title','ObjectType','Manufacturer','ModelNumber','LastDateOfSupport','PlatformEndOfSupport'
$count=0
foreach($m in $models){
  $hasSupport = [bool]$m['LastDateOfSupport']
  $hasCompliance = [bool]$m['PlatformEndOfSupport']
  if($hasSupport -or $hasCompliance){ continue }   # only truly unknown
  $mk=[string]$m['Title']
  $ac= if($active.ContainsKey($mk)){$active[$mk]}else{0}
  if($ac -eq 0){ continue }   # only models with active devices worth researching
  Add-PnPListItem -List 'LC Unknowns' -Values @{
    Title=$mk; ObjectType=[string]$m['ObjectType']; Manufacturer=[string]$m['Manufacturer'];
    ModelNumber=[string]$m['ModelNumber']; ActiveDevices=$ac
  } | Out-Null
  $count++
}
Write-Host ("LC Unknowns populated: {0} models with active devices and no date." -f $count) -ForegroundColor Green
Write-Host "Edit them in place on the site; fill End of support / Replacement, then re-run the promote step." -ForegroundColor Yellow
Disconnect-PnPOnline
