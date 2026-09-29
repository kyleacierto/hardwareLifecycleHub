#Requires -Version 7.4
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$SiteUrl,
    [Parameter(Mandatory)][string]$ClientId
)
$ErrorActionPreference = 'Stop'
Import-Module PnP.PowerShell
Connect-PnPOnline -Url $SiteUrl -ClientId $ClientId -Interactive

$map = [ordered]@{
    'In Production'   = 'Active'
    'Active'          = 'Active'
    'Deployment'      = 'Active'
    'Action Needed'   = 'Active'
    'In Inventory'    = 'Inventory'
    'Decommissioned'  = 'Retired'
    'Disposed'        = 'Retired'
    'Closed'          = 'Retired'
    'Returned (Demo)' = 'Retired'
    'Returned (RMA)'  = 'Retired'
}

$existing = Get-PnPListItem -List 'LC Status Map' -PageSize 100
foreach ($i in $existing) { Remove-PnPListItem -List 'LC Status Map' -Identity $i.Id -Force | Out-Null }

foreach ($raw in $map.Keys) {
    Add-PnPListItem -List 'LC Status Map' -Values @{ Title = $raw; StatusCategory = $map[$raw] } | Out-Null
    Write-Host ("  {0,-16} -> {1}" -f $raw, $map[$raw])
}
Write-Host "`nLC Status Map now has $($map.Count) rows. Blank CI Status is intentionally unmapped." -ForegroundColor Green
Disconnect-PnPOnline
