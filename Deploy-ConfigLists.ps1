#Requires -Version 7.4
<#
.SYNOPSIS
    Build the four hub config lists on a SharePoint site you own, with clean
    internal names and seed rows. Development/prototype use, fake data only.

.DESCRIPTION
    Uses PnP.PowerShell. Idempotent: an existing list is left alone (its rows
    are not touched or duplicated); only missing lists are created and seeded.
    -Remove deletes the four lists.

    First-time setup on your OWN tenant (you are the admin), run once:
        Install-Module PnP.PowerShell -Scope CurrentUser
        Register-PnPEntraIDAppForInteractiveLogin -ApplicationName "PnP-Cheehu" `
            -Tenant <tenant>.onmicrosoft.com -Interactive
    That prints an Application (client) Id. Pass it as -ClientId below.

.PARAMETER SiteUrl
    e.g. https://cheehu.sharepoint.com/sites/qwerty

.PARAMETER ClientId
    The client Id from Register-PnPEntraIDAppForInteractiveLogin.

.PARAMETER Remove
    Delete the four config lists instead of creating them.

.EXAMPLE
    pwsh -File .\Deploy-ConfigLists.ps1 -SiteUrl 'https://cheehu.sharepoint.com/sites/qwerty' -ClientId '<id>'
.EXAMPLE
    pwsh -File .\Deploy-ConfigLists.ps1 -SiteUrl '...' -ClientId '<id>' -Remove
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$SiteUrl,
    [Parameter(Mandatory)][string]$ClientId,
    [switch]$Remove
)

$ErrorActionPreference = 'Stop'

if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
    throw "PnP.PowerShell is not installed. Run: Install-Module PnP.PowerShell -Scope CurrentUser"
}
Import-Module PnP.PowerShell

Write-Host "Connecting to $SiteUrl ..." -ForegroundColor Cyan
Connect-PnPOnline -Url $SiteUrl -ClientId $ClientId -Interactive

# ---- Definitions -----------------------------------------------------------
# Each list: TitleLabel (rename of the built-in Title column), Fields, Rows.
# Field = @{ Internal; Display; Type }  (Type: Text | Number)
$lists = [ordered]@{
    'LC Settings' = @{
        TitleLabel = 'Setting'
        Fields = @(
            @{ Internal='SettingValue'; Display='Value'; Type='Text' }
            @{ Internal='SettingNote';  Display='Note';  Type='Text' }
        )
        Rows = @(
            @{ Title='ThemeBlack';           SettingValue='#1B1A19';           SettingNote='Dashboard background' }
            @{ Title='ThemeYellow';          SettingValue='#FFC83D';           SettingNote='Dashboard accent' }
            @{ Title='ThemeFont';            SettingValue='Segoe UI';          SettingNote='Dashboard font' }
            @{ Title='FiscalYearStartMonth'; SettingValue='10';                SettingNote='October = federal FY' }
            @{ Title='ShowCost';             SettingValue='Yes';               SettingNote='Show MSRP cost row' }
            @{ Title='CostLabel';            SettingValue='MSRP, not a quote'; SettingNote='Printed by every cost figure' }
            @{ Title='CostAsOf';             SettingValue='2026-10-01';        SettingNote='As-of date for MSRP figures' }
            @{ Title='ExportReminderDay';    SettingValue='5';                 SettingNote='Day of month to nudge if no import' }
            @{ Title='ApprovalReminderDays'; SettingValue='2';                 SettingNote='Business days before a reminder' }
        )
    }
    'LC Status Map' = @{
        TitleLabel = 'Raw Status'
        Fields = @(
            @{ Internal='StatusCategory'; Display='Category'; Type='Text' }
        )
        Rows = @(
            @{ Title='In Production'; StatusCategory='Active' }
            @{ Title='Active';        StatusCategory='Active' }
            @{ Title='Deployment';    StatusCategory='Active' }
            @{ Title='Disposed';      StatusCategory='Retired' }
        )
    }
    'LC Lifecycle Bands' = @{
        TitleLabel = 'Band'
        Fields = @(
            @{ Internal='BandOrder'; Display='Order';      Type='Number' }
            @{ Internal='MinMonths'; Display='Min months'; Type='Number' }
            @{ Internal='MaxMonths'; Display='Max months'; Type='Number' }
            @{ Internal='BandColor'; Display='Color';      Type='Text' }
        )
        Rows = @(
            @{ Title='Unsupported'; BandOrder=1;               MaxMonths=0;  BandColor='#E5484D' }
            @{ Title='Replace';     BandOrder=2; MinMonths=0;  MaxMonths=12; BandColor='#F28C28' }
            @{ Title='Plan';        BandOrder=3; MinMonths=12; MaxMonths=36; BandColor='#6F8FB3' }
            @{ Title='Supported';   BandOrder=4; MinMonths=36;               BandColor='#5F5E5C' }
            @{ Title='Unknown';     BandOrder=5;                             BandColor='#8A8886' }
        )
    }
    'LC Report Breakdowns' = @{
        TitleLabel = 'Dimension'
        Fields = @(
            @{ Internal='SourceColumn'; Display='Source column'; Type='Text' }
            @{ Internal='DisplayOrder'; Display='Order';         Type='Number' }
            @{ Internal='ShowIt';       Display='Show';          Type='Text' }
        )
        Rows = @(
            @{ Title='Enclave';           SourceColumn='Enclave';          DisplayOrder=1; ShowIt='Yes' }
            @{ Title='Base';              SourceColumn='Base';             DisplayOrder=2; ShowIt='Yes' }
            @{ Title='Service component'; SourceColumn='ServiceComponent'; DisplayOrder=3; ShowIt='Yes' }
            @{ Title='Object type';       SourceColumn='ObjectType';       DisplayOrder=4; ShowIt='Yes' }
        )
    }
}

# ---- Remove mode -----------------------------------------------------------
if ($Remove) {
    foreach ($name in $lists.Keys) {
        $existing = Get-PnPList -Identity $name -ErrorAction SilentlyContinue
        if ($existing) {
            Remove-PnPList -Identity $name -Force
            Write-Host "Removed $name" -ForegroundColor Yellow
        } else {
            Write-Host "Skip $name (not present)"
        }
    }
    Write-Host "Done (remove)." -ForegroundColor Green
    Disconnect-PnPOnline
    return
}

# ---- Create mode -----------------------------------------------------------
foreach ($name in $lists.Keys) {
    $spec = $lists[$name]
    $existing = Get-PnPList -Identity $name -ErrorAction SilentlyContinue
    if ($existing) {
        Write-Host "Skip $name (already exists; left as-is)" -ForegroundColor Yellow
        continue
    }

    Write-Host "Creating $name ..." -ForegroundColor Cyan
    New-PnPList -Title $name -Template GenericList -OnQuickLaunch:$false | Out-Null

    # Rename the built-in Title column's display label (internal name stays 'Title')
    Set-PnPField -List $name -Identity 'Title' -Values @{ Title = $spec.TitleLabel } | Out-Null

    foreach ($f in $spec.Fields) {
        $type = if ($f.Type -eq 'Number') { 'Number' } else { 'Text' }
        Add-PnPField -List $name -DisplayName $f.Display -InternalName $f.Internal `
            -Type $type -AddToDefaultView | Out-Null
        Write-Host "  + $($f.Internal) ($type)"
    }

    foreach ($row in $spec.Rows) {
        Add-PnPListItem -List $name -Values $row | Out-Null
    }
    Write-Host "  seeded $($spec.Rows.Count) rows" -ForegroundColor Green
}

Write-Host "`nVerifying..." -ForegroundColor Cyan
foreach ($name in $lists.Keys) {
    $items = Get-PnPListItem -List $name -PageSize 100
    $fields = (Get-PnPField -List $name | Where-Object { -not $_.Hidden -and -not $_.ReadOnlyField } |
        Select-Object -ExpandProperty InternalName)
    Write-Host ("{0}: {1} rows; fields: {2}" -f $name, $items.Count, ($fields -join ', '))
}

Write-Host "`nDone. Check values: LC Settings 9 rows, LC Status Map 4, LC Lifecycle Bands 5, LC Report Breakdowns 4." -ForegroundColor Green
Disconnect-PnPOnline
