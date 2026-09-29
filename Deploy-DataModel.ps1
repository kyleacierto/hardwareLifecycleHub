#Requires -Version 7.4
<#
.SYNOPSIS
    Build the rest of the hub data model (beyond the four config lists) on a
    SharePoint site you own. Prototype use, fake data only.

.DESCRIPTION
    Uses PnP.PowerShell. Idempotent: an existing list/library is left alone.
    Creates the Devices, Models, Replacement Catalog, reference, matching,
    summary, and operations lists, plus the Imports and Briefs libraries.
    Seeds ONLY Models and Replacement Catalog (from the dashboard mockup) so the
    dashboard has real rows to bind to; everything else is created empty and
    filled by the import in phase 2.

    Adds list-view indexes on the Devices columns used for filtering/grouping.

    Run Deploy-ConfigLists.ps1 first. Same one-time app registration applies.

.PARAMETER SiteUrl
.PARAMETER ClientId
.PARAMETER Remove   Delete everything this script creates.

.EXAMPLE
    pwsh -File .\Deploy-DataModel.ps1 -SiteUrl 'https://cheehu.sharepoint.com/sites/qwerty' -ClientId '<id>'
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

# Field helper: @{ Internal; Display; Type; Indexed(optional) }
# Types used: Text, Number, DateTime, Note (multi-line), Boolean

$genericLists = [ordered]@{

  'LC Devices' = @{
    TitleLabel = 'Asset Key'
    Fields = @(
      @{ Internal='ObjectType';       Display='Object type';        Type='Text';     Indexed=$true }
      @{ Internal='Base';             Display='Base';               Type='Text';     Indexed=$true }
      @{ Internal='Building';         Display='Building';           Type='Text' }
      @{ Internal='Enclave';          Display='Enclave';            Type='Text';     Indexed=$true }
      @{ Internal='ServiceComponent'; Display='Service component';  Type='Text';     Indexed=$true }
      @{ Internal='ModelKey';         Display='Model key';          Type='Text';     Indexed=$true }
      @{ Internal='Manufacturer';     Display='Manufacturer';       Type='Text' }
      @{ Internal='ModelNumber';      Display='Model number';       Type='Text' }
      @{ Internal='SerialNumber';     Display='Serial number';      Type='Text' }
      @{ Internal='OperatingSystem';  Display='Operating system';   Type='Text' }
      @{ Internal='CIStatusRaw';      Display='CI status (raw)';    Type='Text' }
      @{ Internal='StatusCategory';   Display='Status category';    Type='Text';     Indexed=$true }
      @{ Internal='LifecycleBand';    Display='Lifecycle band';     Type='Text';     Indexed=$true }
      @{ Internal='FirstSeen';        Display='First seen';         Type='DateTime' }
      @{ Internal='LastSeen';         Display='Last seen';          Type='DateTime' }
      @{ Internal='RetiredOn';        Display='Retired on';         Type='DateTime' }
      @{ Internal='LastChangeType';   Display='Last change';        Type='Text' }
    )
    Rows = @()
  }

  'LC Models' = @{
    TitleLabel = 'Model Key'
    Fields = @(
      @{ Internal='ObjectType';            Display='Object type';           Type='Text' }
      @{ Internal='Manufacturer';          Display='Manufacturer';          Type='Text' }
      @{ Internal='ModelNumber';           Display='Model number';          Type='Text' }
      @{ Internal='EndOfSale';             Display='End of sale';           Type='DateTime' }
      @{ Internal='EndOfSecuritySupport';  Display='End of security';       Type='DateTime' }
      @{ Internal='LastDateOfSupport';     Display='Last date of support';  Type='DateTime' }
      @{ Internal='PlatformEndOfSupport';  Display='Platform EoS';          Type='DateTime' }
      @{ Internal='EffectiveLDoS';         Display='Effective LDoS';        Type='DateTime' }
      @{ Internal='LifecycleBand';         Display='Lifecycle band';        Type='Text' }
      @{ Internal='DatesSource';           Display='Dates source';          Type='Text' }
      @{ Internal='DatesConfirmedBy';      Display='Confirmed by';          Type='Text' }
      @{ Internal='DatesConfirmedOn';      Display='Confirmed on';          Type='DateTime' }
      @{ Internal='ReplacementModel';      Display='Replacement model';     Type='Text' }
      @{ Internal='ActiveCount';           Display='Active count';          Type='Number' }
    )
    Rows = @(
      @{ Title='Cisco|WS-C3750X-24T-S'; ObjectType='Switch'; Manufacturer='Cisco'; ModelNumber='WS-C3750X-24T-S'; LastDateOfSupport='2026-10-31'; LifecycleBand='Unsupported'; DatesSource='Vendor (Cisco EoL)'; ReplacementModel='C9300X-48HX'; ActiveCount=186 }
      @{ Title='Cisco|C9300-48P';       ObjectType='Switch'; Manufacturer='Cisco'; ModelNumber='C9300-48P';       LastDateOfSupport='2029-01-31'; LifecycleBand='Supported';   DatesSource='Vendor (Cisco EoL)'; ReplacementModel='C9300X-48HX'; ActiveCount=1240 }
      @{ Title='Cisco|C9166';           ObjectType='Wireless AP'; Manufacturer='Cisco'; ModelNumber='CW9166';     LastDateOfSupport='2031-06-30'; LifecycleBand='Supported';   DatesSource='Vendor (Cisco EoL)'; ActiveCount=430 }
      @{ Title='Palo Alto|PA-3260';     ObjectType='Firewall'; Manufacturer='Palo Alto'; ModelNumber='PA-3260';   LastDateOfSupport='2027-06-30'; LifecycleBand='Replace';     DatesSource='Vendor'; ReplacementModel='PA-3440'; ActiveCount=41 }
    )
  }

  'LC Replacement Catalog' = @{
    TitleLabel = 'Replacement Model'
    Fields = @(
      @{ Internal='ObjectType';          Display='Object type';       Type='Text' }
      @{ Internal='MSRPUSD';             Display='MSRP (USD)';        Type='Number' }
      @{ Internal='MSRPAsOf';            Display='MSRP as of';        Type='DateTime' }
      @{ Internal='OnAPL';               Display='On DoDIN APL';      Type='Text' }
      @{ Internal='CSfCListed';          Display='CSfC listed';       Type='Text' }
      @{ Internal='NIAPMaintenanceDate'; Display='NIAP maint. date';  Type='DateTime' }
      @{ Internal='VendorSTIG';          Display='Vendor STIG';       Type='Text' }
      @{ Internal='TestBedValidated';    Display='Test bed validated';Type='Text' }
    )
    Rows = @(
      @{ Title='C9300X-48HX';      ObjectType='Switch';      MSRPUSD=6100; MSRPAsOf='2026-10-01'; OnAPL='Yes'; CSfCListed='No';  VendorSTIG='Yes'; TestBedValidated='No' }
      @{ Title='CW9166';           ObjectType='Wireless AP'; MSRPUSD=1400; MSRPAsOf='2026-10-01'; OnAPL='Yes'; CSfCListed='No';  VendorSTIG='Yes'; TestBedValidated='No' }
      @{ Title='C8300-2N2S-6T';    ObjectType='Router';      MSRPUSD=9800; MSRPAsOf='2026-10-01'; OnAPL='Yes'; CSfCListed='No';  VendorSTIG='Yes'; TestBedValidated='No' }
      @{ Title='PA-3440';          ObjectType='Firewall';    MSRPUSD=28000;MSRPAsOf='2026-10-01'; OnAPL='Yes'; CSfCListed='No';  VendorSTIG='Yes'; TestBedValidated='No' }
    )
  }

  'LC Ref APL' = @{
    TitleLabel = 'TN'
    Fields = @(
      @{ Internal='Vendor';        Display='Vendor';        Type='Text' }
      @{ Internal='Product';       Display='Product';       Type='Text' }
      @{ Internal='Version';       Display='Version';       Type='Text' }
      @{ Internal='Expiration';    Display='Expiration';    Type='DateTime' }
      @{ Internal='DeviceType';    Display='Device type';   Type='Text' }
      @{ Internal='CoveredModels'; Display='Covered models';Type='Note' }
    )
    Rows = @()
  }

  'LC Ref NIAP' = @{
    TitleLabel = 'Ref Key'
    Fields = @(
      @{ Internal='Product';                Display='Product';                Type='Text' }
      @{ Internal='CertificationDate';      Display='Certification date';     Type='DateTime' }
      @{ Internal='AssuranceMaintenanceDate';Display='Assurance maint. date'; Type='DateTime' }
      @{ Internal='TechTypes';              Display='Tech types';             Type='Text' }
    )
    Rows = @()
  }

  'LC Ref CSfC' = @{
    TitleLabel = 'Category'
    Fields = @(
      @{ Internal='Vendor';           Display='Vendor';            Type='Text' }
      @{ Internal='Model';            Display='Model';             Type='Text' }
      @{ Internal='Version';          Display='Version';           Type='Text' }
      @{ Internal='CertificationDate';Display='Certification date';Type='DateTime' }
      @{ Internal='RefreshedOn';      Display='Refreshed on';      Type='DateTime' }
    )
    Rows = @()
  }

  'LC Model Matches' = @{
    TitleLabel = 'Model Key'
    Fields = @(
      @{ Internal='RefList';     Display='Reference list'; Type='Text' }
      @{ Internal='RefKey';      Display='Reference key';  Type='Text' }
      @{ Internal='MatchMethod'; Display='Match method';   Type='Text' }
      @{ Internal='Confirmed';   Display='Confirmed';      Type='Text' }
    )
    Rows = @()
  }

  'LC Sum Headline' = @{
    TitleLabel = 'Period'
    Fields = @(
      @{ Internal='MetricKey';   Display='Metric';   Type='Text' }
      @{ Internal='MetricValue'; Display='Value';    Type='Number' }
      @{ Internal='MetricText';  Display='Text';     Type='Text' }
    )
    Rows = @()
  }

  'LC Sum Counts' = @{
    TitleLabel = 'Period'
    Fields = @(
      @{ Internal='Dimension';   Display='Dimension';   Type='Text' }
      @{ Internal='MemberValue'; Display='Member';      Type='Text' }
      @{ Internal='ActiveCount'; Display='Active';      Type='Number' }
      @{ Internal='UnsupCount';  Display='Unsupported'; Type='Number' }
      @{ Internal='SoonCount';   Display='FY27';        Type='Number' }
    )
    Rows = @()
  }

  'LC Sum Outlook' = @{
    TitleLabel = 'Period'
    Fields = @(
      @{ Internal='FiscalYear'; Display='Fiscal year'; Type='Text' }
      @{ Internal='DeviceCount';Display='Devices';     Type='Number' }
      @{ Internal='MSRPCost';   Display='MSRP cost';   Type='Number' }
      @{ Internal='BandName';   Display='Band';        Type='Text' }
    )
    Rows = @()
  }

  'LC Attention' = @{
    TitleLabel = 'Title'
    Fields = @(
      @{ Internal='Category';   Display='Category';   Type='Text' }
      @{ Internal='Detail';     Display='Detail';     Type='Note' }
      @{ Internal='RankOrder';  Display='Rank';       Type='Number' }
      @{ Internal='LinkUrl';    Display='Link';       Type='Text' }
    )
    Rows = @()
  }

  'LC Import Log' = @{
    TitleLabel = 'Run'
    Fields = @(
      @{ Internal='RunOn';       Display='Run on';      Type='DateTime' }
      @{ Internal='RowsRead';    Display='Rows read';   Type='Number' }
      @{ Internal='ActiveRows';  Display='Active';      Type='Number' }
      @{ Internal='RetiredRows'; Display='Retired';     Type='Number' }
      @{ Internal='UnmappedRows';Display='Unmapped';    Type='Number' }
      @{ Internal='NewRows';     Display='New';         Type='Number' }
      @{ Internal='ChangedRows'; Display='Changed';     Type='Number' }
      @{ Internal='ErrorCount';  Display='Errors';      Type='Number' }
    )
    Rows = @()
  }

  'LC Data Quality' = @{
    TitleLabel = 'Issue'
    Fields = @(
      @{ Internal='IssueType';  Display='Type';       Type='Text' }
      @{ Internal='Detail';     Display='Detail';     Type='Note' }
      @{ Internal='FoundOn';    Display='Found on';   Type='DateTime' }
      @{ Internal='Resolved';   Display='Resolved';   Type='Text' }
    )
    Rows = @()
  }

  'LC Vendor Inquiries' = @{
    TitleLabel = 'Model Key'
    Fields = @(
      @{ Internal='Vendor';      Display='Vendor';      Type='Text' }
      @{ Internal='SentOn';      Display='Sent on';     Type='DateTime' }
      @{ Internal='Status';      Display='Status';      Type='Text' }
      @{ Internal='ReplyOn';     Display='Reply on';    Type='DateTime' }
      @{ Internal='FollowUpOn';  Display='Follow up on';Type='DateTime' }
    )
    Rows = @()
  }
}

$libraries = @('LC Imports','LC Briefs')

# ---- Remove mode -----------------------------------------------------------
if ($Remove) {
    foreach ($name in @($genericLists.Keys) + $libraries) {
        if (Get-PnPList -Identity $name -ErrorAction SilentlyContinue) {
            Remove-PnPList -Identity $name -Force
            Write-Host "Removed $name" -ForegroundColor Yellow
        } else { Write-Host "Skip $name (not present)" }
    }
    Write-Host "Done (remove)." -ForegroundColor Green
    Disconnect-PnPOnline; return
}

# ---- Create generic lists --------------------------------------------------
foreach ($name in $genericLists.Keys) {
    $spec = $genericLists[$name]
    if (Get-PnPList -Identity $name -ErrorAction SilentlyContinue) {
        Write-Host "Skip $name (already exists)" -ForegroundColor Yellow; continue
    }
    Write-Host "Creating $name ..." -ForegroundColor Cyan
    New-PnPList -Title $name -Template GenericList -OnQuickLaunch:$false | Out-Null
    if ($spec.TitleLabel -ne 'Title') {
        Set-PnPField -List $name -Identity 'Title' -Values @{ Title = $spec.TitleLabel } | Out-Null
    }
    foreach ($f in $spec.Fields) {
        $type = switch ($f.Type) {
            'Number'   { 'Number' }
            'DateTime' { 'DateTime' }
            'Note'     { 'Note' }
            'Boolean'  { 'Boolean' }
            default    { 'Text' }
        }
        Add-PnPField -List $name -DisplayName $f.Display -InternalName $f.Internal -Type $type -AddToDefaultView | Out-Null
        if ($f.Indexed) {
            try { Set-PnPField -List $name -Identity $f.Internal -Values @{ Indexed = $true } | Out-Null } catch {}
        }
    }
    foreach ($row in $spec.Rows) { Add-PnPListItem -List $name -Values $row | Out-Null }
    $n = @($spec.Rows).Count
    Write-Host ("  {0} fields, {1} rows" -f $spec.Fields.Count, $n) -ForegroundColor Green
}

# ---- Create libraries ------------------------------------------------------
foreach ($lib in $libraries) {
    if (Get-PnPList -Identity $lib -ErrorAction SilentlyContinue) {
        Write-Host "Skip $lib (already exists)" -ForegroundColor Yellow; continue
    }
    Write-Host "Creating library $lib ..." -ForegroundColor Cyan
    New-PnPList -Title $lib -Template DocumentLibrary -OnQuickLaunch:$false | Out-Null
    if ($lib -eq 'LC Imports') {
        Add-PnPField -List $lib -DisplayName 'Run status' -InternalName 'RunStatus' -Type Text -AddToDefaultView | Out-Null
    }
    if ($lib -eq 'LC Briefs') {
        foreach ($f in @(
            @{ i='ReleaseStatus'; d='Release status' }
            @{ i='ApprovedBy';    d='Approved by' }
            @{ i='ApprovedOn';    d='Approved on' }
            @{ i='SentOn';        d='Sent on' }
        )) {
            $t = if ($f.i -like '*On') { 'DateTime' } else { 'Text' }
            Add-PnPField -List $lib -DisplayName $f.d -InternalName $f.i -Type $t -AddToDefaultView | Out-Null
        }
    }
    Write-Host "  library created" -ForegroundColor Green
}

# ---- Verify ----------------------------------------------------------------
Write-Host "`nVerifying..." -ForegroundColor Cyan
foreach ($name in @($genericLists.Keys) + $libraries) {
    $l = Get-PnPList -Identity $name -ErrorAction SilentlyContinue
    if ($l) {
        $custom = (Get-PnPField -List $name | Where-Object { -not $_.Hidden -and -not $_.ReadOnlyField -and $_.InternalName -notin @('Title','ContentType','Attachments','_ColorTag','ComplianceAssetId','LinkTitle','DocIcon','FileLeafRef','Modified','Created','Author','Editor','_CopySource','_CheckinComment','FileSizeDisplay') } | Measure-Object).Count
        Write-Host ("{0}: {1} items, {2} custom fields" -f $name, $l.ItemCount, $custom)
    } else {
        Write-Host ("{0}: MISSING" -f $name) -ForegroundColor Red
    }
}
Write-Host "`nDone. Expected: 14 lists + 2 libraries. Models 4 rows, Replacement Catalog 4 rows, all others empty." -ForegroundColor Green
Disconnect-PnPOnline
