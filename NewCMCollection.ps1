[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$CsvPath,

    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]$ADSite,

    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]$CollName,
    
    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]$SiteCode,
    
    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]$LimitingCollectionId = 'LEA013D0',

    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]$SiteAdminGrp,
    
    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]$SystemOU,

    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]$LogPath = 'C:\Temp\CMCollectionScript.log',

    # --- Defaults kept but not used directly below; we rebuild queries per CSV row for clarity ---
    [Parameter(Mandatory = $false)]
    [string]$Query1 = @"
select SMS_R_SYSTEM.ResourceID,SMS_R_SYSTEM.ResourceType,SMS_R_SYSTEM.Name,SMS_R_SYSTEM.SMSUniqueIdentifier,SMS_R_SYSTEM.ResourceDomainORWorkgroup,SMS_R_SYSTEM.Client from SMS_R_System inner join SMS_G_System_LOCAL_ADMINS on SMS_G_System_LOCAL_ADMINS.ResourceID = SMS_R_System.ResourceId inner join SMS_G_System_SYSTEM on SMS_G_System_SYSTEM.ResourceID = SMS_R_System.ResourceId where SMS_R_System.ADSiteName = '$ADSite' and SMS_G_System_SYSTEM.SystemRole = 'Server' and SMS_G_System_LOCAL_ADMINS.AccountName like '%Win32_Group.Domain=\"CORPLEAR\",Name=\"$SiteAdminGrp\"%'
"@,

    [Parameter(Mandatory = $false)]
    [string]$Query2 = @"
select SMS_R_SYSTEM.ResourceID,SMS_R_SYSTEM.ResourceType,SMS_R_SYSTEM.Name,SMS_R_SYSTEM.SMSUniqueIdentifier,SMS_R_SYSTEM.ResourceDomainORWorkgroup,SMS_R_SYSTEM.Client from SMS_R_System where SMS_R_System.NetbiosName like '$SiteCode'
"@,

    [Parameter(Mandatory = $false)]
    [string]$Query3 = @"
select SMS_R_SYSTEM.ResourceID,SMS_R_SYSTEM.ResourceType,SMS_R_SYSTEM.Name,SMS_R_SYSTEM.SMSUniqueIdentifier,SMS_R_SYSTEM.ResourceDomainORWorkgroup,SMS_R_SYSTEM.Client from SMS_R_System where SMS_R_System.SystemOUName = '$SystemOU'
"@,

    # Rule names (must be unique within the collection)
    [Parameter(Mandatory = $false)]
    [string]$RuleName1 = 'Servers',

    [Parameter(Mandatory = $false)]
    [string]$RuleName2 = 'NetBIOSName',

    [Parameter(Mandatory = $false)]
    [string]$RuleName3 = 'ComputersOU'
)

# --- Load ConfigMgr module and move into the site drive (auto-detect) ---
if (-not (Get-Module -Name ConfigurationManager -ListAvailable)) {
    $cmModulePath = Join-Path -Path (Split-Path $env:SMS_ADMIN_UI_PATH -Parent) -ChildPath 'ConfigurationManager.psd1'
    Import-Module $cmModulePath -ErrorAction Stop
} else {
    Import-Module ConfigurationManager -ErrorAction Stop
}

$siteDrive = (Get-PSDrive -PSProvider CMSite | Select-Object -First 1).Name
if ($siteDrive) { Set-Location "$siteDrive`:" }

# --- Create a 3-hour periodic refresh schedule once and reuse ---
$schedule = New-CMSchedule -Start (Get-Date) -RecurInterval Hours -RecurCount 3

# --- Read CSV and loop rows ---
$rows = Import-Csv -Path $CsvPath

foreach ($row in $rows) {
    # Pull per-row values (trim to be safe)
    $rowADSite       = ($row.ADSite       | ForEach-Object { $_.ToString().Trim() })
    $rowSiteCode     = ($row.SiteCode     | ForEach-Object { $_.ToString().Trim() })
    $rowCollName     = ($row.CollName     | ForEach-Object { $_.ToString().Trim() })
    $rowSystemOU     = ($row.SystemOU     | ForEach-Object { $_.ToString().Trim() })
    $rowSiteAdminGrp = ($row.SiteAdminGrp | ForEach-Object { $_.ToString().Trim() })

    if (-not $rowCollName) { continue }

    # Build the three WQL queries per the CSV data
    $q1 = @"
select SMS_R_SYSTEM.ResourceID,SMS_R_SYSTEM.ResourceType,SMS_R_SYSTEM.Name,SMS_R_SYSTEM.SMSUniqueIdentifier,SMS_R_SYSTEM.ResourceDomainORWorkgroup,SMS_R_SYSTEM.Client
from SMS_R_System inner join SMS_G_System_LOCAL_ADMINS on SMS_G_System_LOCAL_ADMINS.ResourceID = SMS_R_System.ResourceId inner join SMS_G_System_SYSTEM on SMS_G_System_SYSTEM.ResourceID = SMS_R_System.ResourceId
where SMS_R_System.ADSiteName = "$rowADSite" and SMS_G_System_SYSTEM.SystemRole = "Server" and SMS_G_System_LOCAL_ADMINS.AccountName like "%Win32_Group.Domain=\"CORPLEAR\",Name=\"$rowSiteAdminGrp\"%"
"@

    $q2 = @"
select SMS_R_SYSTEM.ResourceID,SMS_R_SYSTEM.ResourceType,SMS_R_SYSTEM.Name,SMS_R_SYSTEM.SMSUniqueIdentifier,SMS_R_SYSTEM.ResourceDomainORWorkgroup,SMS_R_SYSTEM.Client
from SMS_R_System where SMS_R_System.NetbiosName like "$rowSiteCode"
"@

    $q3 = @"
select SMS_R_SYSTEM.ResourceID,SMS_R_SYSTEM.ResourceType,SMS_R_SYSTEM.Name,SMS_R_SYSTEM.SMSUniqueIdentifier,SMS_R_SYSTEM.ResourceDomainORWorkgroup,SMS_R_SYSTEM.Client
from SMS_R_System where SMS_R_System.SystemOUName = "$rowSystemOU"
"@

    # Create or get the device collection
    $existing = Get-CMDeviceCollection -Name $rowCollName -ErrorAction SilentlyContinue
    if (-not $existing) {
        New-CMDeviceCollection -Name $rowCollName -LimitingCollectionId $LimitingCollectionId | Out-Null
    }

    # Ensure 3-hour periodic refresh is set
    Set-CMDeviceCollection -Name $rowCollName -RefreshType Periodic -RefreshSchedule $schedule | Out-Null

    # Remove any existing rules with the same names (keeps things idempotent and simple)
    Remove-CMDeviceCollectionQueryMembershipRule -CollectionName $rowCollName -RuleName $RuleName1 -ErrorAction SilentlyContinue
    Remove-CMDeviceCollectionQueryMembershipRule -CollectionName $rowCollName -RuleName $RuleName2 -ErrorAction SilentlyContinue
    Remove-CMDeviceCollectionQueryMembershipRule -CollectionName $rowCollName -RuleName $RuleName3 -ErrorAction SilentlyContinue

    # Add the three query membership rules
    Add-CMDeviceCollectionQueryMembershipRule -CollectionName $rowCollName -RuleName $RuleName1 -QueryExpression $q1 | Out-Null
    Add-CMDeviceCollectionQueryMembershipRule -CollectionName $rowCollName -RuleName $RuleName2 -QueryExpression $q2 | Out-Null
    Add-CMDeviceCollectionQueryMembershipRule -CollectionName $rowCollName -RuleName $RuleName3 -QueryExpression $q3 | Out-Null
    
# Set Target Folder Path
$TargetFolderPath = "LEA:\DeviceCollection\Lear Security"  # Folder path under the Device Collections node

# Get the collection and move it
$coll = Get-CMDeviceCollection -Name $rowCollName -ErrorAction Stop
Move-CMObject -InputObject $coll -FolderPath $TargetFolderPath

}

# From a ConfigMgr admin PowerShell session (or PS inside the CM console)
#.\New-CollectionsFromCsv.ps1 -CsvPath .\sites.csv -LimitingCollectionId 'LEA013D0'
