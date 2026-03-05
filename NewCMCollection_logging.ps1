<#
.Synopsis
 This script will create CM collections with 3 queries, and move them into a specified folder.

.Description
This script will create CM collections with 3 queries - servers, NetBIOSName and ComputersOU. It populates the variables/parameters through a cvs file import. The cvs file should be formatted as follows:

CSV Headers:
ADSite,SiteCode,CollName,SystemOU,SiteAdminGrp

Example Data:
PH-Lapu05e,PET-,PH-Lapu05e-SEC,corp.contoso.com/contoso/GR/APAC/PH-Lapu/Computers,PH Admins 
 
Logs to: C:\Temp\CollectionScript.log
Transcript file: C:\Temp\CollectionScript_transcript.log

.Example
From a ConfigMgr admin PowerShell session (or PS inside the CM console)
.\New-CollectionsFromCsv.ps1 -CsvPath .\sites.csv -LimitingCollectionId 'CEA013D0'
Full path script run:
& 'C:\scripts\sccm ps\collections\createnewcollections.ps1' -csvpath "C:\scripts\sccm ps\collections\newcollections1.csv" -LimitingCollectionId 'CEA013D0'
 
.PARAMETER InputFile
 The CSV stores all the required values, except for the limiting collection ID, target folder path (path to CM collection folder) and logfile+transcription file - change as needed.

.Notes
 Created on:  03/05/2025
 Created by:  TS
 Filename:    New-CMCollection_logging.ps1
 Version:     1.0
#>

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

    # Rule names (must be unique within the collection)
    [Parameter(Mandatory = $false)]
    [string]$RuleName1 = 'Servers',

    [Parameter(Mandatory = $false)]
    [string]$RuleName2 = 'NetBIOSName',

    [Parameter(Mandatory = $false)]
    [string]$RuleName3 = 'ComputersOU'
)

# ----------------------
# Logging Setup (fix)
# ----------------------
$LogFile = 'C:\Temp\CollectionScript.log'                 # primary step-by-step log (your requested file)
$TranscriptFile = 'C:\Temp\CollectionScript_transcript.log'  # separate transcript to avoid lock contention

try {
    $logDir = Split-Path $LogFile
    if (-not (Test-Path -Path $logDir)) {
        New-Item -Path $logDir -ItemType Directory -Force | Out-Null
    }
} catch {
    # last-resort: ignore
}



function Write-Log {
    param(
        [Parameter(Mandatory)]
        [string]$Message,
        [ValidateSet('INFO','WARN','ERROR')]
        [string]$Level = 'INFO'
    )

    $timestamp = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss.fff')
    $line = "[${timestamp}] [$Level] $Message"

    $maxRetries = 5
    $delayMs = 100

    for ($i = 1; $i -le $maxRetries; $i++) {
        try {
            # Open the file allowing other readers/writers (including transcript)
            $fs = [System.IO.File]::Open($LogFile,
                                         [System.IO.FileMode]::Append,
                                         [System.IO.FileAccess]::Write,
                                         [System.IO.FileShare]::ReadWrite)
            try {
                $sw = New-Object System.IO.StreamWriter($fs, [System.Text.Encoding]::UTF8)
                $sw.WriteLine($line)
                $sw.Flush()
                $sw.Dispose()
            } finally {
                $fs.Dispose()
            }
            break  # success, exit retry loop
        } catch {
            if ($i -eq $maxRetries) {
                # last attempt failed; give up silently or emit to console
                Write-Verbose "Write-Log failed after $maxRetries attempts: $($_.Exception.Message)"
            } else {
                Start-Sleep -Milliseconds $delayMs
            }
        }
    }
}

# Start transcript to capture all verbose/implicit outputs as well
try {
    Start-Transcript -Path $TranscriptFile -Append | Out-Null
} catch {
    Write-Log -Level 'WARN' -Message "Start-Transcript failed: $($_.Exception.Message)"
}

Write-Log "Script start. CSV: '$CsvPath'. LimitingCollectionId: '$LimitingCollectionId'."

# --- Load ConfigMgr module and move into the site drive (auto-detect) ---
Write-Log "Checking for ConfigurationManager module."
if (-not (Get-Module -Name ConfigurationManager -ListAvailable)) {
    Write-Log "ConfigurationManager module not found in session. Importing via SMS_ADMIN_UI_PATH."
    $cmModulePath = Join-Path -Path (Split-Path $env:SMS_ADMIN_UI_PATH -Parent) -ChildPath 'ConfigurationManager.psd1'
    Import-Module $cmModulePath -ErrorAction Stop
    Write-Log "ConfigurationManager module imported from '$cmModulePath'."
} else {
    Write-Log "ConfigurationManager module available. Importing."
    Import-Module ConfigurationManager -ErrorAction Stop
    Write-Log "ConfigurationManager module imported."
}

$siteDrive = (Get-PSDrive -PSProvider CMSite | Select-Object -First 1).Name
Write-Log "Detected site drive: '$siteDrive'."
if ($siteDrive) {
    Set-Location "$siteDrive`:"
    Write-Log "Set-Location to '${siteDrive}:'."
}

# --- Create a 3-hour periodic refresh schedule once and reuse ---
Write-Log "Creating 3-hour periodic refresh schedule."
$schedule = New-CMSchedule -Start (Get-Date) -RecurInterval Hours -RecurCount 3
Write-Log "Refresh schedule object created."

# --- Read CSV and loop rows ---
Write-Log "Importing CSV from '$CsvPath'."
$rows = Import-Csv -Path $CsvPath
Write-Log "Imported $($rows.Count) row(s) from CSV."

foreach ($row in $rows) {

    Write-Log "Processing row: $(($row | ConvertTo-Json -Compress))"

    # Pull per-row values (trim to be safe)
    $rowADSite       = ($row.ADSite       | ForEach-Object { $_.ToString().Trim() })
    $rowSiteCode     = ($row.SiteCode     | ForEach-Object { $_.ToString().Trim() })
    $rowCollName     = ($row.CollName     | ForEach-Object { $_.ToString().Trim() })
    $rowSystemOU     = ($row.SystemOU     | ForEach-Object { $_.ToString().Trim() })
    $rowSiteAdminGrp = ($row.SiteAdminGrp | ForEach-Object { $_.ToString().Trim() })

    Write-Log "Row values — CollName:'$rowCollName', ADSite:'$rowADSite', SiteCode:'$rowSiteCode', SystemOU:'$rowSystemOU', SiteAdminGrp:'$rowSiteAdminGrp'."
    if (-not $rowCollName) { Write-Log "Row skipped: CollName is empty." ; continue }

    # Build the three WQL queries per the CSV data
    Write-Log "Building WQL queries for collection '$rowCollName'."
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
    Write-Log "Queries built for '$rowCollName'."

    # Create or get the device collection
    Write-Log "Checking if collection '$rowCollName' exists."
    $existing = Get-CMDeviceCollection -Name $rowCollName -ErrorAction SilentlyContinue
    if (-not $existing) {
        Write-Log "Creating device collection '$rowCollName' with limiting collection '$LimitingCollectionId'."
        New-CMDeviceCollection -Name $rowCollName -LimitingCollectionId $LimitingCollectionId | Out-Null
        Write-Log "Collection '$rowCollName' created."
    } else {
        Write-Log "Collection '$rowCollName' already exists."
    }

    # Ensure 3-hour periodic refresh is set
    Write-Log "Setting periodic refresh (3 hours) on '$rowCollName'."
    Set-CMDeviceCollection -Name $rowCollName -RefreshType Periodic -RefreshSchedule $schedule | Out-Null
    Write-Log "Refresh schedule applied to '$rowCollName'."

    # Remove any existing rules with the same names (keeps things idempotent and simple)
    Write-Log "Removing existing query rules ('$RuleName1', '$RuleName2', '$RuleName3') from '$rowCollName' (if present)."
    Remove-CMDeviceCollectionQueryMembershipRule -CollectionName $rowCollName -RuleName $RuleName1 -ErrorAction SilentlyContinue
    Remove-CMDeviceCollectionQueryMembershipRule -CollectionName $rowCollName -RuleName $RuleName2 -ErrorAction SilentlyContinue
    Remove-CMDeviceCollectionQueryMembershipRule -CollectionName $rowCollName -RuleName $RuleName3 -ErrorAction SilentlyContinue
    Write-Log "Removal complete (if any existed)."

    # Add the three query membership rules
    Write-Log "Adding query rule '$RuleName1' to '$rowCollName'."
    Add-CMDeviceCollectionQueryMembershipRule -CollectionName $rowCollName -RuleName $RuleName1 -QueryExpression $q1 | Out-Null
    Write-Log "Rule '$RuleName1' added."

    Write-Log "Adding query rule '$RuleName2' to '$rowCollName'."
    Add-CMDeviceCollectionQueryMembershipRule -CollectionName $rowCollName -RuleName $RuleName2 -QueryExpression $q2 | Out-Null
    Write-Log "Rule '$RuleName2' added."

    Write-Log "Adding query rule '$RuleName3' to '$rowCollName'."
    Add-CMDeviceCollectionQueryMembershipRule -CollectionName $rowCollName -RuleName $RuleName3 -QueryExpression $q3 | Out-Null
    Write-Log "Rule '$RuleName3' added."
    
    # Set Target Folder Path
    $TargetFolderPath = "CEA:\DeviceCollection\Lear Security"  # Folder path under the Device Collections node
    Write-Log "Target folder path set to '$TargetFolderPath'."

    # Get the collection and move it
    Write-Log "Retrieving collection '$rowCollName' for move."
    $coll = Get-CMDeviceCollection -Name $rowCollName -ErrorAction Stop
    Write-Log "Moving collection '$rowCollName' to '$TargetFolderPath'."
    Move-CMObject -InputObject $coll -FolderPath $TargetFolderPath
    Write-Log "Collection '$rowCollName' moved to '$TargetFolderPath'."
}

Write-Log "Script completed."

try { Stop-Transcript | Out-Null } catch {}
