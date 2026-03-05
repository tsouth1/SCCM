<#
.Synopsis
 This script will create boundaries and boundary groups. This script works with SCCM 1802 and later.

.Description
 This script will create boundaries, boundary groups, add the site systems to the boundary group.  This will work for IP address range,
 IP subnet or Active Directory site boundaries.  Enter the desired values in the example CSV file that is included in the download.  
 The values in this example are IP range.  Change the boundary type and value as desired.
 
 Check c:\ProvisionServer for the output file boundary.log.

.Example
 New-BoundaryAndGroup -Inputfile "C:\Scripts\InputFiles\BoundryInputFile.csv"
 
.PARAMETER InputFile
 Stores all the required values

.Notes
 Created on:  07/31/2018
 Created by:  Lynford Heron
 Filename:    New-BoundaryAndGroup.ps1
 Version:     1.0
#>

Function New-BoundaryAndGroup
{
  Param( 
    [string]$InputFile
    )

  Begin
    {
      $StartTime = Get-date
      If (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(`
        [Security.Principal.WindowsBuiltInRole] "Administrator"))
        {
          Write-Warning "You do not have Administrator rights to run this script!`nPlease re-run this script as an Administrator!"
          Break
        }

      $ADmodule = (Get-module -Name Activedirectory).Name
      If($ADmodule){Write-Host "The AD module exist." -ForegroundColor Yellow}
      Else { Write-host "The AD module does not exist.  Importing, please wait..." -ForegroundColor Cyan
            Install-WindowsFeature RSAT-AD-PowerShell | Out-Null 
           }
      Write-host "Importing Active Directory module, please wait..." -ForegroundColor Yellow
      Import-Module -Name ActiveDirectory
      
      #Import the required values from the ini file
      $csv = Import-Csv $InputFile
      $sitecode = $csv.sitecode[1]

      Import-module (join-path $(Split-path $env:SMS_ADMIN_UI_PATH) ConfigurationManager.psd1)
      $SetSideCode = $SiteCode + ":"
      Set-Location $SetSideCode

      #Create log file
      $date = get-date
      $LogFolder = "c:\ProvisionServer"
      If (!(Test-Path $LogFolder)) {New-Item $LogFolder -ItemType directory -Force}
      $logfile = $LogFolder + "\" + "Boundary.log"
      new-item -ItemType file $logfile -Force
      Write-Host "The log file $logfile was created"; $date = get-date; add-content $logfile "$date  -  The log file $logfile was created"
      
    }

  Process
    {
      Foreach($item in $csv)
        {
          $BoundaryName = $item.BoundaryName; $BoundaryType = $item.BoundaryType; $BoundaryValue = $item.BoundaryValue
          $BoundaryGname = $item.BoundaryGname; [string[]]$SiteSystemName = $item.SiteSystemName.split(","); $Desc = $item.Desc
          # Boundary type 3 = IPRange
          Write-host "Checking if the boundary - $BoundaryName - exist." -ForegroundColor Yellow; $date = get-date; add-content $logfile " $date  -  Checking if the boundary - $BoundaryName - exist."
          $BoundaryExist = (Get-CMBoundary -BoundaryName $BoundaryName -ErrorAction SilentlyContinue).DisplayName
          If($BoundaryExist)
            {
              Write-Host "The boundary - $BoundaryName - already exist."; $date = get-date; add-content $logfile "$date  -  The boundary - $BoundaryName - already exist."
              Write-Host "Checking if the boundary group - $BoundaryGname - exist."; $date = get-date; add-content $logfile "$date  -  Checking if the boundary group - $BoundaryGname - exist."
              $BoundaryGroupExist = Get-CMBoundaryGroup -Name $BoundaryGname -ErrorAction SilentlyContinue
                  If($BoundaryGroupExist)
                    {
                      Write-Host "The boundary group exist.  Adding boundary - $BoundaryName - to the boundary group - $BoundaryGname."; $date = get-date; add-content $logfile "$date  -  The boundary group exist.  Adding boundary - $BoundaryName - to the boundary group - $BoundaryGname."
                      Try
                        { Add-CMBoundaryToGroup -BoundaryGroupID $BoundaryGroupExist.GroupID -BoundaryName $BoundaryName -ErrorAction SilentlyContinue -Verbose
                          Set-CMBoundaryGroup -Id $BoundaryGroupExist.GroupID -AddSiteSystemServerName $SiteSystemName
                          Write-Host "Boundary - $BoundaryName - was added to the boundary group - $BoundaryGname." -ForegroundColor Yellow ; $date = get-date; add-content $logfile "$date  -  Boundary - $BoundaryName - was added to the boundary group - $BoundaryGname."
                        }
                      Catch
                        { $ErrorMessage = $_.Exception.Message
                          Write-Warning "There was a problem adding the boundary to the boundary group. Check the error and resolve the issue. Error: $ErrorMessage. Exiting the script."; $date = get-date; add-content $logfile "$date  -  There was a problem adding the boundary to the boundary group. Check the error and resolve the issue. Error: $ErrorMessage. Exiting the script."
                          Break
                        }
                    }
                  Else
                    {
                      Write-Host "The boundary group - $BoundaryGname - does not exist.  Creating the boundary group, please wait..." -ForegroundColor Yellow; $date = get-date; add-content $logfile "$date  -  The boundary group - $BoundaryGname - does not exist.  Creating the boundary group, please wait..."
                      Try
                         { 
                           New-CMBoundaryGroup -Name $BoundaryGname -Description $Desc -DefaultSiteCode $sitecode -AddSiteSystemServerName $SiteSystemName -ErrorAction SilentlyContinue -Verbose
                         }
                      Catch
                         { $ErrorMessage = $_.Exception.Message
                           Write-Warning "There was a problem creating the boundary group. Check the error and resolve the issue. Error: $ErrorMessage. Exiting the script."; $date = get-date; add-content $logfile "$date  -  There was a problem creating the boundary group. Check the error and resolve the issue. Error: $ErrorMessage. Exiting the script."
                           Break
                         } 
                      $BoundaryGroupExist = Get-CMBoundaryGroup -Name $BoundaryGname -ErrorAction SilentlyContinue
                      If($BoundaryGroupExist)
                         {
                           Write-Host "The creation of the boundary group - $BoundaryGname - was successful." -ForegroundColor Yellow; $date = get-date; add-content $logfile "$date  -  The creation of the boundary group - $BoundaryGname - was successful."
                           Write-Host "Adding the boundary - $BoundaryName - to the boundary group - $BoundaryGname."; $date = get-date; add-content $logfile "$date  -  Adding boundary - $BoundaryName - to the boundary group - $BoundaryGname."
                           Try
                             { Add-CMBoundaryToGroup -BoundaryGroupID $BoundaryGroupExist.GroupID -BoundaryName $BoundaryName -ErrorAction SilentlyContinue -Verbose
                               Set-CMBoundaryGroup -Id $BoundaryGroupExist.GroupID -AddSiteSystemServerName $SiteSystemName
                               Write-Host "Boundary - $BoundaryName - was added to the boundary group - $BoundaryGname." -ForegroundColor Yellow ; $date = get-date; add-content $logfile "$date  -  Boundary - $BoundaryName - was added to the boundary group - $BoundaryGname."
                             }
                           Catch
                                { $ErrorMessage = $_.Exception.Message
                                   Write-Warning "There was a problem adding the boundary to the boundary group. Check the error and resolve the issue. Error: $ErrorMessage. Exiting the script."; $date = get-date; add-content $logfile "$date  -  There was a problem adding the boundary to the boundary group. Check the error and resolve the issue. Error: $ErrorMessage. Exiting the script."
                                   Break
                                }

                         }

                    }
            } 
          Else
            {
              Write-Host "The boundary - $BoundaryName - does not exist.  Creating the boundary, please wait..." -ForegroundColor Yellow; $date = get-date; add-content $logfile "$date  -  The boundary - $BoundaryName - does not exist.  Creating the boundary, please wait..."
              Try
                 { New-CMBoundary -Name $BoundaryName -Type $BoundaryType -Value $BoundaryValue -ErrorAction SilentlyContinue -Verbose}
              Catch
                 { $ErrorMessage = $_.Exception.Message
                   Write-Warning "There was a problem creating the boundary. Check the error and resolve the issue. Error: $ErrorMessage. Exiting the script."; $date = get-date; add-content $logfile "$date  -  There was a problem creating the boundary. Check the error and resolve the issue. Error: $ErrorMessage. Exiting the script."
                   Break
                 } 
              $BoundaryExist = (Get-CMBoundary -BoundaryName $BoundaryName -ErrorAction SilentlyContinue).DisplayName
              If($BoundaryExist)
                {
                  Write-Host "The creation of the boundary - $BoundaryName - was successful.  Adding boundary to boundary group." -ForegroundColor Yellow; $date = get-date; add-content $logfile "$date  -  The creation of the boundar - $BoundaryName - was successful. Adding boundary to boundary group."
                  Write-Host "Checking if the boundary group - $BoundaryGname - exist."; $date = get-date; add-content $logfile "$date  -  Checking if the boundary group - $BoundaryGname - exist."
                  $BoundaryGroupExist = Get-CMBoundaryGroup -Name $BoundaryGname -ErrorAction SilentlyContinue
                  If($BoundaryGroupExist)
                    {
                      Write-Host "The boundary group exist.  Adding boundary - $BoundaryName - to the boundary group - $BoundaryGname."; $date = get-date; add-content $logfile "$date  -  The boundary group exist.  Adding boundary - $BoundaryName - to the boundary group - $BoundaryGname."
                      Try
                        { Add-CMBoundaryToGroup -BoundaryGroupID $BoundaryGroupExist.GroupID -BoundaryName $BoundaryName -ErrorAction SilentlyContinue -Verbose
                          Set-CMBoundaryGroup -Id $BoundaryGroupExist.GroupID -AddSiteSystemServerName $SiteSystemName
                          Write-Host "Boundary - $BoundaryName - was added to the boundary group - $BoundaryGname." -ForegroundColor Yellow ; $date = get-date; add-content $logfile "$date  -  Boundary - $BoundaryName - was added to the boundary group - $BoundaryGname."
                        }
                      Catch
                        { $ErrorMessage = $_.Exception.Message
                          Write-Warning "There was a problem adding the boundary to the boundary group. Check the error and resolve the issue. Error: $ErrorMessage. Exiting the script."; $date = get-date; add-content $logfile "$date  -  There was a problem adding the boundary to the boundary group. Check the error and resolve the issue. Error: $ErrorMessage. Exiting the script."
                          Break
                        }
                    }
                  Else
                    {
                      Write-Host "The boundary group - $BoundaryGname - does not exist.  Creating the boundary group, please wait..." -ForegroundColor Yellow; $date = get-date; add-content $logfile "$date  -  The boundary group - $BoundaryGname - does not exist.  Creating the boundary group, please wait..."
                      Try
                         { New-CMBoundaryGroup -Name $BoundaryGname -Description $Desc -DefaultSiteCode $sitecode -AddSiteSystemServerName $SiteSystemName -ErrorAction SilentlyContinue -Verbose}
                      Catch
                         { $ErrorMessage = $_.Exception.Message
                           Write-Warning "There was a problem creating the boundary group. Check the error and resolve the issue. Error: $ErrorMessage. Exiting the script."; $date = get-date; add-content $logfile "$date  -  There was a problem creating the boundary group. Check the error and resolve the issue. Error: $ErrorMessage. Exiting the script."
                           Break
                         } 
                      $BoundaryGroupExist = Get-CMBoundaryGroup -Name $BoundaryGname -ErrorAction SilentlyContinue
                      If($BoundaryGroupExist)
                         {
                           Write-Host "The creation of the boundary group - $BoundaryGname - was successful." -ForegroundColor Yellow; $date = get-date; add-content $logfile "$date  -  The creation of the boundary group - $BoundaryGname - was successful."
                           Write-Host "Adding the boundary - $BoundaryName - to the boundary group - $BoundaryGname."; $date = get-date; add-content $logfile "$date  -  Adding boundary - $BoundaryName - to the boundary group - $BoundaryGname."
                           Try
                             { Add-CMBoundaryToGroup -BoundaryGroupID $BoundaryGroupExist.GroupID -BoundaryName $BoundaryName -ErrorAction SilentlyContinue -Verbose
                               Set-CMBoundaryGroup -Id $BoundaryGroupExist.GroupID -AddSiteSystemServerName $SiteSystemName
                               Write-Host "Boundary - $BoundaryName - was added to the boundary group - $BoundaryGname." -ForegroundColor Yellow ; $date = get-date; add-content $logfile "$date  -  Boundary - $BoundaryName - was added to the boundary group - $BoundaryGname."
                             }
                           Catch
                                { $ErrorMessage = $_.Exception.Message
                                   Write-Warning "There was a problem adding the boundary to the boundary group. Check the error and resolve the issue. Error: $ErrorMessage. Exiting the script."; $date = get-date; add-content $logfile "$date  -  There was a problem adding the boundary to the boundary group. Check the error and resolve the issue. Error: $ErrorMessage. Exiting the script."
                                   Break
                                }

                         }
                }

            }
        
        }

    }
 
    #Display script run time  
    $endTime = Get-Date
      $TotalRuntime = $endTime - $StartTime
      Write-Host "`n Start Time:" $StartTime
      Write-Host "`n End Time:" (Get-Date)
      Write-host " Script execution time: $TotalRunTime"
  }
}

New-BoundaryAndGroup -InputFile "C:\temp\BoundryInputFile.csv"