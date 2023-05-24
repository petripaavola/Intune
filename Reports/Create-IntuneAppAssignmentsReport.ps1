<#PSScriptInfo

.VERSION 3.0

.GUID 4de3698e-22c8-48d7-96f4-84f09ebb822a

.AUTHOR Petri.Paavola@yodamiitti.fi

.COMPANYNAME Yodamiitti Oy

.COPYRIGHT Petri.Paavola@yodamiitti.fi

.TAGS Intune Application Assignment Report

.LICENSEURI

.PROJECTURI https://github.com/petripaavola/Intune/tree/master/Reports

.ICONURI

.EXTERNALMODULEDEPENDENCIES

.REQUIREDSCRIPTS

.EXTERNALSCRIPTDEPENDENCIES

.RELEASENOTES
Version 1.0:  Original published version
Version 2.0:  Support for Intune Filters
Version 3.0:  New UI with realtime search, sorting and filtering

#>

<#
.Synopsis
   This script creates HTML Report from all Intune App Assignments. This report shows information not easily available in Intune.
   
   Version 3.0

.DESCRIPTION
   This script does HTML Report from Intune Application assigments.
   Report shows information which is not available in Intune UI without making tens or hunders of clicks to all Apps
   
   Report shows also impact of Assignments - number of devices and users targeted to Application.
   
   Report has many options for filtering and search. All columns can be sorted by clicking on column.

   Changelog v2.0:
   - added Application Assignment Filter information to report

   Changelog v3.0
   - Huge update to script and UI
   - There is now one table with all App Assignments
   - And there are multiple ways to search, sort and filter table
     - sort by clicking any column. Text, dates and numbers sorting works in right way
     - realtime free text search
     - OS selection (Windows, Android, iOS, macOS)
     - App type dropdown (multiselection)
     - Assignment group dropdown selection
     - Intune Filters dropdown selection
     - Detailed information shown in ToolTips
     - Web Links to Intune Apps, AssignmentGroups and Filters

   Script downloads AzureADGroups, IntuneApps and IntuneFilters information from Graph API to local cache   (.\cache folder)

   Script also downloads Application icons to local cache to get better looking report.

   You can work with cached data without network connection with parameter -UseOfflineCache


   Author:
   Petri.Paavola@yodamiitti.fi
   Senior Modern Management Principal
   Microsoft MVP - Windows and Devices for IT
   
   2023-05-24

   https://github.com/petripaavola/Intune/tree/master/Reports

.PARAMETER ExportCSV
Export report as ; limited CSV file.

.PARAMETER ExportJSON
Export report as JSON file

.PARAMETER ExportToExcelCopyPaste
Export report to Clipboard. You can paste it to Excel and excel will paste data to columns automatically.

.PARAMETER UseOfflineCache
Create report using files from cache folder.

.PARAMETER DoNotOpenReportAutomatically
Do not automatically open HTML report to Web browser. Can be used when automating report creation.

.PARAMETER UpdateIconsCache
Update App icon cache. New Apps will always get icons downloaded automatically but existing icons are not automatically updated

.PARAMETER IncludeAppsWithoutAssignments
Include Intune Application without Assignments. This will get a lot of Apps you didn't even know exists inside Intune/Graph API.

.PARAMETER DoNotDownloadAppIcons
Do not download Application icons.

.PARAMETER IncludeIdsInReport
Include Appication Ids in report. This makes wider so it is disabled by default.

.PARAMETER IncludeBase64ImagesInReport
Includes Application icons inside HTML file so report will have icons if HTML if copied somewhere else. Note! This is slow and creates huge HTML file.

.EXAMPLE
   .\Create-IntuneAppAssignmentsReport.ps1
.EXAMPLE
   .\Create-IntuneAppAssignmentsReport.ps1 -UseOfflineCache
.EXAMPLE
   .\Create-IntuneAppAssignmentsReport.ps1 -ExportCSV -ExportJSON
.EXAMPLE
   .\Create-IntuneAppAssignmentsReport.ps1 -ExportToExcelCopyPaste
.EXAMPLE
   .\Create-IntuneAppAssignmentsReport.ps1 -UpdateIconsCache
.EXAMPLE
   .\Create-IntuneAppAssignmentsReport.ps1 -DoNotDownloadAppIcons
.EXAMPLE
   .\Create-IntuneAppAssignmentsReport.ps1 -DoNotOpenReportAutomatically
.EXAMPLE
   .\Create-IntuneAppAssignmentsReport.ps1 -IncludeAppsWithoutAssignments

.INPUTS
   None
.OUTPUTS
   Creates report files
    .\yyyyMMddHHmmss_Intune_Application_Assignments_report.html
	.\yyyyMMddHHmmss_Intune_Application_Assignments_report.csv (optional)
	.\yyyyMMddHHmmss_Intune_Application_Assignments_report.json (optional)

   Creates cache files
	.\cache\AllGroups.json
    .\cache\AllApps.json
    .\cache\IntuneFilters.json
	.\cache\appicons

.NOTES
.LINK
   https://github.com/petripaavola/Intune/tree/master/Reports
#>


[CmdletBinding()]
Param(
	[Parameter(Mandatory=$false)]
	[Switch]$ExportCSV,
	[Parameter(Mandatory=$false)]
	[Switch]$ExportJSON,
	[Parameter(Mandatory=$false)]
	[Switch]$ExportToExcelCopyPaste,
	[Parameter(Mandatory=$false)]
	[Switch]$UseOfflineCache,
	[Parameter(Mandatory=$false)]
	[Switch]$DoNotOpenReportAutomatically,
    [Parameter(Mandatory=$false)]
	[Switch]$UpdateIconsCache,
    [Parameter(Mandatory=$false)]
	[Switch]$IncludeAppsWithoutAssignments,
    [Parameter(Mandatory=$false)]
	[Switch]$DoNotDownloadAppIcons,
    [Parameter(Mandatory=$false)]
	[Switch]$IncludeIdsInReport,
    [Parameter(Mandatory=$false)]
	[Switch]$IncludeBase64ImagesInReport
)

# Do not download App Icons if we specify to use cached files
if ($UseOfflineCache) {
    $DoNotDownloadAppIcons = $true
}

$ScriptVersion = "3.0"


function Verify-IntuneModuleExistence {

    # Check that we have Intune Powershell Module available
    # either in installed modules or .psd1 file in same directory
    $IntunePowershellModulePath = "$PSScriptRoot/Microsoft.Graph.Intune.psd1"

    # If we don't have Intune module installed
    if (-not (Get-Module -ListAvailable -Name 'Microsoft.Graph.Intune')) {

        # As backup check if we happen to have module as file in our current directory
        if (-Not (Test-Path $IntunePowershellModulePath)) {

            # Intune Powershell module not found!
            Write-Host "Could not find Intune Powershell module (Microsoft.Graph.Intune)"
            Write-Host "You can install Intune module with command: Install-Module -Name Microsoft.Graph.Intune"
            Write-Host "More information: https://github.com/microsoft/Intune-PowerShell-SDK"
            Write-Host "Script will exit..."
            Pause
            Exit 0

        }
        else {
            Import-Module "$IntunePowershellModulePath"
            $Success = $?

            # If import-module import failed
            if (-not ($Success)) {
                # Intune Powershell module import failed!
                Write-Error "Could not load Intune module from file $IntunePowershellModulePath"
                Write-Host "Double check you have Intune Powershell Module installed and try again."
                Write-Host "You can install Intune module with command: Install-Module -Name Microsoft.Graph.Intune"
                Write-Host "More information: https://github.com/microsoft/Intune-PowerShell-SDK"
                Pause
                Exit 0
            }
            else {
                Write-Host "Successfully imported Intune-module from file: $IntunePowershellModulePath"
            }
        }
    }
    else {
        Write-Host "Microsoft.Graph.Intune module found."
        return $true
    }
}

function Convert-Base64ToFile {
    Param(
        [String]$base64,
        $filepath
    )

    $bytes = [Convert]::FromBase64String($base64)
    [IO.File]::WriteAllBytes($filepath, $bytes)
    $Success = $?

    return $Success
}

function Fix-HTMLSyntax {
    Param(
        $html
    )

    $html = $html.Replace('&lt;', '<')
    $html = $html.Replace('&gt;', '>')
    $html = $html.Replace('&quot;', '"')

    return $html
}

function Fix-HTMLColumns {
    Param(
        $html
    )

    # Rename column headers
    $html = $html -replace '<th>@odata.type</th>','<th>App type</th>'
    $html = $html -replace '<th>displayname</th>','<th>App name</th>'
    $html = $html -replace '<th>assignmentIntent</th>','<th>Assignment Intent</th>'
    $html = $html -replace '<th>assignmentTargetGroupDisplayName</th>','<th>Target Group</th>'
	$html = $html -replace '<th>assignmentFilterDisplayName</th>','<th>Filter name</th>'
	$html = $html -replace '<th>FilterIncludeExclude</th>','<th>Filter Intent</th>'
    $html = $html -replace '<th>publisher</th>','<th>Publisher</th>'
    $html = $html -replace '<th>productVersion</th>','<th>Version</th>'
    $html = $html -replace '<th>filename</th>','<th>Filename</th>'
    $html = $html -replace '<th>createdDateTime</th>','<th>Created</th>'
    $html = $html -replace '<th>lastModifiedDateTime</th>','<th>Modified</th>'

    return $html
}

function Add-AzureADGroupGroupTypeExtraProperties {
	Param(
		[Parameter(Mandatory=$true,
			ValueFromPipeline=$true,
			ValueFromPipelineByPropertyName=$true, 
			Position=0)]
		$AzureADGroups
	)

	# DEBUG Export $group to clipboard for testing off the script
	#$AzureADGroups | ConvertTo-Json -Depth 5 | Set-Clipboard

	# Add new properties groupType and MembershipType
	foreach($group in $AzureADGroups) {

		$GroupType = 'unknown'
		if($group.groupTypes -contains 'Unified') {
			# Group is Office365 group
			$group | Add-Member -MemberType noteProperty -Name YodamiittiCustomGroupType -Value 'Office365'
		} else {
			# Group is either security group or distribution group

			if($group.securityEnabled -and (-not $group.mailEnabled)) {
				# Group is security group
				$group | Add-Member -MemberType noteProperty -Name YodamiittiCustomGroupType -Value 'Security'
			}
			
			if((-not $group.securityEnabled) -and $group.mailEnabled) {
				# Group is Distribution group
				$group | Add-Member -MemberType noteProperty -Name YodamiittiCustomGroupType -Value 'Distribution'
			}
		}


		# Check if group is directoryRole which is not actual AzureAD Group
		if($group.'@odata.type' -eq '#microsoft.graph.directoryRole') {
			# Group is NOT security group at all
			# DirectoryRoles are not Azure AD groups
			$group | Add-Member -MemberType noteProperty -Name YodamiittiCustomGroupType -Value 'DirectoryRole'
		}


		if($group.groupTypes -contains 'DynamicMembership') {
			# Dynamic group
			$group | Add-Member -MemberType noteProperty -Name YodamiittiCustomMembershipType -Value 'Dynamic'
		} else {
			# Static group
			$group | Add-Member -MemberType noteProperty -Name YodamiittiCustomMembershipType -Value 'Static'
		}
	}

	return $AzureADGroups
}


function Add-AzureADGroupDevicesAndUserMemberCountExtraProperties {
	Param(
		[Parameter(Mandatory=$true,
			ValueFromPipeline=$true,
			ValueFromPipelineByPropertyName=$true, 
			Position=0)]
		$AzureADGroups
	)

	Write-Verbose "Getting AzureAD groups membercount for $($AzureADGroups.Count) groups"

	for ($i=0; $i -lt $AzureADGroups.count; $i+=20){

		# Create requests hashtable
		$requests_devices_count = @{}
		$requests_users_count = @{}

		# Create elements array inside hashtable
		$requests_devices_count.requests = @()
		$requests_users_count.requests = @()

		# Create max 20 requests in for-loop
		# For-loop will end automatically when loop counter is same as total count of $AzureADGroups
		for ($a=$i; (($a -lt $i+20) -and ($a -lt $AzureADGroups.count)); $a+=1) {

			if(($AzureADGroups[$a]).'@odata.type' -eq '#microsoft.graph.directoryRole') {
				# Azure DirectoryRole is not AzureAD Group
				$GraphAPIBatchEntry_DevicesCount = @{
					id = ($a+1).ToString()
					"method" = "GET"
					"url" = "/directoryRoles/$(($AzureADGroups[$a]).id)"
				}

			} else {
				# We should have AzureAD Group
				$GraphAPIBatchEntry_DevicesCount = @{
					id = ($a+1).ToString()
					"method" = "GET"
					"url" = "/groups/$(($AzureADGroups[$a]).id)/transitivemembers/microsoft.graph.device/`$count?ConsistencyLevel=eventual"
				}
			}

			# Add GraphAPI Batch entry to requests array
			$requests_devices_count.requests += $GraphAPIBatchEntry_DevicesCount

			if(($AzureADGroups[$a]).'@odata.type' -eq '#microsoft.graph.directoryRole') {
				# Azure DirectoryRole is not AzureAD Group
				$GraphAPIBatchEntry_UsersCount = @{
					id = ($a+1).ToString()
					"method" = "GET"
					"url" = "/directoryRoles/$(($AzureADGroups[$a]).id)"
				}
			} else {
				# We should have AzureAD Group
				$GraphAPIBatchEntry_UsersCount = @{
					id = ($a+1).ToString()
					"method" = "GET"
					"url" = "/groups/$(($AzureADGroups[$a]).id)/transitivemembers/microsoft.graph.user/`$count?ConsistencyLevel=eventual"
				}
			}

			
			# Add GraphAPI Batch entry to requests array
			$requests_users_count.requests += $GraphAPIBatchEntry_UsersCount
			
			# DEBUG/double check index numbers and groupNames
			#Write-Host "`$a=$a   `$i=$i    GroupName=$($AzureADGroups[$a].displayName)"
		}

		# DEBUG
		#$requests_devices_count | ConvertTo-Json
		$requests_devices_count_JSON = $requests_devices_count | ConvertTo-Json

		$url = 'https://graph.microsoft.com/beta/$batch'
		$MSGraphRequest = Invoke-MSGraphRequest -Url $url -Content $requests_devices_count_JSON.ToString() -HttpMethod 'POST'
		$Success = $?

		if($Success) {
			#Write-Host "Success"
		} else {
			# Invoke-MSGraphRequest failed
			Write-Error "Error getting AzureAD groups devices count"
			return 1
		}

		# Get AllMSGraph pages
		# This is also workaround to get objects without assigning them from .Value attribute
		$AzureADGroups_Devices_MemberCount_Batch_Result = Get-MSGraphAllPages -SearchResult $MSGraphRequest
		$Success = $?

		if($Success) {
			#Write-Host "Success"
		} else {
			# Invoke-MSGraphRequest failed
			Write-Error "Error getting AzureAD groups devices count"
			return 1
		}
		
		# DEBUG
		#$AzureADGroups_Devices_MemberCount_Batch_Result

		# Process results for devices count batch requests
		Foreach ($response in $AzureADGroups_Devices_MemberCount_Batch_Result.responses) {
			$GroupArrayIndex = $response.id - 1
			if($response.status -eq 200) {
				
				if(($AzureADGroups[$GroupArrayIndex]).'@odata.type' -eq '#microsoft.graph.directoryRole') {
					# DEBUG
					#Write-Verbose "AzureAD directoryRole (arrayIndex=$GroupArrayIndex) $($AzureADGroups[$GroupArrayIndex].displayName)"

					$AzureADGroups[$GroupArrayIndex] | Add-Member -MemberType noteProperty -Name YodamiittiCustomGroupMembersCountDevices -Value 'N/A'
				} else {
					# DEBUG
					#Write-Verbose "AzureAD group (arrayIndex=$GroupArrayIndex) $($AzureADGroups[$GroupArrayIndex].displayName) adding devices count property: $($response.body)"
					
					$AzureADGroups[$GroupArrayIndex] | Add-Member -MemberType noteProperty -Name YodamiittiCustomGroupMembersCountDevices -Value $response.body					
				}
			} else {
				Write-Error "Error getting devices count for AzureAD group $($AzureADGroups[$GroupArrayIndex].displayName)"
				Write-Error "$($response | ConvertTo-Json)"
			}
		}


		$requests_users_count_JSON = $requests_users_count | ConvertTo-Json

		$url = 'https://graph.microsoft.com/beta/$batch'
		$MSGraphRequest = Invoke-MSGraphRequest -Url $url -Content $requests_users_count_JSON.ToString() -HttpMethod 'POST'
		$Success = $?

		if($Success) {
			#Write-Host "Success"
		} else {
			# Invoke-MSGraphRequest failed
			Write-Error "Error getting AzureAD groups users count"
			return 1
		}

		# Get AllMSGraph pages
		# This is also workaround to get objects without assigning them from .Value attribute
		$AzureADGroups_Users_MemberCount_Batch_Result = Get-MSGraphAllPages -SearchResult $MSGraphRequest
		$Success = $?

		if($Success) {
			#Write-Host "Success"
		} else {
			# Invoke-MSGraphRequest failed
			Write-Error "Error getting AzureAD groups users count"
			return 1
		}
		
		# DEBUG
		#$AzureADGroups_Users_MemberCount_Batch_Result

		# Process results for devices count batch requests
		Foreach ($response in $AzureADGroups_Users_MemberCount_Batch_Result.responses) {
			$GroupArrayIndex = $response.id - 1
			if($response.status -eq 200) {
				
				if(($AzureADGroups[$GroupArrayIndex]).'@odata.type' -eq '#microsoft.graph.directoryRole') {
					# DEBUG
					#Write-Verbose "AzureAD directoryRole (arrayIndex=$GroupArrayIndex) $($AzureADGroups[$GroupArrayIndex].displayName)"

					# Change "AzureAD Group" json to actual real directoryRole json which we just got from batch request
					$AzureADGroups[$GroupArrayIndex] = $response.body
					
					$AzureADGroups[$GroupArrayIndex] | Add-Member -MemberType noteProperty -Name YodamiittiCustomGroupMembersCountUsers -Value 'N/A'
					
					# We need to add below properties again because we just replace whole object so we lost earlier customProperties
					$AzureADGroups[$GroupArrayIndex] | Add-Member -MemberType noteProperty -Name YodamiittiCustomGroupMembersCountDevices -Value 'N/A' -Force
					$AzureADGroups[$GroupArrayIndex] | Add-Member -MemberType noteProperty -Name YodamiittiCustomGroupType -Value 'DirectoryRole' -Force
				} else {
					# DEBUG
					#Write-Verbose "AzureAD group (arrayIndex=$GroupArrayIndex) $($AzureADGroups[$GroupArrayIndex].displayName) adding users count property: $($response.body)"
					
					$AzureADGroups[$GroupArrayIndex] | Add-Member -MemberType noteProperty -Name YodamiittiCustomGroupMembersCountUsers -Value $response.body
				}
				
				
			} else {
				Write-Error "Error getting users count for AzureAD group $($AzureADGroups[$GroupArrayIndex].displayName)"
				Write-Error "$($response | ConvertTo-Json)"
			}
		}
		
	}     

	return $AzureADGroups
}


####################################################################################################
# Main starts here

# Really quick and dirty error handling with huge try-catch block
try {

    # Yes we should take true/false return value back here and do messaging and exiting here based on that value
    # But it was so quick to write that staff to function so let's fix that on some later version (like never ?-)
	if (-not ($UseOfflineCache)) {
		$return = Verify-IntuneModuleExistence
	}

    # Create cache folder if it does not exist
    if (-not (Test-Path "$PSScriptRoot\cache")) {
        Write-Host "Creating cache directory: $PSScriptRoot\cache"
        New-Item -ItemType Directory "$PSScriptRoot\cache"
        $Success = $?

        if (-not ($Success)) {
            Write-Error "Could not create cache directory ($PSScriptRoot\cache). Check file system rights and try again."
            Write-Host "Script will exit..."
            Pause
            Exit 1
        }
    }

    ###################################################
    # Change schema and get Tenant info

    if (-not ($UseOfflineCache)) {
        try {
            Write-Host "Get tenant information from Graph API and change Graph API schema to beta"

            # We have variables
            # $ConnectMSGraph.UPN
            # $ConnectMSGraph.TenantId

            # Update Graph API schema to beta to get Win32LobApps and possible other new apptypes also
            Update-MSGraphEnvironment -SchemaVersion 'beta'
            $ConnectMSGraph = Connect-MSGraph
            $Success = $?

            if (-not ($Success)) {
                Write-Error "Error connecting to Microsoft Graph API with command Connect-MSGraph."
                Write-Host "Check you have Intune Powershell Cmdlest installed with commmand: Install-Module -Name Microsoft.Graph.Intune"
                Write-Host "Script will exit..."
                Pause
                Exit 1
            }
            else {
                Write-Verbose "Connect-MSGraph succeeded after schema change to beta"
            }

            $TenantDisplayName = Invoke-MSGraphRequest -url "https://graph.microsoft.com/beta/organization/$($ConnectMSGraph.TenantId)?`$select=displayName" | Select-Object -ExpandProperty displayName
            Write-Verbose "Get tenant information with Invoke-MSGraphRequest success: $?"
        }
        catch {
            Write-Error "$($_.Exception.GetType().FullName)"
            Write-Error "$($_.Exception.Message)"
            Write-Error "Error trying to connect to MSGraph."
            Write-Host "Script will exit..."
            Pause
            Exit 1
        }
    }
    else {
        $TenantDisplayName = "(offline)"
    }


    ###################################################
    # Get Intune filters

	# Test if we have IntuneFilters.json file
	if (-not (Test-Path "$PSScriptRoot\cache\IntuneFilters.json")) {
		Write-Host "Did NOT find IntuneFilters.json file. We have to download Intune Filters from Graph API"
		if ($UseOfflineCache) {
			Write-Host "Run script without option -UseOfflineCache to download necessary Intune Filters information`n" -ForegroundColor "Yellow"
			Exit 0
		}
	}

	try {

		if (-not ($UseOfflineCache)) {
			Write-Host "Downloading Intune filters information"
			$url = 'https://graph.microsoft.com/beta/deviceManagement/assignmentFilters'
			$MSGraphRequest = Invoke-MSGraphRequest -Url $url -HttpMethod 'GET'
			$Success = $?

			if (-not ($Success)) {
				Write-Error "Error downloading Intune filters information"
				Write-Host "Script will exit..."
				Pause
				Exit 1
			}

			$AllIntuneFilters = Get-MSGraphAllPages -SearchResult $MSGraphRequest
			
			Write-Host "Found $($AllIntuneFilters.Count) Intune filters"
			
			# Save to local cache -Depth 3 is default value
            $AllIntuneFilters | ConvertTo-Json -Depth 3 | Out-File "$PSScriptRoot\cache\IntuneFilters.json" -Force

        }
        else {
            Write-Host "Using cached Intune Filters information from file: $PSScriptRoot\cache\IntuneFilters.json"
        }

        # Get Intune Filters information from cached file always
        $AllIntuneFilters = Get-Content "$PSScriptRoot\cache\IntuneFilters.json" | ConvertFrom-Json


    } catch {
        Write-Error "$($_.Exception.GetType().FullName)"
        Write-Error "$($_.Exception.Message)"
        Write-Error "Error trying to download Intune filters information"
        Write-Host "Script will exit..."
        Pause
        Exit 1
    }

    ###################################################
    # Get AzureADGroups. This should be more efficient than getting AzureADGroup for every assignment one by one

    # Test if we have AllGroups.json file
    if (-not (Test-Path "$PSScriptRoot\cache\AllGroups.json")) {
        Write-Host "Did NOT find AllGroups.json file. We have to get AzureAD Group information from Graph API"
        if ($UseOfflineCache) {
            Write-Host "Run script without option -UseOfflineCache to download necessary AllGroups information`n" -ForegroundColor "Yellow"
            Exit 0
        }
    }

    try {
        if (-not ($UseOfflineCache)) {
            Write-Host "Downloading all AzureAD Security Groups from Graph API (this might take a while)..."
            Write-Verbose "Downloading all AzureAD Security Groups from Graph API"

            # Notice we probably get only part of groups because GraphAPI returns limited number of groups
            $groups = Get-AADGroup -Filter 'securityEnabled eq true'
            $Success = $?

            if (-not ($Success)) {
                Write-Error "Error downloading AzureAD Security Groups"
                Write-Host "Script will exit..."
                Pause
                Exit 1
            }

            # Check if we have value starting https:// in attribute @odate.nextLink
            # If we have nextLink then we get all groups
            if ($groups.'@odata.nextLink' -like "https://*") {
                # Get all groups because we got paged result
                $AllGroups = Get-MSGraphAllPages -SearchResult $groups
                $Success = $?
                if (-not ($Success)) {
                    Write-Error "Error downloading all AzureAD Security Groups with command Get-MSGraphAllPages"
                    Write-Host "Script will exit..."
                    Pause
                    Exit 1
                }
            }
            else {
                $AllGroups = $groups
            }
            Write-Host "AzureAD Group information downloaded."

            # Save to local cache -Depth 2 is default value
            $AllGroups | ConvertTo-Json -Depth 2 | Out-File "$PSScriptRoot\cache\AllGroups.json" -Force
        }
        else {
            Write-Host "Using cached AzureAD Group information from file: $PSScriptRoot\cache\AllGroups.json"
        }

        # Get Group information from cached file always
        $AllGroups = Get-Content "$PSScriptRoot\cache\AllGroups.json" | ConvertFrom-Json
		
		# Create $AllGroupsHashtable for quicker search
		$AllGroupsHashTable = @{}
		$AllGroups | Foreach-Object { $id = $_.id; $value=$_; $AllGroupsHashTable["$id"] = $value }
		

    }
    catch {
        Write-Error "$($_.Exception.GetType().FullName)"
        Write-Error "$($_.Exception.Message)"
        Write-Error "Error trying to download AzureAD Group information"
        Write-Host "Script will exit..."
        Pause
        Exit 1
    }

    ###################################################
    # Test if we have AllApps.json file
    if (-not (Test-Path "$PSScriptRoot\cache\AllApps.json")) {
        Write-Host "Could NOT find AllApps.json file. We have to get Apps information from Graph API"
        if ($UseOfflineCache) {
            Write-Host "Run script without option -UseOfflineCache to download necessary AllApps information`n" -ForegroundColor "Yellow"
            Exit 0
        }
    }

    try {
        # Get App information from Graph API
        if (-not ($UseOfflineCache)) {
            # Get App information from Graph API
            Write-Host "Downloading Intune App information from Graph API (this might take a while)..."

            # We need assignments info so -Expand assignment option is needed here
            $Apps = Get-DeviceAppManagement_MobileApps -Expand assignments
            $Success = $?

            if (-not ($Success)) {
                Write-Error "Error downloading Intune Applications information"
                Write-Host "Script will exit..."
                Pause
                Exit 1
            }

            # Check if we have value starting https:// in attribute @odate.nextLink
            # If we have nextLink then we get all Apps
            if ($Apps.'@odata.nextLink' -like "https://*") {
                # Get all Apps because we got paged result
                $AllApps = Get-MSGraphAllPages -SearchResult $Apps
                $Success = $?
                if (-not ($Success)) {
                    Write-Error "Error downloading Intune Applications information with command Get-MSGraphAllPages"
                    Write-Host "Script will exit..."
                    Pause
                    Exit 1
                }
            }
            else {
                $AllApps = $Apps
            }
            Write-Host "Intune Application information downloaded"

            # Save to local cache
            # Really important parameter!!! Specify -Depth 4 because Application assignment data will be nested down 4 levels
            $AllApps | ConvertTo-Json -Depth 4 | Out-File "$PSScriptRoot\cache\AllApps.json" -Force

        }
        else {
            Write-Host "Using cached Intune App information from file: $PSScriptRoot\cache\AllApps.json"
        }

        # Get App information from cached file always
        $AllApps = Get-Content "$PSScriptRoot\cache\AllApps.json" | ConvertFrom-Json
		
		# Create Hashtable for quicker search
		$AllAppsHashTable = @{}
		$AllApps | Foreach-Object { $id = $_.id; $value=$_; $AllAppsHashTable["$id"] = $value }

    }
    catch {
        Write-Error "$($_.Exception.GetType().FullName)"
        Write-Error "$($_.Exception.Message)"
        Write-Error "Error trying to download Intune Application information"
        Write-Host "Script will exit..."
        Pause
        Exit 1
    }

    ###################################################
    # Find apps which have assignments
    # Check data syntax from GraphAPI with request: https://graph.microsoft.com/v1.0/deviceAppManagement/mobileApps?$expand=assignments
    # or convert $Apps to json to get more human readable format: $Apps | ConvertTo-JSON
    # $AppsWithAssignments = $AllApps | Where-Object { $_.assignments.target.groupid -like "*" }

    # Create custom object array and gather necessary app and assignment information.
    $AppsWithAssignmentInformation = @()
	
	# Create custom object array to list which AzureAD Groups had assignments
	# We will get more information for these groups only
	$AzureADGroupsWithAssignments = @()

    try {
        Write-Host "Creating Application custom object array"

        # Go through each app and save necessary information to custom object
        Foreach ($App in $AllApps) {
			
			if ($App.licenseType -eq 'offline') {
				$displayName = "$($App.displayname) (offline)"
			}
			else {
				$displayName = "$($App.displayname)"
			}

			# Set placeholder value
			# Win32LOB (IntuneWin32)
			$AppVersion = $App.displayVersion
			
			# At least #microsoft.graph.managedAndroidStoreApp
			if($App.version) {
				$AppVersion = $App.version
			}
			
			# MSI LOB App
			if($App.productVersion) {
				$AppVersion = $App.productVersion
			}
			
			# MSIX
			if($App.identityVersion) {
				$AppVersion = $App.identityVersion
			}
			
			# Remove #microsoft.graph. from @odata.type
			$odatatype = $App.'@odata.type'.Replace('#microsoft.graph.', '')

            Foreach ($Assignment in $App.Assignments) {
			
                $assignmentId = $Assignment.id
                $assignmentIntent = $Assignment.intent

				# Change first letter to capitalized
				$assignmentIntent = (Get-Culture).TextInfo.ToTitleCase($assignmentIntent.ToLower())
				
                $assignmentTargetGroupId = $Assignment.target.groupid
				
				# Get Assignment group object
				# Slower
				#$assignmentTargetGroupObject = $AllGroups | Where-Object { $_.id -eq $assignmentTargetGroupId }
				
				# Superfast with Hashtable. Hashtables FTW! ;)
				# Builtin All Users and All Devices are not real AzureAD Groups so they don't exist in Hashtable
				if(($assignmentTargetGroupId) -and ($AllGroupsHashTable.ContainsKey($assignmentTargetGroupId))) {
					$assignmentTargetGroupObject = $AllGroupsHashTable[$assignmentTargetGroupId]
				} else {
					$assignmentTargetGroupObject = $null	
				}
				
				$assignmentTargetGroupDisplayName = $assignmentTargetGroupObject | Select-Object -ExpandProperty displayName

				# Add group to another array so we can get Devices and Users count to each AzureAD Group
                $AzureADGroupsWithAssignments += $assignmentTargetGroupObject
				
				$assignmentFilterId = $Assignment.target.deviceAndAppManagementAssignmentFilterId
				
				$assignmentFilterDisplayName = $AllIntuneFilters | Where-Object { $_.id -eq $assignmentFilterId } | Select-Object -ExpandProperty displayName
				
				$FilterIncludeExclude = $Assignment.target.deviceAndAppManagementAssignmentFilterType
				if($FilterIncludeExclude -eq 'None') {
					$FilterIncludeExclude = $null
				}
				
                # Special case for All Users
                if ($Assignment.target.'@odata.type' -eq '#microsoft.graph.allLicensedUsersAssignmentTarget') {
                    $assignmentTargetGroupDisplayName = 'All Users'
                }

                # Special case for All Devices
                if ($Assignment.target.'@odata.type' -eq '#microsoft.graph.allDevicesAssignmentTarget') {
                    $assignmentTargetGroupDisplayName = 'All Devices'
                }

                # Set included/excluded attribute
                $AppIncludeExclude = ''
                if ($Assignment.target.'@odata.type' -eq '#microsoft.graph.groupAssignmentTarget') {
                    $AppIncludeExclude = 'Included'
                }
                if ($Assignment.target.'@odata.type' -eq '#microsoft.graph.exclusionGroupAssignmentTarget') {
                    $AppIncludeExclude = 'Excluded'
                }

				$assignmentIntent = "$assignmentIntent $AppIncludeExclude"

                $properties = @{
                    '@odata.type'                    = $odatatype
                    displayname                      = $displayName
                    productVersion                   = $AppVersion
                    publisher                        = $App.publisher
                    filename                         = $App.filename
                    createdDateTime                  = $App.createdDateTime
                    lastModifiedDateTime             = $App.lastModifiedDateTime
                    id                               = $App.id
                    licenseType                      = $App.licenseType
                    assignmentId                     = $assignmentId
                    assignmentIntent                 = $assignmentIntent
                    assignmentTargetGroupId          = $assignmentTargetGroupId
                    assignmentTargetGroupDisplayName = $assignmentTargetGroupDisplayName
					devices							 = [int]$null
					users							 = [int]$null
                    AppIncludeExclude                = $AppIncludeExclude
					assignmentFilterId				 = $assignmentFilterId
					assignmentFilterDisplayName      = $assignmentFilterDisplayName
					FilterIncludeExclude             = $FilterIncludeExclude
                    icon                             = ""
                }

                # Create new custom object every time inside foreach-loop
                # This is really important step to do inside foreach-loop!
                # If you create custom object outside of foreach then you would edit same custom object on every foreach cycle resulting only 1 app in custom object array
                $CustomObject = New-Object -TypeName PSObject -Prop $properties

                # Add custom object to our custom object array.
                $AppsWithAssignmentInformation += $CustomObject
            }
			
			# Include Apps without assignments
			if((-not $App.Assignments) -and $IncludeAppsWithoutAssignments) {
				
				$properties = @{
                    '@odata.type'                    = $odatatype
                    displayname                      = $displayName
                    productVersion                   = $AppVersion
                    publisher                        = $App.publisher
                    filename                         = $App.filename
                    createdDateTime                  = $App.createdDateTime
                    lastModifiedDateTime             = $App.lastModifiedDateTime
                    id                               = $App.id
                    licenseType                      = $App.licenseType
                    assignmentId                     = $null
                    assignmentIntent                 = $null
                    assignmentTargetGroupId          = $null
                    assignmentTargetGroupDisplayName = $null
					devices							 = [int]$null
					users							 = [int]$null
                    AppIncludeExclude                = $null
					assignmentFilterId				 = $null
					assignmentFilterDisplayName      = $null
					FilterIncludeExclude             = $null
					icon                             = ""
                }

                # Create new custom object every time inside foreach-loop
                # This is really important step to do inside foreach-loop!
                # If you create custom object outside of foreach then you would edit same custom object on every foreach cycle resulting only 1 app in custom object array
                $CustomObject = New-Object -TypeName PSObject -Prop $properties

                # Add custom object to our custom object array.
                $AppsWithAssignmentInformation += $CustomObject
				
			}
			
        }
		
		# Remove duplicate AzureADGroups
		$AzureADGroupsWithAssignments = $AzureADGroupsWithAssignments | Sort-Object -Property id -Unique
		
		# Add groupType information
		$AzureADGroupsWithAssignments = Add-AzureADGroupGroupTypeExtraProperties $AzureADGroupsWithAssignments
		
		# Get count of Devices and Users for AssignmentGroups
		$AzureADGroupsWithAssignments = Add-AzureADGroupDevicesAndUserMemberCountExtraProperties $AzureADGroupsWithAssignments

		# Create Hashtable for quicker search
		$AzureADGroupsWithAssignmentsHashTable = @{}
		$AzureADGroupsWithAssignments | Foreach-Object { $id = $_.id; $value=$_; $AzureADGroupsWithAssignmentsHashTable["$id"] = $value }
		
		# Add Devices and Users count to $AppsWithAssignmentInformation array
		Foreach ($AppAssignment in $AppsWithAssignmentInformation) {
			if($AppAssignment.assignmentTargetGroupId) {
				if($AzureADGroupsWithAssignmentsHashTable.ContainsKey($AppAssignment.assignmentTargetGroupId)) {
					$AppAssignment.devices = $AzureADGroupsWithAssignmentsHashTable[$AppAssignment.assignmentTargetGroupId].YodamiittiCustomGroupMembersCountDevices
					$AppAssignment.users = $AzureADGroupsWithAssignmentsHashTable[$AppAssignment.assignmentTargetGroupId].YodamiittiCustomGroupMembersCountUsers
				}
			}
		}


		# Create copy of Assignment array for exporting later in the script
		# We do copy now when data is pure. In the next steps we'll add html specific information to columns
		# which would break export data
		if($ExportCSV -or $ExportJSON -or $ExportToExcelCopyPaste -or $Passthru) {
			
			# We need to do deep copy to make unique array
			# Otherwise changes made next on the script would change this copied array also
			
			# This does not work!
			#$AzureADGroupsWithAssignmentsForExport = $AppsWithAssignmentInformation
			
			# Deep copy array of objects (this does not work for nested objects)
			$AzureADGroupsWithAssignmentsForExport = $AppsWithAssignmentInformation | ForEach-Object { $_.psobject.Copy() }

		}
		
		# DEBUG - copy to clipboard
		#$AzureADGroupsWithAssignments | ConvertTo-Json -Depth 5 | Set-Clipboard
		
		
		# Add additional html table information to values for example App link, AssignmentGroup link&Tooltip and Filter link
		$i = 0
		Foreach($AppAssignment in $AppsWithAssignmentInformation) {
			# Add IntuneApp, AssignmentGroup and IntuneFilter information
			
			######################################
			# App displayName
			[String]$AppToolTip = ""
			$AppToolTip += "<h3>$($AppAssignment.displayName)</h2>"

			$version = $AppAssignment.productVersion
			$AppToolTip +=	"Version: $version<br>"
			
			if($AppAssignment.id -and ($AllAppsHashTable.ContainsKey($AppAssignment.id))) {
				$AppObject = $AllAppsHashTable[$AppAssignment.id]
				
				$publisher = $AppObject.publisher
				$AppToolTip +=	"Publisher: $publisher<br>"
				
				$description = $AppObject.description
				# Convert newlines to <br>
				# If we don't remove newlines then Intent column background coloring regex fails
				if($description) {
					$description = $description.Replace("`r`n",'<br>')
					$description = $description.Replace("`n",'<br>')
				}
				$AppToolTip +=	"Description: $description<br><br>"
				
				# IntuneWin32 and WindowsMobileMSI
				if($fileName = $AppObject.fileName) {
					$AppToolTip +=	"FileName: $fileName<br>"
				}
				
				# IntuneWin32
				if($size = $AppObject.size) {
					$size = $size / 1MB
					$size = [math]::round($size, 2)
					$AppToolTip +=	"Size: $size MB<br><br>"
				}

				# IntuneWin32
				if($installCommandLine = $AppObject.installCommandLine) {
					$AppToolTip +=	"InstallCommandLine: $installCommandLine<br>"
				}
				
				# IntuneWin32
				if($uninstallCommandLine = $AppObject.uninstallCommandLine) {
					$AppToolTip +=	"UninstallCommandLine: $uninstallCommandLine<br><br>"	
				}
				
				# WinGet
				if($packageIdentifier = $AppObject.packageIdentifier) {
					$AppToolTip +=	"PackageIdentifier: $packageIdentifier<br>"
				}
				
				# WinGet
				if($runAsAccount = $AppObject.installExperience.runAsAccount) {
					$AppToolTip +=	"RunAsAccount: $runAsAccount<br><br>"
				}
				
				# iOS VPP Apps
				if($vppTokenOrganizationName = $AppObject.vppTokenOrganizationName) {
					$AppToolTip +=	"vppTokenOrganizationName: $vppTokenOrganizationName<br>"
				}
				
				# iOS VPP Apps
				if($vppTokenAccountType = $AppObject.vppTokenAccountType) {
					$AppToolTip +=	"vppTokenAccountType: $vppTokenAccountType<br>"
				}
				
				# iOS VPP Apps
				if($vppTokenAppleId = $AppObject.vppTokenAppleId) {
					$AppToolTip +=	"vppTokenAppleId: $vppTokenAppleId<br>"
				}

				# iOS VPP Apps and microsoftStoreForBusinessApp
				if($totalLicenseCount = $AppObject.totalLicenseCount) {
					$AppToolTip +=	"TotalLicenseCount: $totalLicenseCount<br>"
				}		
			}


			$AppAssignment.displayName = "<!-- $($AppAssignment.displayName) --><a href=`"https://intune.microsoft.com/#view/Microsoft_Intune_Apps/SettingsMenu/~/2/appId/$($AppAssignment.id)`" target=`"_blank`" class=`"tooltip`">$($AppAssignment.displayName)<span class=`"tooltiptext`">$AppToolTip</span></a>"
			
			######################################
			# AssignmentGroup displayName
			
			# Set initial value to empty string just in case we would miss this later
			[String]$AzureADGroupToolTip = ""

			if((-not $AppAssignment.assignmentTargetGroupId) -and ($AppAssignment.assignmentTargetGroupDisplayName -eq 'All Users')) {
				$AzureADGroupToolTip = "Intune Built-in <strong>All Users</strong> group"
				
				$AppAssignment.assignmentTargetGroupDisplayName = "<!-- $($AppAssignment.assignmentTargetGroupDisplayName) --><a href=`"#`" class=`"tooltip`">$($AppAssignment.assignmentTargetGroupDisplayName)<span class=`"tooltiptext`">$AzureADGroupToolTip</span></a>"
			}
			
			if((-not $AppAssignment.assignmentTargetGroupId) -and ($AppAssignment.assignmentTargetGroupDisplayName -eq 'All Devices')) {
				$AzureADGroupToolTip = "Intune Built-in <strong>All Devices</strong> group"
				
				$AppAssignment.assignmentTargetGroupDisplayName = "<!-- $($AppAssignment.assignmentTargetGroupDisplayName) --><a href=`"#`" class=`"tooltip`">$($AppAssignment.assignmentTargetGroupDisplayName)<span class=`"tooltiptext`">$AzureADGroupToolTip</span></a>"
			}
			
			# Builtin All Users and All Devices are not real AzureAD Groups so they don't have group id value
			if($AppAssignment.assignmentTargetGroupId) {
				
				$AzureADGroupToolTip += "<h3>$($AppAssignment.assignmentTargetGroupDisplayName)</h2>"
				
				$description = $AzureADGroupsWithAssignmentsHashTable[$AppAssignment.assignmentTargetGroupId].description
				$AzureADGroupToolTip +=	"Description: $description<br><br>"
				
				$groupType = "$($AzureADGroupsWithAssignmentsHashTable[$AppAssignment.assignmentTargetGroupId].YodamiittiCustomMembershipType) $($AzureADGroupsWithAssignmentsHashTable[$AppAssignment.assignmentTargetGroupId].YodamiittiCustomGroupType)"
				$AzureADGroupToolTip += "<strong>$groupType</strong> AzureAD group<br><br>"
				
				# Add Dynamic membershipRule if it exists
				if($AzureADGroupsWithAssignmentsHashTable[$AppAssignment.assignmentTargetGroupId].membershipRule) {
					$membershipRule = $AzureADGroupsWithAssignmentsHashTable[$AppAssignment.assignmentTargetGroupId].membershipRule
					$AzureADGroupToolTip += "DynamicRule: <strong>$membershipRule</strong><br><br>"
				}					
												
				$devicesCount = $AzureADGroupsWithAssignmentsHashTable[$AppAssignment.assignmentTargetGroupId].YodamiittiCustomGroupMembersCountDevices
				$usersCount = $AzureADGroupsWithAssignmentsHashTable[$AppAssignment.assignmentTargetGroupId].YodamiittiCustomGroupMembersCountUsers
				$AzureADGroupToolTip += "$devicesCount devices, $usersCount users"
				
				$AppAssignment.assignmentTargetGroupDisplayName = "<!-- $($AppAssignment.assignmentTargetGroupDisplayName) --><a href=`"https://intune.microsoft.com/#view/Microsoft_AAD_IAM/GroupDetailsMenuBlade/~/Overview/groupId/$($AppAssignment.assignmentTargetGroupId)`" target=`"_blank`" class=`"tooltip`">$($AppAssignment.assignmentTargetGroupDisplayName)<span class=`"tooltiptext`">$AzureADGroupToolTip</span></a>"
			}
			


			######################################
			# Filter displayName
			
			# Get Intune Filter object
			# Yes, this is not optimal but is fast enough when there are not thousands of filters
			# Hashtable would be fastest but maybe next time we do that here :)
			$IntuneFilterObject = $AllIntuneFilters | Where-Object id -eq $AppAssignment.assignmentFilterId
			
			$IntuneFilterToolTip = ""
			$IntuneFilterToolTip += "<h3>$($AppAssignment.assignmentFilterDisplayName)</h2>"
			$IntuneFilterToolTip += "Description: $($IntuneFilterObject.description)<br><br>"
			$IntuneFilterToolTip += "Platform:    $($IntuneFilterObject.platform)<br><br>"
			$IntuneFilterToolTip += "Rule:          <strong>$($IntuneFilterObject.rule)</strong>"
			
			$AppAssignment.assignmentFilterDisplayName = "<!-- $($AppAssignment.assignmentFilterDisplayName) --><a href=`"https://intune.microsoft.com/#view/Microsoft_Intune_DeviceSettings/AssignmentFilterSummaryBlade/assignmentFilterId/$($AppAssignment.assignmentFilterId)/filterType~/2`" target=`"_blank`" class=`"tooltip`">$($AppAssignment.assignmentFilterDisplayName)<span class=`"tooltiptext`">$IntuneFilterToolTip</span></a>"
			
			$i++	
		}
	
    }
    catch {
        Write-Error "$($_.Exception.GetType().FullName)"
        Write-Error "$($_.Exception.Message)"
        Write-Error "Error creating Application custom object"
        Write-Host "Script will exit..."
        Pause
        Exit 1
    }

    ###################################################
    # Get App Icons information from cache

    # Left option to NOT to download App Icons
    # Why would anyone would not to have good looking report ?-)
    if ((-not ($DoNotDownloadAppIcons))) {

        $CacheIconFiles = $false
        $CacheIconFiles = Get-ChildItem -File -Path "$PSScriptRoot\cache" -Include '*.jpg', '*.jpeg', '*.png' -Recurse | Select-Object -Property Name, FullName, BaseName, Extension, Length

        # Create array of objects which have App.DisplayName and App.id
        # Sort array and get unique to download icon only once per application (id)
        $AppIconDownloadList = $AppsWithAssignmentInformation | Select-Object -Property id, displayName | Sort-Object id -Unique

        Write-Host "Checking if we need to download App Icons..."
        foreach ($ManagedApp in $AppIconDownloadList) {

            #Write-Verbose "Processing App: $($ManagedApp.displayName)"

            # Initialize variable
            $LargeIconFullPathAndFileName = $null

            #Write-Host "Downloading App icon files to local cache from Graph API for App $($ManagedApp.displayName)"
            
            # Check if we have application icon already in cache folder
            # Use existing icon if found
            $IconFileObject = $null
            $IconFileObject = $CacheIconFiles | Where-Object { $_.BaseName -eq $ManagedApp.id }
            if ($IconFileObject -and (-not ($UpdateIconsCache))) {
                $LargeIconFullPathAndFileName = $IconFileObject.FullName
                #Write-Verbose "Found icon file ($LargeIconFullPathAndFileName) for Application id $($ManagedApp.id)"
                #Write-Host "Found cached App icon for App $($ManagedApp.displayName)"
            }
            else {
                try {
                    # Try to download largeIcon-attribute from application and save file to cache folder
                    
					# Extract Application name from $ManagedApp.displayName
					# That includes HTML Table formatting so we don't have just AppName anymore
					$Matches = $null
					if($ManagedApp.displayName -Match '^.*\<\!-- (.*) --\>.*$') {
						$displayName = $Matches[1]
					} else {
						$displayName = 'N/A'
					}
					
					Write-Host "Downloading icon for Application $displayName (id:$($ManagedApp.id))"

                    # Get application largeIcon attribute
                    $AppId = $ManagedApp.id
                    $url = "https://graph.microsoft.com/beta/deviceAppManagement/mobileApps/$($AppId)?`$select=largeIcon"

                    $appLargeIcon = Invoke-MSGraphRequest -Url $url -HttpMethod GET

                    #Write-Host "Invoke-MSGraphRequest succeeded: $?"

                    #largeIcon      : @{type=image/png; value=iVBORw0KGg
                    #$appLargeIcon.largeIcon.type

                    if (($appLargeIcon.largeIcon.type -ne $null) -and ($appLargeIcon.largeIcon.value -ne $null)) {
                        $filetype = ($appLargeIcon.largeIcon.type).Split('/')[1]
                        $largeIconBase64 = $appLargeIcon.largeIcon.value
                    }
                    else {
                        # There is no largeIcon attribute so we create empty file
                        # We create empty file so we know next time that there was no icon in Graph API
                        # This is workaround not to try find non-existing icons over and over again
                        # To check if icon has been added to Intune requires manual deletion of all zero sized "icon" files and running this script
                        # Or specify option -UpdateIconsCache
                        $filetype = "png"
                        $largeIconBase64 = ''
                    }
        
                    $LargeIconFilename = "$($AppId).$($filetype)"
                    $LargeIconFullPathAndFileName = "$PSScriptRoot\cache\$LargeIconFilename"
                    
                    try {
                        $return = Convert-Base64ToFile $largeIconBase64 $LargeIconFullPathAndFileName
                        Write-Verbose "Convert-Base64ToFile ApplicationId:$AppId $LargeIconFullPathAndFileName success: $return"
                    }
                    catch {
                        Write-Host "Error converting Base64 to file. Continuing to next app..." -ForegroundColor "Red"
                    }
                }
                catch {
                    Write-Host "Error downloading icon for app: $($ManagedApp.displayName). Continuing to next application..." -ForegroundColor "Red"
                }
            }
        }
    } else {
        Write-Host "Skipping App Icon downloads..."
    }
    $AppIconDownloadList = $null

    ########################################################################################################################
    # Create HTML report

$head = @'
<style>
    body {
        background-color: #FFFFFF;
        font-family: Arial, sans-serif;
    }

	  header {
		background-color: #444;
		color: white;
		padding: 10px;
		display: flex;
		align-items: center;
	  }

	  header h1 {
		margin: 0;
		font-size: 24px;
		margin-right: 20px;
	  }

	  header .additional-info {
		display: flex;
		flex-direction: column;
		align-items: flex-start;
		justify-content: center;
	  }

	  header .additional-info p {
		margin: 0;
		line-height: 1.2;
	  }

	  header .author-info {
		display: flex;
		flex-direction: column;
		align-items: flex-end;
		justify-content: center;
		margin-left: auto;
	  }

	  header .author-info p {
		margin: 0;
		line-height: 1.2;
	  }
	  
	  header .author-info a {
		color: white;
		text-decoration: none;
	  }

	  header .author-info a:hover {
		text-decoration: underline;
	  }

    table {
        border-collapse: collapse;
        width: 100%;
        text-align: left;
    }

    table, table#TopTable {
        border: 2px solid #1C6EA4;
        background-color: #f7f7f4;
    }

    table td, table th {
        border: 2px solid #AAAAAA;
        padding: 5px;
    }

    table td {
        font-size: 15px;
    }

    table th {
        font-size: 18px;
        font-weight: bold;
        color: #FFFFFF;
        background: #1C6EA4;
        background: -moz-linear-gradient(top, #5592bb 0%, #327cad 66%, #1C6EA4 100%);
        background: -webkit-linear-gradient(top, #5592bb 0%, #327cad 66%, #1C6EA4 100%);
        background: linear-gradient(to bottom, #5592bb 0%, #327cad 66%, #1C6EA4 100%);
    }

    table#TopTable td, table#TopTable th {
        vertical-align: top;
        text-align: center;
    }

    table thead th:first-child {
        border-left: none;
    }

	table thead th span {
		font-size: 14px;
		margin-left: 4px;
		opacity: 0.7;
	}

    table tfoot {
        font-size: 16px;
        font-weight: bold;
        color: #FFFFFF;
        background: #D0E4F5;
        background: -moz-linear-gradient(top, #dcebf7 0%, #d4e6f6 66%, #D0E4F5 100%);
        background: -webkit-linear-gradient(top, #dcebf7 0%, #d4e6f6 66%, #D0E4F5 100%);
        background: linear-gradient(to bottom, #dcebf7 0%, #d4e6f6 66%, #D0E4F5 100%);
        border-top: 2px solid #444444;
    }

    table tfoot .links {
        text-align: right;
    }

    table tfoot .links a {
        display: inline-block;
        background: #1C6EA4;
        color: #FFFFFF;
        padding: 2px 8px;
        border-radius: 5px;
    }

	table tbody tr:nth-child(even) {
	  background-color: #D0E4F5;
	}

	select {
	  font-family: "Courier New", monospace;
	}

	
  footer {
	background-color: #444;
	color: white;
	padding: 10px;
	display: flex;
	align-items: center;
	justify-content: center;
  }


  footer .creator-info {
	display: flex;
	flex-direction: row;
	align-items: center;
	margin-right: 20px;
  }

  footer .creator-info p {
	line-height: 1.2;
	margin: 0;
  }

  footer .creator-info p.author-text {
    margin-right: 20px; /* Add margin-right rule here */
  }

  .profile-container {
	position: relative;
	width: 50px;
	height: 50px;
	border-radius: 50%;
	overflow: hidden;
	margin-right: 10px;
  }

  .profile-container img {
	width: 100%;
	height: 100%;
	object-fit: cover;
	transition: opacity 0.3s;
  }

  .profile-container img.black-profile {
	position: absolute;
	top: 0;
	left: 0;
	z-index: 1;
  }

  .profile-container:hover img.black-profile {
	opacity: 0;
  }

  footer .company-logo {
	width: 100px;
	height: auto;
	margin: 0 20px;
  }

  footer a {
	color: white;
	text-decoration: none;
  }

  footer a:hover {
	text-decoration: underline;
  }

  
	.filter-row {
	  display: flex;
	  align-items: center;
	}

	.control-group {
	  display: flex;
	  flex-direction: column;
	  align-items: flex-start;
	  margin-right: 16px;
	}

	.control-group label {
	  font-weight: bold;
	}

    /* Tooltip container */
    .tooltip {
      position: relative;
      display: inline-block;
      cursor: pointer;
	  /* text-decoration: none; */ /* Remove underline from hyperlink */
	  color: inherit; /* Make the hyperlink have the same color as the text */
    }

    /* Tooltip text */
    .tooltip .tooltiptext {
      visibility: hidden;
      /* width: 120px; */
      background-color: #555;
      color: #fff;
      text-align: left;
      border-radius: 6px;
      position: absolute;
      z-index: 1;
      bottom: 125%;
      left: 50%;
      margin-left: -60px;
      opacity: 0;
      transition: opacity 1s;
	  white-space: pre;
	  padding: 10px; /* Change this value to suit your needs */
    }

    /* Show tooltip text when hovering */
    .tooltip:hover .tooltiptext {
      visibility: visible;
      opacity: 1;
    }
</style>
'@

    ############################################################
    # Application Summary
    # Create Application summary object for Apps which have assignments
    $ApplicationSummary = $AllApps | Where-object { $_.Assignments } | Select-Object -Property '@odata.type' | Sort-Object -Property '@odata.type' | Group-Object '@odata.type' | Select-Object -Property name, count

    # Remove #microsoft.graph.
    $ApplicationSummary | ForEach-Object { $_.Name = ($_.Name).Replace('#microsoft.graph.', '') }

    $ApplicationSummaryHTML = $ApplicationSummary | ConvertTo-Html -Fragment -PreContent "<h2 id=`"AppAssignmentSummary`">Apps Assignments Summary</h2>" | Out-String

    #### Set icon file path to every app ####
    Foreach ($App in $AppsWithAssignmentInformation) {
        $IconFileName = $App.id

        # Check if icon file exist and it is not zero size
        $IconFile = $null
        $IconFile = Get-ChildItem "$PSScriptRoot\cache\$($IconFileName)*"

        if (($IconFile) -and ($IconFile.Length -gt 0)) {

            $ImageFilePath = $IconFile.FullName
            $ImageType = ($IconFile.Extension).Replace('.', '')

            if ($IncludeBase64ImagesInReport) {
                # Include base64 encoded image in HTML report
                # Creates huge HTML files!!!
                $IconBase64 = [convert]::ToBase64String((Get-Content $ImageFilePath -encoding byte))
                $App.icon = "<img src=`"data:image/$ImageType;base64,$IconBase64`" height=`"25`" />"
            }
            else {
                # Add icon relative path from cache folder
                $App.icon = "<img src=`"./cache/$($IconFile.Name)`" height=`"25`" />"
            }            
        }
        else {
            # There is no icon file so we leave value empty
            $App.icon = 'no_icon'
        }
    }

    ######################
    Write-Host "Create Application Assignment information HTML fragment."

    try {

        # All Applications sorted by displayName by default

		if ($IncludeIdsInReport) {
			$AllAppsByDisplayName = $AppsWithAssignmentInformation | Select-Object -Property icon, '@odata.type', publisher, displayName, productVersion, assignmentIntent, assignmentTargetGroupDisplayName, devices, users, assignmentFilterDisplayName, FilterIncludeExclude, createdDateTime, lastModifiedDateTime, filename, id | Sort-Object displayName, id, assignmentIntent
		} else {
			# This default action is not to include id in HTML report
			
			$AllAppsByDisplayName = $AppsWithAssignmentInformation | Select-Object -Property icon, '@odata.type', publisher, displayName, productVersion, assignmentIntent, assignmentTargetGroupDisplayName, devices, users, assignmentFilterDisplayName, FilterIncludeExclude, createdDateTime, lastModifiedDateTime, filename, id | Sort-Object displayName, id, assignmentIntent
			
			# Remove id object
			$AllAppsByDisplayName = $AllAppsByDisplayName | Select-Object -Property * -ExcludeProperty id
		}
		
		
		$PreContent = @"
			<div class="filter-row">
				<div class="control-group">
				  <label><input type="checkbox" class="filterCheckbox" value="microsoftStoreForBusinessApp,officeSuiteApp,win32LobApp,windowsMicrosoftEdgeApp,windowsMobileMSI,windowsUniversalAppX,winGetApp,webApp" onclick="toggleCheckboxes(this)"> Windows</label>
				  <label><input type="checkbox" class="filterCheckbox" value="androidManagedStoreApp,managedAndroidStoreApp,managedAndroidLobApp,androidManagedStoreWebApp" onclick="toggleCheckboxes(this)"> Android</label>
				  <label><input type="checkbox" class="filterCheckbox" value="iosStoreApp,iosVppApp,managedIOSLobApp,managedIOSStoreApp,webApp" onclick="toggleCheckboxes(this)"> iOS</label>
				  <label><input type="checkbox" class="filterCheckbox" value="macOSOfficeSuiteApp,macOSLobApp,macOSMicrosoftEdgeApp,macOsVppApp" onclick="toggleCheckboxes(this)"> macOS</label>
				</div>
				<!-- Dropdown 1 -->
				<div class="control-group">
					<label for="dropdown1">App Type</label>
					<select id="filterDropdown1" multiple>
					  <option value="all" selected>All</option>
					</select>
				</div>
				<!-- Dropdown 2 -->
				<div class="control-group">
					<label for="dropdown1">Target Group</label>
					<select id="filterDropdown2">
					  <option value="all">All</option>
					</select>
				</div>
				<!-- Dropdown 3 -->
				<div class="control-group">
				<label for="dropdown1">Filter name</label>
				<select id="filterDropdown3">
					  <option value="all">All</option>
					</select>
				</div>
			</div>
			<br>
			<div>
				<input type="text" id="searchInput" placeholder="Search...">
				<button id="clearSearch" onclick="clearSearch()">X</button>
				<button id="resetFilter" onclick="resetFilters()">Reset filters</button>
			</div>
"@
		
		$AllAppsByDisplayNameHTML = $AllAppsByDisplayName | ConvertTo-Html -As Table -Fragment -PreContent $PreContent

        # Fix &lt; &quot; etc...
        $AllAppsByDisplayNameHTML = Fix-HTMLSyntax $AllAppsByDisplayNameHTML

        # Fix column names
        $AllAppsByDisplayNameHTML = Fix-HTMLColumns $AllAppsByDisplayNameHTML

		# Add TableId
		$TableId = 'IntuneApps'
		$AllAppsByDisplayNameHTML = $AllAppsByDisplayNameHTML.Replace('<table>',"<table id=`"$TableId`">")

		# Add Column on-click sorting and up/down arrows showing column is sortable
		#<th onclick="sortTable(1, 'IntuneApps')">App type <span>&#8597;</span></th>
		
		#DEBUG
		# <tr><th>icon</th><th>@odata.type</th><th>displayname</th><th>assignmentIntent</th><th>assignmentTargetGroupDisplayName</th><th>assignmentFilterDisplayName</th><th>FilterIncludeExclude</th><th>publisher</th><th>productVersion</th><th>filename</th><th>createdDateTime</th><th>lastModifiedDateTime</th><th>id</th></tr>
		
		# $i is array index
		$i = 0
		Foreach ($Line in $AllAppsByDisplayNameHTML) {
			# Add arrow
			if($Line -like '<tr><th>*') {
				# This is HTML Table header line
				
				# Create table header line start value which we will update on following section
				$TableHeaderLine = '<tr>'
				
				# Example string for regex testing in regex101.com
				# <tr><th>icon</th><th>App type</th><th>App name</th><th>Assignment Intent</th><th>Target Group</th><th>Filter name</th><th>Filter Intent</th><th>Publisher</th><th>Version</th><th>Filename</th><th>Created</th><th>Modified</th><th>id</th></tr>
				
				# This catches <th>...</th> to regex group so it includes <th> and </th> tags
				$regex = '(?i)<\s*th[^>]*>.*?<\s*/\s*th>'

				# I tried to use -Match operator but that didn't work.
				# GPT-4 gave following answer to use Select-String to get all matches
				#
				# "I see the confusion now. The -Match operator in PowerShell is used for matching a single instance,
				# and it populates the $Matches automatic variable with the match groups.
				# However, since you want to match all instances of the <th> elements,
				# you should use the -AllMatches option with the Select-String cmdlet instead."

				$b = 0
				if($Matches = $Line | Select-String -Pattern $regex -AllMatches) {
					Foreach ($Match in $Matches.Matches) {
						# Add ColumnSort and Arrow icon
						# Example:
						# <th>App type</th>
						# <th onclick="sortTable(1, 'IntuneApps')">App type <span>&#8597;</span></th>

						[String]$String = $Match.Value

						# Add sorting Arrow span value
						$String = $String.Replace('</th>',' <span>&#8597;</span></th>')
						
						# Add ColumnSort value
						$String = $String.Replace('<th>',"<th onclick=`"sortTable($b, '$($TableId)', [11,12])`">")
						
						$TableHeaderLine += $String
					
						$b++
					}
					$TableHeaderLine += '</tr>'
					$AllAppsByDisplayNameHTML[$i] = $TableHeaderLine
					
				} else {
					Write-Host "`nRegex failed, fix script!`n" -ForegroundColor Red
				}
				
				#Write-Host "DEBUG: $Line"
				#Write-Host "DEBUG2: $($AllAppsByDisplayNameHTML[$i])"
				
			} else {
				# Change Intent column background color based on Intent value

				# Test string for regex for testing in https://regex101.com
				#                 <td>Available Included</td>
				#                 <td>Required Included</td>
				
				# We specify ? to make lazy regex. For example: Uninstall.*?
				# Otherwise regexp will be greedy and try to match for the end of the line (string
				$regex = '^(.*)(<td>)((?:Required.*?|Uninstall.*?|Available.*?|AvailableWithoutEnrollment.*?))(<\/td>)(<td>.*)$'
				if($Line -Match $regex) {

					$LineStart = $Matches[1]
					# $Matches[2] is <td>
					$assignmentIntent = $Matches[3]
					# $Matches[4] is </td>
					$LineEnd = $Matches[5]
					
					#Write-Host "DEBUG: `$assignmentIntent = $assignmentIntent"
					
					
					# Change color for Intent column based on Intent value
					if ($assignmentIntent -like "Required*") {
						$AllAppsByDisplayNameHTML[$i] = "$LineStart<td bgcolor=`"lightgreen`"><font color=`"black`">$assignmentIntent</font></td>$LineEnd"
					}
					
					if ($assignmentIntent -like "Uninstall*") {
						$AllAppsByDisplayNameHTML[$i] = "$LineStart<td bgcolor=`"lightSalmon`"><font color=`"black`">$assignmentIntent</font></td>$LineEnd"
					}
					
					if ($assignmentIntent -like "Available*") {
						$AllAppsByDisplayNameHTML[$i] = "$LineStart<td bgcolor=`"lightyellow`"><font color=`"black`">$assignmentIntent</font></td>$LineEnd"
					}
					
					if ($assignmentIntent -like "AvailableWithoutEnrollment*") {
						$AllAppsByDisplayNameHTML[$i] = "$LineStart<td bgcolor=`"lightyellow`"><font color=`"black`">$assignmentIntent</font></td>$LineEnd"
					}
				}
			}

			$Matches = $null
			$i++
		}
			
		# Convert HTML Array to String which is requirement for HTTP PostContent
		$AllAppsByDisplayNameHTML = $AllAppsByDisplayNameHTML | Out-String


        # Debug- save $html1 to file
        #$AllAppsByDisplayNameHTML | Out-File "$PSScriptRoot\AllAppsByDisplayNameHTML.html"

    }
    catch {
        Write-Error "$($_.Exception.GetType().FullName)"
        Write-Error "$($_.Exception.Message)"
        Write-Error "Error creating OS specific HTML fragment information"
        Write-Host "Script will exit..."
        Pause
        Exit 1        
    }
    #############################
    # Create html
    Write-Host "Creating HTML report..."

    try {

        $ReportRunDateTime = (Get-Date).ToString("yyyyMMddHHmm")
        $ReportRunDateTimeHumanReadable = (Get-Date).ToString("yyyy-MM-dd HH:mm")
        $ReportRunDateFileName = (Get-Date).ToString("yyyyMMddHHmm")

        $ReportSavePath = $PSScriptRoot
        $HTMLFileName = "$($ReportRunDateFileName)_Intune_Application_Assignments_report.html"
		$CSVFileName = "$($ReportRunDateFileName)_Intune_Application_Assignments_report.csv"
		$JSONFileName = "$($ReportRunDateFileName)_Intune_Application_Assignments_report.json"

        $PreContent = @"
		<header>
		  <h1>Intune Application Assignments Report ver $ScriptVersion</h1>
		  <div class="additional-info">
			<p><strong>Report run:</strong> $ReportRunDateTimeHumanReadable <strong>by</strong> $($ConnectMSGraph.UPN)</p>
			<p><strong>Tenant name:</strong> $TenantDisplayName</p>
			<!-- <p><strong>Tenant id:</strong> $($ConnectMSGraph.TenantId)</p> -->
		  </div>
		  <div class="author-info">
			<p><a href="https://github.com/petripaavola/Intune/tree/master/Reports" target="_blank"><strong>Download Report tool from GitHub</strong></a><br>Author: Petri Paavola - Microsoft MVP</p>
		  </div>
		</header>
		<br>
"@

		$JavascriptPostContent = @'
		<p id="noResultsMessage" style="display: none;">No results found.</p>
		<p><br></p>
		<footer>
			<div class="creator-info">
			<p class="author-text">Author:</p>
			  <div class="profile-container">
				<img src="data:image/png;base64,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" alt="Profile Picture Petri Paavola">
				<img class="black-profile" src="data:image/png;base64,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" alt="Black Profile Picture Petri Paavola">
			  </div>
			  <p><strong>Petri Paavola</strong><br>
				<a href="mailto:Petri.Paavola@yodamiitti.fi">Petri.Paavola@yodamiitti.fi</a><br>
				Senior Modern Management Principal
			  </p>
			</div>
			<img class="company-logo" src="data:image/png;base64,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" alt="Microsoft MVP">
			<p style="margin: 0;">
			  <a href="https://github.com/petripaavola/Intune/tree/master/Reports" target="_blank"><strong>Download report tool from GitHub</strong<</a>
			</p>
		</footer>
		<script>
			function updateRowBackgroundOnColumnValueChange(tableId, columnIndex) {
			  let table = document.getElementById(tableId);
			  let rows = table.getElementsByTagName("tr");
			  let previousValue = null;
			  let currentColor = "rgba(208, 228, 245, 1)";

			  table.setAttribute("data-last-color", currentColor);

			  for (let i = 1; i < rows.length; i++) {
				let row = rows[i];

				// Skip hidden rows
				if (row.style.display === "none") {
				  continue;
				}

				let currentValue = row.getElementsByTagName("td")[columnIndex].textContent;

				if (previousValue !== null && currentValue !== previousValue) {
				  currentColor = table.getAttribute("data-last-color") === "rgba(242, 242, 242, 1)" ? "rgba(208, 228, 245, 1)" : "rgba(242, 242, 242, 1)";
				  table.setAttribute("data-last-color", currentColor);
				}

				row.style.backgroundColor = currentColor;
				previousValue = currentValue;
			  }
			}


			function setColumnBold(tableId, columnIndex) {
			  let table = document.getElementById(tableId);
			  let rows = table.getElementsByTagName("tr");

			  // Unbold previously selected column
			  if (table.hasAttribute("data-bold-column")) {
				let previousBoldColumn = parseInt(table.getAttribute("data-bold-column"));

				rows[0].getElementsByTagName("th")[previousBoldColumn].style.fontWeight = "normal";
				for (let i = 1; i < rows.length; i++) {
				  rows[i].getElementsByTagName("td")[previousBoldColumn].style.fontWeight = "normal";
				}
			  }

			  // Set header text to bold
			  rows[0].getElementsByTagName("th")[columnIndex].style.fontWeight = "bold";

			  // Set column values to bold
			  for (let i = 1; i < rows.length; i++) {
				rows[i].getElementsByTagName("td")[columnIndex].style.fontWeight = "bold";
			  }

			  // Save current bold column index
			  table.setAttribute("data-bold-column", columnIndex);
			}


			function mergeSort(arr, comparator) {
			  if (arr.length <= 1) {
				return arr;
			  }

			  const mid = Math.floor(arr.length / 2);
			  const left = mergeSort(arr.slice(0, mid), comparator);
			  const right = mergeSort(arr.slice(mid), comparator);

			  return merge(left, right, comparator);
			}

			function merge(left, right, comparator) {
			  let result = [];
			  let i = 0;
			  let j = 0;

			  while (i < left.length && j < right.length) {
				if (comparator(left[i], right[j]) <= 0) {
				  result.push(left[i]);
				  i++;
				} else {
				  result.push(right[j]);
				  j++;
				}
			  }

			  return result.concat(left.slice(i)).concat(right.slice(j));
			}

			// Declare a sortingDirections object outside the sortTable function to store the sorting directions for each column.
			const sortingDirections = {};

			// Specify which columns will be sorted as integers
			const integerColumns = [7, 8];

			function sortTable(n, tableId, dateColumns = []) {
			  let table, rows;
			  table = document.getElementById(tableId);

			  // Initialize the sorting direction for the column if it hasn't been set yet
			  if (!(n in sortingDirections)) {
				sortingDirections[n] = "asc";
			  }

			  // Remove existing arrow icons
			  let headerRow = table.getElementsByTagName("th");
			  for (let i = 0; i < headerRow.length; i++) {
				headerRow[i].innerHTML = headerRow[i].innerHTML.replace(
				  /<span>.*<\/span>/,
				  "<span>&#8597;</span>"
				);
			  }

			  const isDateColumn = dateColumns.includes(n);
			  rows = Array.from(table.rows).slice(1);

			  const comparator = (a, b) => {
				const x = a.cells[n].innerHTML.toLowerCase();
				const y = b.cells[n].innerHTML.toLowerCase();
				const isIntegerColumn = integerColumns.includes(n);

				if (isDateColumn) {
				  const xDate = getDateFromString(x);
				  const yDate = getDateFromString(y);

				  if (sortingDirections[n] === "asc") {
					return xDate - yDate;
				  } else {
					return yDate - xDate;
				  }
				} else if (isIntegerColumn) {
					const xInt = parseInt(x, 10) || 0;  // Use 0 if parsing fails
					const yInt = parseInt(y, 10) || 0;  // Use 0 if parsing fails
					if (sortingDirections[n] === "asc") {
					  return xInt - yInt;
					} else {
					  return yInt - xInt;
					}
				} else {
				  if (sortingDirections[n] === "asc") {
					return x.localeCompare(y);
				  } else {
					return y.localeCompare(x);
				  }
				}
			  };

			  const sortedRows = mergeSort(rows, comparator);

			  // Reinsert sorted rows into the table
			  for (let i = 0; i < sortedRows.length; i++) {
				table.tBodies[0].appendChild(sortedRows[i]);
			  }

			  // Update arrow icon for the last sorted column
			  if (sortingDirections[n] === "asc") {
				headerRow[n].innerHTML = headerRow[n].innerHTML.replace(
				  /<span>.*<\/span>/,
				  "<span>&#x25B2;</span>"
				);
			  } else {
				headerRow[n].innerHTML = headerRow[n].innerHTML.replace(
				  /<span>.*<\/span>/,
				  "<span>&#x25BC;</span>"
				);
			  }

			  // Create row coloring based on selected column
			  updateRowBackgroundOnColumnValueChange("IntuneApps", n);

			  // Bold selected column header and text
			  setColumnBold("IntuneApps", n);

			  // Toggle sorting direction for the next click
			  sortingDirections[n] = sortingDirections[n] === "asc" ? "desc" : "asc";
			}


			function getDateFromString(dateStr) {
				let [day, month, year, hours, minutes, seconds] = dateStr.split(/[. :]/);
				return new Date(year, month - 1, day, hours, minutes, seconds);
			}

			// Populates the given select element with unique values from the specified table and column
			function populateSelectWithUniqueColumnValues(tableId, column, selectId) {
			  let table = document.getElementById(tableId);
			  let rows = table.getElementsByTagName("tr");
			  let uniqueValues = {};

			  for (let i = 1; i < rows.length; i++) {
				let cellValue = rows[i].getElementsByTagName("td")[column].innerText;

				if (uniqueValues[cellValue]) {
				  uniqueValues[cellValue]++;
				} else {
				  uniqueValues[cellValue] = 1;
				}
			  }

			  let select = document.getElementById(selectId);

			  // Convert the uniqueValues object to an array of key-value pairs
			  let uniqueValuesArray = Object.entries(uniqueValues);

			  // Sort the array by the keys (unique column values)
			  uniqueValuesArray.sort((a, b) => a[0].localeCompare(b[0]));

			  // Find the longest text
			  let longestTextLength = Math.max(...uniqueValuesArray.map(([value, count]) => (value + " (" + count + ")").length));

			  // Loop through the sorted array to create the options with padded number values
			  for (let [value, count] of uniqueValuesArray) {
				let optionText = value + " (" + count + ")";
				let paddingLength = longestTextLength - optionText.length;
				let padding = "\u00A0".repeat(paddingLength);
				let option = document.createElement("option");
				option.value = value;
				option.text = value + padding + " (" + count + ")";
				select.add(option);
			  }
			}

			// This is used to extract AssignmentGroupDisplayName from complex a tag with span
			function getDirectChildTextNotWorking(parentNode) {
				let childNodes = parentNode.childNodes;
				let textContent = '';

				for(let i = 0; i < childNodes.length; i++) {
					if(childNodes[i].nodeType === Node.TEXT_NODE) {
						textContent += childNodes[i].nodeValue;
					}
				}

				return textContent.trim();  // Remove leading/trailing whitespaces
			}


			// Returns the textContent of the node if the node doesn't have any child nodes
			// In some Columns we have a href tag so we in those cases we need to get display value from <a> tag
			function getDirectChildText(node) {
			  if (!node || !node.hasChildNodes()) return node.textContent.trim();
			  return Array.from(node.childNodes)
				.filter(child => child.nodeType === Node.TEXT_NODE)
				.map(textNode => textNode.textContent)
				.join("");
			}


			// Filters the table based on the selected dropdown values, checkboxes, and search input
			// Should find AssignmentGroup and Filter displayNames for filtering from a href tag
			function combinedFilter(tableId, columnIndexForDropdown1, columnIndexForDropdown2, columnIndexForDropdown3, columnIndexForCheckboxes) {
			  let table = document.getElementById(tableId);
			  let rows = table.getElementsByTagName("tr");
			  let checkboxes = document.getElementsByClassName("filterCheckbox");
			  let dropdown1 = document.getElementById("filterDropdown1");
			  let dropdown2 = document.getElementById("filterDropdown2");
			  let dropdown3 = document.getElementById("filterDropdown3");
			  let searchInput = document.getElementById("searchInput");
			  let searchText = searchInput.value.toLowerCase();

			  let selectedDropdownValues1 = Array.from(dropdown1.selectedOptions).map(option => option.value);

			  for (let i = 1; i < rows.length; i++) {
				let row = rows[i];
				let cell1 = row.getElementsByTagName("td")[columnIndexForDropdown1];
				let cell2 = row.getElementsByTagName("td")[columnIndexForDropdown2];
				let cell3 = row.getElementsByTagName("td")[columnIndexForDropdown3];

				let cellValueDropdown1 = getDirectChildText(cell1.querySelector('a') || cell1);
				let cellValueDropdown2 = getDirectChildText(cell2.querySelector('a') || cell2);
				let cellValueDropdown3 = getDirectChildText(cell3.querySelector('a') || cell3);

				let showRowByDropdown1 = selectedDropdownValues1.includes("all") || selectedDropdownValues1.includes(cellValueDropdown1);
				let showRowByDropdown2 = dropdown2.value === "all" || cellValueDropdown2 === dropdown2.value;
				let showRowByDropdown3 = dropdown3.value === "all" || cellValueDropdown3 === dropdown3.value;

				let showRowByCheckboxes = true;
				for (let checkbox of checkboxes) {
				  if (checkbox.checked) {
					let cellValue = row.getElementsByTagName("td")[columnIndexForCheckboxes].textContent;
					let checkboxValues = checkbox.value.split(",");
					if (!checkboxValues.includes(cellValue)) {
					  showRowByCheckboxes = false;
					  break;
					}
				  }
				}

				let showRowBySearch = true;
				if (searchText) {
				  showRowBySearch = false;
				  let cells = row.getElementsByTagName("td");
				  for (let cell of cells) {
					if (getDirectChildText(cell.querySelector('a') || cell).toLowerCase().includes(searchText)) {
					  showRowBySearch = true;
					  break;
					}
				  }
				}

				row.style.display = (showRowByDropdown1 && showRowByDropdown2 && showRowByDropdown3 && showRowByCheckboxes && showRowBySearch) ? "" : "none";
			  }

			  let visibleRowCount = 0;
			  for (let i = 1; i < rows.length; i++) {
				if (rows[i].style.display !== 'none') {
				  visibleRowCount++;
				}
			  }

			  const noResultsMessage = document.getElementById('noResultsMessage');
			  if (visibleRowCount === 0) {
				noResultsMessage.style.display = 'block';
			  } else {
				noResultsMessage.style.display = 'none';
			  }
			}
			// function combinedFilter ends


			// Unchecks the other checkboxes in the group and updates the table filters
			function toggleCheckboxes(checkbox) {
			  let checkboxes = document.getElementsByClassName("filterCheckbox");
			  
			  for (let cb of checkboxes) {
				if (cb !== checkbox) {
				  cb.checked = false;
				}
			  }
			  
			  combinedFilter('IntuneApps', 1, 6, 9, 1);
			}


			// Clears the search input and updates the table filters
			function clearSearch() {
			  let searchInput = document.getElementById("searchInput");
			  searchInput.value = "";
			  combinedFilter('IntuneApps', 1, 6, 9, 1);
			}

			// Resets all filters and updates the table
			function resetFilters() {
			  let searchInput = document.getElementById("searchInput");
			  searchInput.value = "";
			  
			  let checkboxes = document.getElementsByClassName("filterCheckbox");
			  for (let checkbox of checkboxes) {
				checkbox.checked = false;
			  }
			  
			  let filterDropdown1 = document.getElementById("filterDropdown1");
			  filterDropdown1.value = "all";
			  
			  let filterDropdown2 = document.getElementById("filterDropdown2");
			  filterDropdown2.value = "all";
			  
			  let filterDropdown3 = document.getElementById("filterDropdown3");
			  filterDropdown3.value = "all";
			  
			  combinedFilter('IntuneApps', 1, 6, 9, 1);
			}



			// Event listeners for the dropdowns and checkboxes
			document.getElementById("filterDropdown1").addEventListener("change", function() {
			  combinedFilter("IntuneApps", 1, 6, 9, 1);
			});

			document.getElementById("filterDropdown2").addEventListener("change", function() {
			  combinedFilter("IntuneApps", 1, 6, 9, 1);
			});
			
			document.getElementById("filterDropdown3").addEventListener("change", function() {
			  combinedFilter("IntuneApps", 1, 6, 9, 1);
			});

			let checkboxes = document.getElementsByClassName("filterCheckbox");
			for (let checkbox of checkboxes) {
			  checkbox.addEventListener("change", function() {
				combinedFilter("IntuneApps", 1, 6, 9, 1);
			  });
			}

			// Add an event listener for the search input
			document.getElementById("searchInput").addEventListener("input", function() {
			  combinedFilter("IntuneApps", 1, 6, 9, 1);
			});

		// Another approach to get function to run when loading page
		// This is not needed but left here on purpose just in case needed in the future
		//window.addEventListener('load', function() {
		//  sortTable(2, 'IntuneApps', [10, 11]);
		//});		

		
		window.onload = function() {
			
			// Call this function to populate the first dropdown with unique values from the specified table and column
			populateSelectWithUniqueColumnValues("IntuneApps", 1, "filterDropdown1");

			// Call this function to populate the second dropdown with unique values from the specified table and column
			populateSelectWithUniqueColumnValues("IntuneApps", 6, "filterDropdown2");
			
			// Call this function to populate the third dropdown with unique values from the specified table and column
			populateSelectWithUniqueColumnValues("IntuneApps", 9, "filterDropdown3");

			// Not needed anymore because sorting will also do this automatically
			//updateRowBackgroundOnColumnValueChange('IntuneApps', 2);

			// Sort table by App name so user knowns which column was sorted
			sortTable(3, 'IntuneApps', [11,12]);
		};
		</script>
'@

        $Title = "Intune Application Assignment report"
        ConvertTo-HTML -head $head -PostContent $AllAppsByDisplayNameHTML, $JavascriptPostContent -PreContent $PreContent -Title $Title | Out-File "$ReportSavePath\$HTMLFileName"
        $Success = $?

        if (-not ($Success)) {
            Write-Error "Error creating HTML file."
            Write-Host "Script will exit..."
            Pause
            Exit 1
        }
        else {
            Write-Host "Intune Application Assignment report HTML file created:`n`t$ReportSavePath\$HTMLFileName`n" -ForegroundColor Green
        }
		
		# Exporting data to other formats if selected
		if($ExportCSV) {
			Write-Host "Export data to CSV file: $ReportSavePath\$CSVFileName"
			$AzureADGroupsWithAssignmentsForExport | Select-Object -Property '@odata.type', publisher, displayName, productVersion, assignmentIntent, assignmentTargetGroupDisplayName, devices, users, assignmentFilterDisplayName, FilterIncludeExclude, createdDateTime, lastModifiedDateTime, filename, id | Sort-Object displayName, id, assignmentIntent | Export-Csv -Path "$ReportSavePath\$CSVFileName" -Delimiter ";" -Encoding UTF8 -NoTypeInformation -NoClobber
			$Success = $?
			if($Success) {
				Write-Host "Success: OK`n" -ForegroundColor Green
			} else {
				Write-Host "Success: Failed`n" -ForegroundColor Red
			}
		}
		
		if($ExportJSON) {
			Write-Host "Export data to JSON file: $ReportSavePath\$JSONFileName"
			$AzureADGroupsWithAssignmentsForExport | Select-Object -Property '@odata.type', publisher, displayName, productVersion, assignmentIntent, assignmentTargetGroupDisplayName, devices, users, assignmentFilterDisplayName, FilterIncludeExclude, createdDateTime, lastModifiedDateTime, filename, id | Sort-Object displayName, id, assignmentIntent | ConvertTo-Json -Depth 3 | Out-File -FilePath "$ReportSavePath\$JSONFileName" -Encoding UTF8
			$Success = $?
			if($Success) {
				Write-Host "Success: OK`n" -ForegroundColor Green
			} else {
				Write-Host "Success: Failed`n" -ForegroundColor Red
			}
		}
		
		if($ExportToExcelCopyPaste) {
			# Thanks Juha-Matti for this tip ;)
			
			Write-Host "Export data to Clipboard so you can paste it to Excel"
			$AzureADGroupsWithAssignmentsForExport | Select-Object -Property '@odata.type', publisher, displayName, productVersion, assignmentIntent, assignmentTargetGroupDisplayName, devices, users, assignmentFilterDisplayName, FilterIncludeExclude, createdDateTime, lastModifiedDateTime, filename, id | Sort-Object displayName, id, assignmentIntent | ConvertTo-Csv -Delimiter "`t" -NoTypeInformation | Set-Clipboard
			$Success = $?
			if($Success) {
				Write-Host "Success: OK`n" -ForegroundColor Green
			} else {
				Write-Host "Success: Failed`n" -ForegroundColor Red
			}
		}
    }
    catch {
        Write-Error "$($_.Exception.GetType().FullName)"
        Write-Error "$($_.Exception.Message)"
        Write-Error "Error creating HTML report: $ReportSavePath\$HTMLFileName"
        Write-Host "Script will exit..."
        Pause
        Exit 1
    }

    ############################################################
    # Open HTML file

    # Check file exists and is bigger than 0
    # File should exist already but years ago slow computer/disk caused some problems
    # so this is hopefully not needed workaround
    # Wait max. of 20 seconds

	if(-not $DoNotOpenReportAutomatically) {
		$i = 0
		$filesize = 0
		do {
			Write-Host "Double check HTML file creation is really done (round $i)"
			$filesize = 0
			Start-Sleep -Seconds 2
			try {
				$HTMLFile = Get-ChildItem "$ReportSavePath\$HTMLFileName"
				$filesize = $HTMLFile.Length
			}
			catch {
				# Something went wrong, waiting for next round.
				Write-Host "Trouble getting file size, waiting 2 seconds and trying again..."
			}
			if ($filesize -eq 0) { Write-Host "Filesize is 0kB so waiting for a while for file creation to finish" }

			$i += 1
		} while (($i -lt 10) -and ($filesize -eq 0))

		Write-Host "Opening created file:`n$ReportSavePath\$HTMLFileName`n"
		try {
			Invoke-Item "$ReportSavePath\$HTMLFileName"
		}
		catch {
			Write-Host "Error opening file automatically to browser. Open file manually:`n$ReportSavePath\$HTMLFileName`n" -ForegroundColor Red
		}
	} else {
		Write-Host "`nNote! Parameter -DoNotOpenReportAutomatically specified. Report was not opened automatically to web browser`n" -ForegroundColor Yellow
	}
}
catch {
    Write-Error "Uups! Something happened and we failed. Try again..."

    Write-Error "$($_.Exception.GetType().FullName)"
    Write-Error "$($_.Exception.Message)"
}
