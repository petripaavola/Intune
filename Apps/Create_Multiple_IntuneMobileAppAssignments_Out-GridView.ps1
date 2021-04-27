# Create Multiple Intune Application assignments for selected Apps
# Uses Out-GridView to show GUI to select App(s), AzureADGroup and Intent
#
# Petri.Paavola@yodamiitti.fi
# Microsoft MVP - Windows and Devices for IT
#
# 27.4.2021


function Create-IntuneApplicationAssignment {
    Param(
        $AzureADGroupId,
        $mobileAppId,
        $AssignmentIntent
    )

    # Choices are: 'available', 'required', 'uninstall', 'availableWithoutEnrollment'

    $TargetObject = $null
    $TargetObject = New-Object PSObject
    $TargetObject | Add-Member NoteProperty '@odata.type' '#microsoft.graph.groupAssignmentTarget'
    $TargetObject | Add-Member NoteProperty 'groupId' $AzureADGroupId

    # Fix json - custom-object problem
    # This is not needed anymore 20190220
    #$TargetObject = $TargetObject |ConvertTo-json |ConvertFrom-json

    # Create Intune Application assignment
    New-DeviceAppManagement_MobileApps_Assignments -mobileAppId $mobileAppId -intent $AssignmentIntent -target $TargetObject

    $Success = $?
    return $Success
}

#####

# Update Graph API schema to beta to get Win32LobApps and possible other new features also
Update-MSGraphEnvironment -SchemaVersion 'beta'

$MSGraphEnvironment = Connect-MSGraph
$Success = $?

if (-not $Success) {
	Write-Error "Error connecting to Intune!"
	Exit 1
}

# Get AzureAD groups and show them in Out-GridView. User can select only 1 group

# Notice we probably get only part of groups because GraphAPI returns limited number of groups
$groups = Get-AADGroup -Filter 'securityEnabled eq true'

# Quick workaround to get all Security Groups and also get actual objects and not .Value -attribute
$AllGroups = Get-MSGraphAllPages -SearchResult $groups

if(-not $AllGroups) {
	Write-Error "Could not find any AzureAD Groups!"
	Exit 1
}

# Show AzureAD groups in Out-GridView with selected properties and save selected AzureADGroup to variable
Write-Output "Select AzureAD Group from Out-GridView"
$SelectedAzureADGroup = $AllGroups | Select displayName, description, createdDateTime, id, groupTypes | Sort displayName | Out-GridView -Title "Select AzureAD Group to make assignment" -OutputMode Single

# Exit if nothing was selected in Out-GridView
if (-not $SelectedAzureADGroup) { Write-Output "No groups selected, exiting..."; Exit 0 }

#####

# Get App information and show Apps in Out-GridView

# We need assignments info to check existing assignments so -Expand assignment option is needed here
$Apps = Get-DeviceAppManagement_MobileApps -Expand assignments

# Quick workaround to get all Apps and also get actual objects and not .Value -attribute
$AllApps = Get-MSGraphAllPages -SearchResult $Apps

if(-not $AllApps) {
	Write-Error "Could not find any Intune Apps!"
	Exit 1
}

# Show Apps in Out-GridView, show only specified app properties
Write-Output "Select Apps in Out-GridView to App assignments creation for group: $($SelectedApp.assignmentTargetGroupDisplayName)"
$SelectedApps = $AllApps | Select '@odata.type', vppTokenAppleId, displayName, productVersion, publisher, fileName, size, commandLine, productCode, publishingState, createdDateTime, lastModifiedDateTime, id| Sort displayName | Out-GridView -PassThru -Title 'Select Application(s) to make assignment'

# Exit if nothing was selected in Out-GridView
if (-not $SelectedApps) { Write-Output "No Apps selected, exiting..."; Exit 0 }

#####

# Get Intent with Out-GridView
Write-Output "Select Intent for App Assignment"
$SelectedIntent = @('required', 'available', 'uninstall', 'availableWithoutEnrollment') | Out-GridView -Title 'Select Intent for application assignment' -OutputMode Single

# Exit if nothing was selected in Out-GridView
if (-not $SelectedIntent) { Write-Output "No App assignment intent selected, exiting..."; Exit 0 }


# Make assignments for all selected Apps
foreach ($SelectedApp in $SelectedApps) {

    # Filter out those Apps that already have assignment for selected AzureADGroup
    $ExistingAssignment = $AllApps | Where-Object { ($_.id -eq $SelectedApp.id) -and ($_.Assignments.target.groupid -eq $SelectedAzureADGroup.id) }

    # If there is already is existing Assignment then show warning and continue to next App in foreach loop
    if ($ExistingAssignment) {
        # There is already assignment for this AzureADGroup
        # Show warning and skip this App assignment

        $intent = $ExistingAssignment | Where-Object { $_.Assignments.target.groupid -eq $SelectedAzureADGroup.id } | Select -ExpandProperty Assignments | Select -ExpandProperty Intent
        `Write-Host "Warning: Existing $intent Assignment for App: `"$($SelectedApp.displayName)`"" found for AzureADGroup: `"$($SelectedAzureADGroup.displayName)`"". Skipping this App assignment..." -ForegroundColor Yellow

        # Skip this App in foreach loop and process to next App
        Continue
    }

    # Create Application assignment - most important line in this script
    $Success = Create-IntuneApplicationAssignment $SelectedAzureADGroup.Id $SelectedApp.Id $SelectedIntent
    if ($Success) {
        Write-Host "Success: Created $SelectedIntent Application assignment for Application: `"$($SelectedApp.DisplayName)`" for AzureADGroup: `"$($SelectedAzureADGroup.DisplayName)`"" -Foreground Green
    }
    else {
        Write-Host "Failed: Error creating $SelectedIntent Application assignment for Application: `"$($SelectedApp.DisplayName)`" for AzureADGroup: `"$($SelectedAzureADGroup.DisplayName)`"" -Foreground Red
    }
}
