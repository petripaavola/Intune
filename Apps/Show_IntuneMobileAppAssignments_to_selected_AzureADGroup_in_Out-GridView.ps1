# Show Intune Application assignments for selected AzureADGroup
# Uses Out-GridView to show GUI to select AzureADGroup and App(s)
#
# Petri.Paavola@yodamiitti.fi
# Microsoft MVP - Windows and Devices for IT
#
# 27.4.2021


######

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
$SelectedAzureADGroup = $AllGroups | Select displayName, description, createdDateTime, id, groupTypes | Sort displayName | Out-GridView -Title "Select AzureAD Group to show it's App Assignments"  -OutputMode Single

# Exit if nothing was selected in Out-GridView
if (-not $SelectedAzureADGroup) { Write-Output "No groups selected, exiting..."; Exit 0 }

######

# Get App information and show Apps in Out-GridView

# We need assignments info so -Expand assignment option is needed here
$Apps = Get-DeviceAppManagement_MobileApps -Expand assignments

# Quick workaround to get all Apps and also get actual objects and not .Value -attribute
$AllApps = Get-MSGraphAllPages -SearchResult $Apps

if(-not $AllApps) {
	Write-Error "Could not find any Intune Apps!"
	Exit 1
}

# Find apps which have assignments to our selected AzureAD Group id
# Check data syntax from GraphAPI with request: https://graph.microsoft.com/v1.0/deviceAppManagement/mobileApps?$expand=assignments
# or convert $Apps to json to get more human readable format: $Apps | ConvertTo-JSON
$AppsWithAssignmentToSelectedGroup = $AllApps | Where-Object { $_.assignments.target.groupid -eq $SelectedAzureADGroup.id }

# Exit if there were no assignments to selected AzureAD Group
if (-not $AppsWithAssignmentToSelectedGroup) {
    Write-Output "No Application assignments found for AzureAD Group: $($SelectedAzureADGroup.displayName)"
    Exit 0
}

# Create custom object array and gather necessary app and assignment information. We need appId and assignmentId to remove selected assignments. All other attributes are for humans
$AppsWithAssignmentInformation = @()

# Go through each app and save necessary information to custom object
Foreach ($App in $AppsWithAssignmentToSelectedGroup) {
    $Assignment = $App.Assignments | Where-Object { $_.target.groupid -eq $SelectedAzureADGroup.id }

	if($App.vppTokenAppleId) { $vppTokenAppleId = $App.vppTokenAppleId} else { $vppTokenAppleId = $null }

    $properties = @{
        '@odata.type'                    = $App.'@odata.type'
		vppTokenAppleId					 = $vppTokenAppleId
        displayname                      = $App.displayname
        productVersion                   = $App.productVersion
        publisher                        = $App.publisher
        filename                         = $App.filename
        createdDateTime                  = $App.createdDateTime
        lastModifiedDateTime             = $App.lastModifiedDateTime
        id                               = $App.id
        assignmentId                     = $Assignment.id
        assignmentIntent                 = $Assignment.intent
        assignmentTargetGroupId          = $Assignment.target.groupid
        assignmentTargetGroupDisplayName = $SelectedAzureADGroup.displayName
    }

    # Create new custom object every time inside foreach-loop
    # This is really important step to do inside foreach-loop!
    # If you create custom object outside of foreach then you would edit same custom object on every foreach cycle resulting only 1 app in custom object array
    $CustomObject = New-Object -TypeName PSObject -Prop $properties

    # Add custom object to our custom object array.
    $AppsWithAssignmentInformation += $CustomObject
}


# Show Apps in Out-GridView
$SelectedApps = $AppsWithAssignmentInformation | Select '@odata.type', vppTokenAppleId, displayName, assignmentIntent, assignmentTargetGroupDisplayName, publisher, productVersion, filename, createdDateTime, lastModifiedDateTime, id, assignmentId, assignmentTargetGroupId | Sort displayName | Out-GridView -Title "Found these App Assignment for AzureADGroup: $($SelectedAzureADGroup.displayName)"
