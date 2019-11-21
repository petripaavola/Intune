# Create HTML Report from all Intune App Assignments
# and show information on Out-GridView also (default)
#
# Petri.Paavola@yodamiitti.fi
# 20191121

# Configure if you want to see information on Out-GridView also
$ShowInformationInOutGridView = $true
#$ShowInformationInOutGridView = $false


#####

# Get All AzureAD security Groups
# This should be more efficient than searching AzureADGroup for every assignment one by one
# Even though we may end up getting quite a lot of security groups

# Notice we probably get only part of groups because GraphAPI returns limited number of groups
$groups = Get-AADGroup -Filter 'securityEnabled eq true'

# Check if we have value starting https:// in attribute @odate.nextLink
# If we have nextLink then we get all groups
if ($groups.'@odata.nextLink' -like "https://*") {
    # Get all groups because we got paged result
    $AllGroups = Get-MSGraphAllPages -SearchResult $groups
}
else {
    $AllGroups = $groups
}

######

# Get App information

# Update Graph API schema to beta to get Win32LobApps also
Update-MSGraphEnvironment -SchemaVersion 'beta'
Connect-MSGraph

# We need assignments info so -Expand assignment option is needed here
$Apps = Get-DeviceAppManagement_MobileApps -Expand assignments

# Check if we have value starting https:// in attribute @odate.nextLink
# If we have nextLink then we get all Apps
if ($Apps.'@odata.nextLink' -like "https://*") {
    # Get all Apps because we got paged result
    $AllApps = Get-MSGraphAllPages -SearchResult $Apps
}
else {
    $AllApps = $Apps
}

# Find apps which have assignments to our selected AzureAD Group id
# Check data syntax from GraphAPI with request: https://graph.microsoft.com/v1.0/deviceAppManagement/mobileApps?$expand=assignments
# or convert $Apps to json to get more human readable format: $Apps | ConvertTo-JSON
#$AppsWithAssignmentToSelectedGroup = $AllApps | Where-Object { $_.assignments.target.groupid -like "*" }

# Create custom object array and gather necessary app and assignment information. We need appId and assignmentId to remove selected assignments. All other attributes are for humans
$AppsWithAssignmentInformation = @()

# Go through each app and save necessary information to custom object
Foreach ($App in $AllApps) {
    Foreach ($Assignment in $App.Assignments) {
    
        #$Assignment = $Assignments | Where-Object { $_.target.groupid -eq $SelectedAzureADGroup.id }

        $properties = @{
            '@odata.type'                    = $App.'@odata.type'
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
            assignmentTargetGroupDisplayName = $AllGroups | Where { $_.id -eq $Assignment.target.groupid } | Select -ExpandProperty displayName
        }

        # Create new custom object every time inside foreach-loop
        # This is really important step to do inside foreach-loop!
        # If you create custom object outside of foreach then you would edit same custom object on every foreach cycle resulting only 1 app in custom object array
        $CustomObject = New-Object -TypeName PSObject -Prop $properties

        # Add custom object to our custom object array.
        $AppsWithAssignmentInformation += $CustomObject
    }
}


if ($ShowInformationInOutGridView) {
    # Show All Apps and Assignments in Out-GridView
    $AppsWithAssignmentInformation | Select '@odata.type', displayName, assignmentIntent, assignmentTargetGroupDisplayName, publisher, productVersion, filename, createdDateTime, lastModifiedDateTime, id, assignmentId, assignmentTargetGroupId | Sort displayName | Out-GridView -Title "Found these App Assignments"
}


# Create HTML report

# Commented out to create different color on every second table line because it is harder to read in this case
# We should group coloring based on application and/or AzureADGroup (one color per application or one color per AzureADGroup)
<#
table tr:nth-child(even) {
  background: #D0E4F5;
}
#>

$head = @'
<style>
table {
  border: 1px solid #1C6EA4;
  background-color: #EEEEEE;
  width: 100%;
  text-align: left;
  border-collapse: collapse;
}
table td, table th {
  border: 1px solid #AAAAAA;
  padding: 3px 2px;
}
table tbody td {
  font-size: 13px;
}
table thead {
  background: #1C6EA4;
  background: -moz-linear-gradient(top, #5592bb 0%, #327cad 66%, #1C6EA4 100%);
  background: -webkit-linear-gradient(top, #5592bb 0%, #327cad 66%, #1C6EA4 100%);
  background: linear-gradient(to bottom, #5592bb 0%, #327cad 66%, #1C6EA4 100%);
  border-bottom: 2px solid #444444;
}
table thead th {
  font-size: 15px;
  font-weight: bold;
  color: #FFFFFF;
  border-left: 2px solid #D0E4F5;
}
table thead th:first-child {
  border-left: none;
}
table tfoot {
  font-size: 14px;
  font-weight: bold;
  color: #FFFFFF;
  background: #D0E4F5;
  background: -moz-linear-gradient(top, #dcebf7 0%, #d4e6f6 66%, #D0E4F5 100%);
  background: -webkit-linear-gradient(top, #dcebf7 0%, #d4e6f6 66%, #D0E4F5 100%);
  background: linear-gradient(to bottom, #dcebf7 0%, #d4e6f6 66%, #D0E4F5 100%);
  border-top: 2px solid #444444;
}
table tfoot td {
  font-size: 14px;
}
table tfoot .links {
  text-align: right;
}
table tfoot .links a{
  display: inline-block;
  background: #1C6EA4;
  color: #FFFFFF;
  padding: 2px 8px;
  border-radius: 5px;
}
</style>
'@

$ReportRunDateTime = (Get-Date).ToString("yyyyMMddHHmm")
$ReportRunDateTimeHumanReadable = (Get-Date).ToString("yyyy-MM-dd HH:mm")
$ReportRunDateFileName = (Get-Date).ToString("yyyyMMddHHmm")

$ReportSavePath = $PSScriptRoot
$HTMLFileName = "$($ReportRunDateFileName)_Intune_Application_Assignments_report.html"

# All attributes
#$html1 = $AppsWithAssignmentInformation | Select '@odata.type', displayName, assignmentIntent, assignmentTargetGroupDisplayName, publisher, productVersion, filename, createdDateTime, lastModifiedDateTime, id, assignmentId, assignmentTargetGroupId | ConvertTo-Html -Fragment -PreContent "<h2>Assignments</h2>" | Out-String

# Less attributes
$html1 = $AppsWithAssignmentInformation | Select '@odata.type', displayName, assignmentIntent, assignmentTargetGroupDisplayName, publisher, productVersion, filename, createdDateTime, lastModifiedDateTime | Sort displayName | ConvertTo-Html -Fragment -PreContent "<h2>App Assignments sorted with App displayName</h2>" | Out-String
$html2 = $AppsWithAssignmentInformation | Select '@odata.type', displayName, assignmentIntent, assignmentTargetGroupDisplayName, publisher, productVersion, filename, createdDateTime, lastModifiedDateTime | Sort assignmentTargetGroupDisplayName | ConvertTo-Html -Fragment -PreContent "<h2>App Assignments sorted with assignmentTargetGroupDisplayName</h2>" | Out-String

# Create and save html file
ConvertTo-HTML -head $head -PostContent $html1, $html2 -PreContent "<h1>Intune Application Assignments report<br>Report run: $ReportRunDateTimeHumanReadable</h1>" -Title "Intune Application Assignment report" | Out-File "$ReportSavePath\$HTMLFileName"
    
# Wait to make sure file is really wrote to disk (slow disks may cause problems)
# May not a good "workaround" but still used here :)
Start-Sleep -Seconds 2

# For manual testing open html file
Invoke-Item "$ReportSavePath\$HTMLFileName"
