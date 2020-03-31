# Create HTML Report from all Intune Configuration Assignments for devices
#
# Petri.Paavola@yodamiitti.fi
# Microsoft MVP - Windows and Devices for IT
# 20200330
#
# https://github.com/petripaavola/Intune
#
# Script downloads AzureADGroup and Intune Configuration information from Graph API to local cache folder (.\cache)
#
# Created files are:
# .\cache\AdministrativeTemplates.json
# .\cache\AllGroups.json
# .\cache\CompliancePolicies.json
# .\cache\ConfigurationProfiles.json
# .\cache\windowsAutopilotDeploymentProfiles
# .\cache\deviceEnrollmentConfigurations.json
# .\cache\deviceManagementScripts.json
# .\cache\WindowsFeatureUpdatesPolicies.json

# You can work with cached data without network connection with command
# Create_IntuneConfigurationAssignments_HTML_Report.ps1 -UseOfflineCache
#
# To include id use parameter
# Create_IntuneConfigurationAssignments_HTML_Report.ps1 -IncludeIdsInReport

Param(
    [Switch]$UseOfflineCache,
    [Switch]$IncludeIdsInReport
)

$ScriptVersion = "ver 1.61"


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

function Invoke-MSGraphGetRequestWithMSGraphAllPages {
    param (
        [Parameter(Mandatory = $true)]
        [String]$url
    )

    $MSGraphRequest = $null
    $AllMSGraphRequest = $null

    try {
        $MSGraphRequest = Invoke-MSGraphRequest -Url $url -HttpMethod 'GET'
        $Success = $?

        if($Success) {
            
            # This does not work because we won't catch this if there is Value-attribute which is null
            #if ($MSGraphRequest.Value) {

            # Test if object has attribute named Value (whether value is null or not)
            if((Get-Member -inputobject $MSGraphRequest -name 'Value' -Membertype Properties) -and (Get-Member -inputobject $MSGraphRequest -name '@odata.context' -Membertype Properties)) {

                # Value property exists. We should get here most of the time
                $returnObject = $MSGraphRequest.Value

            } else {
                # Sometimes we get results without Value-attribute (eg. getting user details)
                # We will return all we got as is

                $returnObject = $MSGraphRequest
            }
        } else {
            # Invoke-MSGraphRequest failed so we return false
            return -1
        }

        # Check if we have value starting https:// in attribute @odate.nextLink
        # If we have nextLink then we get GraphAllPages
        if ($MSGraphRequest.'@odata.nextLink' -like 'https://*') {

            # Get AllMSGraph pages
            # This is also workaround to get objects without assigning them from .Value attribute
            $AllMSGraphRequest = Get-MSGraphAllPages -SearchResult $MSGraphRequest
            $Success = $?

            if($Success) {
                $returnObject = $AllMSGraphRequest
            } else {
                # Getting Get-MSGraphAllPages failed
                return -1
            }
        }

        return $returnObject

    } catch {
        Write-Error "There was error with MSGraphRequest with url $url!"
        return -1
    }
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

    # Remove c column
    $html = $html -replace '<th>c</th>',''
    $html = $html -replace '<th>@odata.type</th>','<th>Profile type</th>'
    $html = $html -replace '<th>displayname</th>','<th>Profile name</th>'
    $html = $html -replace '<th>assignmentIntent</th>','<th>Assignment Intent</th>'
    $html = $html -replace '<th>assignmentTargetGroupDisplayName</th>','<th>Target Group</th>'
    $html = $html -replace '<th>createdDateTime</th>','<th>Created</th>'
    $html = $html -replace '<th>lastModifiedDateTime</th>','<th>Modified</th>'

    return $html
}


function Create-GroupingRowColors {
    Param(
        $htmlObject,
        $GroupByAttributeName = 'id'
    )

    $TableRowColor = 'D0E4F5'
    $PreviousRowID = $null

    Foreach ($Row in $htmlObject) {

        $CurrentRowID = $Row.$GroupByAttributeName

        # Check if current ConfigurationProfile is with same id than previous
        if (($CurrentRowID -eq $PreviousRowID) -or ($PreviousRowID -eq $null)) {
            # Same ConfigurationProfile, no need to change table color (whatever that color is)
            #Write-Verbose "Same, not changing color"
        }
        else {
            # Current row ConfigurationProfile is different than previous one. We need to change row color.

            # We need to change table color now because ConfigurationProfile is changing
            if ($TableRowColor -eq 'D0E4F5') {
                $TableRowColor = 'EEEEEE'
            }
            else {
                $TableRowColor = 'D0E4F5'
            }
            #Write-Verbose "Change color to: $TableRowColor"
        }

        # Set attribute c with color information
        # We will use this later with MatchEvaluator regexp
        $Row.c = $TableRowColor

        $PreviousRowID = $CurrentRowID
    }

    return $htmlObject
}

function Change-HTMLTableSyntaxWithRegexpForConfigurationsSortedByDisplayName {
    Param(
        $html
    )

    $MatchEvaluatorSortedByConfigurationName = {
        param($match)

        # Change intent cell background color
        $intent = $match.Groups[10].Value

        # Set first letter to capital
        $intent = (Get-Culture).TextInfo.ToTitleCase($intent.tolower())

        # Set Default
        $IntentTD = "<td>$intent</td>"

        if ($intent -eq "included") { $IntentTD = "<td bgcolor=`"lightgreen`"><font color=`"black`">$intent</font></td>" }
        if ($intent -eq "excluded") { $IntentTD = "<td bgcolor=`"lightSalmon`"><font color=`"black`">$intent</font></td>" }
        
        $RowColor = $match.Groups[3].Value
        $odatatype = $match.Groups[5].Value
        $ConfigurationDisplayName = $match.Groups[7].Value
        $assignmentTargetGroupDisplayName = $match.Groups[13].Value

        # We may want to include ConfigurationProfileId for manual debugging if we have several Configurations with same names
        # We do not include this by default because last column ConfigurationProfileId is so long that it will grow each row height
        # and that doesn't look that good
        if ($IncludeIdsInReport) {
            $id = $match.Groups[17].Value
        }
        else {
            $id = ''
        }

        # This is returned from MatchEvaluator
        #"<tr bgcolor=`"$RowColor`"><td></td><td>$odatatype</td><td style=`"font-weight:bold`">$ConfigurationDisplayName</font></td>$IntentTD<td>$assignmentTargetGroupDisplayName</td>$($match.Groups[15].Value)</td><td>$id</td></tr>"

        # Removed column c
        "<tr bgcolor=`"$RowColor`"><td>$odatatype</td><td style=`"font-weight:bold`">$ConfigurationDisplayName</font></td>$IntentTD<td>$assignmentTargetGroupDisplayName</td>$($match.Groups[15].Value)</td><td>$id</td></tr>"
    }

    # $html1 is now single string object with multiple lines separated by newline
    # We need to convert $html1 to array of String objects so we can do foreach
    # otherwise foreach only sees 1 string and we can't acccess every line individually
    $html = @($html -split '[\r\n]+')

    # Example string. Use string and regex in https://regex101.com
    #<tr><td>EEEEEE</td><td>windowsDeliveryOptimizationConfiguration</td><td>Delivery Optimization</td><td>Included</td><td>DynDev - AllWindowsMDMDevices</td><td>10.12.2018 10.15.48</td><td>10.12.2018 10.15.48</td><td>f99f3e16-8592-476e-a986-2b84b58987d4</td></tr>
   
    # Original
    #$regex = '^(<tr>)(<td>)(.*?)(<\/td><td>)(.*?)(<\/td><td>)(.*?)(<\/td>)(<td>)(.*?)(<\/td>)(<td>)(.*?)(<\/td>)(.*)(<td>)([a-f0-9]{8}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{12})(<\/td><\/tr>)$'

    # Enrollment configurations have different id :/
    # for example: a0e55368-ef3c-4c2e-be11-25bb6952f34b_Windows10EnrollmentCompletionPageConfiguration
    #<tr><td>EEEEEE</td><td>windows10EnrollmentCompletionPageConfiguration</td><td>Enrollment Status Page</td><td>Included</td><td>DynDev - AllWindowsMDMDevices</td><td>10.12.2018 10.15.48</td><td>10.12.2018 10.15.48</td><td>a0e55368-ef3c-4c2e-be11-25bb6952f34b_Windows10EnrollmentCompletionPageConfiguration</td></tr>
    $regex = '^(<tr>)(<td>)(.*?)(<\/td><td>)(.*?)(<\/td><td>)(.*?)(<\/td>)(<td>)(.*?)(<\/td>)(<td>)(.*?)(<\/td>)(.*)(<td>)([a-f0-9]{8}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{12})(.*)(<\/td><\/tr>)$'

    # Do match evaluator regex -> tweak table icon and colors
    # Convert array of string objects back to String with | Out-String
    $html = $html | ForEach-Object { [regex]::Replace($_, $regex, $MatchEvaluatorSortedByConfigurationName) } | Out-String

    return $html
}

function Change-HTMLTableSyntaxWithRegexpForConfigurationProfilesSortedByAssignmentTargetGroupDisplayName {
    Param(
        $html
    )

    # Regexp magic :)
    $MatchEvaluatorSortedByAssignmentTargetGroupDisplayName = {
        param($match)
       
        # Change intent cell background color
        $intent = $match.Groups[10].Value

        # Set first letter to capital
        $intent = (Get-Culture).TextInfo.ToTitleCase($intent.tolower())

        # Set Default
        $IntentTD = "<td>$intent</td>"

        if ($intent -eq "included") { $IntentTD = "<td bgcolor=`"lightgreen`"><font color=`"black`">$intent</font></td>" }
        if ($intent -eq "excluded") { $IntentTD = "<td bgcolor=`"lightSalmon`"><font color=`"black`">$intent</font></td>" }
        
        $RowColor = $match.Groups[3].Value
        $odatatype = $match.Groups[5].Value
        $ConfigurationDisplayName = $match.Groups[7].Value
        $assignmentTargetGroupDisplayName = $match.Groups[13].Value
        $createdDateTime = $match.Groups[16].Value
        $lastModifiedDateTime = $match.Groups[19].Value

        # We may want to include ConfigurationProfileId for manual debugging
        # We do not include this by default because last column Id is so long that it will grow each row height
        # and that doesn't look that good
        if ($IncludeIdsInReport) {
            $id = $match.Groups[22].Value
        }
        else {
            $id = ''
        }

        # This is returned from MatchEvaluator
        #"<tr bgcolor=`"$RowColor`"><td></td><td>$odatatype</td><td>$ConfigurationDisplayName</td>$IntentTD<td style=`"font-weight:bold`">$assignmentTargetGroupDisplayName</font></td><td>$createdDateTime</td><td>$lastModifiedDateTime</td><td>$id</td></tr>"

        # Remove column c
        "<tr bgcolor=`"$RowColor`"><td>$odatatype</td><td>$ConfigurationDisplayName</td>$IntentTD<td style=`"font-weight:bold`">$assignmentTargetGroupDisplayName</font></td><td>$createdDateTime</td><td>$lastModifiedDateTime</td><td>$id</td></tr>"
    }

    # $html1 is now single string object with multiple lines separated by newline
    # We need to convert $html1 to array of String objects so we can do foreach
    # otherwise foreach only sees 1 string and we can't acccess every line individually
    $html = @($html -split '[\r\n]+')

    # Example string. Use string and regex in https://regex101.com
    #<tr><td>EEEEEE</td><td>windows10GeneralConfiguration</td><td>Windows 10 Device Restrictions Baseline for Autopilot Devices</td><td>Included</td><td>DynDev_Autopilot_GroupTag_AutopilotVM</td><td>20.11.2019 14.11.21</td><td>20.11.2019 14.12.08</td><td>0b49e17b-f3ef-4057-b161-37a24d7e7cfa</td></tr>
    #$regex = '^(<tr>)(<td>)(.*?)(<\/td><td>)(.*?)(<\/td><td>)(.*?)(<\/td>)(<td>)(.*?)(<\/td>)(<td>)(.*?)(<\/td>)(<td>)(.*?)(<\/td>)(<td>)(.*?)(<\/td>)(<td>)([a-f0-9]{8}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{12})(<\/td><\/tr>)$'

    # Enrollment configurations have different id :/
    # for example: a0e55368-ef3c-4c2e-be11-25bb6952f34b_Windows10EnrollmentCompletionPageConfiguration
    #<tr><td>EEEEEE</td><td>windows10EnrollmentCompletionPageConfiguration</td><td>Enrollment Status Page</td><td>Included</td><td>DynDev - AllWindowsMDMDevices</td><td>10.12.2018 10.15.48</td><td>10.12.2018 10.15.48</td><td>a0e55368-ef3c-4c2e-be11-25bb6952f34b_Windows10EnrollmentCompletionPageConfiguration</td></tr>
    $regex = '^(<tr>)(<td>)(.*?)(<\/td><td>)(.*?)(<\/td><td>)(.*?)(<\/td>)(<td>)(.*?)(<\/td>)(<td>)(.*?)(<\/td>)(<td>)(.*?)(<\/td>)(<td>)(.*?)(<\/td>)(<td>)([a-f0-9]{8}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{12})(.*)(<\/td><\/tr>)$'

    # Do match evaluator regex -> tweak table icon and colors
    # Convert array of string objects back to String with | Out-String
    $html = $html | ForEach-Object { [regex]::Replace($_, $regex, $MatchEvaluatorSortedByAssignmentTargetGroupDisplayName) } | Out-String

    return $html
}

####################################################################################################
# Main starts here

# Really quick and dirty error handling with huge try-catch block
try {

    # Yes we should take true/false return value back here and do messaging and exiting here based on that value
    # But it was so quick to write that staff to function so let's fix that on some later version (like never ?-)
    $return = Verify-IntuneModuleExistence

    # Create cache folder if it does not exist
    if (-not (Test-Path "$PSScriptRoot\cache")) {
        Write-Output "Creating cache directory: $PSScriptRoot\cache"
        New-Item -ItemType Directory "$PSScriptRoot\cache"
        $Success = $?

        if (-not ($Success)) {
            Write-Error "Could not create cache directory ($PSScriptRoot\cache). Check file system rights and try again."
            Write-Output "Script will exit..."
            Pause
            Exit 1
        }
    }

    #####
    # Change schema and get Tenant info

    if (-not ($UseOfflineCache)) {
        try {
            Write-Output "Get tenant information from Graph API and change Graph API schema to beta"

            # We have variables
            # $ConnectMSGraph.UPN
            # $ConnectMSGraph.TenantId

            # Update Graph API schema to beta to get latest and greatest results
            Update-MSGraphEnvironment -SchemaVersion 'beta'
            $ConnectMSGraph = Connect-MSGraph
            $Success = $?

            if (-not ($Success)) {
                Write-Error "Error connecting to Microsoft Graph API with command Connect-MSGraph."
                Write-Output "Check you have Intune Powershell Cmdlest installed with commmand: Install-Module -Name Microsoft.Graph.Intune"
                Write-Output "Script will exit..."
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
            Write-Output "Script will exit..."
            Pause
            Exit 1
        }
    }
    else {
        $TenantDisplayName = "(offline)"
    }

    #####

    # Get AzureADGroups. This should be more efficient than getting AzureADGroup for every assignment one by one

    # Test if we have AllGroups.json file
    if (-not (Test-Path "$PSScriptRoot\cache\AllGroups.json")) {
        Write-Output "Did NOT find AllGroups.json file. We have to get AzureAD Group information from Graph API"
        if ($UseOfflineCache) {
            Write-Host "Run script without option -UseOfflineCache to download necessary AllGroups information`n" -ForegroundColor "Yellow"
            Exit 0
        }
    }

    try {
        if (-not ($UseOfflineCache)) {
            Write-Output "Downloading all AzureAD Security Groups from Graph API (this might take a while)..."
            Write-Verbose "Downloading all AzureAD Security Groups from Graph API"

            # Notice we probably get only part of groups because GraphAPI returns limited number of groups
            $groups = Get-AADGroup -Filter 'securityEnabled eq true'
            $Success = $?

            if (-not ($Success)) {
                Write-Error "Error downloading AzureAD Security Groups"
                Write-Output "Script will exit..."
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
                    Write-Output "Script will exit..."
                    Pause
                    Exit 1
                }
            }
            else {
                $AllGroups = $groups
            }
            Write-Output "AzureAD Group information downloaded."

            # Save to local cache -Depth 2 is default value
            $AllGroups | ConvertTo-Json -Depth 2 | Out-File "$PSScriptRoot\cache\AllGroups.json" -Force
        }
        else {
            Write-Output "Using cached AzureAD Group information from file "$PSScriptRoot\cache\AllGroups.json""
        }

        # Get Group information from cached file always
        $AllGroups = Get-Content "$PSScriptRoot\cache\AllGroups.json" | ConvertFrom-Json

    }
    catch {
        Write-Error "$($_.Exception.GetType().FullName)"
        Write-Error "$($_.Exception.Message)"
        Write-Error "Error trying to download AzureAD Group information"
        Write-Output "Script will exit..."
        Pause
        Exit 1
    }

    ######
    # Test if we have AllGroups.json file
    if (-not (Test-Path "$PSScriptRoot\cache\AllConfigurationProfiles.json")) {
        Write-Output "Could NOT find AllConfigurationProfiles.json file. We have to get ConfigurationProfiles information from Graph API"
        if ($UseOfflineCache) {
            Write-Host "Run script without option -UseOfflineCache to download necessary AllConfigurationProfiles information`n" -ForegroundColor "Yellow"
            Exit 0
        }
    }

    $AllConfigurationProfiles = New-Object -TypeName "System.Collections.ArrayList"

    try {
        # Get ConfigurationProfiles information from Graph API
        if (-not ($UseOfflineCache)) {
            # Get ConfigurationProfiles information from Graph API
            Write-Output "Downloading Intune Configuration Profiles information from Graph API (this might take a while)..."

            # We need assignments info so -Expand assignment option is needed here
            #$ConfigurationProfiles = Get-IntuneDeviceConfigurationPolicy -Expand assignments
            
            # Intune uses this url for Windows Configuration profiles
            #https://graph.microsoft.com/beta/deviceManagement/deviceConfigurations?$select=id,displayName,lastModifiedDateTime,roleScopeTagIds,microsoft.graph.unsupportedDeviceConfiguration/originalEntityTypeName&$expand=assignments&top=500
            
            # We need assignments info so -Expand assignment option is needed here
            $url = "https://graph.microsoft.com/beta/deviceManagement/deviceConfigurations?`$expand=assignments"
            $ConfigurationProfiles = Invoke-MSGraphGetRequestWithMSGraphAllPages $url

            if ($ConfigurationProfiles -eq -1) {
                Write-Error "Error downloading Intune Configuration Profiles information"
                Write-Output "Script will exit..."
                Pause
                Exit 1
            }

            Write-Output "Intune Configuration Profiles information downloaded"

            # Save to local cache
            # Really important parameter!!! Specify -Depth 4 because assignment data will be nested down 4 levels
            $ConfigurationProfiles | ConvertTo-Json -Depth 4 | Out-File "$PSScriptRoot\cache\ConfigurationProfiles.json" -Force

        }
        else {
            Write-Output "Using cached Intune Configuration Profiles information from file $PSScriptRoot\cache\ConfigurationProfiles.json"
        }

        # Get Configuration information from cached file always
        $ConfigurationProfiles = Get-Content "$PSScriptRoot\cache\ConfigurationProfiles.json" | ConvertFrom-Json

        # Add ConfigurationProfiles to $AllConfigurationProfiles variable
        $AllConfigurationProfiles += $ConfigurationProfiles

    }
    catch {
        Write-Error "$($_.Exception.GetType().FullName)"
        Write-Error "$($_.Exception.Message)"
        Write-Error "Error trying to download Intune Configuration Profiles information"
        Write-Output "Script will exit..."
        Pause
        Exit 1
    }


    try {
        # Get AdministrativeTemplate (aka GroupPolicy) information from Graph API
        if (-not ($UseOfflineCache)) {
            # Get AdministrativeTemplate (aka GroupPolicy) from Graph API
            Write-Output "Downloading Intune AdministrativeTemplate (aka GroupPolicy) information from Graph API (this might take a while)..."

            # We need assignments info so -Expand assignment option is needed here
            $url = "https://graph.microsoft.com/beta/deviceManagement/groupPolicyConfigurations?`$expand=assignments"
            $AdminTemplates = Invoke-MSGraphGetRequestWithMSGraphAllPages $url

            if ($AdminTemplates -eq -1) {
                Write-Error "Error downloading Intune Administrative Templates information"
                Write-Output "Script will exit..."
                Pause
                Exit 1
            }

            Write-Output "Intune Administrative Templates information downloaded"

            # Save to local cache
            # Really important parameter!!! Specify -Depth 4 because assignment data will be nested down 4 levels
            $AdminTemplates | ConvertTo-Json -Depth 4 | Out-File "$PSScriptRoot\cache\AdministrativeTemplates.json" -Force

        }
        else {
            Write-Output "Using cached Intune Configuration Profiles information from file $PSScriptRoot\cache\AdministrativeTemplates.json"
        }

        # Get Configuration information from cached file always
        $AdminTemplates = Get-Content "$PSScriptRoot\cache\AdministrativeTemplates.json" | ConvertFrom-Json

        # Add odatatype manually (because we use this information later)
        #$odatatype = "#microsoft.graph.groupPolicyConfiguration"
        $AdminTemplates | Foreach-Object { $_ | Add-Member -MemberType NoteProperty -Name '@odata.type' -Value "#microsoft.graph.groupPolicyConfiguration" -Force }

        # Add Administrative Templates to $AllConfigurationProfiles variable
        $AllConfigurationProfiles += $AdminTemplates

    }
    catch {
        Write-Error "$($_.Exception.GetType().FullName)"
        Write-Error "$($_.Exception.Message)"
        Write-Error "Error trying to download Intune Administrative Templates information"
        Write-Output "Script will exit..."
        Pause
        Exit 1
    }


    try {
        # Get CompliancePolicies information from Graph API
        if (-not ($UseOfflineCache)) {
            # Get CompliancePolicies from Graph API
            Write-Output "Downloading Intune CompliancePolicies information from Graph API (this might take a while)..."

            # We need assignments info so -Expand assignment option is needed here
            $url = "https://graph.microsoft.com/beta/deviceManagement/deviceCompliancePolicies?`$expand=assignments"
            $CompliancePolicies = Invoke-MSGraphGetRequestWithMSGraphAllPages $url

            if ($CompliancePolicies -eq -1) {
                Write-Error "Error downloading Intune CompliancePolicies information"
                Write-Output "Script will exit..."
                Pause
                Exit 1
            }

            Write-Output "Intune CompliancePolicies information downloaded"

            # Save to local cache
            # Really important parameter!!! Specify -Depth 4 because assignment data will be nested down 4 levels
            $CompliancePolicies | ConvertTo-Json -Depth 4 | Out-File "$PSScriptRoot\cache\CompliancePolicies.json" -Force

        }
        else {
            Write-Output "Using cached Intune Configuration Profiles information from file $PSScriptRoot\cache\CompliancePolicies.json"
        }

        # Get Configuration information from cached file always
        $CompliancePolicies = Get-Content "$PSScriptRoot\cache\CompliancePolicies.json" | ConvertFrom-Json

        # Add CompliancePolicies to $AllConfigurationProfiles variable
        $AllConfigurationProfiles += $CompliancePolicies

    }
    catch {
        Write-Error "$($_.Exception.GetType().FullName)"
        Write-Error "$($_.Exception.Message)"
        Write-Error "Error trying to download Intune CompliancePolicies information"
        Write-Output "Script will exit..."
        Pause
        Exit 1
    }

    try {
        # Get WindowsFeatureUpdates information from Graph API
        if (-not ($UseOfflineCache)) {
            # Get WindowsFeatureUpdates from Graph API
            Write-Output "Downloading Intune WindowsFeatureUpdates information from Graph API (this might take a while)..."

            # We need assignments info so -Expand assignment option is needed here
            $url = "https://graph.microsoft.com/beta/deviceManagement/windowsFeatureUpdateProfiles?`$expand=assignments"
            $WindowsFeatureUpdatesPolicies = Invoke-MSGraphGetRequestWithMSGraphAllPages $url

            if ($WindowsFeatureUpdatesPolicies -eq -1) {
                Write-Error "Error downloading Intune WindowsFeatureUpdatesPolicies information"
                Write-Output "Script will exit..."
                Pause
                Exit 1
            }

            Write-Output "Intune WindowsFeatureUpdatesPolicies information downloaded"

            # Save to local cache
            # Really important parameter!!! Specify -Depth 4 because assignment data will be nested down 4 levels
            $WindowsFeatureUpdatesPolicies | ConvertTo-Json -Depth 4 | Out-File "$PSScriptRoot\cache\WindowsFeatureUpdatesPolicies.json" -Force

        }
        else {
            Write-Output "Using cached Intune WindowsFeatureUpdates information from file $PSScriptRoot\cache\WindowsFeatureUpdatesPolicies.json"
        }

        # Get WindowsFeatureUpdates information from cached file always
        $WindowsFeatureUpdatesPolicies = Get-Content "$PSScriptRoot\cache\WindowsFeatureUpdatesPolicies.json" | ConvertFrom-Json

        # Add odatatype manually (because we use this information later)
        #$odatatype = "#microsoft.graph.windowsFeatureUpdateProfile"
        $WindowsFeatureUpdatesPolicies | Foreach-Object { $_ | Add-Member -MemberType NoteProperty -Name '@odata.type' -Value "#microsoft.graph.windowsFeatureUpdateProfile" -Force }

        # Add WindowsFeatureUpdates to $AllConfigurationProfiles variable
        $AllConfigurationProfiles += $WindowsFeatureUpdatesPolicies

    }
    catch {
        Write-Error "$($_.Exception.GetType().FullName)"
        Write-Error "$($_.Exception.Message)"
        Write-Error "Error trying to download Intune WindowsFeatureUpdates information"
        Write-Output "Script will exit..."
        Pause
        Exit 1
    }

    try {
        # Get deviceManagementScripts information from Graph API
        if (-not ($UseOfflineCache)) {
            # Get deviceManagementScripts from Graph API
            Write-Output "Downloading Intune deviceManagementScripts from Graph API (this might take a while)..."

            # We need assignments info so -Expand assignment option is needed here
            $url = "https://graph.microsoft.com/beta/deviceManagement/deviceManagementScripts?`$expand=assignments"
            $deviceManagementScripts = Invoke-MSGraphGetRequestWithMSGraphAllPages $url

            if ($deviceManagementScripts -eq -1) {
                Write-Error "Error downloading Intune deviceManagementScripts"
                Write-Output "Script will exit..."
                Pause
                Exit 1
            }

            Write-Output "Intune deviceManagementScripts downloaded"

            # Save to local cache
            # Really important parameter!!! Specify -Depth 4 because assignment data will be nested down 4 levels
            $deviceManagementScripts | ConvertTo-Json -Depth 4 | Out-File "$PSScriptRoot\cache\deviceManagementScripts.json" -Force

        }
        else {
            Write-Output "Using cached Intune deviceManagementScripts from file $PSScriptRoot\cache\deviceManagementScripts.json"
        }

        # Get deviceManagementScripts information from cached file always
        $deviceManagementScripts = Get-Content "$PSScriptRoot\cache\deviceManagementScripts.json" | ConvertFrom-Json

        # Add odatatype manually (because we use this information later)
        #$odatatype = "#microsoft.graph.deviceManagementScript"
        $deviceManagementScripts | Foreach-Object { $_ | Add-Member -MemberType NoteProperty -Name '@odata.type' -Value "#microsoft.graph.deviceManagementScript" -Force }

        # Add deviceManagementScripts to $AllConfigurationProfiles variable
        $AllConfigurationProfiles += $deviceManagementScripts

    }
    catch {
        Write-Error "$($_.Exception.GetType().FullName)"
        Write-Error "$($_.Exception.Message)"
        Write-Error "Error trying to download Intune deviceManagementScripts"
        Write-Output "Script will exit..."
        Pause
        Exit 1
    }

    try {
        # Get deviceEnrollmentConfigurations information from Graph API
        if (-not ($UseOfflineCache)) {
            # Get deviceEnrollmentConfigurations from Graph API
            Write-Output "Downloading Intune deviceEnrollmentConfigurations from Graph API (this might take a while)..."

            # We need assignments info so -Expand assignment option is needed here
            $url = "https://graph.microsoft.com/beta/deviceManagement/deviceEnrollmentConfigurations?`$expand=assignments"
            $deviceEnrollmentConfigurations = Invoke-MSGraphGetRequestWithMSGraphAllPages $url

            if ($deviceEnrollmentConfigurations -eq -1) {
                Write-Error "Error downloading Intune deviceEnrollmentConfigurations"
                Write-Output "Script will exit..."
                Pause
                Exit 1
            }

            Write-Output "Intune deviceEnrollmentConfigurations downloaded"

            # Save to local cache
            # Really important parameter!!! Specify -Depth 4 because assignment data will be nested down 4 levels
            $deviceEnrollmentConfigurations | ConvertTo-Json -Depth 4 | Out-File "$PSScriptRoot\cache\deviceEnrollmentConfigurations.json" -Force

        }
        else {
            Write-Output "Using cached Intune deviceEnrollmentConfigurations from file $PSScriptRoot\cache\deviceEnrollmentConfigurations.json"
        }

        # Get deviceEnrollmentConfigurations information from cached file always
        $deviceEnrollmentConfigurations = Get-Content "$PSScriptRoot\cache\deviceEnrollmentConfigurations.json" | ConvertFrom-Json

        # Add deviceEnrollmentConfigurations to $AllConfigurationProfiles variable
        $AllConfigurationProfiles += $deviceEnrollmentConfigurations

    }
    catch {
        Write-Error "$($_.Exception.GetType().FullName)"
        Write-Error "$($_.Exception.Message)"
        Write-Error "Error trying to download Intune deviceEnrollmentConfigurations information"
        Write-Output "Script will exit..."
        Pause
        Exit 1
    }

    try {
        # Get windowsAutopilotDeploymentProfiles information from Graph API
        if (-not ($UseOfflineCache)) {
            # Get windowsAutopilotDeploymentProfiles from Graph API
            Write-Output "Downloading Intune windowsAutopilotDeploymentProfiles from Graph API (this might take a while)..."

            # We need assignments info so -Expand assignment option is needed here
            $url = "https://graph.microsoft.com/beta/deviceManagement/windowsAutopilotDeploymentProfiles?`$expand=assignments"
            $windowsAutopilotDeploymentProfiles = Invoke-MSGraphGetRequestWithMSGraphAllPages $url

            if ($windowsAutopilotDeploymentProfiles -eq -1) {
                Write-Error "Error downloading Intune windowsAutopilotDeploymentProfiles"
                Write-Output "Script will exit..."
                Pause
                Exit 1
            }

            Write-Output "Intune windowsAutopilotDeploymentProfiles downloaded"

            # Save to local cache
            # Really important parameter!!! Specify -Depth 4 because assignment data will be nested down 4 levels
            $windowsAutopilotDeploymentProfiles | ConvertTo-Json -Depth 4 | Out-File "$PSScriptRoot\cache\windowsAutopilotDeploymentProfiles.json" -Force

        }
        else {
            Write-Output "Using cached Intune windowsAutopilotDeploymentProfiles from file $PSScriptRoot\cache\windowsAutopilotDeploymentProfiles.json"
        }

        # Get windowsAutopilotDeploymentProfiles information from cached file always
        $windowsAutopilotDeploymentProfiles = Get-Content "$PSScriptRoot\cache\windowsAutopilotDeploymentProfiles.json" | ConvertFrom-Json

        # Add windowsAutopilotDeploymentProfiles to $AllConfigurationProfiles variable
        $AllConfigurationProfiles += $windowsAutopilotDeploymentProfiles

    }
    catch {
        Write-Error "$($_.Exception.GetType().FullName)"
        Write-Error "$($_.Exception.Message)"
        Write-Error "Error trying to download Intune windowsAutopilotDeploymentProfiles information"
        Write-Output "Script will exit..."
        Pause
        Exit 1
    }

    # Debug export variable to json for debugging
    #$AllConfigurationProfiles | ConvertTo-Json -Depth 4 | Out-File "$PSScriptRoot\cache\AllConfigurationProfiles_DEBUG.json" -Force

    #####

    # Find configurations which have assignments
    # or convert $AllConfigurationProfiles to json to get more human readable format: $AllConfigurationProfiles | ConvertTo-JSON
    # $ConfigurationsWithAssignments = $AllConfigurationProfiles | Where-Object { $_.assignments.target.groupid -like "*" }

    # Create custom object array and gather necessary ConfigurationProfile and assignment information.
    $ConfigurationProfilesWithAssignmentInformation = @()

    try {
        Write-Output "Creating Device Configuration custom object array"

        # Go through each DeviceConfiguration and save necessary information to custom object
        Foreach ($ConfigurationProfile in $AllConfigurationProfiles) {
            Foreach ($Assignment in $ConfigurationProfile.Assignments) {
            
                $assignmentId = $Assignment.id
                #$assignmentIntent = $Assignment.intent
                $assignmentTargetGroupId = $Assignment.target.groupid
                $assignmentTargetGroupDisplayName = $AllGroups | Where { $_.id -eq $assignmentTargetGroupId } | Select -ExpandProperty displayName
                
                # Special case for All Users
                if ($Assignment.target.'@odata.type' -eq '#microsoft.graph.allLicensedUsersAssignmentTarget') {
                    $assignmentTargetGroupDisplayName = 'All Users'
                }

                # Special case for All Devices
                if ($Assignment.target.'@odata.type' -eq '#microsoft.graph.allDevicesAssignmentTarget') {
                    $assignmentTargetGroupDisplayName = 'All Devices'
                }

                # Set included/excluded attribute
                $DeviceConfigurationExclude = ''
                if ($Assignment.target.'@odata.type' -eq '#microsoft.graph.groupAssignmentTarget') {
                    $DeviceConfigurationExclude = 'Included'
                }
                if ($Assignment.target.'@odata.type' -eq '#microsoft.graph.exclusionGroupAssignmentTarget') {
                    $DeviceConfigurationExclude = 'Excluded'
                }

                # Remove #microsoft.graph. from @odata.type
                $odatatype = $ConfigurationProfile.'@odata.type'.Replace('#microsoft.graph.', '')

                $properties = @{
                    '@odata.type'                    = $odatatype
                    displayname                      = $ConfigurationProfile.displayName
                    createdDateTime                  = $ConfigurationProfile.createdDateTime
                    lastModifiedDateTime             = $ConfigurationProfile.lastModifiedDateTime
                    id                               = $ConfigurationProfile.id
                    assignmentId                     = $assignmentId
                    assignmentIntent                 = "$DeviceConfigurationExclude"
                    assignmentTargetGroupId          = $assignmentTargetGroupId
                    assignmentTargetGroupDisplayName = $assignmentTargetGroupDisplayName
                    IncludeExclude                   = $DeviceConfigurationExclude
                    c                                = "D0E4F5"
                }

                # Create new custom object every time inside foreach-loop
                # This is really important step to do inside foreach-loop!
                # If you create custom object outside of foreach then you would edit same custom object on every foreach cycle resulting only 1 ConfigurationProfile in custom object array
                $CustomObject = New-Object -TypeName PSObject -Prop $properties

                # Add custom object to our custom object array.
                $ConfigurationProfilesWithAssignmentInformation += $CustomObject
            }
        }
    }
    catch {
        Write-Error "$($_.Exception.GetType().FullName)"
        Write-Error "$($_.Exception.Message)"
        Write-Error "Error creating Device Configuration Profiles custom object"
        Write-Output "Script will exit..."
        Pause
        Exit 1
    }

    ########################################################################################################################

    # Create HTML report

    $head = @'
    <style>
    body {
        background-color: #FFFFFF;
    }
    table#TopTable {
        border: 2px solid #1C6EA4;
        background-color: #f7f7f4;
        #width: 100%;
        text-align: left;
        border-collapse: separate;
    }
    table {
    border: 2px solid #1C6EA4;
    background-color: #EEEEEE;
    #width: 100%;
    text-align: left;
    border-collapse: collapse;
    #table-layout: fixed;
    }
    table td {
    border: 2px solid #AAAAAA;
    padding: 5px;
    max-width: 100;
    white-space:nowrap;
    overflow: hidden;
    text-overflow: ellipsis;
    #word-wrap:break-word;
    }
    table th {
    border: 2px solid #AAAAAA;
    padding: 5px;
    }
    table#TopTable td, table#TopTable th {
        vertical-align: top;
        text-align: center;
        bgcolor: f7f7f4;
        border: 2px solid #AAAAAA;
        padding: 5px;
        border-collapse: border;
    }
    table tbody td {
    font-size: 15px;
    }
    table thead {
    background: #1C6EA4;
    background: -moz-linear-gradient(top, #5592bb 0%, #327cad 66%, #1C6EA4 100%);
    background: -webkit-linear-gradient(top, #5592bb 0%, #327cad 66%, #1C6EA4 100%);
    background: linear-gradient(to bottom, #5592bb 0%, #327cad 66%, #1C6EA4 100%);
    border-bottom: 2px solid #444444;
    }
    table thead th {
    font-size: 18px;
    font-weight: bold;
    color: #FFFFFF;
    border-left: 2px solid #D0E4F5;
    }
    table#TopTable thead th {
        font-size: 18px;
        font-weight: bold;
        color: #FFFFFF;
        border-left: 2px solid #D0E4F5;
    }
    table thead th:first-child {
    border-left: none;
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
    table tfoot td {
    font-size: 16px;
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
    .author .img-top {
        display: none;
		position: relative;
        top: 0;
        left: 0;
        z-index: 99;
    }
    .author:hover .img-top {
        display: inline;
		position: relative;
        top: 0;
        left: 0;
        z-index: 99;
    }
	.author:hover .img-back {
        display: none;
		position: relative;
        top: 0;
        left: 0;
        z-index: 99;
    }
    .author .img-back {
        display: inline;
		position: relative;
        top: 0;
        left: 0;
        z-index: 99;
    }
    </style>
'@

    ############################################################

    # Configurations Summary
    # Create Configurations summary object for ConfigurationProfiles which have assignments
    $AllConfigurationProfilesSummary = $AllConfigurationProfiles | Where-object { $_.Assignments } | Select -Property '@odata.type' | Sort-Object -Property '@odata.type' | Group-Object '@odata.type' | Select-Object name, count

    # Remove #microsoft.graph.
    $AllConfigurationProfilesSummary | ForEach-Object { $_.Name = ($_.Name).Replace('#microsoft.graph.', '') }

    $AllConfigurationProfilesSummaryHTML = $AllConfigurationProfilesSummary | ConvertTo-Html -Fragment -PreContent "<h2 id=`"DeviceConfigurationAssignmentSummary`">Device Configuration Profiles Assignments Summary</h2>" | Out-String

    ######################
    Write-Output "Create OS specific Device Configuration Profiles Assignment information HTML fragments."

<#
    @odata.type
    ---------- -
    #microsoft.graph.androidDeviceOwnerGeneralDeviceConfiguration
    #microsoft.graph.androidDeviceOwnerWiFiConfiguration
    #microsoft.graph.androidGeneralDeviceConfiguration
    #microsoft.graph.androidWorkProfileGeneralDeviceConfiguration

    #microsoft.graph.iosDeviceFeaturesConfiguration
    #microsoft.graph.iosGeneralDeviceConfiguration
    #microsoft.graph.iosWiFiConfiguration

    #microsoft.graph.macOSDeviceFeaturesConfiguration
    #microsoft.graph.macOSGeneralDeviceConfiguration

    #microsoft.graph.sharedPCConfiguration
    #microsoft.graph.windows10CustomConfiguration
    #microsoft.graph.windows10EasEmailProfileConfiguration
    #microsoft.graph.windows10EndpointProtectionConfiguration
    #microsoft.graph.windows10EnterpriseModernAppManagementConfiguration
    #microsoft.graph.windows10GeneralConfiguration
    #microsoft.graph.windows10SecureAssessmentConfiguration
    #microsoft.graph.windows81TrustedRootCertificate
    #microsoft.graph.windows81WifiImportConfiguration
    #microsoft.graph.windowsDeliveryOptimizationConfiguration
    #microsoft.graph.windowsUpdateForBusinessConfiguration
    #microsoft.graph.windowsWifiConfiguration

    #microsoft.graph.groupPolicyConfiguration
#>

    try {

        # Windows policies
        $WindowsConfigurationsByDisplayName = $ConfigurationProfilesWithAssignmentInformation | Where-object { ($_.'@odata.type' -eq 'sharedPCConfiguration') -or ($_.'@odata.type' -like 'windows*') -or ($_.'@odata.type' -eq 'groupPolicyConfiguration') -or ($_.'@odata.type' -eq 'deviceManagementScript') -or ($_.'@odata.type' -eq 'windows10EnrollmentCompletionPageConfiguration') -or ($_.'@odata.type' -eq 'azureADWindowsAutopilotDeploymentProfile') } | Select c, '@odata.type', displayName, assignmentIntent, assignmentTargetGroupDisplayName, createdDateTime, lastModifiedDateTime, id | Sort-Object displayName, id, assignmentIntent

        # Create grouping colors by id attribute
        $WindowsConfigurationsByDisplayName = Create-GroupingRowColors $WindowsConfigurationsByDisplayName 'id'

        # Create grouping colors by assignmentTargetGroupDisplayName attribute
        #$WindowsConfigurationsByDisplayName = Create-GroupingRowColors $WindowsConfigurationsByDisplayName 'displayName'

        $PreContent = "<h2 id=`"WindowsConfigurationsSortedByConfigurationdisplayName`">Windows configuration assignments by profile</h2>"
        $WindowsConfigurationsByDisplayNameHTML = $WindowsConfigurationsByDisplayName | ConvertTo-Html -Fragment -PreContent $PreContent | Out-String

        # Fix &lt; &quot; etc...
        $WindowsConfigurationsByDisplayNameHTML = Fix-HTMLSyntax $WindowsConfigurationsByDisplayNameHTML
        
        # Change HTML Table TD values with regexp
        # We bold DisplayName, set Intent TD backgroundcolor and set grouped row coloring
        $WindowsConfigurationsByDisplayNameHTML = Change-HTMLTableSyntaxWithRegexpForConfigurationsSortedByDisplayName $WindowsConfigurationsByDisplayNameHTML

        # Fix column names and remove c column
        $WindowsConfigurationsByDisplayNameHTML = Fix-HTMLColumns $WindowsConfigurationsByDisplayNameHTML

        ######
        
        # Android policies
        $AndroidConfigurationsByDisplayName = $ConfigurationProfilesWithAssignmentInformation | Where-object { ($_.'@odata.type' -like 'android*') } | Select c, '@odata.type', displayName, assignmentIntent, assignmentTargetGroupDisplayName, createdDateTime, lastModifiedDateTime, id | Sort-Object displayName, id, assignmentIntent

        # Create grouping colors by id attribute
        $AndroidConfigurationsByDisplayName = Create-GroupingRowColors $AndroidConfigurationsByDisplayName 'id'

        # Create grouping colors by assignmentTargetGroupDisplayName attribute
        #$AndroidConfigurationsByDisplayName = Create-GroupingRowColors $AndroidConfigurationsByDisplayName 'displayName'

        $PreContent = "<h2 id=`"AndroidConfigurationsSortedByConfigurationdisplayName`">Android configuration assignments by profile</h2>"
        $AndroidConfigurationsByDisplayNameHTML = $AndroidConfigurationsByDisplayName | ConvertTo-Html -Fragment -PreContent $PreContent | Out-String

        # Fix &lt; &quot; etc...
        $AndroidConfigurationsByDisplayNameHTML = Fix-HTMLSyntax $AndroidConfigurationsByDisplayNameHTML

        # Change HTML Table TD values with regexp
        # We bold DisplayName, set Intent TD backgroundcolor and set grouped row coloring
        $AndroidConfigurationsByDisplayNameHTML = Change-HTMLTableSyntaxWithRegexpForConfigurationsSortedByDisplayName $AndroidConfigurationsByDisplayNameHTML

        # Fix column names and remove c column
        $AndroidConfigurationsByDisplayNameHTML = Fix-HTMLColumns $AndroidConfigurationsByDisplayNameHTML

        ######
        
        # iOS policies
        $iOSConfigurationsByDisplayName = $ConfigurationProfilesWithAssignmentInformation | Where-object { ($_.'@odata.type' -like 'ios*') } | Select c, '@odata.type', displayName, assignmentIntent, assignmentTargetGroupDisplayName, createdDateTime, lastModifiedDateTime, id | Sort-Object displayName, id, assignmentIntent

        # Create grouping colors by id attribute
        $iOSConfigurationsByDisplayName = Create-GroupingRowColors $iOSConfigurationsByDisplayName 'id'

        # Create grouping colors by assignmentTargetGroupDisplayName attribute
        #$iOSConfigurationsByDisplayName = Create-GroupingRowColors $iOSConfigurationsByDisplayName 'displayName'

        $PreContent = "<h2 id=`"iOSConfigurationsSortedByConfigurationdisplayName`">iOS configuration assignments by profile</h2>"
        $iOSConfigurationsByDisplayNameHTML = $iOSConfigurationsByDisplayName | ConvertTo-Html -Fragment -PreContent $PreContent | Out-String

        # Fix &lt; &quot; etc...
        $iOSConfigurationsByDisplayNameHTML = Fix-HTMLSyntax $iOSConfigurationsByDisplayNameHTML

        # Change HTML Table TD values with regexp
        # We bold DisplayName, set Intent TD backgroundcolor and set grouped row coloring
        $iOSConfigurationsByDisplayNameHTML = Change-HTMLTableSyntaxWithRegexpForConfigurationsSortedByDisplayName $iOSConfigurationsByDisplayNameHTML

        # Fix column names and remove c column
        $iOSConfigurationsByDisplayNameHTML = Fix-HTMLColumns $iOSConfigurationsByDisplayNameHTML

        ######
        
        # macOS policies
        $macOSConfigurationsByDisplayName = $ConfigurationProfilesWithAssignmentInformation | Where-object { ($_.'@odata.type' -like 'macos*') } | Select c, '@odata.type', displayName, assignmentIntent, assignmentTargetGroupDisplayName, createdDateTime, lastModifiedDateTime, id | Sort-Object displayName, id, assignmentIntent

        # Create grouping colors by id attribute
        $macOSConfigurationsByDisplayName = Create-GroupingRowColors $macOSConfigurationsByDisplayName 'id'

        # Create grouping colors by assignmentTargetGroupDisplayName attribute
        #$macOSConfigurationsByDisplayName = Create-GroupingRowColors $macOSConfigurationsByDisplayName 'displayName'

        $PreContent = "<h2 id=`"macOSConfigurationsSortedByConfigurationdisplayName`">macOS configuration assignments by profile</h2>"
        $macOSConfigurationsByDisplayNameHTML = $macOSConfigurationsByDisplayName | ConvertTo-Html -Fragment -PreContent $PreContent | Out-String

        # Fix &lt; &quot; etc...
        $macOSConfigurationsByDisplayNameHTML = Fix-HTMLSyntax $macOSConfigurationsByDisplayNameHTML

        # Change HTML Table TD values with regexp
        # We bold DisplayName, set Intent TD backgroundcolor and set grouped row coloring
        $macOSConfigurationsByDisplayNameHTML = Change-HTMLTableSyntaxWithRegexpForConfigurationsSortedByDisplayName $macOSConfigurationsByDisplayNameHTML

        # Fix column names and remove c column
        $macOSConfigurationsByDisplayNameHTML = Fix-HTMLColumns $macOSConfigurationsByDisplayNameHTML

        ######################
        # Other configuration profile types
        # If there are @odata.types that weren't specified earlier
        # This makes sure to catch all configuration types which might be released in the future

        # Create object with specified attributes and sorting
        $OtherConfigurationsByDisplayName = $ConfigurationProfilesWithAssignmentInformation | Where-object {`
            ($_.'@odata.type' -notlike 'windows*') -and `
            ($_.'@odata.type' -ne 'groupPolicyConfiguration') -and `
            ($_.'@odata.type' -ne 'deviceManagementScript') -and `
            ($_.'@odata.type' -ne 'windows10EnrollmentCompletionPageConfiguration') -and `
            ($_.'@odata.type' -ne 'deviceEnrollmentLimitConfiguration') -and `
            ($_.'@odata.type' -ne 'deviceEnrollmentPlatformRestrictionsConfiguration') -and `
            ($_.'@odata.type' -ne 'deviceEnrollmentWindowsHelloForBusinessConfiguration') -and `
            ($_.'@odata.type' -ne 'azureADWindowsAutopilotDeploymentProfile') -and `
            ($_.'@odata.type' -notlike 'android*') -and `
            ($_.'@odata.type' -notlike 'ios*') -and `
            ($_.'@odata.type' -notlike 'macos*') -and `
            ($_.'@odata.type' -ne 'sharedPCConfiguration') }`
        | Select c, '@odata.type', displayName, assignmentIntent, assignmentTargetGroupDisplayName, createdDateTime, lastModifiedDateTime, id | Sort-Object displayName, id, assignmentIntent
            
        # Create grouping colors by id attribute
        $OtherConfigurationsByDisplayName = Create-GroupingRowColors $OtherConfigurationsByDisplayName 'id'

        $PreContent = "<h2 id=`"OtherConfigurationsSortedByConfigurationdisplayName`">Other configuration assignments by profile</h2>"
        $OtherConfigurationsByDisplayNameHTML = $OtherConfigurationsByDisplayName | ConvertTo-Html -Fragment -PreContent $PreContent | Out-String

        # Fix &lt; &quot; etc...
        $OtherConfigurationsByDisplayNameHTML = Fix-HTMLSyntax $OtherConfigurationsByDisplayNameHTML

        # Change HTML Table TD values with regexp
        # We bold DisplayName, set Intent TD backgroundcolor and set grouped row coloring
        $OtherConfigurationsByDisplayNameHTML = Change-HTMLTableSyntaxWithRegexpForConfigurationsSortedByDisplayName $OtherConfigurationsByDisplayNameHTML

        # Fix column names and remove c column
        $OtherConfigurationsByDisplayNameHTML = Fix-HTMLColumns $OtherConfigurationsByDisplayNameHTML

        ######################
        # All Configurations sorted by assignmentTargetGroupDisplayName

        # All Configurations
        #$htmlObjectAllDeviceConfigurationsSortedByAssignmentTargetGroupDisplayName = $ConfigurationProfilesWithAssignmentInformation | Select c, '@odata.type', displayName, assignmentIntent, assignmentTargetGroupDisplayName, createdDateTime, lastModifiedDateTime, id | Sort-Object assignmentTargetGroupDisplayName, assignmentIntent, displayName, id

        # Exclude few deviceEnrollment policies
        # 'deviceEnrollmentLimitConfiguration'
        # 'deviceEnrollmentPlatformRestrictionsConfiguration'
        # 'deviceEnrollmentWindowsHelloForBusinessConfiguration'
        #
        # If we want to exclude default ESP targeting
        #-and (-not (($_.'@odata.type' -eq 'windows10EnrollmentCompletionPageConfiguration') -and ($_.displayName -eq 'All users and all devices')))
        $htmlObjectAllDeviceConfigurationsSortedByAssignmentTargetGroupDisplayName = $ConfigurationProfilesWithAssignmentInformation | Where-Object { (($_.'@odata.type' -ne 'deviceEnrollmentLimitConfiguration') -and ($_.'@odata.type' -ne 'deviceEnrollmentPlatformRestrictionsConfiguration') -and ($_.'@odata.type' -ne 'deviceEnrollmentWindowsHelloForBusinessConfiguration')) } | Select c, '@odata.type', displayName, assignmentIntent, assignmentTargetGroupDisplayName, createdDateTime, lastModifiedDateTime, id | Sort-Object assignmentTargetGroupDisplayName, assignmentIntent, displayName, id

        # Create grouping colors by assignmentTargetGroupDisplayName attribute
        $htmlObjectAllDeviceConfigurationsSortedByAssignmentTargetGroupDisplayName = Create-GroupingRowColors $htmlObjectAllDeviceConfigurationsSortedByAssignmentTargetGroupDisplayName 'assignmentTargetGroupDisplayName'

        # Working
        $htmlAllDeviceConfigurationsSortedByAssignmentTargetGroupDisplayName = $htmlObjectAllDeviceConfigurationsSortedByAssignmentTargetGroupDisplayName | ConvertTo-Html -Fragment -PreContent "<br><hr><h2 id=`"AllDeviceConfigurationsSortedByAssignmentTargetGroupDisplayName`">Profile assignments by target group</h2>" | Out-String

        # Fix html syntax
        $htmlAllDeviceConfigurationsSortedByAssignmentTargetGroupDisplayName = Fix-HTMLSyntax $htmlAllDeviceConfigurationsSortedByAssignmentTargetGroupDisplayName

        # Change HTML Table TD values with regexp
        # We bold AssignmentTargetGroupDisplayName, set Intent TD backgroundcolor and set grouped row coloring
        $htmlAllDeviceConfigurationsSortedByAssignmentTargetGroupDisplayName = Change-HTMLTableSyntaxWithRegexpForConfigurationProfilesSortedByAssignmentTargetGroupDisplayName $htmlAllDeviceConfigurationsSortedByAssignmentTargetGroupDisplayName

        # Fix column names and remove c column
        $htmlAllDeviceConfigurationsSortedByAssignmentTargetGroupDisplayName = Fix-HTMLColumns $htmlAllDeviceConfigurationsSortedByAssignmentTargetGroupDisplayName

    }
    catch {
        Write-Error "$($_.Exception.GetType().FullName)"
        Write-Error "$($_.Exception.Message)"
        Write-Error "Error creating OS specific HTML fragment information"
        Write-Output "Script will exit..."
        Pause
        Exit 1        
    }
    #############################
    # Create html
    Write-Output "Creating HTML report..."

    try {

        $ReportRunDateTime = (Get-Date).ToString("yyyyMMddHHmm")
        $ReportRunDateTimeHumanReadable = (Get-Date).ToString("yyyy-MM-dd HH:mm")
        $ReportRunDateFileName = (Get-Date).ToString("yyyyMMddHHmm")

        $ReportSavePath = $PSScriptRoot
        $HTMLFileName = "$($ReportRunDateFileName)_Intune_DeviceConfiguration_Assignments_report.html"

        #$PreContent = "<table id=`"TopTable`"><tr align=`"center`"><td valign=`"top`" bgcolor=`"f7f7f4`" border=`"5`">`
        $PreContent = "<table id=`"TopTable`"><tr><td>`
        <h1>Intune Device<br>`
        Configurations Report</h1>`
        <p align=`"left`">`
        &nbsp;&nbsp;&nbsp;<strong>Report run:</strong> $ReportRunDateTimeHumanReadable<br>`
        &nbsp;&nbsp;&nbsp;<strong>By:</strong> $($ConnectMSGraph.UPN)<br>`
        &nbsp;&nbsp;&nbsp;<strong>Tenant name:</strong> $TenantDisplayName<br>`
        &nbsp;&nbsp;&nbsp;<strong>Tenant id:</strong> $($ConnectMSGraph.TenantId)`
        </p>`
        <h3>Configurations sorted by profile name</h3>`
        <string>`
        <a href=`"#WindowsConfigurationsSortedByConfigurationdisplayName`">Windows Configurations</a><br>`
        <a href=`"#AndroidConfigurationsSortedByConfigurationdisplayName`">Android Configurations</a><br>`
        <a href=`"#iOSConfigurationsSortedByConfigurationdisplayName`">iOS Configurations</a><br>`
        <a href=`"#macOSConfigurationsSortedByConfigurationdisplayName`">macOS Configurations</a><br>`
        <a href=`"#OtherConfigurationsSortedByConfigurationdisplayName`">Other Configurations</a><br>`
        </string>`
        <br><h3><a href=`"#AllDeviceConfigurationsSortedByAssignmentTargetGroupDisplayName`">All configurations sorted by target group name</a></h3></td>`
        <!-- <td bgcolor=`"f7f7f4`" valign=`"top`" align=`"center`" border=`"5`">$AllConfigurationProfilesSummaryHTML</td> -->`
        <!-- <td bgcolor=`"f7f7f4`" valign=`"top`" border=`"5`"> -->`
        <td>$AllConfigurationProfilesSummaryHTML</td>`
        <td>`
        <h2>Author:</h2>`
        <p><strong>`
        Get more Intune reports and tools from<br>`
        <a href=`"https://github.com/petripaavola/Intune`" target=`"_blank`">https://github.com/petripaavola/Intune</a>`
        <br><br><br><br><br><br>`
        <div class=`"author`">`
            <img src=`"data:image/png;base64,/9j/4AAQSkZJRgABAQEAeAB4AAD/4QBoRXhpZgAATU0AKgAAAAgABAEaAAUAAAABAAAAPgEbAAUAAAABAAAARgEoAAMAAAABAAIAAAExAAIAAAARAAAATgAAAAAAAAB4AAAAAQAAAHgAAAABcGFpbnQubmV0IDQuMC4yMQAA/9sAQwACAQECAQECAgICAgICAgMFAwMDAwMGBAQDBQcGBwcHBgcHCAkLCQgICggHBwoNCgoLDAwMDAcJDg8NDA4LDAwM/9sAQwECAgIDAwMGAwMGDAgHCAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwM/8AAEQgAZABkAwEiAAIRAQMRAf/EAB8AAAEFAQEBAQEBAAAAAAAAAAABAgMEBQYHCAkKC//EALUQAAIBAwMCBAMFBQQEAAABfQECAwAEEQUSITFBBhNRYQcicRQygZGhCCNCscEVUtHwJDNicoIJChYXGBkaJSYnKCkqNDU2Nzg5OkNERUZHSElKU1RVVldYWVpjZGVmZ2hpanN0dXZ3eHl6g4SFhoeIiYqSk5SVlpeYmZqio6Slpqeoqaqys7S1tre4ubrCw8TFxsfIycrS09TV1tfY2drh4uPk5ebn6Onq8fLz9PX29/j5+v/EAB8BAAMBAQEBAQEBAQEAAAAAAAABAgMEBQYHCAkKC//EALURAAIBAgQEAwQHBQQEAAECdwABAgMRBAUhMQYSQVEHYXETIjKBCBRCkaGxwQkjM1LwFWJy0QoWJDThJfEXGBkaJicoKSo1Njc4OTpDREVGR0hJSlNUVVZXWFlaY2RlZmdoaWpzdHV2d3h5eoKDhIWGh4iJipKTlJWWl5iZmqKjpKWmp6ipqrKztLW2t7i5usLDxMXGx8jJytLT1NXW19jZ2uLj5OXm5+jp6vLz9PX29/j5+v/aAAwDAQACEQMRAD8A/fyiiigAoor4X/4LSf8ABXaz/wCCdPw/g8N+HYf7Q+Jniqykm0/eoa30aDJQXcoP323BhGnRijFsAYbOpUjCPNI0pUpVJckTp/8AgpX/AMFn/hb/AME5dNm02+uF8U+PpI98Hh6znCmDIJV7mXDCFTgcAM5yCFwdw/ED9q//AILvftY/theJbiy8NeOj4E8PzsVj07wpCdOlQHjm4y1wxx3EoGf4V6VJ8H/2DNU/annm8f8Aj3xNqF7qXiOZr6R5GMs8zOxZmdm7knJ69a+xfgf+xX4L+F2mW/2HRLWe4gwRcTRB5CR3ya+Rx3EkINqOr7dP+CfdZbwjOcVOrZLv1/4B+Uk/xh/aA0u+XVf+FxfE1dQUiQXP/CS3qyB84+/5mcjOa+uP2Cf+Di39pT9lvxLZaf8AETUP+Ft+B45FF2mtSY1OGM9WivQN7N3xL5gOCBtzkfb/AIy/Z48J+PdFWy1rw/p1xCqlVxCFZB7EYI/CvFvib/wTK+F/ivQprO30abTGkBCzwzuzxk45wxK9vSuShxNH7at6HoYng+Ml+6f3/wBM/Yv9jr9tn4d/t1/CiDxd8PNcj1OzOEu7STCXmmSn/llPHklW4OCMq2MqSOa9Zr+a34MWHxM/4I0/H2w8feCb6617wmZRb6rZglY7y3JGYp16bTnKtyVbBHI5/ou+DvxV0n45fCrw74x0GZptH8TafDqVozDDCOVAwDDswzgjsQRX1uX4+GJheLufB5lltXCT5aisdJRRRXoHmhRRRQAUUUUAFfzz/wDBQyyX9sX/AIKG/FDxJqkiz6PpOrHR9OVXLIYbMC2Rl9mKO+OmXJr97vjn8Qf+FTfBXxd4owrN4d0a71JVbo7QwvIB+JUD8a/AP4KeDbzxLof9o3l03m3szyTtIcl3J5J9+/vXzfEWL9lTjFOx9Vwrg/bVnJrbQ+ivgV4Sh0XwTptqo2pbxAJ8uMj6V654djVYVDfd6cDFeSeFviHpHhlbWO+vI4VUBFDN1AH+etej6R8Y/Cklv+71zTGfH+rS5Uvz/s5zX5uuaT5mfq/I4pRR2EtlDIuNsnzDHrXP69aqqNiNl25B4zWrp+u291Z+ZHcRzDg7g3tkVS1O+tbuDakiu38RQ5waJbChzJnlPj/w5B4l0+6sbqFJ7W6iaOVGHDKRivpD/g3y+K99P+z34s+FurP/AKV8N9akOnqT/wAw+6Z5EA+ky3B9g6j0rwbxdcxwxSbZBlT0zzXpv/BGS3utC/a1+IVuVC2er6At5u7u8dzGo/ISn86+j4YxMoYpU+jPleMMKqmEdXsfpdRRRX6SflIUUUUAFFFFAHyL/wAFjvFXiTS/2eNJ0fw/cQ29r4g1T7Nq/mOyLc2YicyQEqQcPn6fLzkZB/KvTvDniH4e+B20/wAPw2d5cLvkt4r6SUpGW5CO67m9BuAOPSv2K/4KdaPHd/soanqMlv8AaG0O8gulGMlN7G3J/ATHPtX5c+CdUD6+zDDY4bPrX53xVKcMVrqmk0mfrHBtOnWwMUlZxck2t23Z/kfPd7qfiDTb7SpDos19farbpPMEkC21uzKCy5ILnByOo6ZwKr/DyDxR4nvRqEnhu3sJRcJALO5gZTKpBJdZCNwC45JBHI4xyPqab4Qyaxqk11pM1uyySNKbadWCI7Es21lIK7mJJzuGTwB3k1rwnrGk6PM0GjaXaXEaEfa5757hIAerBPLBbHXGVz614ixXuNcq16n10cHOM4vmenTTU5PwL+1z4as/hJ4kvZo9ehk8LgxahJDplxcojrkna8aMGXAzkcAdcHIHhXxC/aWub/R4des5vElrp94sbqonNuZFkDOjAc4JVScHHvXvPwZ+GlvZfCXUvDmjw3FxpMkcsUjFQPND53MQAANxZjhQBknFch8Ivg+3hb4ZWvhifR9Uvl0cNZwXlp5brNErHCyK7gq46HAKnqCM7VuHsIe8k3r36f16mlSji+WzktVrZXs9LadVb0287Hk3hj4uatr99aw2V5qjalsjuo7e8uDJKyOMqR8oQgg8gsO/cHH6Df8ABGz9oXSdV/aWt9PSOa8v9cs7/Q2eMFRZT2wSecOuOgMITdnbuYAE5r5atvhVeaJr0dzb6HeWsyKUS4vhGqRqcA4CFiexxx9RX2r/AMEaPhLZWfxw1zVFdpH0DSDFHv8AmcyTuqs+cekbZ9S31r0Mrkp46n7NW1/Dr+B83xDSdPAVHWd1Zra2vS3z337H6TUUUV+mH4yFFFFABRRRQB5n+2T8N5/iz+zB400O2vJ7G4uNNeaOSJQzO0WJRGR6OU2n2Y1+M/hKRxqEzREYYKwPpxX7xSxLPE0ciq6OCrKwyGB6g1+F2o6ZbfDT9oPxh4PklVpPC+s3emIwOd8cczojY91Cn8a+L4uw+kKy80/zX6n6FwLjOWc6LfZr8n+h6X4H14W1mq8KzDPT/P8Ak1N8Ube58Q+D7y3huFjkmixGCSFY9cEjsemcd6zY7Jb3T/Mt/wDWQnPHcVwfjP4meKLa5WS38Ivexx/Kj/bVCtjjJVQxAP418VRlzOx+sU5uc1yrU8/m8CfE/Q7LWNQ03xClmt4hjtrVbdGWyVV6r3kYk5O44yAAAM59Y/ZbtdW0Twht1Wdbq6Lb3Jxuf5VBY44BJBJA6ZrndS+L/iuz8OtNP4V0uUshCiPUT+6B65j27t35Vn/B/wCKOreJb14/+ET1jTtpI87zEaEt1yMNux/wHFddSDUL6fgdmIpzhBymvxv+p65441iO5gY7Rux6/d6V9Of8EY9Le5134halyIY4rK1X3ZjMx/IKv/fVfJmt2zJpXmXTfvXXcR6V+gP/AASS+Hs3hT9ma41i5hMUnijVZbyEkYZoEVYk/wDHkkI9mFenwzB1Mapfypv8LfqfnfHGJUcA4fzNL9f0PqSiiiv0s/HQooooAKyfHXj3Rfhl4VvNc8QanZ6PpOnoZLi6upRHHGPcnuewHJPAr8k/+CqP/Byt4g/Zw/aI1r4a/Bnw34X1STwpObLWNe17zZ4Zbpf9ZFbRRSIcRtlTI5O5gwCgAM35p/tUf8FdPjN+2jr9vdePtas7zT7MAW2kWMb22m27YwXEQb5nOT80m5ucZA4rjrYtRTUdWddHCylrLRH21/wWM/4LFeM/ijpuq2/w11bWNE8G2cyWEAs5mt5tSLNtMsxUhtpPRDwABkZJr5ZsvFmreAdS8MeINQmuJ5rywga+ndizTybAJGYnkktySeea5P4V+MdB+N2nLpkbxLrEk0Uo0m4QRtcOrBswN92QgqvyYWQnG1WwTX1Fd/AOH4h/C2OyWM+ZDHiIgcjjp/Kvi83xlnGNXW97/wBeR9/w7gU4ynSdmrW/rzHeOv2iItG+Hdvd2d0qz3lzDGBnK8sM59QQO1enfDb4g6Z4q8O2/n3kYuph5SqgC7iOOBn8voelfnP+0P4W8UeB/DN14fuftEfkyiWynyVVtpyFJ9f0rkvgZ+2vqfw+uhZ6z9qTyXLCQ5cqSR1/ID259TXmwyZ1KXPRd3f8D6L+3I0a/s665U1v5n6E6v4Tvj8QDMuoXUdiHLctlsHPXnvjOOwrQ+Jnxd0r4S+FLma3vla8hQjzFP3fX6/zP518f63/AMFFtLkuZLhrmWWSZMBTn5jjAOOxAyK8q0X4q6x+0D4std0lxHp9jN5k0pP+twwYLjuTxz7fSqp5TVlrV0ijbGcQYdLkoPmk+x+qv7Nvga//AGyPjnovhPTWl8h1+0ardRjixtFI8xz23HhVz1Zl7Zr9kvDHhux8G+HLDSdMt47PTtMt0tbaBB8sUaKFVR9ABX87f7G3/BWDx3+w54q8ceDvB+j+ENQ1aJIdYuJNWsJZptSRLTzvsYkjlRkHJCHkB3ZiCMg/tx/wT4/4KG+Af+CivwRs/FXg++hj1SGCH+3NDeTddaHcOpzG+QNyEq2yQDa4U9CGUfXcP4Olh6Vl8Utfl0sfmvFGYVcViNfgjovXq2e9UUUV9EfLhRRRQB/FTEtnNJtVt0n92TIerlvGkJ+72xzmqusadFeDZIqtznB6j3FU4IdQ0RM20jXkI6xTN8wH+y3+OfwrwbX2PdWnQ6OAAYaNthXkYNfe/wDwT2/b+s/EF7D4P+IWoeTq0hEVhq05x9tPQRTk/wDLboBIfv8ARvmwW/PHTfE8csg82GSykXtJgA/Q8g161+yn4n8JeGv2gfCeoeNtHs/EPhRb1YdVsrklYpbeQGNmO0g5QN5gwRkoB0rhxmEhWp8lRf15Ho4HGzoVFUpP/gn6tfGT9nnRvirocnmQ29xHcpvVh8ySAjIYEV8G/tAf8E4rvR9Ukm02HzIc5CMOfpn/ABrf+Lv7SvxC/wCCWP7Uvij4fqreLvBemXn2m0sLy4ZpJtOnAmgkgmIJWURuAwwUZg3G75q+ov2bv23vhT+2bpscGg61DDrTx5l0W/It9QiPU4QnEgH96MsB3I6V87LC4vBfvKWse6/VdD7CjmWDxy9nV0l2e/y7n5pv+ynqml6iFutLuPl7nofxAz0r2b4J/Bv+xJYPMt1giVtxAXAr7r8bfAuzuwzRLsz0G0cfpXlPxj8Gaf8ACTwBqeuahcLbWOmQPNPNKdoVQP5ngAdyQKJZlVre4zoWW0KPvxPhnXPFMS/t9eONW0/Y1j4f8LX7XTdnkj04xRfX/SZIU+hNUf2Vfj141/Yi+NXg/wAaaV/ami6rps0Gq28MpltYtVtfMUmN8YMlvMqshIyrAt3FeU+AvGVxrGnfEfxAqlZPEV1bafLuP/LGWd7xgPfzLOD8M+tfZ3wekj/4KUfsMXng66hgk+MXwE086j4cnVcTeIfD8eBNZN/feAbSgxk/uwBlpGP1zh7KMY9kl+B+d1Kiq1JT/mbf3s/YD9hb/g4x+Av7XOn2On+I9Rf4W+MJtqSafrb7rGRz/wA8r0ARlf8ArqIj7HrX33bXMd5bxzQyJLDKodHRtyup5BB7g+tfxXT2jaNqPmRu21sSxSKSu5TyDx/Sv0G/4Jpf8F0/iZ+xD4Qn0GRY/H3hOFF+z6Jq128baedwybacBjGrDIKEMmTuwDkt3Qxtvj27nnSwV17m/Y/pOor5Z/Zi/wCCyXwD/aV+Etn4m/4TrQ/Bt1I5t7zRvEV9DY31jOoUuhVmw6/MMOuVPsQygruVaDV00cTozTtZn8sMlujq2V9aoWDE3rRH5lPrRRXhx2PaZb+xwhtvloQeDkZrl/Eun/8ACOX8cun3FzZ7nAKRv+7OSP4TkDr2oorSjrOzM62kbo/QH/gsPEus+MfgNqtwoa91/wCEehXN7J3llAl+f1zzj8B6V+cnxOsF8N+L4bqxaW1nb98HicoyOrcMpHIPGcjvRRRg/isVivgv5nonhH/gpv8AHj4f2a2tj8SdcuIUGxf7RSHUGA/3p0dv1rkPjL+1z8Sv2i7ZLfxl4w1bWrSNw62rFYbfcOjGKNVQsMnBIyMmiiu2OFoxlzxgr97K5zTxmIlHklOTXa7sdL8KrKOX4F3mV+7ridP4v9HPX6c/ma9g/YZ+KmtfAz9r/wCGOveHbn7LfDxNp2nuCMxzQXVxHbTxuBjIaKVx17g9QKKK4sR8TOqn8KOo/wCCmfwg0P4Nftc/ETw3oFqbTR9G1iUWUGRi2SQLL5a8fcUyFVHZQBknk/P3h+dotTjRfuyny2Hqp4P86KKxp/Aay3N/T4/7T06CaX/WMuGO0fNgkZORRRRU9QP/2Q==`" class=`"img-top`" alt=`"Petri Paavola`">`
            <img src=`"data:image/png;base64,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`" class=`"img-back`" alt=`"Petri Paavola`">`
        </div>`
        Petri.Paavola@yodamiitti.fi<br><br>`
        Microsoft MVP<br>`
        Windows and Devices for IT<br><br>`
        <img src=`"data:image/png;base64,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`" alt=`"Microsoft MVP`"><br>`
        </strong></p>`
        </td>`
        </tr></table>"

        $Title = "Intune Configuration Assignment report"
        ConvertTo-HTML -head $head -PostContent $WindowsConfigurationsByDisplayNameHTML, $AndroidConfigurationsByDisplayNameHTML, $iOSConfigurationsByDisplayNameHTML, $macOSConfigurationsByDisplayNameHTML, $OtherConfigurationsByDisplayNameHTML, $htmlAllDeviceConfigurationsSortedByAssignmentTargetGroupDisplayName -PreContent $PreContent -Title $Title | Out-File "$ReportSavePath\$HTMLFileName"
        $Success = $?

        if (-not ($Success)) {
            Write-Error "Error creating HTML file."
            Write-Output "Script will exit..."
            Pause
            Exit 1
        }
        else {
            Write-Output "Intune Device Configurations Assignment report HTML file created`n($ReportSavePath\$HTMLFileName)`n"
        }
    }
    catch {
        Write-Error "$($_.Exception.GetType().FullName)"
        Write-Error "$($_.Exception.Message)"
        Write-Error "Error creating HTML report: $ReportSavePath\$HTMLFileName"
        Write-Output "Script will exit..."
        Pause
        Exit 1
    }

    ############################################################

    # Open html file

    # Check file exists and is bigger than 0
    # File should exist already but years ago slow computer/disk caused some problems
    # so this is hopefully not needed workaround
    # Wait max. of 20 seconds

    $i = 0
    $filesize = 0
    do {
        Write-Output "Double check HTML file creation is really done (round $i)"
        $filesize = 0
        Start-Sleep -Seconds 2
        try {
            $HTMLFile = Get-ChildItem "$ReportSavePath\$HTMLFileName"
            $filesize = $HTMLFile.Length
        }
        catch {
            # Something went wrong, waiting for next round.
            Write-Output "Trouble getting file size, waiting 2 seconds and trying again..."
        }
        if ($filesize -eq 0) { Write-Output "Filesize is 0kB so waiting for a while for file creation to finish" }

        $i += 1
    } while (($i -lt 10) -and ($filesize -eq 0))

    Write-Output "Opening created file:`n$ReportSavePath\$HTMLFileName`n"
    try {
        Invoke-Item "$ReportSavePath\$HTMLFileName"
    }
    catch {
        Write-Output "Error opening file automatically to browser. Open file manually:`n$ReportSavePath\$HTMLFileName`n"
    }

}
catch {
    Write-Error "Uups! Something happened and we failed. Try again..."

    Write-Error "$($_.Exception.GetType().FullName)"
    Write-Error "$($_.Exception.Message)"
}


# SIG # Begin signature block
# MIIh1wYJKoZIhvcNAQcCoIIhyDCCIcQCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUVNR2mRNDJ+Wtb5tOicETFQOG
# 5Miggh1EMIIDxTCCAq2gAwIBAgIQAqxcJmoLQJuPC3nyrkYldzANBgkqhkiG9w0B
# AQUFADBsMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYD
# VQQLExB3d3cuZGlnaWNlcnQuY29tMSswKQYDVQQDEyJEaWdpQ2VydCBIaWdoIEFz
# c3VyYW5jZSBFViBSb290IENBMB4XDTA2MTExMDAwMDAwMFoXDTMxMTExMDAwMDAw
# MFowbDELMAkGA1UEBhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UE
# CxMQd3d3LmRpZ2ljZXJ0LmNvbTErMCkGA1UEAxMiRGlnaUNlcnQgSGlnaCBBc3N1
# cmFuY2UgRVYgUm9vdCBDQTCCASIwDQYJKoZIhvcNAQEBBQADggEPADCCAQoCggEB
# AMbM5XPm+9S75S0tMqbf5YE/yc0lSbZxKsPVlDRnogocsF9ppkCxxLeyj9CYpKlB
# WTrT3JTWPNt0OKRKzE0lgvdKpVMSOO7zSW1xkX5jtqumX8OkhPhPYlG++MXs2ziS
# 4wblCJEMxChBVfvLWokVfnHoNb9Ncgk9vjo4UFt3MRuNs8ckRZqnrG0AFFoEt7oT
# 61EKmEFBIk5lYYeBQVCmeVyJ3hlKV9Uu5l0cUyx+mM0aBhakaHPQNAQTXKFx01p8
# VdteZOE3hzBWBOURtCmAEvF5OYiiAhF8J2a3iLd48soKqDirCmTCv2ZdlYTBoSUe
# h10aUAsgEsxBu24LUTi4S8sCAwEAAaNjMGEwDgYDVR0PAQH/BAQDAgGGMA8GA1Ud
# EwEB/wQFMAMBAf8wHQYDVR0OBBYEFLE+w2kD+L9HAdSYJhoIAu9jZCvDMB8GA1Ud
# IwQYMBaAFLE+w2kD+L9HAdSYJhoIAu9jZCvDMA0GCSqGSIb3DQEBBQUAA4IBAQAc
# GgaX3NecnzyIZgYIVyHbIUf4KmeqvxgydkAQV8GK83rZEWWONfqe/EW1ntlMMUu4
# kehDLI6zeM7b41N5cdblIZQB2lWHmiRk9opmzN6cN82oNLFpmyPInngiK3BD41VH
# MWEZ71jFhS9OMPagMRYjyOfiZRYzy78aG6A9+MpeizGLYAiJLQwGXFK3xPkKmNEV
# X58Svnw2Yzi9RKR/5CYrCsSXaQ3pjOLAEFe4yHYSkVXySGnYvCoCWw9E1CAx2/S6
# cCZdkGCevEsXCS+0yx5DaMkHJ8HSXPfqIbloEpw8nL+e/IBcm2PN7EeqJSdnoDfz
# AIJ9VNep+OkuE6N36B9KMIIFeDCCBGCgAwIBAgIQBQbkfYp2NP3HW6pLDphGODAN
# BgkqhkiG9w0BAQsFADBsMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQg
# SW5jMRkwFwYDVQQLExB3d3cuZGlnaWNlcnQuY29tMSswKQYDVQQDEyJEaWdpQ2Vy
# dCBFViBDb2RlIFNpZ25pbmcgQ0EgKFNIQTIpMB4XDTE3MTAxMTAwMDAwMFoXDTIw
# MTAxNDEyMDAwMFowgZUxEzARBgsrBgEEAYI3PAIBAxMCRkkxHTAbBgNVBA8MFFBy
# aXZhdGUgT3JnYW5pemF0aW9uMRIwEAYDVQQFEwkyNTQzMTQ0LTgxCzAJBgNVBAYT
# AkZJMQ4wDAYDVQQHEwVFc3BvbzEWMBQGA1UEChMNWW9kYW1paXR0aSBPeTEWMBQG
# A1UEAxMNWW9kYW1paXR0aSBPeTCCASIwDQYJKoZIhvcNAQEBBQADggEPADCCAQoC
# ggEBAKT23in8UmTBz9Gz0DqujCVYyRZlJLH+hid+XUtouAFQ6KegkazzQgMcKi05
# 5pGZ1PFG0cURqNYjDpxaG5fmsWFWFGj63NRtXkDCNN12lJJA96sLxfcX9AsjSxaj
# LkZ8Cje+TAbV3j1Jf4+mnMeh/y3NnNsGkP7Ii9ZsNFBZvo+ipv/OGUryfcDi5sZt
# WXb7w/x/4i5LVyC607ThQF3BOYxo2FWSVWrT3qvN4jntpkU3pHLvl0ktvCigXPKs
# vn1DKG5A2UlsKj3HWRGImAI/wbyP0/LQPLgqdxTtt8anHI3NrEI2+WQhrxLED+FZ
# xKNmaLgxI8giI5GqKWtJlGviXAkCAwEAAaOCAeowggHmMB8GA1UdIwQYMBaAFI/o
# fvBtMmoABSPHcJdqOpD/a+rUMB0GA1UdDgQWBBSx3zynPP0jv/yzR+jzY2PIsHEY
# +TAnBgNVHREEIDAeoBwGCCsGAQUFBwgDoBAwDgwMRkktMjU0MzE0NC04MA4GA1Ud
# DwEB/wQEAwIHgDATBgNVHSUEDDAKBggrBgEFBQcDAzB7BgNVHR8EdDByMDegNaAz
# hjFodHRwOi8vY3JsMy5kaWdpY2VydC5jb20vRVZDb2RlU2lnbmluZ1NIQTItZzEu
# Y3JsMDegNaAzhjFodHRwOi8vY3JsNC5kaWdpY2VydC5jb20vRVZDb2RlU2lnbmlu
# Z1NIQTItZzEuY3JsMEsGA1UdIAREMEIwNwYJYIZIAYb9bAMCMCowKAYIKwYBBQUH
# AgEWHGh0dHBzOi8vd3d3LmRpZ2ljZXJ0LmNvbS9DUFMwBwYFZ4EMAQMwfgYIKwYB
# BQUHAQEEcjBwMCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdpY2VydC5jb20w
# SAYIKwYBBQUHMAKGPGh0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2Vy
# dEVWQ29kZVNpZ25pbmdDQS1TSEEyLmNydDAMBgNVHRMBAf8EAjAAMA0GCSqGSIb3
# DQEBCwUAA4IBAQAiZKBW4WeI27Dua1pfjCH0FhuUrEv/jDR5uBOyGt6DEslF/O2K
# 91RsKUU4Z6HEOsIGC+mNmrlg0PPQquU7mvudexoo1QXVZW4NQrQc9dJuuwqgyk56
# PA8l0S6JDUwNX2UgIjfz3oSeDQqxcR1V7UvyAzuxOg/zCUqv8FL4iuIvlCBNxJ8o
# f02/vzsRwI8YPRvT6Xh2zVIygpPip/r4MTuPOfvSEK3Id2WmLNT8YLH7Er1Laum9
# p22FM7wpml43qRkbMjnOst949kUZ/DGUwcMDQSLKp5z4suz4578is3L1VMmnsIHc
# g0Li8TtAypPpcA4RBnP7u6+GuIldKeAPqs5SMIIGajCCBVKgAwIBAgIQAwGaAjr/
# WLFr1tXq5hfwZjANBgkqhkiG9w0BAQUFADBiMQswCQYDVQQGEwJVUzEVMBMGA1UE
# ChMMRGlnaUNlcnQgSW5jMRkwFwYDVQQLExB3d3cuZGlnaWNlcnQuY29tMSEwHwYD
# VQQDExhEaWdpQ2VydCBBc3N1cmVkIElEIENBLTEwHhcNMTQxMDIyMDAwMDAwWhcN
# MjQxMDIyMDAwMDAwWjBHMQswCQYDVQQGEwJVUzERMA8GA1UEChMIRGlnaUNlcnQx
# JTAjBgNVBAMTHERpZ2lDZXJ0IFRpbWVzdGFtcCBSZXNwb25kZXIwggEiMA0GCSqG
# SIb3DQEBAQUAA4IBDwAwggEKAoIBAQCjZF38fLPggjXg4PbGKuZJdTvMbuBTqZ8f
# ZFnmfGt/a4ydVfiS457VWmNbAklQ2YPOb2bu3cuF6V+l+dSHdIhEOxnJ5fWRn8YU
# Oawk6qhLLJGJzF4o9GS2ULf1ErNzlgpno75hn67z/RJ4dQ6mWxT9RSOOhkRVfRiG
# BYxVh3lIRvfKDo2n3k5f4qi2LVkCYYhhchhoubh87ubnNC8xd4EwH7s2AY3vJ+P3
# mvBMMWSN4+v6GYeofs/sjAw2W3rBerh4x8kGLkYQyI3oBGDbvHN0+k7Y/qpA8bLO
# cEaD6dpAoVk62RUJV5lWMJPzyWHM0AjMa+xiQpGsAsDvpPCJEY93AgMBAAGjggM1
# MIIDMTAOBgNVHQ8BAf8EBAMCB4AwDAYDVR0TAQH/BAIwADAWBgNVHSUBAf8EDDAK
# BggrBgEFBQcDCDCCAb8GA1UdIASCAbYwggGyMIIBoQYJYIZIAYb9bAcBMIIBkjAo
# BggrBgEFBQcCARYcaHR0cHM6Ly93d3cuZGlnaWNlcnQuY29tL0NQUzCCAWQGCCsG
# AQUFBwICMIIBVh6CAVIAQQBuAHkAIAB1AHMAZQAgAG8AZgAgAHQAaABpAHMAIABD
# AGUAcgB0AGkAZgBpAGMAYQB0AGUAIABjAG8AbgBzAHQAaQB0AHUAdABlAHMAIABh
# AGMAYwBlAHAAdABhAG4AYwBlACAAbwBmACAAdABoAGUAIABEAGkAZwBpAEMAZQBy
# AHQAIABDAFAALwBDAFAAUwAgAGEAbgBkACAAdABoAGUAIABSAGUAbAB5AGkAbgBn
# ACAAUABhAHIAdAB5ACAAQQBnAHIAZQBlAG0AZQBuAHQAIAB3AGgAaQBjAGgAIABs
# AGkAbQBpAHQAIABsAGkAYQBiAGkAbABpAHQAeQAgAGEAbgBkACAAYQByAGUAIABp
# AG4AYwBvAHIAcABvAHIAYQB0AGUAZAAgAGgAZQByAGUAaQBuACAAYgB5ACAAcgBl
# AGYAZQByAGUAbgBjAGUALjALBglghkgBhv1sAxUwHwYDVR0jBBgwFoAUFQASKxOY
# spkH7R7for5XDStnAs0wHQYDVR0OBBYEFGFaTSS2STKdSip5GoNL9B6Jwcp9MH0G
# A1UdHwR2MHQwOKA2oDSGMmh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2Vy
# dEFzc3VyZWRJRENBLTEuY3JsMDigNqA0hjJodHRwOi8vY3JsNC5kaWdpY2VydC5j
# b20vRGlnaUNlcnRBc3N1cmVkSURDQS0xLmNybDB3BggrBgEFBQcBAQRrMGkwJAYI
# KwYBBQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBBBggrBgEFBQcwAoY1
# aHR0cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEQ0Et
# MS5jcnQwDQYJKoZIhvcNAQEFBQADggEBAJ0lfhszTbImgVybhs4jIA+Ah+WI//+x
# 1GosMe06FxlxF82pG7xaFjkAneNshORaQPveBgGMN/qbsZ0kfv4gpFetW7easGAm
# 6mlXIV00Lx9xsIOUGQVrNZAQoHuXx/Y/5+IRQaa9YtnwJz04HShvOlIJ8OxwYtNi
# S7Dgc6aSwNOOMdgv420XEwbu5AO2FKvzj0OncZ0h3RTKFV2SQdr5D4HRmXQNJsQO
# fxu19aDxxncGKBXp2JPlVRbwuwqrHNtcSCdmyKOLChzlldquxC5ZoGHd2vNtomHp
# igtt7BIYvfdVVEADkitrwlHCCkivsNRu4PQUCjob4489yq9qjXvc2EQwgga8MIIF
# pKADAgECAhAD8bThXzqC8RSWeLPX2EdcMA0GCSqGSIb3DQEBCwUAMGwxCzAJBgNV
# BAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdp
# Y2VydC5jb20xKzApBgNVBAMTIkRpZ2lDZXJ0IEhpZ2ggQXNzdXJhbmNlIEVWIFJv
# b3QgQ0EwHhcNMTIwNDE4MTIwMDAwWhcNMjcwNDE4MTIwMDAwWjBsMQswCQYDVQQG
# EwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYDVQQLExB3d3cuZGlnaWNl
# cnQuY29tMSswKQYDVQQDEyJEaWdpQ2VydCBFViBDb2RlIFNpZ25pbmcgQ0EgKFNI
# QTIpMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAp1P6D7K1E/Fkz4SA
# /K6ANdG218ejLKwaLKzxhKw6NRI6kpG6V+TEyfMvqEg8t9Zu3JciulF5Ya9DLw23
# m7RJMa5EWD6koZanh08jfsNsZSSQVT6hyiN8xULpxHpiRZt93mN0y55jJfiEmpqt
# RU+ufR/IE8t1m8nh4Yr4CwyY9Mo+0EWqeh6lWJM2NL4rLisxWGa0MhCfnfBSoe/o
# PtN28kBa3PpqPRtLrXawjFzuNrqD6jCoTN7xCypYQYiuAImrA9EWgiAiduteVDgS
# YuHScCTb7R9w0mQJgC3itp3OH/K7IfNs29izGXuKUJ/v7DYKXJq3StMIoDl5/d2/
# PToJJQIDAQABo4IDWDCCA1QwEgYDVR0TAQH/BAgwBgEB/wIBADAOBgNVHQ8BAf8E
# BAMCAYYwEwYDVR0lBAwwCgYIKwYBBQUHAwMwfwYIKwYBBQUHAQEEczBxMCQGCCsG
# AQUFBzABhhhodHRwOi8vb2NzcC5kaWdpY2VydC5jb20wSQYIKwYBBQUHMAKGPWh0
# dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydEhpZ2hBc3N1cmFuY2VF
# VlJvb3RDQS5jcnQwgY8GA1UdHwSBhzCBhDBAoD6gPIY6aHR0cDovL2NybDMuZGln
# aWNlcnQuY29tL0RpZ2lDZXJ0SGlnaEFzc3VyYW5jZUVWUm9vdENBLmNybDBAoD6g
# PIY6aHR0cDovL2NybDQuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0SGlnaEFzc3VyYW5j
# ZUVWUm9vdENBLmNybDCCAcQGA1UdIASCAbswggG3MIIBswYJYIZIAYb9bAMCMIIB
# pDA6BggrBgEFBQcCARYuaHR0cDovL3d3dy5kaWdpY2VydC5jb20vc3NsLWNwcy1y
# ZXBvc2l0b3J5Lmh0bTCCAWQGCCsGAQUFBwICMIIBVh6CAVIAQQBuAHkAIAB1AHMA
# ZQAgAG8AZgAgAHQAaABpAHMAIABDAGUAcgB0AGkAZgBpAGMAYQB0AGUAIABjAG8A
# bgBzAHQAaQB0AHUAdABlAHMAIABhAGMAYwBlAHAAdABhAG4AYwBlACAAbwBmACAA
# dABoAGUAIABEAGkAZwBpAEMAZQByAHQAIABDAFAALwBDAFAAUwAgAGEAbgBkACAA
# dABoAGUAIABSAGUAbAB5AGkAbgBnACAAUABhAHIAdAB5ACAAQQBnAHIAZQBlAG0A
# ZQBuAHQAIAB3AGgAaQBjAGgAIABsAGkAbQBpAHQAIABsAGkAYQBiAGkAbABpAHQA
# eQAgAGEAbgBkACAAYQByAGUAIABpAG4AYwBvAHIAcABvAHIAYQB0AGUAZAAgAGgA
# ZQByAGUAaQBuACAAYgB5ACAAcgBlAGYAZQByAGUAbgBjAGUALjAdBgNVHQ4EFgQU
# j+h+8G0yagAFI8dwl2o6kP9r6tQwHwYDVR0jBBgwFoAUsT7DaQP4v0cB1JgmGggC
# 72NkK8MwDQYJKoZIhvcNAQELBQADggEBABkzSgyBMzfbrTbJ5Mk6u7UbLnqi4vRD
# Qheev06hTeGx2+mB3Z8B8uSI1en+Cf0hwexdgNLw1sFDwv53K9v515EzzmzVshk7
# 5i7WyZNPiECOzeH1fvEPxllWcujrakG9HNVG1XxJymY4FcG/4JFwd4fcyY0xyQwp
# ojPtjeKHzYmNPxv/1eAal4t82m37qMayOmZrewGzzdimNOwSAauVWKXEU1eoYObn
# AhKguSNkok27fIElZCG+z+5CGEOXu6U3Bq9N/yalTWFL7EZBuGXOuHmeCJYLgYyK
# O4/HmYyjKm6YbV5hxpa3irlhLZO46w4EQ9f1/qbwYtSZaqXBwfBklIAwggbNMIIF
# taADAgECAhAG/fkDlgOt6gAK6z8nu7obMA0GCSqGSIb3DQEBBQUAMGUxCzAJBgNV
# BAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdp
# Y2VydC5jb20xJDAiBgNVBAMTG0RpZ2lDZXJ0IEFzc3VyZWQgSUQgUm9vdCBDQTAe
# Fw0wNjExMTAwMDAwMDBaFw0yMTExMTAwMDAwMDBaMGIxCzAJBgNVBAYTAlVTMRUw
# EwYDVQQKEwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20x
# ITAfBgNVBAMTGERpZ2lDZXJ0IEFzc3VyZWQgSUQgQ0EtMTCCASIwDQYJKoZIhvcN
# AQEBBQADggEPADCCAQoCggEBAOiCLZn5ysJClaWAc0Bw0p5WVFypxNJBBo/JM/xN
# RZFcgZ/tLJz4FlnfnrUkFcKYubR3SdyJxArar8tea+2tsHEx6886QAxGTZPsi3o2
# CAOrDDT+GEmC/sfHMUiAfB6iD5IOUMnGh+s2P9gww/+m9/uizW9zI/6sVgWQ8DIh
# FonGcIj5BZd9o8dD3QLoOz3tsUGj7T++25VIxO4es/K8DCuZ0MZdEkKB4YNugnM/
# JksUkK5ZZgrEjb7SzgaurYRvSISbT0C58Uzyr5j79s5AXVz2qPEvr+yJIvJrGGWx
# wXOt1/HYzx4KdFxCuGh+t9V3CidWfA9ipD8yFGCV/QcEogkCAwEAAaOCA3owggN2
# MA4GA1UdDwEB/wQEAwIBhjA7BgNVHSUENDAyBggrBgEFBQcDAQYIKwYBBQUHAwIG
# CCsGAQUFBwMDBggrBgEFBQcDBAYIKwYBBQUHAwgwggHSBgNVHSAEggHJMIIBxTCC
# AbQGCmCGSAGG/WwAAQQwggGkMDoGCCsGAQUFBwIBFi5odHRwOi8vd3d3LmRpZ2lj
# ZXJ0LmNvbS9zc2wtY3BzLXJlcG9zaXRvcnkuaHRtMIIBZAYIKwYBBQUHAgIwggFW
# HoIBUgBBAG4AeQAgAHUAcwBlACAAbwBmACAAdABoAGkAcwAgAEMAZQByAHQAaQBm
# AGkAYwBhAHQAZQAgAGMAbwBuAHMAdABpAHQAdQB0AGUAcwAgAGEAYwBjAGUAcAB0
# AGEAbgBjAGUAIABvAGYAIAB0AGgAZQAgAEQAaQBnAGkAQwBlAHIAdAAgAEMAUAAv
# AEMAUABTACAAYQBuAGQAIAB0AGgAZQAgAFIAZQBsAHkAaQBuAGcAIABQAGEAcgB0
# AHkAIABBAGcAcgBlAGUAbQBlAG4AdAAgAHcAaABpAGMAaAAgAGwAaQBtAGkAdAAg
# AGwAaQBhAGIAaQBsAGkAdAB5ACAAYQBuAGQAIABhAHIAZQAgAGkAbgBjAG8AcgBw
# AG8AcgBhAHQAZQBkACAAaABlAHIAZQBpAG4AIABiAHkAIAByAGUAZgBlAHIAZQBu
# AGMAZQAuMAsGCWCGSAGG/WwDFTASBgNVHRMBAf8ECDAGAQH/AgEAMHkGCCsGAQUF
# BwEBBG0wazAkBggrBgEFBQcwAYYYaHR0cDovL29jc3AuZGlnaWNlcnQuY29tMEMG
# CCsGAQUFBzAChjdodHRwOi8vY2FjZXJ0cy5kaWdpY2VydC5jb20vRGlnaUNlcnRB
# c3N1cmVkSURSb290Q0EuY3J0MIGBBgNVHR8EejB4MDqgOKA2hjRodHRwOi8vY3Js
# My5kaWdpY2VydC5jb20vRGlnaUNlcnRBc3N1cmVkSURSb290Q0EuY3JsMDqgOKA2
# hjRodHRwOi8vY3JsNC5kaWdpY2VydC5jb20vRGlnaUNlcnRBc3N1cmVkSURSb290
# Q0EuY3JsMB0GA1UdDgQWBBQVABIrE5iymQftHt+ivlcNK2cCzTAfBgNVHSMEGDAW
# gBRF66Kv9JLLgjEtUYunpyGd823IDzANBgkqhkiG9w0BAQUFAAOCAQEARlA+ybco
# JKc4HbZbKa9Sz1LpMUerVlx71Q0LQbPv7HUfdDjyslxhopyVw1Dkgrkj0bo6hnKt
# OHisdV0XFzRyR4WUVtHruzaEd8wkpfMEGVWp5+Pnq2LN+4stkMLA0rWUvV5PsQXS
# Dj0aqRRbpoYxYqioM+SbOafE9c4deHaUJXPkKqvPnHZL7V/CSxbkS3BMAIke/MV5
# vEwSV/5f4R68Al2o/vsHOE8Nxl2RuQ9nRc3Wg+3nkg2NsWmMT/tZ4CMP0qquAHzu
# nEIOz5HXJ7cW7g/DvXwKoO4sCFWFIrjrGBpN/CohrUkxg0eVd3HcsRtLSxwQnHcU
# wZ1PL1qVCCkQJjGCA/0wggP5AgEBMIGAMGwxCzAJBgNVBAYTAlVTMRUwEwYDVQQK
# EwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xKzApBgNV
# BAMTIkRpZ2lDZXJ0IEVWIENvZGUgU2lnbmluZyBDQSAoU0hBMikCEAUG5H2KdjT9
# x1uqSw6YRjgwCQYFKw4DAhoFAKBAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEE
# MCMGCSqGSIb3DQEJBDEWBBSsFrIi3jOxZBjLGDgJ2ZBwfT47sDANBgkqhkiG9w0B
# AQEFAASCAQAKgQK9Fua6vAQZXGXw+6+q29rleONkXUYntVhGopWGd+haEhs/IjzH
# 0hDDmBpiJDFtmriYC5HG4rbEbT2SY7IZtBXclh6AJC5rHn8KSYjWw4qp4WatfPok
# A1p1xNdoDnmrmTVFBlyM1FD1t4sd937h2jLvq5ehUrPJxfuS1e3rMA3eOTyY/naw
# DAxYo6KbKVBDE/DAib4ITdZPp5/bZFaLVwJgxh8pdCOiMIki8NMC1I40VLUZpNGK
# xTU+nVVxR2vYg3co5YpES0yPDekdRWRBh7DHeCJFfbU5GiCBUl3kIVQ0sgGadKS0
# wgtxVyCUZt5jydOWE3hZiEuti9/ItuWqoYICDzCCAgsGCSqGSIb3DQEJBjGCAfww
# ggH4AgEBMHYwYjELMAkGA1UEBhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZ
# MBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTEhMB8GA1UEAxMYRGlnaUNlcnQgQXNz
# dXJlZCBJRCBDQS0xAhADAZoCOv9YsWvW1ermF/BmMAkGBSsOAwIaBQCgXTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yMDAzMzExMDQ0
# MDlaMCMGCSqGSIb3DQEJBDEWBBTImEPn5EQFEekFv4uZwT4IqFjy6zANBgkqhkiG
# 9w0BAQEFAASCAQAEj7CWz01LexOz2HHOknDBxXWAkebXuegtmnxpipH/FxWpVdZl
# 1bWl+vrDkzf/2ipF8V3CemockEqv1DPmHu42pAccLBE6CrJSR/QmnjBZnWrpZdkj
# dS+j0zzEh3i5KNffRIiLD4nT6Dlx0Pir3pCHXjZ7PpgddYSKlnz20sCcKQa7EMOJ
# Qn1iE5O84lVb5D/aow4k4fYCzL3XqX37GD8D5t6bUOFlNpAIbt8Zr369TaEQJxef
# MWWP2pEo7kDMYSGL+ziJ2da2kHm2ajKiAz4lMpiutEpkuOZLuqvltJvZuCubJTyf
# D+DbHn/Z9IzvsiAxVhI25N+e55ao5L1XAQP8
# SIG # End signature block
