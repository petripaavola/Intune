# Azure Runbook example
# Using
#	Intune Powershell Module
#	Application secret
#	AutomationAccount hidden Credentials variable to get Secret and AppId
#
# 17.3.2022
#
# Petri.Paavola@yodamiitti.fi
# Microsoft MVP - Windows and Devices for IT
# https://www.github.com/petripaavola


Import-Module Microsoft.Graph.Intune

$myCred = Get-AutomationPSCredential -Name 'Intune GraphAPI AzureAutomation'
$clientId = $myCred.UserName
$securePassword = $myCred.Password
$clientSecret = $myCred.GetNetworkCredential().Password

$tenantId = "bef185b9-FIXME"
$authority = "https://login.windows.net/$tenantId"
#$clientId = "F00bar_NOT_USED"
#$clientSecret = "HdV_NOT_USED"

# Authenticate
Update-MSGraphEnvironment -AppId $clientId -Quiet
Update-MSGraphEnvironment -AuthUrl $authority -Quiet
Update-MSGraphEnvironment -SchemaVersion 'beta' -Quiet
Connect-MSGraph -ClientSecret $ClientSecret -Quiet

# Get Autopilot Devices
$Uri = "https://graph.microsoft.com/beta/deviceManagement/windowsAutopilotDeviceIdentities"
$MSGraphRequest = Invoke-MSGraphRequest -Url $Uri -HttpMethod Get

# Get all Graph API pages
$AllAutopilotDevices = Get-MSGraphAllPages -SearchResult $MSGraphRequest

$AllAutopilotDevices | Select-Object -Property serialNumber,GroupTag,id,displayName

