# Intune Powershell Commands tips and tricks
#
# Petri.Paavola@yodamiitti.fi

Write-Host "Do NOT run me!" -ForegroundColor Red; Pause; Exit 0

#############
# Intune Module management

Install-Module -Name Microsoft.Graph.Intune

Get-Module -Name 'Microsoft.Graph.Intune'
Get-Module -All
Get-Module -ListAvailable

Update-Module -Name 'Microsoft.Graph.Intune' -Force

# Remove old duplicate modules
$Latest = Get-InstalledModule 'Microsoft.Graph.Intune'
Get-InstalledModule 'Microsoft.Graph.Intune' -AllVersions | Where-Object {$_.Version -ne $Latest.Version} | Uninstall-Module -WhatIf
Get-InstalledModule 'Microsoft.Graph.Intune' -AllVersions | Where-Object {$_.Version -ne $Latest.Version} | Uninstall-Module -Force


#############
# Show Intune Commands

# Show available Intune commands
Get-Command -Module Microsoft.Graph.Intune

# Show available Intune commands in Out-GridView
Get-Command -Module Microsoft.Graph.Intune | Out-GridView

#############
# Intune Managed Devices

# Get all Intune Managed Devices
Get-DeviceManagement_ManagedDevices

# Get help for commands Get-DeviceManagement_ManagedDevices
Get-Help Get-DeviceManagement_ManagedDevices

# Show Managed Devices in Out-GridView
Get-DeviceManagement_ManagedDevices | Select deviceName, userDisplayName, userPrincipalName, emailAddress, isEncrypted, manufacturer, model, serialNumber, wifiMacAddress, imei, ownerType, managementState, operatingSystem, osVersion, deviceType, complianceState, managementAgent, aadRegistered, deviceEnrollmentType, easActivated, easDeviceId, easActivationTime, lostModeState, enrolledDateTime, lastSyncDateTime, id, azureActiveDirectoryDeviceId | Sort deviceName | Out-GridView -Title "Intune Managed Devices"

Get-DeviceManagement_ManagedDevices -Filter "deviceName eq 'HPZBOOKSTUDIO'"

$device = Get-DeviceManagement_ManagedDevices -Filter "deviceName eq 'HPZBOOKSTUDIO'"
$device.deviceName
$device.wiFiMacAddress


# Measure how long query takes time. Filtering on Get-request is quicker
Measure-Command { Get-DeviceManagement_ManagedDevices -Filter "deviceName eq 'HPZBOOKSTUDIO'" }
Measure-Command { Get-DeviceManagement_ManagedDevices | Where { $_.deviceName -eq 'HPZBOOKSTUDIO' } }


##########

# Intune Apps

Get-DeviceAppManagement_MobileApps
Get-DeviceAppManagement_MobileApps -Filter "displayName eq '7-Zip 16.02 x64'"
Get-DeviceAppManagement_MobileApps -Filter "startswith(displayname,'7-Zip')"

# Measure how long query takes time. Filtering on Get-request is quicker
Measure-Command { Get-DeviceAppManagement_MobileApps -Filter "displayName eq '7-Zip 16.02 x64'" }
Measure-Command { Get-DeviceAppManagement_MobileApps | Where { $_.displayName -eq '7-Zip 16.02 x64'} }


# Get Intune Apps and show in Out-GridView
# By default we don't get Win32LobApps because those are NOT shown with Graph API v1.0 schema
# See next steps to fix this...
Get-DeviceAppManagement_MobileApps | Select -Property * -ExcludeProperty description | Out-GridView -Title "Intune Applications"

# Update Graph API schema to beta to get Win32LobApps and show Apps in Out-GridView
# Thanks Sandy for this tip :)
Update-MSGraphEnvironment -SchemaVersion 'beta'
Connect-MSGraph
Get-DeviceAppManagement_MobileApps | Select -Property * -ExcludeProperty description | Out-GridView -Title "Intune Applications"

# Change Graph API schema back to v1.0
Update-MSGraphEnvironment -SchemaVersion 'v1.0'
Connect-MSGraph


# Another way to get all application from Graph API using Invoke-MSGraphRequest
# Out-GridView with all apps, including Win32LobApp because uses beta schema

# Default url
#$url = "https://graph.microsoft.com/beta/deviceAppManagement/mobileApps"

# Url used in Intune Web UI
$url = "https://graph.microsoft.com/beta/deviceAppManagement/mobileApps?$filter=(microsoft.graph.managedApp/appAvailability%20eq%20null%20or%20microsoft.graph.managedApp/appAvailability%20eq%20%27lineOfBusiness%27%20or%20isAssigned%20eq%20true)&$orderby=displayName&&_=1571905984828"
$ManagedApps = Invoke-MSGraphRequest -Url $url -HttpMethod 'GET'
$AllManagedApps = Get-MSGraphAllPages -SearchResult $ManagedApps
$AllManagedApps | Select -Property * -ExcludeProperty description | Out-GridView -Title "Intune Applications"

