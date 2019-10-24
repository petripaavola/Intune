# Intune Powershell Commands tips and tricks

See all examples from [Intune_Powershell_Commands_Examples.ps1](./Intune_Powershell_Commands_Examples.ps1)

## Quick tips

### Show Intune Managed Devices in Out-GridView
```
Get-DeviceManagement_ManagedDevices | Select deviceName, userDisplayName, userPrincipalName, emailAddress, isEncrypted, manufacturer, model, serialNumber, wifiMacAddress, imei, ownerType, managementState, operatingSystem, osVersion, deviceType, complianceState, managementAgent, aadRegistered, deviceEnrollmentType, easActivated, easDeviceId, easActivationTime, lostModeState, enrolledDateTime, lastSyncDateTime, id, azureActiveDirectoryDeviceId | Sort deviceName | Out-GridView -Title "Intune Managed Devices"

```

### Show Intune Apps in Out-GridView using Graph API request
```
$url = "https://graph.microsoft.com/beta/deviceAppManagement/mobileApps?$filter=(microsoft.graph.managedApp/appAvailability%20eq%20null%20or%20microsoft.graph.managedApp/appAvailability%20eq%20%27lineOfBusiness%27%20or%20isAssigned%20eq%20true)&$orderby=displayName&&_=1571905984828"
$ManagedApps = Invoke-MSGraphRequest -Url $url -HttpMethod 'GET'
$AllManagedApps = Get-MSGraphAllPages -SearchResult $ManagedApps
$AllManagedApps | Select -Property * -ExcludeProperty description | Out-GridView -Title "Intune Applications"

```

### Find device by name and Sync Intune policies
```
$deviceName = 'HPZBOOKSTUDIO'
Get-DeviceManagement_ManagedDevices -Filter "deviceName eq '$deviceName'" | Invoke-DeviceManagement_ManagedDevices_SyncDevice
Write-Output "Sync action succeeded: $?"

```
