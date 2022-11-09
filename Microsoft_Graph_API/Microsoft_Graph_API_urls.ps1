# Microsoft Graph API related tips and tricks
#
# You can use these url examples in Graph Explorer: https://developer.microsoft.com/en-us/graph/graph-explorer
# or with Powershell Cmdlet: Invoke-MSGraphRequest -Url $url -HttpMethod 'GET'
#
# Petri.Paavola@yodamiitti.fi

Write-Host "Do NOT run me!" -ForegroundColor Red; Pause; Exit 0


# Web links
# Graph Explorer: https://developer.microsoft.com/en-us/graph/graph-explorer
# Documentation: https://docs.microsoft.com/en-us/graph/
# Graph API beta reference: https://docs.microsoft.com/en-us/graph/api/overview?toc=./ref/toc.json&view=graph-rest-beta
# Managed Device docs: https://docs.microsoft.com/en-us/graph/api/intune-devices-manageddevice-get?view=graph-rest-beta

#####################
# AzureAD

# This gets AzureAD Devices which are different than Intune Managed Devices!
GET
https://graph.microsoft.com/beta/devices

# AzureAD Users
GET
https://graph.microsoft.com/beta/users

# Use Select to get only specified attributes
GET
https://graph.microsoft.com/beta/users?$select=mail,givenName,surname

# Find user by email address
GET
https://graph.microsoft.com/beta/users?$filter=mail eq 'firstname.lastname@yodamiitti.fi'

# Find user by userPrincipalName
GET
https://graph.microsoft.com/beta/users?$filter=userPrincipalName eq 'firstname.lastname@yodamiitti.fi'

# AzureAD Groups
GET
https://graph.microsoft.com/beta/groups

# AzureAD Group by displayName
GET
https://graph.microsoft.com/beta/groups?$filter=displayName eq 'Test FOOBAR'

# AzureAD Group by displayName starting with DynDev
GET
https://graph.microsoft.com/beta/groups?$filter=startswith(displayName,'DynDev')

# AzureAD Group by displayName starting with DynDev and show only displayname
GET
https://graph.microsoft.com/beta/groups?$filter=startswith(displayName,'DynDev')&$select=displayName

#####################
# Intune Managed Devices

# This gets Intune Managed Devices
GET
https://graph.microsoft.com/beta/devicemanagement/manageddevices

# User filtering to narrow down search results
# Search Intune Device by deviceName
GET
https://graph.microsoft.com/beta/devicemanagement/manageddevices?$filter=deviceName eq 'etunimi.sukunimi_AndroidForWork_9/19/2019_12:10 PM'

# Search Intune Device by deviceName
GET
https://graph.microsoft.com/beta/devicemanagement/manageddevices?$filter=deviceName eq 'HPZBOOKSTUDIO'

# Search Intune Device by deviceName starting with
GET
https://graph.microsoft.com/beta/devicemanagement/manageddevices?$filter=startswith(deviceName,'Win10')

# Search Intune Device by serialNumber
GET
https://graph.microsoft.com/beta/devicemanagement/manageddevices?$filter=serialNumber eq '1205-2767-4568-7306-0982-6763-81'

# Get Intune Managed Devices and use Select to get only specified attributes
GET
https://graph.microsoft.com/beta/devicemanagement/manageddevices?$select=deviceName,id


# Send custom notification to Intune Managed Device
# Use device id here
POST
https://graph.microsoft.com/beta/deviceManagement/managedDevices('034edbcd-a295-4124-829a-e15b50bd003a')/sendCustomNotificationToCompanyPortal

{
  "notificationTitle":"Title",
  "notificationBody":"Body text"
}

#####################
# Intune Apps

# Get Apps - notice we don't get Win32LobApps (intunewin) because we are using v1.0 schema
GET
https://graph.microsoft.com/v1.0/deviceAppManagement/mobileApps

# Get Apps using beta schema to get Win32LobApps also (intunewin)
GET
https://graph.microsoft.com/beta/deviceAppManagement/mobileApps

# Normally we don't get App Assignments so we have to expand our query with expand to Assignments also
GET
https://graph.microsoft.com/beta/deviceAppManagement/mobileApps?$expand=assignments


#####################
# Autopilot


# Get Autopilot devices
GET
https://graph.microsoft.com/beta/deviceManagement/windowsAutopilotDeviceIdentities


#####################
