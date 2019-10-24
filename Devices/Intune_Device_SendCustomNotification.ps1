# Send Custom Notification to Intune Device
#
# .\Intune_Device_SendCustomNotification.ps1 -id 034edbcd-a295-4124-829a-e15b50bd003a -MessageTitle "Important message!" -MessageBody "Hello World!"
#
# You can pass device object in pipeline. id is used automatically
# Example: Get-DeviceManagement_ManagedDevices -Filter "deviceName eq 'etunimi.sukunimi_AndroidForWork_9/19/2019_12:10 PM'" | .\Intune_Device_SendCustomNotification.ps1 -MessageTitle "Important message!" -MessageBody "Hello World!"
#
# TODO: Proper processing if multiple deviceObjects are passed from pipeline
#
# Petri.Paavola@yodamiitti.fi
# 24.10.2019

Param(
    [Parameter(Mandatory = $true, Position = 0, ValueFromPipelineByPropertyName = $true)] [String]$id,
    [Parameter(Mandatory = $true, Position = 1)] [String]$MessageTitle,
    [Parameter(Mandatory = $true, Position = 2)] [String]$MessageBody
)


<# Request body JSON syntax
{
  "notificationTitle":"Message from GraphAPI",
  "notificationBody":"Hello World!"
}
#>

# Create Powershell custom object which will be converted to JSON
$TargetObject = $null
$TargetObject = New-Object PSObject
$TargetObject | Add-Member NoteProperty 'notificationTitle' $MessageTitle
$TargetObject | Add-Member NoteProperty 'notificationBody' $MessageBody

# Create json from Powershell custom object
# Cast variable as String, otherwise Invoke-MSGraphRequest will fail
# This is really important step!
[String]$SendMessageBodyJSON = $TargetObject | ConvertTo-json

# Debug
#$SendMessageBodyJSON

# Device Custom Notification Message url
$url = "https://graph.microsoft.com/beta/deviceManagement/managedDevices/$id/sendCustomNotificationToCompanyPortal"

# Send MSGraph request
Invoke-MSGraphRequest -Url $url -Content $SendMessageBodyJSON -HttpMethod 'POST'
$Success = $?

Write-Output "Send Device (id:$id) Notification success: $Success"
