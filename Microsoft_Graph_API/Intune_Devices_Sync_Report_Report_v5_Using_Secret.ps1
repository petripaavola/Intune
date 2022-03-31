# Intune devices lastSync report and retire lastSyncDateTime over 180 days ago
# Using AppSecret and Intune Powershell module
# NVS_2022_Application
# 17.3.2022
#
# Petri.Paavola@yodamiitti.fi
# Microsoft MVP - Windows and Devices for IT
# https://www.github.com/petripaavola

#############################################################
#region App information
$AppId = 'f779d5d4-170b-FIXME'
$TenantId = 'bef185b9-FIXME'
$Authority = "https://login.windows.net/$tenantId"
# $AppSecret = 'REMOVED'

# Export AppSecret to local disk with Export-CliXML
# This encrypts secret and works only on that specific computer
# $Cred = Get-Credential
# $Cred
# $Cred | Export-CliXml -Path .\AppSecret.xml

# Import App Registration Secret from XML file
Try {
	$Cred = Import-Clixml -Path "$PSScriptRoot\AppSecret.xml"
	$Success = $?
	if(-not $Success) {
		Write-Error "Error importing Application Secret!"
		Exit 1
	}

	# Get App secret in clear text
	$AppSecret = (New-Object PSCredential "user",$Cred.Password).GetNetworkCredential().Password
	#Write-Host "`$AppSecret=$AppSecret"
	
} Catch {
	Write-Error "$($_.Exception.Message)"
	Write-Error "Error importing Application Secret!"
	Exit 2
}

#############################################################
#region Authenticate to Graph using AppSecret and using Intune Powershell module

Write-Host "Connecting to Graph API using AppSecret and Intune Powershell module"

Update-MSGraphEnvironment -AppId $AppId -Quiet
Update-MSGraphEnvironment -AuthUrl $Authority -Quiet
Update-MSGraphEnvironment -SchemaVersion 'beta' -Quiet
$ConnectMSGraph = Connect-MSGraph -ClientSecret $AppSecret
$Success = $?

if($Success) {
	Write-Host "Success`n"  -ForegroundColor 'Green'
} else {
	Write-Host "Failed to authenticate using AppSecret" -ForegroundColor 'Red'
	Exit 1
}

#endregion

#region Get Devices
Write-Host "Get Intune devices"

<#
# Get Intune devices using Graph API url
#$Uri = "https://graph.microsoft.com/beta/deviceManagement/managedDevices"
$Uri = "https://graph.microsoft.com/beta/deviceManagement/managedDevices?$select=Id,DeviceName,OperatingSystem,OSVersion,EnrolledDateTime,LastSyncDateTime,UserPrincipalName"
$MSGraphRequest = Invoke-MSGraphRequest -Url $Uri -HttpMethod Get

# Get all results
# takes care of paging
# also takes care of .Value property
$ManagedDevices = Get-MSGraphAllPages -SearchResult $MSGraphRequest
#>

# Get Intune devices using Intune Powershell commands
$ManagedDevices = Get-IntuneManagedDevice -Select Id,DeviceName,OperatingSystem,OSVersion,EnrolledDateTime,LastSyncDateTime,UserPrincipalName

# Get all results
$ManagedDevices = Get-MSGraphAllPages -SearchResult $ManagedDevices


# DEBUG
#Write-Verbose $ManagedDevices
#$ManagedDevices
#$ManagedDevices | Format-List *
#$ManagedDevices.Count
#$ManagedDevices | ConvertTo-Json -Depth 6 | Set-Clipboard
#Exit 0

#endregion

#region Process by LastSyncDateTime

# Sort devices by lastSyncDateTime (oldest first)
$ManagedDevices = $ManagedDevices | Sort-Object -Property LastSyncDateTime

# Debug sorting
#$ManagedDevices | Select-Object -Property DeviceName,LastSyncDateTime
#exit 0

Foreach($ManagedDevice in $ManagedDevices) {
	$DeviceName = $ManagedDevice.DeviceName
	$LastSyncDateTime = $ManagedDevice.LastSyncDateTime

	# Calculate how many days ago since lastSync
	$TimeSpan = New-TimeSpan (Get-Date) (Get-Date $LastSyncDateTime)
	$LastSyncInDays = $TimeSpan | Select-Object -ExpandProperty Days

	# Add $LastSyncInDays property to existing device object
	$ManagedDevice | Add-Member –Membertype NoteProperty –Name LastSyncDateTimeInDays –Value $LastSyncInDays

	if($LastSyncInDays -lt -180) {
		# Last Sync over 180 days ago
		Write-Host "Device $DeviceName last sync is $LastSyncInDays ago. Device will be retired from Intune"	-ForegroundColor Red
		
		# Retire device
		# run retire action
		
	} else {
		# Last Sync less than 180 days ago
		Write-Host "Device $DeviceName last sync is $LastSyncInDays ago."	-ForegroundColor Green
	}
}

#endregion

#region Create HTML report

# Create HTML Report

$ReportHTML = $ManagedDevices | Select-Object -Property DeviceName,LastSyncDateTime | ConvertTo-Html -Fragment -PreContent "<h2>Intune devices</h2>" | Out-String

# Fix special characters
$ReportHTML = $ReportHTML.Replace('&lt;', '<')
$ReportHTML = $ReportHTML.Replace('&gt;', '>')
$ReportHTML = $ReportHTML.Replace('&quot;', '"')

$head = @'
	<style>
		body {
			background-color:#dddddd;
			font-family:Tahoma;
			font-size:12pt;
		}
		td, th {
			border:1px solid black;
			border-collapse:collapse;
		}
		th {
			color:white;
			background-color:black;
		}
		table, tr, td, th {
			padding: 2px; margin: 0px
		}
		table { 
			margin-left:50px;
		}
	</style>
'@

ConvertTo-HTML -head $head -PostContent $ReportHTML -PreContent "<h1>Intune managed devices LastSyncDateTime report</h1><h2>for tenant $OrganizationName<br>TenantId: $TenantId</h2" -Title "Intune managed device LastSyncDateTime" | Out-File "$PSScriptRoot\IntuneDevicesLastSyncDateTimeReport.html" -Force

#endregion
