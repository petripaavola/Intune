# Intune devices lastSync report and retire (not actually doing retire) lastSyncDateTime over 180 days ago
# Using certificate and Graph Powershell module
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
$CertificateThumbPrint = 'C1066AFE0EB9-FIXME'

#############################################################
#region Authenticate to Graph using Certificate

Write-Host "Connecting to Graph API using certificate"

Connect-MgGraph -ClientID $AppId -TenantId $TenantId -CertificateThumbprint $CertificateThumbPrint
$Success = $?

if($Success) {
	Write-Host "Success`n"  -ForegroundColor 'Green'
} else {
	Write-Host "Failed to authenticate using certificate" -ForegroundColor 'Red'
	Exit 1
}

# Change to beta API
Select-MgProfile -Name "beta"

# Verify connection type
#Get-MgContext

# Get organization name to HTML report
$url = 'https://graph.microsoft.com/beta/organization?$select=displayName'
$Request = Invoke-MgGraphRequest -Uri $url -Method 'GET' -OutputType PSObject
$OrganizationName = $Request.Value.displayName

#endregion

#region Get Devices
Write-Host "Get Intune devices"

# Use parameter -All to get all results (NextLinks)
#$ManagedDevices = Get-MgDeviceManagementManagedDevice -All
$ManagedDevices = Get-MgDeviceManagementManagedDevice -All -Property Id,DeviceName,OperatingSystem,OSVersion,EnrolledDateTime,LastSyncDateTime,UserPrincipalName

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
		Write-Host "Device $DeviceName last sync is $LastSyncInDays ago. Device will be retired from Intune" -ForegroundColor Red
		
		# Retire device
		# run retire action
		
	} else {
		# Last Sync less than 180 days ago
		Write-Host "Device $DeviceName last sync is $LastSyncInDays ago." -ForegroundColor Green
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

# Disconnect MgGraph connection to Graph API
Disconnect-MgGraph
