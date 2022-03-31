# App registration secret expiration report using certificate authentication
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

#region Get Applications
Write-Host "Get App Registrations"
$url = "https://graph.microsoft.com/beta/myorganization/applications?$select=displayName,id,appId,info,createdDateTime,keyCredentials,passwordCredentials,deletedDateTime"

# Note! We are not handling possible Graph API .NextLink in this example !!!
# Note! Specify -OutputType PSObject - otherwise you will get hashtables
$Request = Invoke-MgGraphRequest -Uri $url -Method 'GET' -OutputType PSObject
$Success = $?

# Check if Invoke-MgGraphRequest returned True (which is success)
if(-not $Success) {
	Write-Error "There was error getting Graph API data with url $url."
	Exit 0
}

# DEBUG
#Write-Verbose $Request
#$Request
#$Request.Value
#$Request.Value | ConvertTo-Json -Depth 6 | Set-Clipboard
#Exit 0

# Save .value to variable
$AppRegistrations = $Request.value

#endregion

#region Process date

# DEBUG from script to clipboard
# What data we have now compared to Graph Explorer ????????
#$AppRegistrations | ConvertTo-Json -Depth 5 | Set-Clipboard

# Create array for report
$AppRegistrationSecretReportObjects = @()

Foreach($AppRegistration in $AppRegistrations) {
	$displayName = $AppRegistration.displayName
	$passwordCredentials = $AppRegistration.passwordCredentials

	if($passwordCredentials) {
		$secretExpiresInformationString = $null
		Write-Host "App Registration $displayName secrets expires" -ForegroundColor Cyan
		foreach($passwordCredential in $passwordCredentials) {
			$secretDisplayName = $passwordCredential.displayName
			$endDateTime = $passwordCredential.endDateTime
			
			$TimeSpan = New-TimeSpan (Get-Date) (Get-Date $endDateTime)
			$ExpiresInDays = $TimeSpan | Select-Object -ExpandProperty Days
			
			if($ExpiresInDays -gt 0) {
				# Secret is not expired yet
				Write-Host "$secretDisplayName password expires in $ExpiresInDays days"	-ForegroundColor Green
				$secretExpiresInformationString = $secretExpiresInformationString + "<b>$secretDisplayName</b> password expires in <font color=`"Green`"><b>$ExpiresInDays</b></font> days<br>"
			} else {
				# Secret has expired
				# Do something
				Write-Host "$secretDisplayName password expires in $ExpiresInDays days"	-ForegroundColor Red
				$secretExpiresInformationString = $secretExpiresInformationString + "<font color=`"Red`"><b>$secretDisplayName</b> password expires in<b>$ExpiresInDays</b> days</font><br>"
			}
		}
		Write-Host

		# Create custom object
		$PowershellCustomObject = New-Object -TypeName psobject
		$PowershellCustomObject | Add-Member –Membertype NoteProperty –Name AppRegistrationDisplayName –Value $displayName
		$PowershellCustomObject | Add-Member –Membertype NoteProperty –Name SecretExpires –Value $secretExpiresInformationString

		# Add Custom object to array
		$AppRegistrationSecretReportObjects += $PowershellCustomObject
	}
}

#endregion

#region Create HTML report

# Create HTML Report

$ReportHTML = $AppRegistrationSecretReportObjects | Select-Object -Property AppRegistrationDisplayName,SecretExpires | ConvertTo-Html -Fragment -PreContent "<h2>Secrets expires</h2>" | Out-String

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

ConvertTo-HTML -head $head -PostContent $ReportHTML -PreContent "<h1>Application registrations</h1><h2>for tenant $OrganizationName<br>TenantId: $TenantId</h2" -Title "Application Registrations secret summary" | Out-File "$PSScriptRoot\ApplicationRegistrationSecretReport.html" -Force

#endregion

# Disconnect MgGraph connection to Graph API
Disconnect-MgGraph
