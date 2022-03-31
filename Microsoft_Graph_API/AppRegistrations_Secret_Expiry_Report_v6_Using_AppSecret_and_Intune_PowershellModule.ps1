# App registration secret expiration report using AppSecret
# and Intune Powershell Module command Invoke-MSGraphRequest and Get-MSGraphAllPages
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
#region Authenticate to Graph using AppSecret

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

#region Get Applications
Write-Host "Get App Registrations"
$url = "https://graph.microsoft.com/beta/myorganization/applications?$select=displayName,id,appId,info,createdDateTime,keyCredentials,passwordCredentials,deletedDateTime"
$MSGraphRequest = Invoke-MSGraphRequest -Url $Url -HttpMethod Get
$Success = $?
if(-not $Success) {
	Write-Error "There was error getting Graph API data with url $url"
	Exit 2
}

# Get all results (possible paged results)
$AppRegistrations = Get-MSGraphAllPages -SearchResult $MSGraphRequest
$Success = $?
if(-not $Success) {
	Write-Error "There was error getting all Graph API pages with $url"
	Exit 3
}


# DEBUG
#Write-Verbose $AppRegistrations
#$AppRegistrations
#$AppRegistrations | ConvertTo-Json -Depth 6 | Set-Clipboard
#Exit 0

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

