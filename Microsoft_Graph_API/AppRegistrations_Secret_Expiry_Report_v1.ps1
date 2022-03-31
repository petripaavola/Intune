# App registration secret expiration report console version
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
$Scope = 'https://graph.microsoft.com/.default'
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
#endregion

#############################################################
#region Do Azure AD Application based Authentication to Graph API

$Url = "https://login.microsoftonline.com/$TenantID/oauth2/v2.0/token"

# Add System.Web for urlencode
Add-Type -AssemblyName System.Web

# Create body
$Body = @{
	client_id = $AppId
	client_secret = $AppSecret
	scope = $Scope
	grant_type = 'client_credentials'
}

# Splat the parameters for Invoke-RestMethod for cleaner code
$PostSplat = @{
	ContentType = 'application/x-www-form-urlencoded'
	Method = 'POST'
	# Create string by joining bodylist with '&'
	Body = $Body
	Uri = $Url
}

Write-Host "Authenticate to Graph API using Azure AD Application based authentication"

# Request the token!
$Request = Invoke-RestMethod @PostSplat
$Success = $?

if($Success) {
	Write-Host "Success`n"  -ForegroundColor 'Green'
} else {
	Write-Host "Failed to authenticate using Azure AD Application secret" -ForegroundColor 'Red'
	Exit 3
}


# Create header
$Header = @{
	Authorization = "$($Request.token_type) $($Request.access_token)"
}

#endregion

#region Get date
Write-Host "Get App Registrations"
$url = "https://graph.microsoft.com/beta/myorganization/applications?$select=displayName,id,appId,info,createdDateTime,keyCredentials,passwordCredentials,deletedDateTime"
$Request = Invoke-RestMethod -Uri $url -Headers $Header -Method Get -ContentType "application/json"
$Success = $?

# Check if Invoke-RestMethod returned True (which is success)
if(-not $Success) {
	Write-Error "There was error getting Graph API data with url $url"
	Exit 0
}

# DEBUG
#$Request
#$Request | Format-List *
#Exit 0

# Note we are not handling possible Graph API .NextLink in this example !!!
$AppRegistrations = $Request.Value

#endregion

#region Process date

# DEBUG from script to clipboard
# What data we have now compared to Graph Explorer ????????
#$AppRegistrations | ConvertTo-Json -Depth 5 | Set-Clipboard

Foreach($AppRegistration in $AppRegistrations) {
	$displayName = $AppRegistration.displayName
	$passwordCredentials = $AppRegistration.passwordCredentials

	if($passwordCredentials) {
		Write-Host "App Registration $displayName secrets expires" -ForegroundColor Cyan
		foreach($passwordCredential in $passwordCredentials) {
			$secretDisplayName = $passwordCredential.displayName
			$endDateTime = $passwordCredential.endDateTime
			
			$TimeSpan = New-TimeSpan (Get-Date) (Get-Date $endDateTime)
			$ExpiresInDays = $TimeSpan | Select-Object -ExpandProperty Days
			
			if($ExpiresInDays -gt 0) {
				# Secret is not expired yet
				Write-Host "$secretDisplayName password expires in $ExpiresInDays days"	-ForegroundColor Green
			} else {
				# Secret has expired
				# Do something
				Write-Host "$secretDisplayName password expires in $ExpiresInDays days"	-ForegroundColor Red
			}
		}
		Write-Host
	}
}

#endregion


# DEBUG DATA
#return $AppRegistrations
