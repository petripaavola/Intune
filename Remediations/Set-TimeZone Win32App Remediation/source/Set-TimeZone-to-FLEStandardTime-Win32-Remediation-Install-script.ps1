# Intune Win32App custom "Remediation" script to set device TimeZone to 'FLE Standard Time'
#
#
# Petri.Paavola@yodamiitti.fi
# Microsoft MVP - Windows and Intune
# 2025-02-06


# Set TimeZone to FLE Standard Time
Set-Timezone -Id 'FLE Standard Time'
$Success = $?

if($Success) {
	# Set Timezone succeeded
	
	Write-Host "OK: TimeZone set to FLE Standard Time"
	exit 0
} else {
	# Set TimeZone failed
	Write-Host "ERROR: Failed to set TimeZone to FLE Standard Time"
	Exit 1
}
