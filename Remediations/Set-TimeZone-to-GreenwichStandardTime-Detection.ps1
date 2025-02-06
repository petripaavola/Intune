# Change TimeZone to Greenwich Standard Time
# Intune Remediation script
#
# Petri.Paavola@yodamiitti.fi
# Microsoft MVP - Windows and Intune
# 2025-02-06

# Set TimeZone to Greenwich Standard Time
Set-Timezone -Id 'Greenwich Standard Time'
$Success = $?

if($Success) {
	# Set Timezone succeeded
	
	Write-Host "OK: TimeZone set to Greenwich Standard Time"
	exit 0
} else {
	# Set TimeZone failed
	Write-Host "ERROR: Setting TimeZone to Greenwich Standard Time"
	Exit 1
}
