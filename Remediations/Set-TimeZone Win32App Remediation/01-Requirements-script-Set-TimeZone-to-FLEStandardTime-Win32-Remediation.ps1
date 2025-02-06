# Intune Win32App custom Requirement script will check if device system locale is fi-FI
# so script/Win32App/remediation is only run for fi-FI devices
#
#
# Petri.Paavola@yodamiitti.fi
# Microsoft MVP - Windows and Intune
# 2025-02-06


# Get the system locale (this returns an object with a Name property, e.g. "fi-FI")
$systemLocale = Get-WinSystemLocale

if ($systemLocale.Name -eq "fi-FI") {
	# Requirements are met
	# we can continue installing Win32 application ("Remediation" in this case)
	
	#Write-Host "Finnish language detected. Application is applicable and we can continue to Detection check"

	# Return 
	$true
} else {
    # Requirements are NOT met
	
	#Write-Host "Non-Finnish language detected. Application is not applicable."

	# Return
	$false
}