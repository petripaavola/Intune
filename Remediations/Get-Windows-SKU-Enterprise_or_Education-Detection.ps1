# Intune Remediation Detection script to detect Windows SKU version
# Computer is in Compliance if Windows SKU is Windows 10/11 Enterprise or Windows 10/11 Education
#
# This is mainly used to detect if Windows 10/11 Pro is installed on the computer originally
# and it was NOT upgraded to Enterprise or Education version for some reason
#
# One reason could be that computer did several Workplace joins which may prevent User subscription edition upgrade
# See more information from Rudy Ooms blog post:
#   https://call4cloud.nl/subscription-activation-0x87e10bf2-work-accounts/
#   https://call4cloud.nl/kb5040527-fixes-subscription-activation/
#
#
# Petri.Paavola@yodamiitti.fi
# Microsoft MVP - Windows and Intune
# 6.2.2025


# Get Windows SKU from WMI
# Values we are looking for are
# Windows 10 Enterprise
# Windows 10 Education
# Windows 11 Enterprise
# Windows 11 Education
$OperatingSystemCaption = Get-CimInstance -ClassName Win32_OperatingSystem | Select-Object -ExpandProperty Caption


if($OperatingSystemCaption -like "*Windows 1* Enterprise*" -or $OperatingSystemCaption -like "*Windows 1* Education*") {
	# We are in compliance

	# Print current Windows SKU for reporting
	Write-Host "$OperatingSystemCaption"
	Exit 0
} else {
	# We are NOT in compliance
	
	# Print current Windows SKU for reporting
	Write-Host "$OperatingSystemCaption"
	Exit 1
}
